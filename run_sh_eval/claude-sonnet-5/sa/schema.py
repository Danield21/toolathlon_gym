"""Finite tool schemas, preserving the existing evaluation adapter behavior."""
import json

_FINITE_NOTION_PARENT: dict = {'type': 'object', 'description': 'Parent of a page or block (finite schema).', 'oneOf': [{'type': 'object', 'properties': {'page_id': {'type': 'string'}, 'type': {'type': 'string', 'enum': ['page_id']}}, 'required': ['page_id'], 'additionalProperties': True}, {'type': 'object', 'properties': {'database_id': {'type': 'string'}, 'type': {'type': 'string', 'enum': ['database_id']}}, 'required': ['database_id'], 'additionalProperties': True}, {'type': 'object', 'properties': {'workspace': {'type': 'boolean', 'enum': [True]}, 'type': {'type': 'string', 'enum': ['workspace']}}, 'required': ['workspace'], 'additionalProperties': True}]}
_GENERIC_CYCLE_LEAF: dict = {'type': 'object', 'additionalProperties': True, 'description': 'Cycle-broken schema node (recursive reference removed).'}
_SCHEMA_CHILD_KEYS = ('properties', 'patternProperties', '$defs', 'definitions', 'dependentSchemas', 'items', 'prefixItems', 'additionalProperties', 'additionalItems', 'contains', 'oneOf', 'anyOf', 'allOf', 'not', 'if', 'then', 'else', 'unevaluatedProperties', 'unevaluatedItems', 'contentSchema', 'propertyNames')
_REF_KEYS = ('$ref', '$recursiveRef', '$dynamicRef')

def _node_ref(node: dict) -> str | None:
    for key in _REF_KEYS:
        value = node.get(key)
        if isinstance(value, str) and value:
            return value
    return None

def _looks_like_parent_ref(ref: str, prop_name: str | None) -> bool:
    tail = ref.rsplit('/', 1)[-1].lower()
    return prop_name == 'parent' or 'parent' in tail

def _cycle_leaf(ref: str, prop_name: str | None) -> dict:
    if _looks_like_parent_ref(ref, prop_name):
        return json.loads(json.dumps(_FINITE_NOTION_PARENT))
    return json.loads(json.dumps(_GENERIC_CYCLE_LEAF))

def _strip_remaining_refs(node: object, prop_name: str | None=None) -> object:
    if isinstance(node, list):
        return [_strip_remaining_refs(item, None) for item in node]
    if not isinstance(node, dict):
        return node
    ref = _node_ref(node)
    if ref is not None:
        leaf = _cycle_leaf(ref, prop_name)
        siblings = {k: v for (k, v) in node.items() if k not in _REF_KEYS}
        if siblings:
            leaf = {**leaf, **{k: _strip_remaining_refs(v, None) for (k, v) in siblings.items()}}
        return leaf
    out: dict = {}
    for (key, value) in node.items():
        if key == 'properties' and isinstance(value, dict):
            out[key] = {pk: _strip_remaining_refs(pv, pk) for (pk, pv) in value.items()}
        else:
            out[key] = _strip_remaining_refs(value, key if key == 'parent' else None)
    return out

def _resolve_local_ref(root: dict, ref: str) -> dict | None:
    if not isinstance(ref, str):
        return None
    if ref == '#':
        return root
    if not ref.startswith('#/'):
        return None
    cur: object = root
    for raw in ref[2:].split('/'):
        part = raw.replace('~1', '/').replace('~0', '~')
        if isinstance(cur, dict) and part in cur:
            cur = cur[part]
        elif isinstance(cur, list) and part.isdigit() and (int(part) < len(cur)):
            cur = cur[int(part)]
        else:
            return None
    return cur if isinstance(cur, dict) else None

def _schema_has_ref(node: object) -> bool:
    if isinstance(node, dict):
        if isinstance(node.get('$ref'), str):
            return True
        return any((_schema_has_ref(v) for v in node.values()))
    if isinstance(node, list):
        return any((_schema_has_ref(v) for v in node))
    return False

def break_json_schema_cycles(schema: dict, stats: list[str] | None=None) -> dict:
    if not isinstance(schema, dict):
        return schema

    def walk(node: object, stack: set[str], prop_name: str | None) -> object:
        if isinstance(node, list):
            return [walk(item, stack, None) for item in node]
        if not isinstance(node, dict):
            return node
        ref = _node_ref(node)
        if isinstance(ref, str) and (ref == '#' or ref.startswith('#/') or prop_name == 'parent'):
            if prop_name == 'parent':
                if stats is not None:
                    stats.append(f'parent->{ref}')
                return _cycle_leaf(ref, 'parent')
            if ref in stack:
                leaf = _cycle_leaf(ref, prop_name)
                if stats is not None:
                    stats.append(f"{prop_name or '?'}->{ref}")
                siblings = {k: v for (k, v) in node.items() if k not in _REF_KEYS}
                if siblings:
                    leaf = {**leaf, **{k: walk(v, stack, None) for (k, v) in siblings.items()}}
                return leaf
            resolved = _resolve_local_ref(schema, ref)
            if resolved is None:
                return _cycle_leaf(ref, prop_name) if ref else node
            inner_props = resolved.get('properties') if isinstance(resolved, dict) else None
            if prop_name == 'parent' and isinstance(inner_props, dict) and ('parent' in inner_props):
                if stats is not None:
                    stats.append(f'parent->{ref}')
                return _cycle_leaf(ref, 'parent')
            stack.add(ref)
            out = walk(resolved, stack, prop_name)
            stack.discard(ref)
            if isinstance(out, dict):
                merged = dict(out)
                for (key, value) in node.items():
                    if key in _REF_KEYS:
                        continue
                    merged[key] = walk(value, stack, key if key != 'properties' else None)
                return merged
            return out
        out: dict = {}
        for (key, value) in node.items():
            if key == 'properties' and isinstance(value, dict):
                out[key] = {pk: walk(pv, stack, pk) for (pk, pv) in value.items()}
            elif key in _SCHEMA_CHILD_KEYS:
                out[key] = walk(value, stack, prop_name if key != 'items' else None)
            else:
                out[key] = walk(value, stack, None)
        return out
    result = walk(schema, set(), None)
    if not isinstance(result, dict):
        return schema
    if not _schema_has_ref(result):
        result.pop('$defs', None)
        result.pop('definitions', None)
    return result
