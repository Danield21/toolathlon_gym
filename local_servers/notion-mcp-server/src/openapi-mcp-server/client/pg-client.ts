import type { OpenAPIV3, OpenAPIV3_1 } from 'openapi-types'
import pg from 'pg'
import crypto from 'crypto'
import { Headers } from './polyfill-headers.js'

const { Pool } = pg

/**
 * Normalize a page-parent object so the stored JSON always carries the "type"
 * discriminator that evaluators (and real Notion) expect. Agents frequently
 * omit it: {"database_id": "..."} instead of {"type": "database_id",
 * "database_id": "..."}. See createPage() for the Bug C case-study.
 */
function normalizeParent(parent: Record<string, any>): Record<string, any> {
  if (!parent || typeof parent !== 'object' || Array.isArray(parent)) return parent
  if (typeof parent.type === 'string' && parent.type) return parent
  for (const key of ['database_id', 'page_id', 'block_id', 'workspace']) {
    if (parent[key] !== undefined && parent[key] !== null) {
      return { type: key, ...parent }
    }
  }
  return parent
}

export type HttpClientConfig = {
  baseUrl: string
  headers?: Record<string, string>
}

export type HttpClientResponse<T = any> = {
  data: T
  status: number
  headers: Headers
}

export class PgHttpClient {
  private pool: InstanceType<typeof Pool>

  constructor(
    _config: HttpClientConfig,
    _openApiSpec: OpenAPIV3.Document | OpenAPIV3_1.Document,
  ) {
    this.pool = new Pool({
      host: process.env.PG_HOST || process.env.PGHOST || '127.0.0.1',
      port: parseInt(process.env.PG_PORT || process.env.PGPORT || '5432', 10),
      database: process.env.PG_DATABASE || process.env.PGDATABASE || 'toolathlon_gym',
      user: process.env.PG_USER || process.env.PGUSER || 'eigent',
      password: process.env.PG_PASSWORD || process.env.PGPASSWORD || 'camel',
    })
  }

  private makeResponse<T>(data: T, status = 200): HttpClientResponse<T> {
    return { data, status, headers: new Headers() }
  }

  private wrapList(results: any[]): any {
    return {
      object: 'list',
      results,
      has_more: false,
      next_cursor: null,
    }
  }

  private formatBlock(row: any): any {
    if (!row) return row
    const { block_data, parent_type, parent_id, position, ...rest } = row
    const blockData = block_data || {}
    return {
      ...rest,
      parent: { type: parent_type, [parent_type]: parent_id },
      has_children: rest.has_children ?? false,
      ...blockData,
      [rest.type]: blockData[rest.type] ?? blockData,
    }
  }

  private nowISO(): string {
    return new Date().toISOString()
  }

  async executeOperation<T = any>(
    operation: OpenAPIV3.OperationObject & { method: string; path: string },
    params: Record<string, any> = {},
  ): Promise<HttpClientResponse<T>> {
    const operationId = operation.operationId
    if (!operationId) {
      throw new Error('Operation ID is required')
    }

    switch (operationId) {
      case 'get-self':
        return this.getSelf()
      case 'get-user':
        return this.getUser(params)
      case 'get-users':
        return this.getUsers()
      case 'retrieve-a-database':
        return this.retrieveDatabase(params)
      case 'create-a-database':
        return this.createDatabase(params)
      case 'update-a-database':
        return this.updateDatabase(params)
      case 'post-database-query':
        return this.queryDatabase(params)
      case 'retrieve-a-page':
        return this.retrievePage(params)
      case 'patch-page':
        return this.patchPage(params)
      case 'post-page':
        return this.createPage(params)
      case 'retrieve-a-page-property':
        return this.retrievePageProperty(params)
      case 'retrieve-a-block':
        return this.retrieveBlock(params)
      case 'update-a-block':
        return this.updateBlock(params)
      case 'delete-a-block':
        return this.deleteBlock(params)
      case 'get-block-children':
        return this.getBlockChildren(params)
      case 'patch-block-children':
        return this.patchBlockChildren(params)
      case 'post-search':
        return this.postSearch(params)
      case 'retrieve-a-comment':
        return this.retrieveComments(params)
      case 'create-a-comment':
        return this.createComment(params)
      default:
        throw new Error(`Unknown operation: ${operationId}`)
    }
  }

  // 1. get-self
  private async getSelf(): Promise<HttpClientResponse> {
    const { rows } = await this.pool.query(
      `SELECT * FROM notion.users WHERE type = 'bot' LIMIT 1`,
    )
    if (rows.length === 0) {
      const fallback = await this.pool.query(
        `SELECT * FROM notion.users LIMIT 1`,
      )
      return this.makeResponse(fallback.rows[0] || null)
    }
    return this.makeResponse(rows[0])
  }

  // 2. get-user
  private async getUser(params: Record<string, any>): Promise<HttpClientResponse> {
    const { rows } = await this.pool.query(
      `SELECT * FROM notion.users WHERE id = $1`,
      [params.user_id],
    )
    return this.makeResponse(rows[0] || null)
  }

  // 3. get-users
  private async getUsers(): Promise<HttpClientResponse> {
    const { rows } = await this.pool.query(`SELECT * FROM notion.users`)
    return this.makeResponse(this.wrapList(rows))
  }

  // 4. retrieve-a-database
  private async retrieveDatabase(params: Record<string, any>): Promise<HttpClientResponse> {
    const { rows } = await this.pool.query(
      `SELECT * FROM notion.databases WHERE id = $1`,
      [params.database_id],
    )
    return this.makeResponse(rows[0] || null)
  }

  // 5. create-a-database
  private async createDatabase(params: Record<string, any>): Promise<HttpClientResponse> {
    const id = crypto.randomUUID()
    const now = this.nowISO()
    const {
      title = [],
      description = [],
      icon = null,
      cover = null,
      properties = {},
      parent = {},
      is_inline = false,
    } = params

    const { rows } = await this.pool.query(
      `INSERT INTO notion.databases
        (id, object, created_time, last_edited_time, title, description, icon, cover, properties, parent, is_inline, archived)
       VALUES ($1, 'database', $2, $3, $4, $5, $6, $7, $8, $9, $10, false)
       RETURNING *`,
      [id, now, now, JSON.stringify(title), JSON.stringify(description), JSON.stringify(icon), JSON.stringify(cover), JSON.stringify(properties), JSON.stringify(parent), is_inline],
    )
    return this.makeResponse(rows[0])
  }

  // 6. update-a-database
  // Mirrors the real Notion `PATCH /v1/databases/{database_id}` merge
  // semantics: top-level scalar/array/object fields (title, description,
  // icon, cover, archived, is_inline) are replaced when provided, but
  // `properties` is MERGED at the property-name level — new properties are
  // added, same-name properties are replaced, and unmentioned properties
  // are preserved. The previous implementation replaced the whole
  // `properties` JSON, silently dropping every unmentioned property schema
  // (audit Bug C / §A.5).
  private async updateDatabase(params: Record<string, any>): Promise<HttpClientResponse> {
    const { database_id, properties: incomingProperties, ...body } = params

    // Merge properties first if provided, reading the current row so we can
    // deep-merge by property name.
    let mergedProperties: Record<string, any> | undefined
    if (incomingProperties !== undefined) {
      const { rows: currentRows } = await this.pool.query(
        'SELECT properties FROM notion.databases WHERE id = $1',
        [database_id],
      )
      const currentProps = currentRows[0]?.properties
      const baseProps =
        typeof currentProps === 'string'
          ? safeParseJson(currentProps, {})
          : (currentProps as Record<string, any> | null | undefined) ?? {}
      mergedProperties = { ...baseProps, ...incomingProperties }
    }

    const setClauses: string[] = []
    const values: any[] = []
    let idx = 1

    const allowedFields = ['title', 'description', 'icon', 'cover', 'archived', 'is_inline']
    for (const field of allowedFields) {
      if (body[field] !== undefined) {
        const isJsonField = ['title', 'description', 'icon', 'cover'].includes(field)
        setClauses.push(`${field} = $${idx}`)
        values.push(isJsonField ? JSON.stringify(body[field]) : body[field])
        idx++
      }
    }

    if (mergedProperties !== undefined) {
      setClauses.push(`properties = $${idx}`)
      values.push(JSON.stringify(mergedProperties))
      idx++
    }

    setClauses.push(`last_edited_time = $${idx}`)
    values.push(this.nowISO())
    idx++

    values.push(database_id)

    const { rows } = await this.pool.query(
      `UPDATE notion.databases SET ${setClauses.join(', ')} WHERE id = $${idx} RETURNING *`,
      values,
    )
    return this.makeResponse(rows[0] || null)
  }

  // 7. post-database-query
  private async queryDatabase(params: Record<string, any>): Promise<HttpClientResponse> {
    const { database_id, filter, sorts } = params
    let query = `SELECT * FROM notion.pages WHERE parent->>'database_id' = $1`
    const values: any[] = [database_id]
    let idx = 2

    // Apply basic filter support — including `and` / `or` compound filters.
    // Builds a list of parameterised SQL fragments and joins them with AND/OR
    // so we never emit a stray leading operator (which would be a syntax error).
    const filterTypes = ['rich_text', 'title', 'number', 'checkbox', 'select',
      'multi_select', 'date', 'url', 'email', 'phone_number', 'status']

    const propScalarExpr = (propName: string, ft: string): string | null => {
      switch (ft) {
        case 'number':
          return `(properties->>'${propName}')::numeric`
        case 'checkbox':
          return `(properties->>'${propName}')::boolean`
        case 'date':
          return `(properties->>'${propName}')::date`
        case 'rich_text':
        case 'title':
        case 'select':
        case 'multi_select':
        case 'url':
        case 'email':
        case 'phone_number':
        case 'status':
          return `properties->'${propName}'->'${ft}'->>'content'`
        default:
          return null
      }
    }

    // Returns the SQL fragment for a leaf filter (no leading AND/OR) and pushes
    // any bound parameters into `values`. Returns null if the filter is not a
    // recognised leaf.
    const leafFragment = (f: any): string | null => {
      if (!f || !f.property) return null
      const propName = f.property
      for (const ft of filterTypes) {
        const cond = f[ft]
        if (cond === undefined) continue
        const col = propScalarExpr(propName, ft)
        if (col === null) break
        for (const [op, sqlOp, cast] of [
          ['equals', '=', false],
          ['does_not_equal', '<>', false],
          ['greater_than', '>', true],
          ['less_than', '<', true],
          ['greater_than_or_equal_to', '>=', true],
          ['less_than_or_equal_to', '<=', true],
        ] as [string, string, boolean][]) {
          if (cond[op] !== undefined) {
            if (cast && ft !== 'number') continue
            const frag = `${col} ${sqlOp} $${idx}`
            values.push(cond[op])
            idx++
            return frag
          }
        }
        if (cond.contains !== undefined) {
          const frag = `${col} ILIKE $${idx}`
          values.push(`%${cond.contains}%`)
          idx++
          return frag
        }
        if (cond.does_not_contain !== undefined) {
          const frag = `${col} NOT ILIKE $${idx}`
          values.push(`%${cond.does_not_contain}%`)
          idx++
          return frag
        }
        if (cond.starts_with !== undefined) {
          const frag = `${col} ILIKE $${idx}`
          values.push(`${cond.starts_with}%`)
          idx++
          return frag
        }
        if (cond.ends_with !== undefined) {
          const frag = `${col} ILIKE $${idx}`
          values.push(`%${cond.ends_with}`)
          idx++
          return frag
        }
        if (cond.is_empty !== undefined) {
          return `${col} IS ${cond.is_empty ? 'NULL' : 'NOT NULL'}`
        }
        if (cond.checked !== undefined && ft === 'checkbox') {
          return `${col} = ${cond.checked ? 'TRUE' : 'FALSE'}`
        }
        break
      }
      return null
    }

    const buildFilter = (f: any): string | null => {
      if (!f) return null
      if (Array.isArray(f.and)) {
        const parts = f.and.map((s: any) => buildFilter(s)).filter((x: any) => x)
        return parts.length ? `(${parts.join(' AND ')})` : null
      }
      if (Array.isArray(f.or)) {
        const parts = f.or.map((s: any) => buildFilter(s)).filter((x: any) => x)
        return parts.length ? `(${parts.join(' OR ')})` : null
      }
      return leafFragment(f)
    }

    const filterSql = buildFilter(filter)
    if (filterSql) query += ` AND ${filterSql}`

    // Apply sorts
    if (sorts && Array.isArray(sorts) && sorts.length > 0) {
      const orderClauses: string[] = []
      for (const sort of sorts) {
        const dir = sort.direction === 'descending' ? 'DESC' : 'ASC'
        if (sort.property) {
          orderClauses.push(`properties->'${sort.property}' ${dir}`)
        } else if (sort.timestamp) {
          orderClauses.push(`${sort.timestamp} ${dir}`)
        }
      }
      if (orderClauses.length > 0) {
        query += ` ORDER BY ${orderClauses.join(', ')}`
      }
    }

    const { rows } = await this.pool.query(query, values)
    return this.makeResponse(this.wrapList(rows))
  }

  // 8. retrieve-a-page
  private async retrievePage(params: Record<string, any>): Promise<HttpClientResponse> {
    const { rows } = await this.pool.query(
      `SELECT * FROM notion.pages WHERE id = $1`,
      [params.page_id],
    )
    return this.makeResponse(rows[0] || null)
  }

  // 9. patch-page
  // Mirrors the real Notion `PATCH /v1/pages/{page_id}` merge semantics:
  // scalar/array/object fields (icon, cover, archived, in_trash) are
  // replaced when provided, but `properties` is MERGED at the property-name
  // level — same-name properties are replaced, unmentioned properties are
  // preserved. (Audit Bug C / §A.5 — the previous replace-semantics lost
  // every unmentioned page property, which is the same class of bug as
  // update-a-database.)
  private async patchPage(params: Record<string, any>): Promise<HttpClientResponse> {
    const { page_id, properties: incomingProperties, ...body } = params

    let mergedProperties: Record<string, any> | undefined
    if (incomingProperties !== undefined) {
      const { rows: currentRows } = await this.pool.query(
        'SELECT properties FROM notion.pages WHERE id = $1',
        [page_id],
      )
      const currentProps = currentRows[0]?.properties
      const baseProps =
        typeof currentProps === 'string'
          ? safeParseJson(currentProps, {})
          : (currentProps as Record<string, any> | null | undefined) ?? {}
      mergedProperties = { ...baseProps, ...incomingProperties }
    }

    const setClauses: string[] = []
    const values: any[] = []
    let idx = 1

    const allowedFields = ['icon', 'cover', 'archived', 'in_trash']
    for (const field of allowedFields) {
      if (body[field] !== undefined) {
        const isJsonField = ['icon', 'cover'].includes(field)
        setClauses.push(`${field} = $${idx}`)
        values.push(isJsonField ? JSON.stringify(body[field]) : body[field])
        idx++
      }
    }

    if (mergedProperties !== undefined) {
      setClauses.push(`properties = $${idx}`)
      values.push(JSON.stringify(mergedProperties))
      idx++
    }

    setClauses.push(`last_edited_time = $${idx}`)
    values.push(this.nowISO())
    idx++

    values.push(page_id)

    const { rows } = await this.pool.query(
      `UPDATE notion.pages SET ${setClauses.join(', ')} WHERE id = $${idx} RETURNING *`,
      values,
    )
    return this.makeResponse(rows[0] || null)
  }

  // 10. post-page (create page)
  private async createPage(params: Record<string, any>): Promise<HttpClientResponse> {
    const id = crypto.randomUUID()
    const now = this.nowISO()
    const {
      parent = {},
      properties = {},
      icon = null,
      cover = null,
      children,
    } = params

    // Normalize the parent object before storage (Bug C fix, 2026-08-15):
    // agents legitimately send {"database_id": "..."} / {"page_id": "..."} /
    // {"workspace": true} WITHOUT the "type" discriminator, but evaluators
    // filter on parent->>'type' or parent.type and seed data always carries
    // the field. Without normalization the stored JSON lacks "type" and
    // downstream checks report "0 database-parented pages" (c4 case:
    // yt-transcript-notion-song-report). Infer the type from whichever key
    // is present, mirroring how formatEvent row->parent is reconstructed.
    const parentNormalized = normalizeParent(parent)

    const { rows } = await this.pool.query(
      `INSERT INTO notion.pages
        (id, object, created_time, last_edited_time, parent, properties, icon, cover, archived, in_trash)
       VALUES ($1, 'page', $2, $3, $4, $5, $6, $7, false, false)
       RETURNING *`,
      [id, now, now, JSON.stringify(parentNormalized), JSON.stringify(properties), JSON.stringify(icon), JSON.stringify(cover)],
    )

    // If children blocks are provided, insert them
    if (children && Array.isArray(children)) {
      await this.insertBlocks(id, 'page_id', children)
    }

    return this.makeResponse(rows[0])
  }

  // 11. retrieve-a-page-property
  private async retrievePageProperty(params: Record<string, any>): Promise<HttpClientResponse> {
    const { page_id, property_id } = params
    const { rows } = await this.pool.query(
      `SELECT properties FROM notion.pages WHERE id = $1`,
      [page_id],
    )
    if (rows.length === 0) {
      return this.makeResponse(null, 404)
    }
    const properties = rows[0].properties || {}
    // property_id could be the property name or an actual id
    const value = properties[property_id] ?? null
    return this.makeResponse(value)
  }

  // 12. retrieve-a-block
  private async retrieveBlock(params: Record<string, any>): Promise<HttpClientResponse> {
    const { rows } = await this.pool.query(
      `SELECT * FROM notion.blocks WHERE id = $1`,
      [params.block_id],
    )
    if (rows.length === 0) {
      return this.makeResponse(null, 404)
    }
    return this.makeResponse(this.formatBlock(rows[0]))
  }

  // 13. update-a-block
  private async updateBlock(params: Record<string, any>): Promise<HttpClientResponse> {
    const { block_id, ...body } = params
    const setClauses: string[] = []
    const values: any[] = []
    let idx = 1

    // Handle type-specific block data updates
    const metaFields = ['type', 'archived', 'has_children']
    const blockDataUpdates: Record<string, any> = {}

    for (const [key, value] of Object.entries(body)) {
      if (metaFields.includes(key)) {
        if (key === 'archived' || key === 'has_children') {
          setClauses.push(`${key} = $${idx}`)
          values.push(value)
          idx++
        } else if (key === 'type') {
          setClauses.push(`type = $${idx}`)
          values.push(value)
          idx++
        }
      } else {
        // Assume it's block_data content (e.g., paragraph, heading_1, etc.)
        blockDataUpdates[key] = value
      }
    }

    if (Object.keys(blockDataUpdates).length > 0) {
      setClauses.push(`block_data = block_data || $${idx}::jsonb`)
      values.push(JSON.stringify(blockDataUpdates))
      idx++
    }

    setClauses.push(`last_edited_time = $${idx}`)
    values.push(this.nowISO())
    idx++

    values.push(block_id)

    const { rows } = await this.pool.query(
      `UPDATE notion.blocks SET ${setClauses.join(', ')} WHERE id = $${idx} RETURNING *`,
      values,
    )
    if (rows.length === 0) {
      return this.makeResponse(null, 404)
    }
    return this.makeResponse(this.formatBlock(rows[0]))
  }

  // 14. delete-a-block
  private async deleteBlock(params: Record<string, any>): Promise<HttpClientResponse> {
    const { rows } = await this.pool.query(
      `UPDATE notion.blocks SET archived = true, in_trash = true, last_edited_time = $1 WHERE id = $2 RETURNING *`,
      [this.nowISO(), params.block_id],
    )
    if (rows.length === 0) {
      return this.makeResponse(null, 404)
    }
    return this.makeResponse(this.formatBlock(rows[0]))
  }

  // 15. get-block-children
  private async getBlockChildren(params: Record<string, any>): Promise<HttpClientResponse> {
    const { rows } = await this.pool.query(
      `SELECT * FROM notion.blocks WHERE parent_id = $1 ORDER BY position ASC`,
      [params.block_id],
    )
    const formatted = rows.map((r: any) => this.formatBlock(r))
    return this.makeResponse(this.wrapList(formatted))
  }

  // 16. patch-block-children
  private async patchBlockChildren(params: Record<string, any>): Promise<HttpClientResponse> {
    const { block_id, children } = params
    if (children && Array.isArray(children)) {
      await this.insertBlocks(block_id, 'block_id', children)
    }

    // Update parent's has_children flag
    await this.pool.query(
      `UPDATE notion.blocks SET has_children = true WHERE id = $1`,
      [block_id],
    )

    // Return the children
    const { rows } = await this.pool.query(
      `SELECT * FROM notion.blocks WHERE parent_id = $1 ORDER BY position ASC`,
      [block_id],
    )
    const formatted = rows.map((r: any) => this.formatBlock(r))
    return this.makeResponse(this.wrapList(formatted))
  }

  // 17. post-search
  private async postSearch(params: Record<string, any>): Promise<HttpClientResponse> {
    const { query, filter, sort } = params
    const results: any[] = []

    const shouldSearchPages = !filter || !filter.value || filter.value === 'page'
    const shouldSearchDatabases = !filter || !filter.value || filter.value === 'database'

    if (shouldSearchPages) {
      let pageQuery = `SELECT * FROM notion.pages WHERE archived = false`
      const pageValues: any[] = []
      let idx = 1

      if (query) {
        pageQuery += ` AND properties::text ILIKE $${idx}`
        pageValues.push(`%${query}%`)
        idx++
      }

      const { rows: pageRows } = await this.pool.query(pageQuery, pageValues)
      results.push(...pageRows)
    }

    if (shouldSearchDatabases) {
      let dbQuery = `SELECT * FROM notion.databases WHERE archived = false`
      const dbValues: any[] = []
      let idx = 1

      if (query) {
        dbQuery += ` AND title::text ILIKE $${idx}`
        dbValues.push(`%${query}%`)
        idx++
      }

      const { rows: dbRows } = await this.pool.query(dbQuery, dbValues)
      results.push(...dbRows)
    }

    // Apply sort if provided
    if (sort && sort.direction) {
      const dir = sort.direction === 'ascending' ? 1 : -1
      const field = sort.timestamp || 'last_edited_time'
      results.sort((a: any, b: any) => {
        const aVal = new Date(a[field] || 0).getTime()
        const bVal = new Date(b[field] || 0).getTime()
        return (aVal - bVal) * dir
      })
    }

    return this.makeResponse(this.wrapList(results))
  }

  // 18. retrieve-a-comment
  private async retrieveComments(params: Record<string, any>): Promise<HttpClientResponse> {
    const blockId = params.block_id
    const { rows } = await this.pool.query(
      `SELECT * FROM notion.comments
       WHERE parent->>'block_id' = $1
          OR parent->>'page_id' = $1
       ORDER BY created_time ASC`,
      [blockId],
    )
    return this.makeResponse(this.wrapList(rows))
  }

  // 19. create-a-comment
  private async createComment(params: Record<string, any>): Promise<HttpClientResponse> {
    const id = crypto.randomUUID()
    const now = this.nowISO()
    const {
      parent = {},
      discussion_id = null,
      rich_text = [],
    } = params

    const { rows } = await this.pool.query(
      `INSERT INTO notion.comments
        (id, object, parent, discussion_id, created_time, last_edited_time, rich_text)
       VALUES ($1, 'comment', $2, $3, $4, $5, $6)
       RETURNING *`,
      [id, JSON.stringify(parent), discussion_id, now, now, JSON.stringify(rich_text)],
    )
    return this.makeResponse(rows[0])
  }

  // Helper: insert child blocks
  private async insertBlocks(
    parentId: string,
    parentType: string,
    children: any[],
  ): Promise<void> {
    // Get current max position for this parent
    const { rows: posRows } = await this.pool.query(
      `SELECT COALESCE(MAX(position), -1) AS max_pos FROM notion.blocks WHERE parent_id = $1`,
      [parentId],
    )
    let position = (posRows[0]?.max_pos ?? -1) + 1

    for (const child of children) {
      const id = crypto.randomUUID()
      const now = this.nowISO()

      // Notion API tolerates plain-string children by treating them as paragraph
      // blocks. Normalize before destructuring so `child.type` etc. work below.
      let normalizedChild: any
      if (typeof child === 'string') {
        // Try parsing as JSON first (agent may send stringified block objects)
        try {
          const parsed = JSON.parse(child)
          normalizedChild = typeof parsed === 'object' && parsed !== null ? parsed : child
        } catch {
          normalizedChild = child
        }
        // If still a string, wrap as paragraph
        if (typeof normalizedChild === 'string') {
          normalizedChild = {
            type: 'paragraph',
            paragraph: {
              rich_text: [{ type: 'text', text: { content: normalizedChild } }],
            },
          }
        }
      } else {
        normalizedChild = child
      }

      const blockType = normalizedChild.type || 'paragraph'
      const hasChildren = !!(normalizedChild.children && normalizedChild.children.length > 0)

      // Extract block data: everything except meta fields
      const { type, children: childChildren, object, ...blockData } = normalizedChild

      await this.pool.query(
        `INSERT INTO notion.blocks
          (id, object, parent_type, parent_id, created_time, last_edited_time, type, has_children, archived, in_trash, block_data, position)
         VALUES ($1, 'block', $2, $3, $4, $5, $6, $7, false, false, $8, $9)`,
        [id, parentType, parentId, now, now, blockType, hasChildren, JSON.stringify(blockData), position],
      )

      position++

      // Recursively insert nested children
      if (childChildren && Array.isArray(childChildren) && childChildren.length > 0) {
        await this.insertBlocks(id, 'block_id', childChildren)
      }
    }
  }
}

// Module-level JSON parser used when merging properties read back from PG.
// `pg` returns JSON/JSONB columns as either already-parsed objects or as
// strings depending on the driver configuration, so tolerate both and fall
// back to the caller-provided default on any parse failure.
function safeParseJson<T>(value: unknown, fallback: T): T {
  if (typeof value !== 'string') {
    return (value as T | null | undefined) ?? fallback
  }
  try {
    return JSON.parse(value) as T
  } catch {
    return fallback
  }
}
