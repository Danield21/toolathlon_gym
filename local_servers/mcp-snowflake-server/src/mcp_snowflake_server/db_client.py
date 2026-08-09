"""
PostgreSQL adapter for SnowflakeDB.

Translates Snowflake SQL (DB.SCHEMA.TABLE) to PostgreSQL queries against
the sf_data schema (sf_data."DB__SCHEMA__TABLE"), keeping the same
SnowflakeDB interface so server.py needs no changes.
"""
import asyncio
import logging
import os
import re
import time
import uuid
from typing import Any

import psycopg2
import psycopg2.extras

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    handlers=[logging.StreamHandler()],
)
logger = logging.getLogger("mcp_snowflake_server")


def _get_pg_connection():
    return psycopg2.connect(
        host=os.environ.get("PG_HOST", "localhost"),
        port=int(os.environ.get("PG_PORT", "5432")),
        database=os.environ.get("PG_DATABASE", os.environ.get("PGDATABASE", "toolathlon")),
        user=os.environ.get("PG_USER", "postgres"),
        password=os.environ.get("PG_PASSWORD", "postgres"),
    )


# ---------------------------------------------------------------------------
# SQL translation helpers
# ---------------------------------------------------------------------------

_SUCCESS = [{"status": "success", "message": "Query executed successfully"}]


def _translate_query(query: str):
    """Translate Snowflake SQL to PostgreSQL.

    Returns the translated SQL string, or None for no-op commands (USE, DDL).
    """
    q = query.strip().rstrip(";").strip()
    q_upper = q.upper()

    # ── No-op commands ────────────────────────────────────────────────
    if q_upper.startswith("USE "):
        return None
    if q_upper.startswith("SHOW "):
        return None
    for prefix in ("CREATE DATABASE", "DROP DATABASE",
                    "CREATE SCHEMA", "DROP SCHEMA"):
        if q_upper.startswith(prefix):
            return None

    # ── INFORMATION_SCHEMA.DATABASES (unqualified) ────────────────────
    if re.search(r'\bINFORMATION_SCHEMA\.DATABASES\b', q_upper):
        # Make sure it's NOT a 3-part ref like DB.INFORMATION_SCHEMA.DATABASES
        if not re.search(r'\w+\.INFORMATION_SCHEMA\.DATABASES', q_upper):
            return (
                "SELECT DISTINCT split_part(table_name, '__', 1) "
                'AS "DATABASE_NAME" '
                "FROM information_schema.tables "
                "WHERE table_schema = 'sf_data' ORDER BY 1"
            )

    # ── DB.INFORMATION_SCHEMA.SCHEMATA ────────────────────────────────
    m = re.search(r'(\w+)\.INFORMATION_SCHEMA\.SCHEMATA', q, re.IGNORECASE)
    if m:
        db = m.group(1).upper()
        return (
            "SELECT DISTINCT split_part(table_name, '__', 2) "
            'AS "SCHEMA_NAME" '
            "FROM information_schema.tables "
            f"WHERE table_schema = 'sf_data' "
            f"AND split_part(table_name, '__', 1) = '{db}' "
            "ORDER BY 1"
        )

    # ── DB.information_schema.tables ──────────────────────────────────
    m = re.search(r'(\w+)\.information_schema\.tables\b', q, re.IGNORECASE)
    if m:
        db = m.group(1).upper()
        sm = re.search(r"table_schema\s*=\s*'([^']+)'", q, re.IGNORECASE)
        schema = sm.group(1).upper() if sm else "PUBLIC"
        return (
            f"SELECT '{db}' AS \"TABLE_CATALOG\", "
            f"'{schema}' AS \"TABLE_SCHEMA\", "
            "split_part(table_name, '__', 3) AS \"TABLE_NAME\", "
            "'' AS \"COMMENT\" "
            "FROM information_schema.tables "
            f"WHERE table_schema = 'sf_data' "
            f"AND split_part(table_name, '__', 1) = '{db}' "
            f"AND split_part(table_name, '__', 2) = '{schema}' "
            "ORDER BY table_name"
        )

    # ── DB.information_schema.columns ─────────────────────────────────
    m = re.search(r'(\w+)\.information_schema\.columns\b', q, re.IGNORECASE)
    if m:
        db = m.group(1).upper()
        sm = re.search(r"table_schema\s*=\s*'([^']+)'", q, re.IGNORECASE)
        tm = re.search(r"table_name\s*=\s*'([^']+)'", q, re.IGNORECASE)
        schema = sm.group(1).upper() if sm else "PUBLIC"
        table = tm.group(1).upper() if tm else ""
        pg_table = f"{db}__{schema}__{table}"
        return (
            'SELECT column_name AS "COLUMN_NAME", '
            'column_default AS "COLUMN_DEFAULT", '
            'is_nullable AS "IS_NULLABLE", '
            'data_type AS "DATA_TYPE", '
            "'' AS \"COMMENT\" "
            "FROM information_schema.columns "
            f"WHERE table_schema = 'sf_data' AND table_name = '{pg_table}' "
            "ORDER BY ordinal_position"
        )

    # ── MEDIAN(expr) → PERCENTILE_CONT(0.5) WITHIN GROUP (ORDER BY expr) ──
    # PostgreSQL has no native MEDIAN aggregate; Snowflake does. Translate it so
    # agent queries like SELECT MEDIAN(SALARY) ... work against the PG backing
    # store. Handles a single level of nested parentheses in the argument.
    def _replace_median(m):
        inner = m.group(1).strip()
        return f"PERCENTILE_CONT(0.5) WITHIN GROUP (ORDER BY {inner})"
    q = re.sub(
        r"\bMEDIAN\s*\(([^()]*(?:\([^()]*\)[^()]*)*)\)",
        _replace_median,
        q,
        flags=re.IGNORECASE,
    )

    # ── Regular queries: DB.SCHEMA.TABLE → sf_data."DB__SCHEMA__TABLE" ─
    def _replace_ref(match):
        d, s, t = match.group(1).upper(), match.group(2).upper(), match.group(3).upper()
        if s == "INFORMATION_SCHEMA":
            return match.group(0)          # already handled above
        return f'sf_data."{d}__{s}__{t}"'

    # Match both quoted and unquoted three-part references:
    #   HR_ANALYTICS.PUBLIC.EMPLOYEES  or  "HR_ANALYTICS"."PUBLIC"."EMPLOYEES"
    translated = re.sub(
        r'"?([A-Za-z_]\w*)"?\."?([A-Za-z_]\w*)"?\."?([A-Za-z_]\w*)"?',
        _replace_ref,
        q,
    )

    # ── Quote bare identifiers as uppercase to preserve Snowflake case ──
    # Snowflake stores unquoted identifiers as uppercase; PostgreSQL folds them
    # to lowercase. Our mock tables store columns as uppercase quoted
    # ("ORDER_DATE"), so agent queries like `SELECT ORDER_DATE FROM ...`
    # resolve to lowercase `order_date` in PG and fail with
    # `column does not exist`. Wrap each remaining bare identifier as
    # "UPPERCASE" so PG preserves the case.
    # Run AFTER _replace_ref so three-part table refs (now sf_data."...")
    # and the information_schema branch are not disturbed; only column /
    # alias identifiers are folded here.
    translated = _snowflake_fold(translated)
    return translated


_SQL_KEYWORDS = frozenset({
    "select", "from", "where", "group", "order", "by", "having", "limit",
    "offset", "join", "left", "right", "inner", "outer", "on", "as", "and",
    "or", "not", "in", "is", "null", "distinct", "union", "all", "case",
    "when", "then", "else", "end", "count", "sum", "avg", "min", "max",
    "round", "coalesce", "greatest", "least", "between", "like", "ilike",
    "asc", "desc", "with", "over", "partition", "qualify", "from",
    "cast", "extract", "date", "interval", "true", "false",
    # PG-only schema ref produced by _replace_ref must not be re-quoted.
    "sf_data",
})


def _snowflake_fold(q: str) -> str:
    """Quote bare identifiers as uppercase to preserve Snowflake case.

    Skips: already-quoted identifiers, string literals, SQL keywords, function
    calls (identifier immediately followed by '('), and numeric literals.
    """
    out = []
    i = 0
    while i < len(q):
        # Skip double-quoted identifiers / strings
        if q[i] == '"':
            j = q.index('"', i + 1) + 1 if '"' in q[i + 1:] else len(q)
            out.append(q[i:j] if j <= len(q) else q[i:])
            i = j
            continue
        # Skip single-quoted string literals
        if q[i] == "'":
            j = i + 1
            while j < len(q) and q[j] != "'":
                if q[j] == "\\" and j + 1 < len(q):
                    j += 1
                j += 1
            out.append(q[i:j + 1])
            i = j + 1
            continue
        # Identifier?
        m = re.match(r'[A-Za-z_]\w*', q[i:])
        if m:
            word = m.group(0)
            end = i + len(word)
            # Skip SQL keywords
            if word.lower() in _SQL_KEYWORDS:
                out.append(word)
            # Skip function calls (next non-space char is '(')
            elif end < len(q) and q[end:].lstrip()[:1] == '(':
                out.append(word)
            else:
                out.append(f'"{word.upper()}"')
            i = end
        else:
            out.append(q[i])
            i += 1
    return ''.join(out)


# ---------------------------------------------------------------------------
# SnowflakeDB – same public interface, backed by PostgreSQL
# ---------------------------------------------------------------------------

class SnowflakeDB:
    AUTH_EXPIRATION_TIME = 1800

    def __init__(self, connection_config: dict):
        self.connection_config = connection_config
        self.session = None          # kept for interface compat
        self._conn = None            # persistent pg connection
        self.insights: list[str] = []
        self.auth_time = 0
        self.init_task = None

    def _get_conn(self):
        """Return a live PostgreSQL connection, reconnecting if needed."""
        if self._conn is None or self._conn.closed:
            self._conn = _get_pg_connection()
            self._conn.autocommit = True
        return self._conn

    async def _init_database(self):
        """Verify PostgreSQL connectivity."""
        try:
            self._get_conn()
            self.auth_time = time.time()
            logger.info("Connected to PostgreSQL (sf_data adapter mode)")
        except Exception as e:
            raise ValueError(f"Failed to connect to PostgreSQL: {e}")

    def start_init_connection(self):
        loop = asyncio.get_event_loop()
        self.init_task = loop.create_task(self._init_database())
        return self.init_task

    async def execute_query(self, query: str) -> tuple[list[dict[str, Any]], str]:
        """Execute a SQL query, translating Snowflake syntax to PostgreSQL."""
        if self.init_task and not self.init_task.done():
            await self.init_task

        logger.debug(f"Original query: {query}")
        translated = _translate_query(query)

        if translated is None:
            return list(_SUCCESS), str(uuid.uuid4())

        logger.debug(f"Translated query: {translated}")

        try:
            conn = self._get_conn()
            cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
            cur.execute(translated)

            if cur.description:
                rows = cur.fetchall()
                # Uppercase all keys to match Snowflake convention
                result_rows = [
                    {k.upper(): v for k, v in dict(row).items()} for row in rows
                ]
            else:
                result_rows = list(_SUCCESS)

            cur.close()

            if not result_rows:
                result_rows = list(_SUCCESS)

            data_id = str(uuid.uuid4())
            return result_rows, data_id

        except Exception as e:
            logger.error(
                f'Database error executing "{query}" '
                f'(translated: "{translated}"): {e}'
            )
            raise

    def add_insight(self, insight: str) -> None:
        self.insights.append(insight)

    def get_memo(self) -> str:
        if not self.insights:
            return "No data insights have been discovered yet."

        memo = "📊 Data Intelligence Memo 📊\n\n"
        memo += "Key Insights Discovered:\n\n"
        memo += "\n".join(f"- {insight}" for insight in self.insights)

        if len(self.insights) > 1:
            memo += (
                f"\n\nSummary:\nAnalysis has revealed {len(self.insights)} "
                "key data insights that suggest opportunities for "
                "strategic optimization and growth."
            )
        return memo
