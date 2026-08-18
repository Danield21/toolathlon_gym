import pg from 'pg';
const { Pool } = pg;

export interface CalendarEvent {
    id: string;
    summary: string | null;
    description: string | null;
    location: string | null;
    start: { dateTime: string | null; timeZone?: string };
    end: { dateTime: string | null; timeZone?: string };
    status: string | null;
    htmlLink: string | null;
    creator: any;
    organizer: any;
    attendees: any;
    recurrence: any;
    reminders: any;
    created: string | null;
    updated: string | null;
}

interface EventRow {
    id: string;
    summary: string | null;
    description: string | null;
    location: string | null;
    start_datetime: string | null;
    start_timezone: string | null;
    end_datetime: string | null;
    end_timezone: string | null;
    status: string | null;
    html_link: string | null;
    creator: any;
    organizer: any;
    attendees: any;
    recurrence: any;
    reminders: any;
    created: string | null;
    updated: string | null;
}

/**
 * Convert an agent-supplied dateTime into an explicit-offset ISO instant string
 * (see insert() comment for the timezone bug rationale). Rules:
 *   - already has an offset / is a Z string  -> returned as-is
 *   - naive "2026-04-07T10:00:00"           -> interpreted in `timeZone`
 *                                              (IANA name; default UTC)
 *   - date-only "2026-04-07"                -> interpreted as midnight in
 *                                              `timeZone` (Google Calendar treats
 *                                              date-only values as all-day in the
 *                                              event's timezone)
 * Implementation: compute the wall-clock offset of the target zone for that
 * instant using Intl.DateTimeFormat parts — no external date library needed.
 */
function toInstantString(dateTime: unknown, timeZone?: string | null): string | null {
    if (dateTime === null || dateTime === undefined) return null;
    const raw = String(dateTime);
    // Already offset-aware (Z, +hh:mm, -hh:mm) — trust it verbatim.
    if (/(?:Z|[+-]\d{2}:?\d{2})\s*$/i.test(raw)) return raw;
    // Normalize a space separator (JS Date can parse "2026-04-07 10:00:00").
    const normalized = raw.includes('T') ? raw : raw.replace(' ', 'T');
    // Interpret the naive wall-clock in the requested zone. `timeZone`
    // undefined/empty/invalid falls back to UTC.
    let tz = timeZone && timeZone.trim() ? timeZone.trim() : 'UTC';
    try {
        // First pass: treat the naive string as UTC to get the epoch millis.
        // Date-only strings ("2026-04-07") parse as UTC midnight directly;
        // datetime strings get an explicit Z appended.
        const asUtc = normalized.length === 10
            ? Date.parse(`${normalized}T00:00:00Z`)
            : Date.parse(`${normalized}Z`);
        const epoch = asUtc;
        if (Number.isNaN(epoch)) return raw; // unparseable: let PG see the original
        // For date-only inputs the inserted value must remain a full timestamp,
        // otherwise "2026-04-07" + offset suffix concatenates into an invalid
        // string like "2026-04-07-04:00".
        const body = normalized.length === 10 ? `${normalized}T00:00:00` : normalized;
        // Second pass: find the zone's wall-clock offset AT that epoch by
        // formatting the instant in the zone and re-parsing it as if it were UTC.
        const dtf = new Intl.DateTimeFormat('en-US', {
            timeZone: tz,
            year: 'numeric', month: '2-digit', day: '2-digit',
            hour: '2-digit', minute: '2-digit', second: '2-digit',
            hour12: false,
        });
        const parts = dtf.formatToParts(new Date(epoch));
        const get = (t: string) => parts.find((p) => p.type === t)?.value ?? '00';
        const wallClockUtc = Date.parse(
            `${get('year')}-${get('month')}-${get('day')}T` +
            `${get('hour')}:${get('minute')}:${get('second')}Z`
        );
        const offsetMin = Math.round((wallClockUtc - epoch) / 60000);
        const sign = offsetMin >= 0 ? '+' : '-';
        const absMin = Math.abs(offsetMin);
        const pad = (n: number) => String(n).padStart(2, '0');
        return `${body}${sign}${pad(Math.floor(absMin / 60))}:${pad(absMin % 60)}`;
    } catch {
        return raw; // invalid tz name etc.: pass through, PG will surface the error
    }
}

function formatEvent(row: EventRow): CalendarEvent {
    return {
        id: row.id,
        summary: row.summary,
        description: row.description,
        location: row.location,
        start: {
            dateTime: row.start_datetime ? new Date(row.start_datetime).toISOString() : null,
            timeZone: row.start_timezone || undefined,
        },
        end: {
            dateTime: row.end_datetime ? new Date(row.end_datetime).toISOString() : null,
            timeZone: row.end_timezone || undefined,
        },
        status: row.status,
        htmlLink: row.html_link,
        creator: row.creator,
        organizer: row.organizer,
        attendees: row.attendees,
        recurrence: row.recurrence,
        reminders: row.reminders,
        created: row.created ? new Date(row.created).toISOString() : null,
        updated: row.updated ? new Date(row.updated).toISOString() : null,
    };
}

export class PgCalendar {
    events: {
        insert(params: { calendarId: string; requestBody: any }): Promise<{ data: CalendarEvent }>;
        get(params: { calendarId: string; eventId: string }): Promise<{ data: CalendarEvent }>;
        patch(params: { calendarId: string; eventId: string; requestBody: any }): Promise<{ data: CalendarEvent }>;
        delete(params: { calendarId: string; eventId: string }): Promise<{ data: {} }>;
        list(params: {
            calendarId: string;
            timeMin?: string;
            timeMax?: string;
            maxResults?: number;
            orderBy?: string;
            singleEvents?: boolean;
        }): Promise<{ data: { items: CalendarEvent[] } }>;
    };

    constructor(pool: InstanceType<typeof Pool>) {
        this.events = {
            async insert({ calendarId, requestBody }) {
                // Timezone-correct serialization (c4 case-study 2026-08-15,
                // yt-top-videos-excel-gcal / sf-sales-region-forecast-gcal-excel):
                // the column is TIMESTAMPTZ, but agents send naive dateTime strings
                // like "2026-04-07T10:00:00" plus a timeZone field. Passing the
                // naive string straight to pg makes PostgreSQL interpret it in the
                // SESSION timezone (UTC+8 on this harness) — so "10:00 UTC" became
                // 02:00Z and every evaluator window match failed. Interpret the
                // naive value in the REQUESTED timeZone (default UTC) and insert an
                // explicit-offset ISO string instead.
                const startDateTime = toInstantString(
                    requestBody.start?.dateTime,
                    requestBody.start?.timeZone
                );
                const endDateTime = toInstantString(
                    requestBody.end?.dateTime,
                    requestBody.end?.timeZone
                );
                const startTimeZone = requestBody.start?.timeZone || null;
                const endTimeZone = requestBody.end?.timeZone || null;
                const attendeesJson = requestBody.attendees
                    ? JSON.stringify(requestBody.attendees)
                    : '[]';
                // recurrence is a real column (gcal.events.recurrence jsonb) but the
                // old INSERT never wrote it, silently dropping the agent's RRULE
                // (c4 case-study: yt-fireship-tech-report "recurrence" FAIL).
                const recurrenceJson = requestBody.recurrence
                    ? JSON.stringify(requestBody.recurrence)
                    : null;

                const result = await pool.query(
                    `INSERT INTO gcal.events (summary, description, location, start_datetime, start_timezone, end_datetime, end_timezone, attendees, recurrence)
                     VALUES ($1, $2, $3, $4, $5, $6, $7, $8::jsonb, $9::jsonb)
                     RETURNING *`,
                    [
                        requestBody.summary || null,
                        requestBody.description || null,
                        requestBody.location || null,
                        startDateTime,
                        startTimeZone,
                        endDateTime,
                        endTimeZone,
                        attendeesJson,
                        recurrenceJson,
                    ]
                );
                return { data: formatEvent(result.rows[0]) };
            },

            async get({ calendarId, eventId }) {
                const result = await pool.query(
                    `SELECT * FROM gcal.events WHERE id = $1`,
                    [eventId]
                );
                if (result.rows.length === 0) {
                    throw new Error(`Event not found: ${eventId}`);
                }
                return { data: formatEvent(result.rows[0]) };
            },

            async patch({ calendarId, eventId, requestBody }) {
                const setClauses: string[] = [];
                const values: any[] = [];
                let paramIndex = 1;

                if (requestBody.summary !== undefined) {
                    setClauses.push(`summary = $${paramIndex++}`);
                    values.push(requestBody.summary);
                }
                if (requestBody.description !== undefined) {
                    setClauses.push(`description = $${paramIndex++}`);
                    values.push(requestBody.description);
                }
                if (requestBody.location !== undefined) {
                    setClauses.push(`location = $${paramIndex++}`);
                    values.push(requestBody.location);
                }
                if (requestBody.start?.dateTime !== undefined) {
                    setClauses.push(`start_datetime = $${paramIndex++}`);
                    values.push(toInstantString(
                        requestBody.start.dateTime,
                        requestBody.start.timeZone
                    ));
                }
                if (requestBody.start?.timeZone !== undefined) {
                    setClauses.push(`start_timezone = $${paramIndex++}`);
                    values.push(requestBody.start.timeZone);
                }
                if (requestBody.end?.dateTime !== undefined) {
                    setClauses.push(`end_datetime = $${paramIndex++}`);
                    values.push(toInstantString(
                        requestBody.end.dateTime,
                        requestBody.end.timeZone
                    ));
                }
                if (requestBody.end?.timeZone !== undefined) {
                    setClauses.push(`end_timezone = $${paramIndex++}`);
                    values.push(requestBody.end.timeZone);
                }
                if (requestBody.recurrence !== undefined) {
                    setClauses.push(`recurrence = $${paramIndex++}::jsonb`);
                    values.push(requestBody.recurrence ? JSON.stringify(requestBody.recurrence) : null);
                }

                // Always update the updated timestamp
                setClauses.push(`updated = NOW()`);

                if (setClauses.length === 1) {
                    // Only the updated timestamp, no real changes; just fetch
                    const result = await pool.query(
                        `SELECT * FROM gcal.events WHERE id = $1`,
                        [eventId]
                    );
                    if (result.rows.length === 0) throw new Error(`Event not found: ${eventId}`);
                    return { data: formatEvent(result.rows[0]) };
                }

                values.push(eventId);
                const result = await pool.query(
                    `UPDATE gcal.events SET ${setClauses.join(', ')} WHERE id = $${paramIndex} RETURNING *`,
                    values
                );
                if (result.rows.length === 0) {
                    throw new Error(`Event not found: ${eventId}`);
                }
                return { data: formatEvent(result.rows[0]) };
            },

            async delete({ calendarId, eventId }) {
                const result = await pool.query(
                    `DELETE FROM gcal.events WHERE id = $1`,
                    [eventId]
                );
                if (result.rowCount === 0) {
                    throw new Error(`Event not found: ${eventId}`);
                }
                return { data: {} };
            },

            async list({ calendarId, timeMin, timeMax, maxResults, orderBy, singleEvents }) {
                const conditions: string[] = [];
                const values: any[] = [];
                let paramIndex = 1;

                if (timeMin) {
                    conditions.push(`start_datetime >= $${paramIndex++}`);
                    values.push(timeMin);
                }
                if (timeMax) {
                    conditions.push(`end_datetime <= $${paramIndex++}`);
                    values.push(timeMax);
                }

                const whereClause = conditions.length > 0
                    ? `WHERE ${conditions.join(' AND ')}`
                    : '';

                let orderClause = 'ORDER BY start_datetime ASC';
                if (orderBy === 'updated') {
                    orderClause = 'ORDER BY updated DESC';
                }

                const limitClause = maxResults ? `LIMIT $${paramIndex++}` : '';
                if (maxResults) {
                    values.push(maxResults);
                }

                const result = await pool.query(
                    `SELECT * FROM gcal.events ${whereClause} ${orderClause} ${limitClause}`,
                    values
                );
                return { data: { items: result.rows.map(formatEvent) } };
            },
        };
    }
}

export function createPool(): InstanceType<typeof Pool> {
    return new Pool({
        host: process.env.PG_HOST || 'localhost',
        port: parseInt(process.env.PG_PORT || '5432'),
        database: process.env.PG_DATABASE || 'toolathlon',
        user: process.env.PG_USER || 'postgres',
        password: process.env.PG_PASSWORD || 'postgres',
    });
}
