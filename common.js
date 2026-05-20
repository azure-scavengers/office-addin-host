/* common.js — shared Office.js helpers used by all task-pane HTML pages.
 *
 * Replaces the Apps Script `google.script.run.<fn>()` bridge with direct
 * Excel reads/writes via the Office.js / Excel JavaScript API.
 *
 * The original implementation referenced an external Google Sheet by ID
 * for master data. In Excel we expect a sheet named `Database` inside the
 * same workbook (see Excel_Migration/README.md for the consolidation step).
 */

const POD_BOOKING = (function () {
    const DB_SHEET_NAME = "Database";
    const DB_TABLE_NAME = "tbl_Database"; // exists only in the master S4U CLIENT DB workbook

    const DEFAULT_COLS = {
        user: 0,
        client: 1,
        clientId: 2,
        pod: 3,
        sensitivity: 4,
    };

    function normalizeText(value) {
        return String(value ?? "").trim().toLowerCase();
    }

    function normalizePod(value) {
        return normalizeText(value)
            .replace(/^pod[\s\-_]*/i, "")
            .replace(/[\s\-_]/g, "");
    }

    function detectColumns(rows) {
        if (!rows || !rows.length) {
            return {
                cols: DEFAULT_COLS,
                startRow: 0,
            };
        }

        const header = rows[0].map((v) => normalizeText(v));

        const hasHeader = header.some((h) =>
            [
                "user",
                "username",
                "user name",
                "client",
                "client name",
                "pod",
                "sensitivity",
            ].includes(h)
        );

        const find = (candidates, fallback) => {
            const idx = header.findIndex((h) => candidates.includes(h));
            return idx >= 0 ? idx : fallback;
        };

        const cols = hasHeader
            ? {
                  user: find(
                      ["user", "username", "user name"],
                      DEFAULT_COLS.user
                  ),
                  client: find(
                      ["client", "clientname", "client name"],
                      DEFAULT_COLS.client
                  ),
                  clientId: find(
                      [
                          "client id",
                          "clientid",
                          "client_id",
                          "client code",
                          "clientcode",
                      ],
                      DEFAULT_COLS.clientId
                  ),
                  pod: find(["pod"], DEFAULT_COLS.pod),
                  sensitivity: find(
                      ["sensitivity"],
                      DEFAULT_COLS.sensitivity
                  ),
              }
            : DEFAULT_COLS;

        return {
            cols,
            startRow: hasHeader ? 1 : 0,
        };
    }

    /**
     * Detect the workbook role:
     *   - "master" if the workbook contains an Excel Table named tbl_Database
     *                (this is the S4U CLIENT DB workbook).
     *   - "pod" if it has a Database sheet but no tbl_Database table.
     *   - "unknown" otherwise.
     */
    async function getWorkbookContext() {
        return Excel.run(async (ctx) => {
            const wb = ctx.workbook;

            const table = wb.tables.getItemOrNullObject(DB_TABLE_NAME);
            const sheet = wb.worksheets.getItemOrNullObject(DB_SHEET_NAME);

            table.load("name");
            sheet.load("name");

            await ctx.sync();

            const hasTable = !table.isNullObject;
            const hasSheet = !sheet.isNullObject;

            let role = "unknown";

            if (hasTable) {
                role = "master";
            } else if (hasSheet) {
                role = "pod";
            }

            return {
                role,
                hasTable,
                hasSheet,
            };
        });
    }

    /** Read the Database sheet into an array-of-rows. Returns [] if missing. */
    async function loadDatabase(context) {
        const wb = context.workbook;

        const sheet = wb.worksheets.getItemOrNullObject(DB_SHEET_NAME);

        sheet.load("name");

        await context.sync();

        if (sheet.isNullObject) {
            return [];
        }

        const used = sheet.getUsedRange();

        used.load("values");

        await context.sync();

        return used.values || [];
    }

    /**
     * Returns the unique users for a given POD filter (mirrors the original
     * RoW filter). Pass empty string to disable filtering.
     *
     * Database columns:
     * B=User(1), C=Client(2), D=ClientId(3), E=POD(4), F=Sensitivity(5).
     */
    async function getDropdownData(podFilter) {
        return Excel.run(async (ctx) => {
            const rows = await loadDatabase(ctx);

            const { cols, startRow } = detectColumns(rows);

            const users = new Set();

            const wantedPod = normalizePod(podFilter);

            for (let i = startRow; i < rows.length; i++) {
                const u = String(rows[i][cols.user] || "");
                const p = String(rows[i][cols.pod] || "");

                const podMatch =
                    !wantedPod || normalizePod(p) === wantedPod;

                if (u && podMatch) {
                    users.add(u);
                }
            }

            // Fallback: if POD parsing/filtering mismatches workbook naming,
            // return all users.
            if (!users.size && wantedPod) {
                for (let i = startRow; i < rows.length; i++) {
                    const u = String(rows[i][cols.user] || "");

                    if (u) {
                        users.add(u);
                    }
                }
            }

            return {
                users: Array.from(users).sort(),
            };
        });
    }

    async function getClientsBasedOnUser(user) {
        return Excel.run(async (ctx) => {
            const rows = await loadDatabase(ctx);

            const { cols, startRow } = detectColumns(rows);

            const target = normalizeText(user);

            const clients = [];

            for (let i = startRow; i < rows.length; i++) {
                if (normalizeText(rows[i][cols.user]) === target) {
                    const c = String(
                        rows[i][cols.client] ?? ""
                    ).trim();

                    if (c) {
                        clients.push(c);
                    }
                }
            }

            return Array.from(new Set(clients)).sort();
        });
    }

    async function getClientDetails(client) {
        return Excel.run(async (ctx) => {
            const rows = await loadDatabase(ctx);

            const { cols, startRow } = detectColumns(rows);

            const target = normalizeText(client);

            for (let i = startRow; i < rows.length; i++) {
                if (normalizeText(rows[i][cols.client]) === target) {
                    return {
                        clientId: String(
                            rows[i][cols.clientId] ?? ""
                        ).trim(),

                        pod: String(
                            rows[i][cols.pod] ?? ""
                        ).trim(),
                    };
                }
            }

            return {
                clientId: "",
                pod: "",
            };
        });
    }

    async function getSensitivity(user) {
        return Excel.run(async (ctx) => {
            const rows = await loadDatabase(ctx);

            const { cols, startRow } = detectColumns(rows);

            const target = normalizeText(user);

            for (let i = startRow; i < rows.length; i++) {
                if (normalizeText(rows[i][cols.user]) === target) {
                    return String(
                        rows[i][cols.sensitivity] ?? ""
                    ).trim();
                }
            }

            return "";
        });
    }

    /** Diagnostics: returns Database row count, detected columns, sample data. */
    async function diagnose() {
        return Excel.run(async (ctx) => {
            const wb = ctx.workbook;

            const sheet = wb.worksheets.getItemOrNullObject(
                DB_SHEET_NAME
            );

            sheet.load("name");

            const active = wb.worksheets.getActiveWorksheet();

            active.load("name");

            await ctx.sync();

            const out = {
                databaseSheetFound: !sheet.isNullObject,
                activeSheet: active.name,
                rowCount: 0,
                detectedColumns: null,
                startRow: 0,
                header: [],
                sample: [],
            };

            if (sheet.isNullObject) {
                return out;
            }

            const used = sheet.getUsedRange();

            used.load("values");

            await ctx.sync();

            const rows = used.values || [];

            const det = detectColumns(rows);

            out.rowCount = rows.length;
            out.detectedColumns = det.cols;
            out.startRow = det.startRow;
            out.header = rows[0] || [];
            out.sample = rows.slice(
                det.startRow,
                det.startRow + 3
            );

            return out;
        });
    }

    async function getClientListWithPOD() {
        return Excel.run(async (ctx) => {
            const rows = await loadDatabase(ctx);

            const { cols, startRow } = detectColumns(rows);

            const seen = new Set();

            const out = [];

            for (let i = startRow; i < rows.length; i++) {
                const c = String(rows[i][cols.client] || "");

                const p = String(rows[i][cols.pod] || "POD X");

                if (c && !seen.has(c)) {
                    seen.add(c);

                    out.push({
                        client: c,
                        pod: p,
                    });
                }
            }

            return out;
        });
    }

    /**
     * Append a Fill Row record to the active worksheet.
     * Mirrors `submitForm` from `using form.gs`.
     */
    async function submitForm(form) {
        return Excel.run(async (ctx) => {
            const sheet =
                ctx.workbook.worksheets.getActiveWorksheet();

            const used = sheet.getUsedRangeOrNullObject();

            used.load("rowCount");

            await ctx.sync();

            const nextRowIdx = used.isNullObject
                ? 0
                : used.rowCount; // 0-based

            const range = sheet.getRangeByIndexes(
                nextRowIdx,
                0,
                1,
                17
            );

            const values = [[
                false, // 1 Checkbox placeholder
                form.date, // 2 Date
                form.softwareId, // 3 Software ID
                form.inTime, // 4 In Time
                form.user, // 5 User
                form.client, // 6 Client
                form.clientCode, // 7 Client ID
                form.pod, // 8 POD
                form.sensitivity, // 9 Sensitivity
                form.projectCode, // 10 Charge Code
                form.jobDescription, // 11 Job description
                form.newValue, // 12 New
                form.edits, // 13 Edits
                form.rework, // 14 Formatting
                "", // 15 left blank — matches original
                "TBC", // 16 CSS DL
                "TBC", // 17 VA DL
            ]];

            range.values = values;

            if (form.sensitivity === "High") {
                const userCell = sheet.getRangeByIndexes(
                    nextRowIdx,
                    4,
                    1,
                    1
                );

                userCell.format.font.color = "#FF0000";
            }

            await ctx.sync();

            return {
                ok: true,
                row: nextRowIdx + 1,
            };
        });
    }

    /**
     * Append a new client / user mapping directly to the master Database table
     * (`tbl_Database`) inside S4U CLIENT DB.xlsx.
     *
     * This function MUST run inside the master workbook. The `New Client`
     * task pane is delivered only by manifest-masterdb.xml, but as a defensive
     * measure we also check for the presence of `tbl_Database` and refuse to
     * write otherwise (PODs would lose the write on the next Power Query
     * refresh).
     *
     * POD workbooks pick up the new row automatically via their existing
     * Power Query mirror on next open or Refresh All.
     */
    async function saveNewClient(payload) {
        return Excel.run(async (ctx) => {
            const wb = ctx.workbook;

            const table = wb.tables.getItemOrNullObject(
                DB_TABLE_NAME
            );

            table.load("name");

            await ctx.sync();

            if (table.isNullObject) {
                throw new Error(
                    "Master Database table '" +
                        DB_TABLE_NAME +
                        "' was not found in this workbook. Open S4U CLIENT DB.xlsx and try again."
                );
            }

            // Read the header to map our payload to whatever column order the
            // master table happens to use (User / Client Name / Client Code /
            // POD / Sensitivity / Reviewing or any variation).

            const headerRange = table.getHeaderRowRange();

            headerRange.load("values");

            await ctx.sync();

            const header =
                (headerRange.values && headerRange.values[0]) || [];

            const norm = (v) =>
                String(v ?? "").trim().toLowerCase();

            // Map known fields. Anything we don't recognise stays blank so we
            // don't overwrite custom admin columns.

            const aliases = {
                user: [
                    "user",
                    "username",
                    "user name",
                ],

                client: [
                    "client",
                    "client name",
                    "clientname",
                ],

                clientId: [
                    "client id",
                    "clientid",
                    "client code",
                    "clientcode",
                    "client_id",
                ],

                pod: ["pod"],

                sensitivity: ["sensitivity"],

                reviewing: [
                    "reviewing",
                    "inc",
                    "flag",
                ],
            };

            const idx = {};

            header.forEach((h, i) => {
                const n = norm(h);

                for (const key of Object.keys(aliases)) {
                    if (
                        aliases[key].includes(n) &&
                        idx[key] === undefined
                    ) {
                        idx[key] = i;
                    }
                }
            });

            const row = new Array(header.length).fill("");

            const put = (key, value) => {
                if (idx[key] !== undefined) {
                    row[idx[key]] = value;
                }
            };

            put("user", payload.userName || "");
            put("client", payload.clientName || "");
            put("clientId", payload.clientCode || "");
            put("pod", payload.pod || "POD X");
            put("sensitivity", payload.sensitivity || "High");
            put("reviewing", payload.reviewing || "INC");

            const added = table.rows.add(null, [row]);

            added.load("index");

            await ctx.sync();

            return {
                ok: true,
                rowIndex: added.index,
                columns: header,
            };
        });
    }

    /**
     * Helper used by the New Client form to fetch the master table header so it
     * can render a single-row form aligned to whatever columns exist today.
     */
    async function getMasterTableHeader() {
        return Excel.run(async (ctx) => {
            const wb = ctx.workbook;

            const table = wb.tables.getItemOrNullObject(
                DB_TABLE_NAME
            );

            table.load("name");

            await ctx.sync();

            if (table.isNullObject) {
                return [];
            }

            const headerRange = table.getHeaderRowRange();

            headerRange.load("values");

            await ctx.sync();

            return (
                (headerRange.values && headerRange.values[0]) || []
            );
        });
    }

    /** Append HH:MM to active cell, preserving slash-history (mirror appendTime). */
    async function appendTime(timeValue) {
        return Excel.run(async (ctx) => {
            const cell = ctx.workbook.getActiveCell();

            cell.load("values");

            await ctx.sync();

            const old = String(
                (cell.values &&
                    cell.values[0] &&
                    cell.values[0][0]) ||
                    ""
            );

            const matches =
                old.match(/(\d{1,2}:\d{2})/g) || [];

            matches.push(timeValue);

            cell.values = [[matches.join(" / ")]];

            await ctx.sync();

            return matches.join(" / ");
        });
    }

    /** Save "Mon, 09:00 AM" to active cell (mirror saveDateTimeToSheet). */
    async function saveDateTime(day, hour, minute, ampm) {
        return Excel.run(async (ctx) => {
            const cell = ctx.workbook.getActiveCell();

            const formatted =
                day.substring(0, 3) +
                ", " +
                hour +
                ":" +
                String(minute).padStart(2, "0") +
                " " +
                ampm;

            cell.values = [[formatted]];

            await ctx.sync();

            return formatted;
        });
    }

    return {
        getWorkbookContext,
        getDropdownData,
        getClientsBasedOnUser,
        getClientDetails,
        getSensitivity,
        getClientListWithPOD,
        submitForm,
        saveNewClient,
        getMasterTableHeader,
        appendTime,
        saveDateTime,
        diagnose,
    };
})();