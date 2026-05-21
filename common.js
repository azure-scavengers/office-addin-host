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
   *   - "pod" if it has a Database sheet but no tbl_Database table
   *   - "unknown" otherwise
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

  /**
   * Read the master/POD database into an array-of-rows.
   *
   * Lookup order:
   *   1. Excel Table named tbl_Database
   *   2. Worksheet named Database
   *
   * Returns [] if neither is found.
   */
  async function loadDatabase(context) {
    const wb = context.workbook;

    const table = wb.tables.getItemOrNullObject(DB_TABLE_NAME);
    const sheet = wb.worksheets.getItemOrNullObject(DB_SHEET_NAME);

    table.load("name");
    sheet.load("name");

    await context.sync();

    if (!table.isNullObject) {
      const headerRange = table.getHeaderRowRange();
      const bodyRange = table.getDataBodyRange();

      headerRange.load("values");
      bodyRange.load("values");

      await context.sync();

      const header =
        (headerRange.values && headerRange.values[0]) || [];
      const body = bodyRange.values || [];

      return [header].concat(body);
    }

    if (!sheet.isNullObject) {
      const used = sheet.getUsedRange();

      used.load("values");

      await context.sync();

      return used.values || [];
    }

    return [];
  }

  // Remaining functions:
  // getDropdownData()
  // getClientsBasedOnUser()
  // getClientDetails()
  // getSensitivity()
  // diagnose()
  // getClientListWithPOD()
  // submitForm()
  // saveNewClient()
  // getMasterTableHeader()
  // appendTime()
  // saveDateTime()

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
