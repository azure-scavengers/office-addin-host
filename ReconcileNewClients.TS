 /**
  * Office Script — ReconcileNewClients
  *
  * Runs INSIDE the silent master `S4U CLIENT DB.xlsx`. Invoked from a
  * Power Automate flow once per Pending row collected from the POD queues.
  *
  * Master workbook layout assumed by this script:
  *   - A worksheet named exactly `Database`.
  *   - Row 1 contains headers. Any of these are recognised (case-insensitive,
  *     in any order), additional columns are tolerated and left blank:
  *       User      | Username | User Name
  *       Client    | Client Name
  *       ClientId  | Client Id | Client Code
  *       POD
  *       Sensitivity
  *       Reviewing | INC | Flag
  *   - No Excel Table conversion required. The script appends new rows
  *     directly below the current used range.
  *
  * Power Automate flow shape (scheduled, e.g. every 15 minutes):
  *   For each POD file (POD A.xlsx, POD B.xlsx, POD RoW.xlsx):
  *     Excel Online (Business) -> List rows present in a table
  *       Table: tbl_NewClientQueue
  *       Filter Query: Status eq 'Pending'
  *     For each row:
  *       Excel Online (Business) -> Run script (against MASTER S4U CLIENT DB.xlsx)
  *         Script: Reconcile New Clients
  *         Parameters: payload (object below)
  *       Excel Online (Business) -> Update a row
  *         Table: tbl_NewClientQueue (POD file)
  *         Key: QueuedAt = payload.QueuedAt
  *         Status: <returned by this script>
  *         ProcessedAt: utcNow()
  */

interface ReconcilePayload {
  clientName: string;
  userName: string;
  pod: string;
  sensitivity?: string;
  reviewing?: string;
  clientCode?: string;
}

interface ReconcileResult {
  status: "Processed" | "Skipped" | "Error";
  message: string;
}

const DB_SHEET_NAME = "Database";

/**
 * Normalize values for safe comparison (case-insensitive, trimmed)
 */
function norm(v: string | number | boolean): string {
  return String(v == null ? "" : v).trim().toLowerCase();
}

/**
 * Find column index based on multiple possible header aliases
 */
function findColumn(header: (string | number | boolean)[], aliases: string[]): number {
  for (let i = 0; i < header.length; i++) {
    if (aliases.indexOf(norm(header[i])) !== -1) {
      return i;
    }
  }
  return -1;
}

/**
 * Main reconciliation entry point
 */
function main(
  workbook: ExcelScript.Workbook,
  payload: ReconcilePayload
): ReconcileResult {

  if (!payload || !payload.clientName || !payload.userName) {
    return {
      status: "Error",
      message: "Payload missing clientName or userName."
    };
  }

  const sheet = workbook.getWorksheet(DB_SHEET_NAME);

  if (!sheet) {
    return {
      status: "Error",
      message: "Master workbook is missing worksheet '" + DB_SHEET_NAME + "'."
    };
  }

  const used = sheet.getUsedRange();

  if (!used) {
    return {
      status: "Error",
      message: "'" + DB_SHEET_NAME + "' sheet appears empty. Add a header row first."
    };
  }

  const values = used.getValues();

  if (values.length < 1) {
    return {
      status: "Error",
      message: "'" + DB_SHEET_NAME + "' has no header row."
    };
  }

  const header = values[0];

  const idx = {
    user:        findColumn(header, ["user", "username", "user name"]),
    client:      findColumn(header, ["client", "client name", "clientname"]),
    clientId:    findColumn(header, ["client id", "clientid", "client code", "clientcode", "client_id"]),
    pod:         findColumn(header, ["pod"]),
    sensitivity: findColumn(header, ["sensitivity"]),
    reviewing:   findColumn(header, ["reviewing", "inc", "flag"])
  };

  if (idx.user < 0 || idx.client < 0) {
    return {
      status: "Error",
      message: "Database sheet must have at least 'User' and 'Client' columns in the header row."
    };
  }

  // Duplicate detection on Client + User pair (case-insensitive)
  const wantClient = norm(payload.clientName);
  const wantUser = norm(payload.userName);

  for (let i = 1; i < values.length; i++) {
    const c = norm(values[i][idx.client] as string | number | boolean);
    const u = norm(values[i][idx.user] as string | number | boolean);

    if (c === wantClient && u === wantUser) {
      return {
        status: "Skipped",
        message: "Duplicate (Client + User already exists)."
      };
    }
  }

  // Build new row aligned with header structure
  const colCount = header.length;
  const row: (string | number | boolean)[] = new Array(colCount).fill("");

  const put = (i: number, v: string) => {
    if (i >= 0 && i < colCount) row[i] = v;
  };

  put(idx.user, payload.userName);
  put(idx.client, payload.clientName);
  put(idx.clientId, payload.clientCode || "");
  put(idx.pod, payload.pod || "");
  put(idx.sensitivity, payload.sensitivity || "High");
  put(idx.reviewing, payload.reviewing || "INC");

  // Append row
  const nextRowIndex = used.getRowIndex() + values.length;
  const target = sheet.getRangeByIndexes(
    nextRowIndex,
    used.getColumnIndex(),
    1,
    colCount
  );

  target.setValues([row]);

  return {
    status: "Processed",
    message: "Added to Database sheet."
  };
}