import { google } from "googleapis";

import { getTabConfig, sheetsConfig } from "../config/sheetsConfig.js";

function getPrivateKey() {
  const privateKey = process.env.GOOGLE_PRIVATE_KEY;
  return privateKey ? privateKey.replace(/\\n/g, "\n") : "";
}

export function getSheetsAuth() {
  const email = process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL;
  const privateKey = getPrivateKey();

  if (!email || !privateKey) {
    throw new Error(
      "Google Sheets credentials are not configured. Set GOOGLE_SERVICE_ACCOUNT_EMAIL and GOOGLE_PRIVATE_KEY.",
    );
  }

  return new google.auth.JWT({
    email,
    key: privateKey,
    scopes: ["https://www.googleapis.com/auth/spreadsheets.readonly"],
  });
}

export function getSheetsClient(auth = getSheetsAuth()) {
  return google.sheets({ version: "v4", auth });
}

export function rowsToObjects(values = [], expectedColumns = []) {
  if (!Array.isArray(values) || values.length === 0) {
    return [];
  }

  const headerRow = values[0] || [];
  const headers = headerRow.map((header) => String(header || "").trim());
  const usableHeaders = headers.some(Boolean) ? headers : expectedColumns;

  return values.slice(1).map((row) => {
    const item = {};
    usableHeaders.forEach((header, index) => {
      if (!header) {
        return;
      }
      item[header] = row[index] ?? "";
    });
    return item;
  });
}

export async function readSheetRows(tabKey, options = {}) {
  const tabConfig = options.tabConfig || getTabConfig(tabKey);
  const spreadsheetId =
    options.spreadsheetId || sheetsConfig.spreadsheetId || process.env.GOOGLE_SPREADSHEET_ID;

  if (!spreadsheetId) {
    throw new Error("GOOGLE_SPREADSHEET_ID is not configured.");
  }

  const sheets = options.sheetsClient || getSheetsClient();
  const response = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: tabConfig.range,
    valueRenderOption: "UNFORMATTED_VALUE",
    dateTimeRenderOption: "FORMATTED_STRING",
  });

  return rowsToObjects(response.data.values || [], tabConfig.columns);
}
