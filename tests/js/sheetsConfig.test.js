import assert from "node:assert/strict";
import test from "node:test";

import {
  DEFAULT_GOOGLE_SERVICE_ACCOUNT_EMAIL,
  DEFAULT_GOOGLE_SPREADSHEET_ID,
  DEFAULT_LEADS_TAB,
  quoteSheetName,
  sheetRange,
  sheetsConfig,
} from "../../config/sheetsConfig.js";

test("sheetsConfig uses the provided Google Sheet by default", () => {
  assert.equal(sheetsConfig.spreadsheetId, DEFAULT_GOOGLE_SPREADSHEET_ID);
  assert.equal(sheetsConfig.serviceAccountEmail, DEFAULT_GOOGLE_SERVICE_ACCOUNT_EMAIL);
  assert.equal(sheetsConfig.tabs.leads.name, DEFAULT_LEADS_TAB);
});

test("sheetRange quotes tab names with spaces", () => {
  assert.equal(sheetRange("May 26 Turkey  Leads"), "'May 26 Turkey  Leads'!A:Z");
});

test("quoteSheetName escapes apostrophes for Google A1 notation", () => {
  assert.equal(quoteSheetName("May's Leads"), "'May''s Leads'");
});
