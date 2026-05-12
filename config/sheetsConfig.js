const DEFAULT_LEADS_COLUMNS = [
  "Brand",
  "ID",
  "Created",
  "Department",
  "Status",
  "Country",
  "Campaign",
  "First Call Agent",
  "Team Leader",
  "FTD",
  "Office",
];

const DEFAULT_FTD_COLUMNS = ["Date", "Customer ID", "Agent", "Country", "Amount"];

const DEFAULT_TRANSACTION_COLUMNS = [
  "Date",
  "Customer ID",
  "Amount",
  "Type",
  "Country",
];

export const DEFAULT_GOOGLE_SPREADSHEET_ID = "1cXyL60QniZevYOb06adN5FPHWN5tbYhiHX12yIa6kG4";
export const DEFAULT_GOOGLE_SERVICE_ACCOUNT_EMAIL =
  "ammar-265@rapid-chassis-424212-r3.iam.gserviceaccount.com";
export const DEFAULT_LEADS_TAB = "May 26 Turkey  Leads";

export function quoteSheetName(sheetName) {
  return `'${String(sheetName).replace(/'/g, "''")}'`;
}

export function sheetRange(sheetName, columns = "A:Z") {
  return `${quoteSheetName(sheetName)}!${columns}`;
}

const leadsTabName = process.env.GOOGLE_LEADS_TAB || DEFAULT_LEADS_TAB;
const ftdTabName = process.env.GOOGLE_FTD_TAB || "FTD";
const transactionTabName = process.env.GOOGLE_TRANSACTION_TAB || "TRANSACTION";

export const sheetsConfig = {
  spreadsheetId: process.env.GOOGLE_SPREADSHEET_ID || DEFAULT_GOOGLE_SPREADSHEET_ID,
  serviceAccountEmail:
    process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL || DEFAULT_GOOGLE_SERVICE_ACCOUNT_EMAIL,
  tabs: {
    leads: {
      key: "leads",
      name: leadsTabName,
      range: process.env.GOOGLE_LEADS_RANGE || sheetRange(leadsTabName),
      dateColumn: "Created",
      countryColumn: "Country",
      agentColumn: "First Call Agent",
      statusColumn: "Status",
      amountColumn: null,
      columns: DEFAULT_LEADS_COLUMNS,
    },
    ftd: {
      key: "ftd",
      name: ftdTabName,
      range: process.env.GOOGLE_FTD_RANGE || sheetRange(ftdTabName),
      dateColumn: "Date",
      countryColumn: "Country",
      agentColumn: "Agent",
      statusColumn: null,
      amountColumn: "Amount",
      columns: DEFAULT_FTD_COLUMNS,
    },
    transactions: {
      key: "transactions",
      name: transactionTabName,
      range: process.env.GOOGLE_TRANSACTION_RANGE || sheetRange(transactionTabName),
      dateColumn: "Date",
      countryColumn: "Country",
      agentColumn: null,
      statusColumn: "Type",
      amountColumn: "Amount",
      columns: DEFAULT_TRANSACTION_COLUMNS,
    },
  },
};

export function getTabConfig(tabKey) {
  const tab = sheetsConfig.tabs[tabKey];
  if (!tab) {
    throw new Error(`Unknown sheet tab config: ${tabKey}`);
  }
  return tab;
}
