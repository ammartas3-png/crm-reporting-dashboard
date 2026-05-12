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

export const sheetsConfig = {
  spreadsheetId: process.env.GOOGLE_SPREADSHEET_ID || "",
  tabs: {
    leads: {
      key: "leads",
      name: process.env.GOOGLE_LEADS_TAB || "Leads",
      range: process.env.GOOGLE_LEADS_RANGE || "Leads!A:Z",
      dateColumn: "Created",
      countryColumn: "Country",
      agentColumn: "First Call Agent",
      statusColumn: "Status",
      amountColumn: null,
      columns: DEFAULT_LEADS_COLUMNS,
    },
    ftd: {
      key: "ftd",
      name: process.env.GOOGLE_FTD_TAB || "FTD",
      range: process.env.GOOGLE_FTD_RANGE || "FTD!A:Z",
      dateColumn: "Date",
      countryColumn: "Country",
      agentColumn: "Agent",
      statusColumn: null,
      amountColumn: "Amount",
      columns: DEFAULT_FTD_COLUMNS,
    },
    transactions: {
      key: "transactions",
      name: process.env.GOOGLE_TRANSACTION_TAB || "TRANSACTION",
      range: process.env.GOOGLE_TRANSACTION_RANGE || "TRANSACTION!A:Z",
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
