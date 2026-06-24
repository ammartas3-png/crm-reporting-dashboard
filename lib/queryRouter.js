import { getTabConfig } from "../config/sheetsConfig.js";
import { calculateMetric, parseMonth } from "./calculations.js";
import { readSheetRows } from "./googleSheets.js";
import { START_MESSAGE } from "./telegram.js";

const COUNTRY_ALIASES = new Map([
  ["germany", "Germany"],
  ["deutschland", "Germany"],
  ["de", "Germany"],
  ["turkey", "Turkey"],
  ["turkiye", "Turkey"],
  ["türkiye", "Turkey"],
  ["tr", "Turkey"],
  ["uk", "United Kingdom"],
  ["united kingdom", "United Kingdom"],
  ["england", "United Kingdom"],
  ["spain", "Spain"],
  ["es", "Spain"],
  ["italy", "Italy"],
  ["it", "Italy"],
  ["france", "France"],
  ["fr", "France"],
]);

const HELP_MESSAGE = [
  "I can answer simple CRM reporting questions.",
  "",
  "Examples:",
  "- How many FTD today?",
  "- Germany total leads?",
  "- Ahmet total calls?",
  "- May Turkey leads count?",
  "- Last 4 months agent hire dates",
].join("\n");

function normalize(text) {
  return String(text || "")
    .trim()
    .toLocaleLowerCase("en-US");
}

function titleCase(text) {
  return String(text || "")
    .trim()
    .replace(/\s+/g, " ")
    .replace(/\b\w/g, (char) => char.toUpperCase());
}

function extractCountry(text) {
  const normalized = normalize(text);
  const aliases = [...COUNTRY_ALIASES.keys()].sort((a, b) => b.length - a.length);
  const match = aliases.find((alias) =>
    new RegExp(`\\b${alias.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}\\b`, "i").test(
      normalized,
    ),
  );
  return match ? COUNTRY_ALIASES.get(match) : null;
}

function extractDateFilter(text, now = new Date()) {
  const normalized = normalize(text);
  if (/\btoday\b/.test(normalized)) {
    return { type: "today" };
  }

  const month = parseMonth(normalized);
  if (month !== null) {
    return { type: "month", month, year: now.getUTCFullYear() };
  }

  return null;
}

function extractAgent(text) {
  const trimmed = String(text || "").trim();
  const totalCallsMatch = trimmed.match(/^(.+?)\s+total\s+calls?\b/i);
  if (totalCallsMatch) {
    return titleCase(totalCallsMatch[1]);
  }

  const agentMatch = trimmed.match(/\bagent\s+([a-zğüşöçıİĞÜŞÖÇ]+(?:\s+[a-zğüşöçıİĞÜŞÖÇ]+)*)/i);
  if (agentMatch) {
    return titleCase(agentMatch[1]);
  }

  return null;
}

function extractLastMonths(text) {
  const normalized = normalize(text);
  const englishMatch = normalized.match(/\blast\s+(\d{1,2})\s*(?:months?|month|m\w*)\b/);
  if (englishMatch) {
    return Number(englishMatch[1]);
  }

  const turkishMatch = normalized.match(/\bson\s+(\d{1,2})\s*ay\b/);
  if (turkishMatch) {
    return Number(turkishMatch[1]);
  }

  return null;
}

function isAgentHireDateQuery(text) {
  const normalized = normalize(text);
  const mentionsAgent =
    /\bagents?\b/.test(normalized) || /\bagnet\b/.test(normalized) || /\bisim/.test(normalized);
  const mentionsHireDate =
    /\bhire\b/.test(normalized) ||
    /\bstart\b/.test(normalized) ||
    /\bjoin/.test(normalized) ||
    /\bi[şs]e\b/.test(normalized) ||
    /\bgiri[şs]/.test(normalized);
  return mentionsAgent && mentionsHireDate;
}

function parseDateValue(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return value;
  }

  if (typeof value === "number") {
    const excelEpoch = Date.UTC(1899, 11, 30);
    return new Date(excelEpoch + value * 24 * 60 * 60 * 1000);
  }

  const text = String(value || "").trim();
  if (!text) {
    return null;
  }

  const parsed = new Date(text);
  return Number.isNaN(parsed.getTime()) ? null : parsed;
}

function toIsoDate(date) {
  return date.toISOString().slice(0, 10);
}

function headerCandidates() {
  return [
    "Hire Date",
    "Start Date",
    "Join Date",
    "Joining Date",
    "Date of Joining",
    "Employee Start Date",
    "Agent Start Date",
    "İşe Giriş Tarihi",
    "Ise Giris Tarihi",
  ];
}

function findColumnByCandidates(rows, candidates) {
  if (!Array.isArray(rows) || rows.length === 0) {
    return null;
  }

  const normalizedCandidates = candidates.map((item) => normalize(item));
  const headers = Object.keys(rows[0] || {});
  return (
    headers.find((header) => normalizedCandidates.includes(normalize(header))) || null
  );
}

function buildAgentHireDateReply(rows, tabConfig, months, now = new Date()) {
  const safeMonths = Number.isFinite(months) && months > 0 ? Math.trunc(months) : 4;
  const cutoff = new Date(now);
  cutoff.setUTCMonth(cutoff.getUTCMonth() - safeMonths);

  const agentColumn = tabConfig.agentColumn || findColumnByCandidates(rows, ["Agent", "First Call Agent"]);
  if (!agentColumn) {
    return "I could not find an agent column in the sheet configuration.";
  }

  const hireDateColumn = findColumnByCandidates(rows, headerCandidates());
  if (!hireDateColumn) {
    return (
      "I could not find a hire/start date column. " +
      "Expected headers like Hire Date, Start Date, Join Date, or İşe Giriş Tarihi."
    );
  }

  const agents = new Map();

  for (const row of rows) {
    const agentName = String(row[agentColumn] || "").trim();
    if (!agentName) {
      continue;
    }
    const hireDate = parseDateValue(row[hireDateColumn]);
    if (!hireDate || hireDate < cutoff) {
      continue;
    }

    const key = normalize(agentName);
    const current = agents.get(key);
    if (!current || hireDate < current.hireDate) {
      agents.set(key, { agentName, hireDate });
    }
  }

  const list = [...agents.values()].sort((a, b) => a.hireDate - b.hireDate);
  if (list.length === 0) {
    return `No agents with hire dates found in the last ${safeMonths} months.`;
  }

  const lines = list.map((item) => `- ${item.agentName} — ${toIsoDate(item.hireDate)}`);
  return [
    `Agents and hire dates from the last ${safeMonths} months:`,
    ...lines,
  ].join("\n");
}

function describeFilters(filters = {}) {
  const parts = [];
  if (filters.date?.type === "today") {
    parts.push("today");
  }
  if (filters.date?.type === "month") {
    const monthName = new Date(Date.UTC(filters.date.year, filters.date.month, 1)).toLocaleString(
      "en-US",
      { month: "long" },
    );
    parts.push(monthName);
  }
  if (filters.country) {
    parts.push(filters.country);
  }
  if (filters.agent) {
    parts.push(filters.agent);
  }
  if (filters.status) {
    parts.push(filters.status);
  }
  return parts.length ? ` (${parts.join(", ")})` : "";
}

export function parseQuery(text, now = new Date()) {
  const normalized = normalize(text);
  const lastMonths = extractLastMonths(text);
  const filters = {
    country: extractCountry(text),
    date: extractDateFilter(text, now),
  };

  if (!normalized || normalized === "/help") {
    return { type: "help" };
  }

  if (normalized === "/start") {
    return { type: "start" };
  }

  if (lastMonths && isAgentHireDateQuery(text)) {
    return {
      type: "agentHireDates",
      tabKey: "leads",
      filters: { months: lastMonths },
    };
  }

  if (/\bftd\b/.test(normalized)) {
    return {
      type: "metric",
      metric: { key: "ftdCount", label: "FTD", operation: "count" },
      tabKey: "ftd",
      filters,
    };
  }

  if (/\btotal\s+calls?\b/.test(normalized) || /\bcalls?\b/.test(normalized)) {
    filters.agent = extractAgent(text);
    return {
      type: "metric",
      metric: { key: "agentCalls", label: "total calls", operation: "count" },
      tabKey: "leads",
      filters,
    };
  }

  if (/\bleads?\b/.test(normalized)) {
    return {
      type: "metric",
      metric: { key: "leadsCount", label: "leads", operation: "count" },
      tabKey: "leads",
      filters,
    };
  }

  if (/\btransactions?\b|\bdeposit\b|\bwithdrawal\b/.test(normalized)) {
    if (/\bdeposit\b/.test(normalized)) {
      filters.status = "Deposit";
    }
    if (/\bwithdrawal\b/.test(normalized)) {
      filters.status = "Withdrawal";
    }
    return {
      type: "metric",
      metric: { key: "transactionAmount", label: "transaction amount", operation: "sum" },
      tabKey: "transactions",
      filters,
    };
  }

  return { type: "unknown" };
}

export async function answerQuery(text, options = {}) {
  const now = options.now || new Date();
  const parsed = parseQuery(text, now);

  if (parsed.type === "start") {
    return START_MESSAGE;
  }

  if (parsed.type === "help" || parsed.type === "unknown") {
    return HELP_MESSAGE;
  }

  if (parsed.type === "agentHireDates") {
    const tabConfig = options.getTabConfig ? options.getTabConfig(parsed.tabKey) : getTabConfig(parsed.tabKey);
    const readRows = options.readRows || readSheetRows;
    const rows = await readRows(parsed.tabKey, { tabConfig });
    return buildAgentHireDateReply(rows, tabConfig, parsed.filters?.months, now);
  }

  const tabConfig = options.getTabConfig
    ? options.getTabConfig(parsed.tabKey)
    : getTabConfig(parsed.tabKey);
  const readRows = options.readRows || readSheetRows;
  const rows = await readRows(parsed.tabKey, { tabConfig });
  const value = calculateMetric(parsed.metric, rows, tabConfig, parsed.filters, now);
  const suffix = describeFilters(parsed.filters);

  if (parsed.metric.operation === "sum") {
    return `${parsed.metric.label}${suffix}: ${value.toLocaleString("en-US")}`;
  }

  return `${parsed.metric.label}${suffix}: ${value}`;
}
