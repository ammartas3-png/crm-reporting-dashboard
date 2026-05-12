const MONTHS = {
  january: 0,
  jan: 0,
  february: 1,
  feb: 1,
  march: 2,
  mar: 2,
  april: 3,
  apr: 3,
  may: 4,
  june: 5,
  jun: 5,
  july: 6,
  jul: 6,
  august: 7,
  aug: 7,
  september: 8,
  sep: 8,
  october: 9,
  oct: 9,
  november: 10,
  nov: 10,
  december: 11,
  dec: 11,
};

function normalizeText(value) {
  return String(value || "")
    .trim()
    .toLocaleLowerCase("en-US");
}

function getRowValue(row, columnName) {
  if (!columnName) {
    return "";
  }

  if (Object.prototype.hasOwnProperty.call(row, columnName)) {
    return row[columnName];
  }

  const normalizedColumn = normalizeText(columnName);
  const foundKey = Object.keys(row).find((key) => normalizeText(key) === normalizedColumn);
  return foundKey ? row[foundKey] : "";
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

function sameUtcDate(left, right) {
  return (
    left.getUTCFullYear() === right.getUTCFullYear() &&
    left.getUTCMonth() === right.getUTCMonth() &&
    left.getUTCDate() === right.getUTCDate()
  );
}

export function parseMonth(text) {
  const normalized = normalizeText(text);
  const found = Object.entries(MONTHS).find(([name]) =>
    new RegExp(`\\b${name}\\b`, "i").test(normalized),
  );
  return found ? found[1] : null;
}

export function dateMatches(value, filter, now = new Date()) {
  if (!filter) {
    return true;
  }

  const date = parseDateValue(value);
  if (!date) {
    return false;
  }

  if (filter.type === "today") {
    return sameUtcDate(date, now);
  }

  if (filter.type === "month") {
    const year = filter.year || now.getUTCFullYear();
    return date.getUTCFullYear() === year && date.getUTCMonth() === filter.month;
  }

  return true;
}

export function rowMatchesFilters(row, tabConfig, filters = {}, now = new Date()) {
  if (filters.country) {
    const country = normalizeText(getRowValue(row, tabConfig.countryColumn));
    if (country !== normalizeText(filters.country)) {
      return false;
    }
  }

  if (filters.agent) {
    const agent = normalizeText(getRowValue(row, tabConfig.agentColumn));
    if (!agent.includes(normalizeText(filters.agent))) {
      return false;
    }
  }

  if (filters.status) {
    const status = normalizeText(getRowValue(row, tabConfig.statusColumn));
    if (status !== normalizeText(filters.status)) {
      return false;
    }
  }

  if (filters.date) {
    const dateValue = getRowValue(row, tabConfig.dateColumn);
    if (!dateMatches(dateValue, filters.date, now)) {
      return false;
    }
  }

  return true;
}

export function countRows(rows, tabConfig, filters = {}, now = new Date()) {
  return rows.filter((row) => rowMatchesFilters(row, tabConfig, filters, now)).length;
}

export function sumRows(rows, tabConfig, filters = {}, now = new Date()) {
  return rows
    .filter((row) => rowMatchesFilters(row, tabConfig, filters, now))
    .reduce((total, row) => {
      const rawAmount = getRowValue(row, tabConfig.amountColumn);
      const amount = Number(rawAmount);
      return Number.isFinite(amount) ? total + amount : total;
    }, 0);
}

export function calculateMetric(metric, rows, tabConfig, filters = {}, now = new Date()) {
  if (metric.operation === "sum") {
    return sumRows(rows, tabConfig, filters, now);
  }

  return countRows(rows, tabConfig, filters, now);
}
