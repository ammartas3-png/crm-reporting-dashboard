import assert from "node:assert/strict";
import test from "node:test";

import { answerQuery, parseQuery } from "../../lib/queryRouter.js";

const NOW = new Date("2026-05-12T10:00:00Z");

const tabConfigs = {
  leads: {
    dateColumn: "Created",
    countryColumn: "Country",
    agentColumn: "First Call Agent",
    statusColumn: "Status",
    amountColumn: null,
  },
  ftd: {
    dateColumn: "Date",
    countryColumn: "Country",
    agentColumn: "Agent",
    statusColumn: null,
    amountColumn: "Amount",
  },
  transactions: {
    dateColumn: "Date",
    countryColumn: "Country",
    agentColumn: null,
    statusColumn: "Type",
    amountColumn: "Amount",
  },
};

const data = {
  leads: [
    {
      Created: "2026-05-01",
      Country: "Turkey",
      "First Call Agent": "Ahmet",
      Status: "Potential",
    },
    {
      Created: "2026-05-02",
      Country: "Turkey",
      "First Call Agent": "Ayse",
      Status: "Potential",
    },
    {
      Created: "2026-04-02",
      Country: "Germany",
      "First Call Agent": "Ahmet",
      Status: "Potential",
    },
  ],
  ftd: [
    { Date: "2026-05-12", Country: "Turkey", Agent: "Ahmet", Amount: 100 },
    { Date: "2026-05-11", Country: "Germany", Agent: "Max", Amount: 200 },
  ],
  transactions: [
    { Date: "2026-05-12", Country: "Turkey", Type: "Deposit", Amount: 100 },
    { Date: "2026-05-12", Country: "Turkey", Type: "Withdrawal", Amount: 40 },
  ],
};

function answer(text) {
  return answerQuery(text, {
    now: NOW,
    getTabConfig: (tabKey) => tabConfigs[tabKey],
    readRows: async (tabKey) => data[tabKey],
  });
}

test("parseQuery routes FTD questions to the FTD tab", () => {
  const parsed = parseQuery("How many FTD today?", NOW);

  assert.equal(parsed.type, "metric");
  assert.equal(parsed.tabKey, "ftd");
  assert.equal(parsed.metric.key, "ftdCount");
  assert.deepEqual(parsed.filters.date, { type: "today" });
});

test("answerQuery calculates FTD today count", async () => {
  assert.equal(await answer("How many FTD today?"), "FTD (today): 1");
});

test("answerQuery calculates country leads", async () => {
  assert.equal(await answer("Germany total leads?"), "leads (Germany): 1");
});

test("answerQuery calculates agent total calls", async () => {
  assert.equal(await answer("Ahmet total calls?"), "total calls (Ahmet): 2");
});

test("answerQuery applies month and country filters", async () => {
  assert.equal(await answer("May Turkey leads count?"), "leads (May, Turkey): 2");
});

test("answerQuery can sum simple transaction amounts", async () => {
  assert.equal(await answer("Turkey deposit transactions today"), "transaction amount (today, Turkey, Deposit): 100");
});
