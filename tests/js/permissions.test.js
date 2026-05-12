import assert from "node:assert/strict";
import test from "node:test";

import { isAllowedTelegramUser, parseAllowedUsers } from "../../lib/permissions.js";

test("parseAllowedUsers reads comma-separated Telegram IDs", () => {
  const users = parseAllowedUsers("123, 456,789");

  assert.equal(users.has("123"), true);
  assert.equal(users.has("456"), true);
  assert.equal(users.has("789"), true);
});

test("isAllowedTelegramUser denies empty config by default", () => {
  assert.equal(isAllowedTelegramUser(123, new Set()), false);
});

test("isAllowedTelegramUser compares IDs as strings", () => {
  assert.equal(isAllowedTelegramUser(123, new Set(["123"])), true);
  assert.equal(isAllowedTelegramUser(999, new Set(["123"])), false);
});
