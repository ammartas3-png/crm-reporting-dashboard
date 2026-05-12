export const UNAUTHORIZED_MESSAGE = "You are not authorized to use this bot.";

export function parseAllowedUsers(rawAllowedUsers = process.env.ALLOWED_USERS || "") {
  return new Set(
    String(rawAllowedUsers)
      .split(",")
      .map((value) => value.trim())
      .filter(Boolean),
  );
}

export function isAllowedTelegramUser(userId, allowedUsers = parseAllowedUsers()) {
  if (userId === undefined || userId === null) {
    return false;
  }

  if (allowedUsers.size === 0) {
    return false;
  }

  return allowedUsers.has(String(userId));
}
