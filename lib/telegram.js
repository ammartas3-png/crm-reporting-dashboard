export const START_MESSAGE = [
  "CRM Reporting Bot is ready.",
  "",
  "Ask questions like:",
  "- How many FTD today?",
  "- Germany total leads?",
  "- Ahmet total calls?",
  "- May Turkey leads count?",
  "",
  "Data access is limited to authorized Telegram users.",
].join("\n");

export function extractTelegramMessage(update) {
  if (!update || typeof update !== "object") {
    return null;
  }

  return (
    update.message ||
    update.edited_message ||
    update.channel_post ||
    update.edited_channel_post ||
    null
  );
}

export function getMessageText(message) {
  return String(message?.text || "").trim();
}

export function getTelegramUserId(message) {
  return message?.from?.id ?? null;
}

export async function sendTelegramMessage(chatId, text, options = {}) {
  const token = options.token || process.env.TELEGRAM_BOT_TOKEN;
  if (!token) {
    throw new Error("TELEGRAM_BOT_TOKEN is not configured.");
  }

  const response = await fetch(`https://api.telegram.org/bot${token}/sendMessage`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      chat_id: chatId,
      text,
      disable_web_page_preview: true,
    }),
  });

  const payload = await response.json().catch(() => ({}));
  if (!response.ok || payload.ok === false) {
    const description = payload.description || response.statusText;
    throw new Error(`Telegram sendMessage failed: ${description}`);
  }

  return payload;
}
