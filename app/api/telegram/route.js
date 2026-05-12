import { NextResponse } from "next/server";

import { isAllowedTelegramUser, UNAUTHORIZED_MESSAGE } from "../../../lib/permissions.js";
import {
  extractTelegramMessage,
  getMessageText,
  getTelegramUserId,
  sendTelegramMessage,
} from "../../../lib/telegram.js";
import { answerQuery } from "../../../lib/queryRouter.js";

export const runtime = "nodejs";

export async function GET() {
  return NextResponse.json({
    ok: true,
    service: "telegram-reporting-bot",
  });
}

export async function POST(request) {
  let update;
  try {
    update = await request.json();
  } catch {
    return NextResponse.json({ ok: false, error: "Invalid JSON" }, { status: 400 });
  }

  const message = extractTelegramMessage(update);
  if (!message?.chat?.id) {
    return NextResponse.json({ ok: true, ignored: true });
  }

  const chatId = message.chat.id;
  const userId = getTelegramUserId(message);
  const text = getMessageText(message);

  try {
    if (!isAllowedTelegramUser(userId)) {
      await sendTelegramMessage(chatId, UNAUTHORIZED_MESSAGE);
      return NextResponse.json({ ok: true });
    }

    const answer = await answerQuery(text);
    await sendTelegramMessage(chatId, answer);
    return NextResponse.json({ ok: true });
  } catch (error) {
    console.error("Telegram webhook failed", error);

    try {
      await sendTelegramMessage(
        chatId,
        "Sorry, I could not calculate that report right now. Please try again later.",
      );
    } catch (sendError) {
      console.error("Telegram error reply failed", sendError);
    }

    return NextResponse.json(
      {
        ok: false,
        error: "Webhook processing failed",
      },
      { status: 500 },
    );
  }
}
