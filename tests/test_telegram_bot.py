from __future__ import annotations

import unittest

from api.telegram import build_reply, is_authorized


def _message_update(text: str | None = None, **message_fields: object) -> dict[str, object]:
    message: dict[str, object] = {"chat": {"id": 12345}}
    if text is not None:
        message["text"] = text
    message.update(message_fields)
    return {"message": message}


class TelegramBotTests(unittest.TestCase):
    def test_start_reply_lists_commands_and_dashboard(self) -> None:
        reply = build_reply(_message_update("/start"), dashboard_url="https://example.test/")

        self.assertIsNotNone(reply)
        self.assertEqual(reply["method"], "sendMessage")
        self.assertEqual(reply["chat_id"], 12345)
        self.assertIn("/programa", str(reply["text"]))
        self.assertIn("Panel: https://example.test", str(reply["text"]))

    def test_columns_reply_contains_required_headers(self) -> None:
        reply = build_reply(_message_update("/columns"))

        self.assertIsNotNone(reply)
        text = str(reply["text"])
        self.assertIn("Customer Type", text)
        self.assertIn("Voip Calls Attempts Cnt", text)

    def test_document_reply_directs_user_to_dashboard(self) -> None:
        reply = build_reply(
            _message_update(document={"file_name": "crm.xlsx"}),
            dashboard_url="https://reports.example",
        )

        self.assertIsNotNone(reply)
        text = str(reply["text"])
        self.assertIn("Excel dosyasi aldim", text)
        self.assertIn("https://reports.example", text)

    def test_unknown_message_shows_help_hint(self) -> None:
        reply = build_reply(_message_update("selam"))

        self.assertIsNotNone(reply)
        self.assertIn("/help", str(reply["text"]))

    def test_authorization_uses_telegram_secret_header(self) -> None:
        self.assertTrue(
            is_authorized({"X-Telegram-Bot-Api-Secret-Token": "secret"}, "secret")
        )
        self.assertFalse(
            is_authorized({"X-Telegram-Bot-Api-Secret-Token": "wrong"}, "secret")
        )
        self.assertTrue(is_authorized({}, ""))


if __name__ == "__main__":
    unittest.main()
