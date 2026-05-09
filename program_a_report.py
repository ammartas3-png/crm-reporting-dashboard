"""Program A: CRM output workbook enriched with PowerBI comments.

This module names the currently deployed report generator explicitly so the web
dashboard can route between Program A and Program B without changing Program A's
existing implementation.
"""

from __future__ import annotations

from report_generator import build_output, main

__all__ = ["build_output", "main"]


if __name__ == "__main__":
    raise SystemExit(main())
