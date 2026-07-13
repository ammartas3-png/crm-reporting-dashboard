"""Helpers for chunked large-file uploads on size-limited serverless hosts."""

from __future__ import annotations

import json
import re
import shutil
from pathlib import Path

UPLOAD_ROOT = Path("/tmp/crm-uploads")
CHUNK_UPLOAD_MAX_BYTES = 3 * 1024 * 1024


def sanitize_upload_id(upload_id: str) -> str:
    cleaned = re.sub(r"[^a-zA-Z0-9_-]", "", str(upload_id or "").strip())
    if not cleaned:
        raise ValueError("Invalid upload id.")
    return cleaned


def upload_workspace(upload_id: str) -> Path:
    workspace = UPLOAD_ROOT / sanitize_upload_id(upload_id)
    workspace.mkdir(parents=True, exist_ok=True)
    return workspace


def save_chunk(
    upload_id: str,
    chunk_index: int,
    total_chunks: int,
    filename: str,
    chunk_bytes: bytes,
) -> Path | None:
    if total_chunks < 1:
        raise ValueError("total_chunks must be at least 1.")
    if chunk_index < 0 or chunk_index >= total_chunks:
        raise ValueError("chunk_index is out of range.")
    if len(chunk_bytes) > CHUNK_UPLOAD_MAX_BYTES:
        raise ValueError(
            f"Each upload chunk must be {CHUNK_UPLOAD_MAX_BYTES // (1024 * 1024)} MB or smaller."
        )

    workspace = upload_workspace(upload_id)
    (workspace / f"chunk_{chunk_index}").write_bytes(chunk_bytes)
    (workspace / "meta.json").write_text(
        json.dumps({"filename": filename, "total_chunks": total_chunks}),
        encoding="utf-8",
    )

    if all((workspace / f"chunk_{index}").exists() for index in range(total_chunks)):
        return finalize_upload(workspace)
    return None


def finalize_upload(workspace: Path) -> Path:
    meta = json.loads((workspace / "meta.json").read_text(encoding="utf-8"))
    filename = Path(str(meta["filename"])).name or "upload.xlsx"
    total_chunks = int(meta["total_chunks"])

    output_path = workspace / filename
    with output_path.open("wb") as output_file:
        for index in range(total_chunks):
            chunk_path = workspace / f"chunk_{index}"
            if not chunk_path.exists():
                raise ValueError(f"Missing upload chunk {index}.")
            output_file.write(chunk_path.read_bytes())

    (workspace / "ready").write_text(str(output_path), encoding="utf-8")
    return output_path


def resolve_uploaded_file(upload_id: str) -> Path:
    workspace = UPLOAD_ROOT / sanitize_upload_id(upload_id)
    ready_marker = workspace / "ready"
    if not ready_marker.exists():
        raise ValueError("Uploaded file is not ready. Please finish uploading all chunks first.")

    output_path = Path(ready_marker.read_text(encoding="utf-8").strip())
    if not output_path.exists():
        raise ValueError("Uploaded file is missing on the server.")
    return output_path


def cleanup_upload(upload_id: str) -> None:
    workspace = UPLOAD_ROOT / sanitize_upload_id(upload_id)
    if workspace.exists():
        shutil.rmtree(workspace, ignore_errors=True)
