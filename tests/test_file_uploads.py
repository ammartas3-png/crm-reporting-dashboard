from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from api.file_uploads import (
    cleanup_upload,
    read_upload_metadata,
    resolve_uploaded_file,
    save_chunk,
)


class FileUploadStorageTests(unittest.TestCase):
    def test_save_chunk_assembles_uploaded_file(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            upload_id = "test-upload-123"
            original = Path(temp_dir) / "source.xlsx"
            original.write_bytes(b"monthly-comments-workbook")

            first = b"monthly-"
            second = b"comments-workbook"
            save_chunk(upload_id, 0, 2, "monthly.xlsx", first)
            output_path = save_chunk(upload_id, 1, 2, "monthly.xlsx", second)

            self.assertIsNotNone(output_path)
            self.assertEqual(resolve_uploaded_file(upload_id).read_bytes(), original.read_bytes())
            cleanup_upload(upload_id)


    def test_save_chunk_stores_report_metadata(self) -> None:
        upload_id = "test-upload-meta"
        save_chunk(
            upload_id,
            0,
            1,
            "monthly.xlsx",
            b"workbook",
            pivot_name="ZA July",
            program="program_c",
        )
        self.assertEqual(read_upload_metadata(upload_id)["pivot_name"], "ZA July")
        self.assertEqual(read_upload_metadata(upload_id)["program"], "program_c")
        cleanup_upload(upload_id)


if __name__ == "__main__":
    unittest.main()
