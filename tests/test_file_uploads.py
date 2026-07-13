from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from api.file_uploads import (
    cleanup_upload,
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


if __name__ == "__main__":
    unittest.main()
