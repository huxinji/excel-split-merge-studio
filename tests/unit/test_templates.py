from __future__ import annotations

import json
import tempfile
import unittest
from pathlib import Path

from excel_studio.services.template_service import TemplateService


class TemplateServiceTests(unittest.TestCase):
    def test_template_removes_sensitive_paths(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            path = Path(temporary) / "template.json"
            TemplateService.save(
                path,
                {
                    "operation": "split",
                    "input_files": ["C:/secret/source.xlsx"],
                    "output": {
                        "directory": "C:/secret/output",
                        "naming_template": "{original_name}",
                    },
                },
            )
            payload = json.loads(path.read_text(encoding="utf-8"))
            self.assertNotIn("input_files", payload)
            self.assertNotIn("directory", payload["output"])
            self.assertEqual(payload["output"]["naming_template"], "{original_name}")
            self.assertEqual(TemplateService.load(path), payload)


if __name__ == "__main__":
    unittest.main()
