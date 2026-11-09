"""HTTP contract and local-file isolation checks; Office is never invoked."""

from __future__ import annotations

import base64
import http.client
import importlib.util
import io
import json
from pathlib import Path
import sys
import threading
import unittest
from unittest.mock import patch

from PIL import Image

sys.dont_write_bytecode = True
ROOT = Path(__file__).resolve().parents[1]
SPEC = importlib.util.spec_from_file_location("fastresume_serve_under_test", ROOT / "scripts/serve.py")
if SPEC is None or SPEC.loader is None:
    raise RuntimeError("Cannot load local editor")
server_module = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(server_module)


class EditorHttpTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.server = server_module.create_server(0)
        cls.port = cls.server.server_address[1]
        cls.thread = threading.Thread(target=cls.server.serve_forever, daemon=True)
        cls.thread.start()

    @classmethod
    def tearDownClass(cls):
        cls.server.shutdown()
        cls.server.server_close()
        cls.thread.join(timeout=5)

    def request(self, method, path, payload=None, headers=None):
        connection = http.client.HTTPConnection("127.0.0.1", self.port, timeout=10)
        options = {"Content-Type": "application/json"} if payload is not None else {}
        options.update(headers or {})
        content = None if payload is None else json.dumps(payload, ensure_ascii=False).encode("utf-8")
        try:
            connection.request(method, path, body=content, headers=options)
            response = connection.getresponse()
            return response.status, dict(response.getheaders()), response.read()
        finally:
            connection.close()

    def test_home_and_example_presets_are_available_without_filesystem_paths(self):
        status, headers, body = self.request("GET", "/")
        self.assertEqual(200, status)
        self.assertIn("fastResume", body.decode("utf-8"))
        self.assertIn("no-store", headers["Cache-Control"])
        status, _, body = self.request("GET", "/api/examples")
        self.assertEqual(200, status)
        examples = json.loads(body)["examples"]
        self.assertEqual({"template", "anon", "tomori", "sakiko"}, {item["id"] for item in examples})
        for item in examples:
            self.assertIn("basics", item["resume"])
            self.assertIn("internship_experiences", item["resume"])
            self.assertNotIn("path", item)

    def test_preview_normalizes_yaml_and_rejects_invalid_resume(self):
        payload = {"yaml": "basics:\n  name: 测试候选人\n  details: 每周五天\ninternship_experiences:\n  - name: 测试团队\n    highlights: [完成接口开发]\n"}
        status, _, body = self.request("POST", "/api/preview", payload)
        self.assertEqual(200, status)
        normalized = json.loads(body)
        self.assertEqual("每周五天", normalized["resume"]["basics"]["details"])
        self.assertEqual("测试团队", normalized["resume"]["internship_experiences"][0]["name"])
        self.assertIn("测试候选人", normalized["yaml"])
        status, _, body = self.request("POST", "/api/preview", {"yaml": "basics: {name: ''}"})
        self.assertEqual(400, status)
        self.assertIn("basics.name", json.loads(body)["error"])

    def test_cross_site_and_nonlocal_host_requests_are_rejected(self):
        status, headers, _ = self.request("POST", "/api/preview", {"resume": {"basics": {"name": "测试"}}}, {"Origin": "https://unrelated.example"})
        self.assertEqual(403, status)
        self.assertNotIn("Access-Control-Allow-Origin", headers)
        status, _, _ = self.request("GET", "/api/examples", headers={"Host": f"unrelated.example:{self.port}"})
        self.assertEqual(403, status)
        status, _, _ = self.request("GET", "/../scripts/build_resume.py")
        self.assertEqual(404, status)

    def test_download_uses_temporary_output_and_cleans_it_after_response(self):
        created = []

        def generate(data, output, formats):
            self.assertEqual(("docx",), formats)
            path = Path(str(output) + ".docx")
            path.write_bytes(b"download fixture")
            created.append(path.parent)
            return {"docx": str(path), "pages": None}

        with patch.object(server_module.generator, "generate_resume", side_effect=generate):
            status, headers, body = self.request("POST", "/api/export?format=docx", {"resume": {"basics": {"name": "测试候选人"}}})
        self.assertEqual(200, status)
        self.assertEqual(b"download fixture", body)
        self.assertIn("filename*=UTF-8''", headers["Content-Disposition"])
        self.assertTrue(created)
        self.assertTrue(all(not path.exists() for path in created))

    def test_failed_export_cleans_partial_output_and_unlocks_next_export(self):
        created = []

        def fail(data, output, formats):
            Path(str(output) + ".pdf").write_bytes(b"partial fixture")
            created.append(Path(output).parent)
            raise server_module.generator.ResumeError("模拟转换失败")

        with patch.object(server_module.generator, "generate_resume", side_effect=fail):
            status, _, body = self.request("POST", "/api/export?format=pdf", {"resume": {"basics": {"name": "测试候选人"}}})
        self.assertEqual(500, status)
        self.assertIn("模拟转换失败", json.loads(body)["error"])
        self.assertTrue(all(not path.exists() for path in created))
        self.assertTrue(self.server.export_lock.acquire(blocking=False))
        self.server.export_lock.release()

    def test_photo_round_trip_is_portable_and_can_be_removed(self):
        stream = io.BytesIO()
        Image.new("RGB", (3, 4), "#e5eef1").save(stream, format="PNG")
        uri = "data:image/png;base64," + base64.b64encode(stream.getvalue()).decode("ascii")
        data = {"basics": {"name": "测试候选人", "photo": uri}}
        status, _, body = self.request("POST", "/api/preview", {"resume": data})
        self.assertEqual(200, status)
        result = json.loads(body)
        self.assertEqual(uri, result["resume"]["basics"]["photo"])
        self.assertIn(uri, result["yaml"])
        data["basics"]["photo"] = ""
        status, _, body = self.request("POST", "/api/preview", {"resume": data})
        self.assertEqual(200, status)
        self.assertEqual("", json.loads(body)["resume"]["basics"]["photo"])

    def test_invalid_photo_is_rejected_before_any_export(self):
        with patch.object(server_module.generator, "generate_resume") as generate:
            for photo in ("file:///private/photo.png", "https://example.com/photo.png", "data:image/svg+xml;base64,PHN2Zz48L3N2Zz4=", "data:image/png;base64,bm90LWFuLWltYWdl"):
                with self.subTest(photo=photo):
                    status, _, body = self.request("POST", "/api/export?format=docx", {"resume": {"basics": {"name": "测试候选人", "photo": photo}}})
                    self.assertEqual(400, status)
                    self.assertIn("basics.photo", json.loads(body)["error"])
            generate.assert_not_called()


if __name__ == "__main__":
    unittest.main()
