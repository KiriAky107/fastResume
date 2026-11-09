#!/usr/bin/env python3
"""Run the private, local fastResume editor at http://127.0.0.1:8765."""

from __future__ import annotations

import argparse
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
import importlib.util
import json
from pathlib import Path
import re
import sys
import tempfile
import threading
from urllib.parse import parse_qs, quote, urlsplit

import yaml

sys.dont_write_bytecode = True
ROOT = Path(__file__).resolve().parents[1]
SPEC = importlib.util.spec_from_file_location("fastresume_web_generator", ROOT / "scripts/build_resume.py")
if SPEC is None or SPEC.loader is None:
    raise RuntimeError("Cannot load scripts/build_resume.py")
generator = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(generator)

MAX_BODY = 6 * 1024 * 1024
PRESETS = (
    ("template", "通用模板", "resume.template.yaml"),
    ("anon", "千早爱音 · 粉色提案", "anon.yaml"),
    ("tomori", "高松灯 · 蓝色记录", "tomori.yaml"),
    ("sakiko", "丰川祥子 · 紫色编排", "sakiko.yaml"),
)
ASSETS = {
    "/": ("index.html", "text/html; charset=utf-8"),
    "/index.html": ("index.html", "text/html; charset=utf-8"),
    "/app.js": ("app.js", "text/javascript; charset=utf-8"),
    "/style.css": ("style.css", "text/css; charset=utf-8"),
}


class ApiError(ValueError):
    def __init__(self, message: str, status: int = 400):
        super().__init__(message)
        self.status = status


def _canonical(data: dict) -> str:
    return yaml.safe_dump(data, allow_unicode=True, sort_keys=False, width=100)


def _resume(payload: dict) -> dict:
    if not isinstance(payload, dict):
        raise ApiError("请求内容必须是 JSON 对象。")
    if ("yaml" in payload) == ("resume" in payload):
        raise ApiError("请提供 yaml 文本或 resume 对象中的一项。")
    try:
        if "yaml" in payload:
            if not isinstance(payload["yaml"], str):
                raise ApiError("yaml 必须是文本。")
            raw = yaml.safe_load(payload["yaml"])
        else:
            raw = payload["resume"]
        return generator.normalize_resume(raw)
    except (yaml.YAMLError, generator.ResumeError, RecursionError) as exc:
        raise ApiError(str(exc)) from exc


def _filename(data: dict, format_name: str) -> str:
    name = re.sub(r'[\x00-\x1f\x7f<>:"/\\|?*]', "_", data["basics"]["name"])
    return (name.strip(" .")[:80] or "resume") + "." + format_name


class EditorServer(ThreadingHTTPServer):
    daemon_threads = True

    def __init__(self, address):
        self.export_lock = threading.Lock()
        super().__init__(address, EditorHandler)


class EditorHandler(BaseHTTPRequestHandler):
    server_version = "fastResume"

    def setup(self):
        super().setup()
        self.connection.settimeout(15)

    def _check_origin(self):
        expected_port = self.server.server_address[1]
        host = self.headers.get("Host", "")
        try:
            parsed = urlsplit("//" + host)
            safe_host = (parsed.hostname in {"127.0.0.1", "localhost"}
                         and parsed.port == expected_port and not parsed.username
                         and not parsed.password and not parsed.path and not parsed.query
                         and not parsed.fragment)
        except ValueError:
            safe_host = False
        if not safe_host:
            raise ApiError("仅允许通过本机地址访问编辑器。", 403)
        origin = self.headers.get("Origin")
        if origin is not None and origin != "http://" + host:
            raise ApiError("已拒绝来自其他网页的请求，请直接打开本地编辑器。", 403)
        top_level_page = (self.command == "GET" and urlsplit(self.path).path in {"/", "/index.html"}
                          and self.headers.get("Sec-Fetch-Mode") == "navigate")
        if self.headers.get("Sec-Fetch-Site") == "cross-site" and not top_level_page:
            raise ApiError("已拒绝跨站请求。", 403)

    def _send(self, body: bytes, content_type: str, status=200, filename=None):
        self.send_response(status)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.send_header("X-Content-Type-Options", "nosniff")
        self.send_header("Referrer-Policy", "no-referrer")
        self.send_header("Content-Security-Policy", "default-src 'self'; img-src 'self' data:; object-src 'none'; base-uri 'none'; frame-ancestors 'none'; form-action 'self'")
        if filename:
            extension = Path(filename).suffix
            self.send_header("Content-Disposition", f"attachment; filename=\"resume{extension}\"; filename*=UTF-8''{quote(filename, safe='')}")
        self.send_header("Connection", "close")
        self.end_headers()
        self.close_connection = True
        try:
            self.wfile.write(body)
        except (BrokenPipeError, ConnectionResetError):
            pass

    def _json(self, value, status=200):
        self._send(json.dumps(value, ensure_ascii=False).encode("utf-8"), "application/json; charset=utf-8", status)

    def _body(self):
        if self.headers.get("Transfer-Encoding"):
            raise ApiError("不支持分块请求。")
        if self.headers.get_content_type() != "application/json":
            raise ApiError("请使用 application/json 请求。", 415)
        lengths = self.headers.get_all("Content-Length", [])
        if len(lengths) != 1 or not lengths[0].isdigit():
            raise ApiError("请求缺少有效的 Content-Length。")
        try:
            length = int(lengths[0])
        except ValueError as exc:
            raise ApiError("Content-Length 数值无效。") from exc
        if length > MAX_BODY:
            raise ApiError("配置内容不能超过 6 MB。", 413)
        try:
            raw = self.rfile.read(length)
            if len(raw) != length:
                raise ApiError("请求内容不完整。")
            return json.loads(raw.decode("utf-8"))
        except (UnicodeError, json.JSONDecodeError, RecursionError) as exc:
            raise ApiError("JSON 格式无效，请检查配置内容。") from exc

    def do_GET(self):
        try:
            self._check_origin()
            path = urlsplit(self.path).path
            if path in ASSETS:
                filename, content_type = ASSETS[path]
                self._send((ROOT / "web" / filename).read_bytes(), content_type)
            elif path == "/api/examples":
                entries = []
                for key, label, filename in PRESETS:
                    data = generator.load_resume(ROOT / "examples" / filename)
                    entries.append({"id": key, "name": label, "resume": data, "yaml": _canonical(data)})
                self._json({"examples": entries})
            elif path == "/favicon.ico":
                self._send(b"", "image/x-icon", 204)
            else:
                raise ApiError("没有这个页面。", 404)
        except ApiError as exc:
            self._json({"error": str(exc)}, exc.status)
        except (OSError, generator.ResumeError) as exc:
            self._json({"error": f"无法加载本地文件：{exc}"}, 500)

    def do_POST(self):
        try:
            self._check_origin()
            route = urlsplit(self.path)
            if route.path not in {"/api/preview", "/api/export"}:
                raise ApiError("没有这个接口。", 404)
            data = _resume(self._body())
            if route.path == "/api/preview":
                self._json({"resume": data, "yaml": _canonical(data)})
                return
            query = parse_qs(route.query)
            formats = query.get("format", [])
            if len(formats) != 1 or formats[0] not in {"docx", "pdf"}:
                raise ApiError("导出格式只能是 docx 或 pdf。")
            format_name = formats[0]
            if not self.server.export_lock.acquire(blocking=False):
                raise ApiError("另一个导出正在进行，请稍后重试。", 409)
            try:
                with tempfile.TemporaryDirectory(prefix="fastresume-web-") as temporary:
                    result = generator.generate_resume(data, Path(temporary) / "resume", formats=(format_name,))
                    contents = Path(result[format_name]).read_bytes()
                mime = "application/pdf" if format_name == "pdf" else "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                self._send(contents, mime, filename=_filename(data, format_name))
            finally:
                self.server.export_lock.release()
        except ApiError as exc:
            self._json({"error": str(exc)}, exc.status)
        except (generator.ResumeError, OSError) as exc:
            self._json({"error": f"导出失败：{exc}"}, 500)

    def do_OPTIONS(self):
        self._json({"error": "不接受跨站请求。"}, 405)


def create_server(port=8765):
    return EditorServer(("127.0.0.1", port))


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--port", type=int, default=8765, help="本机端口，默认 8765")
    args = parser.parse_args(argv)
    if not 1 <= args.port <= 65535:
        parser.error("--port 必须在 1–65535 之间")
    try:
        server = create_server(args.port)
    except OSError as exc:
        parser.exit(2, f"无法启动编辑器：{exc}\n可使用 --port 8766 换一个端口。\n")
    print(f"fastResume 本地编辑器：http://127.0.0.1:{args.port}", flush=True)
    print("简历仅在本机处理。按 Ctrl+C 停止。", flush=True)
    try:
        server.serve_forever(poll_interval=0.25)
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
