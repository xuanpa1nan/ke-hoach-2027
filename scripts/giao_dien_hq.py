"""Chạy công cụ quản trị kênh trên trình duyệt, chỉ lắng nghe tại localhost."""

import argparse
import hmac
import json
import os
import re
import secrets
import subprocess
import sys
import uuid
from datetime import date
from email import policy
from email.parser import BytesParser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlsplit


ROOT = Path(__file__).resolve().parent.parent
RUNS = ROOT / "local" / "giao-dien-hq"
MAX_UPLOAD = 100 * 1024 * 1024
FIELDS = {"ims", "bris", "smcs", "reference", "template", "td49", "sim_tsd"}
OUTPUTS = {
    "hq": ("HQ và BSC", ".xlsx"),
    "excel": ("BC SXKD", ".xlsx"),
    "word": ("BC Điều hành SXKD", ".docx"),
}


def invoke(command):
    result = subprocess.run(
        command, cwd=ROOT, capture_output=True, encoding="utf-8", errors="replace",
        timeout=180, env=dict(os.environ, PYTHONUTF8="1", PYTHONDONTWRITEBYTECODE="1"),
    )
    if result.returncode:
        detail = (result.stderr or result.stdout).strip().splitlines()
        raise ValueError(detail[-1] if detail else "Chương trình xử lý trả về lỗi.")


class App(BaseHTTPRequestHandler):
    def same_host(self):
        return self.headers.get("Host", "") == f"127.0.0.1:{self.server.server_port}"

    def reply(self, status, content, mime, filename=None):
        self.send_response(status)
        self.send_header("Content-Type", mime)
        self.send_header("Content-Length", str(len(content)))
        self.send_header("Cache-Control", "no-store")
        self.send_header("X-Content-Type-Options", "nosniff")
        self.send_header("Content-Security-Policy",
                         "default-src 'none'; style-src 'unsafe-inline'; script-src 'unsafe-inline'; "
                         "connect-src 'self'; frame-ancestors 'none'; base-uri 'none'")
        if filename:
            from urllib.parse import quote
            self.send_header("Content-Disposition", f"attachment; filename*=UTF-8''{quote(filename)}")
        self.end_headers()
        self.wfile.write(content)

    def json(self, status, body):
        self.reply(status, json.dumps(body, ensure_ascii=False).encode(), "application/json; charset=utf-8")

    def do_GET(self):
        if not self.same_host():
            return self.json(403, {"error": "Chỉ dùng địa chỉ localhost của tool."})
        path = urlsplit(self.path).path
        if path == "/":
            html = (ROOT / "app" / "index.html").read_text(encoding="utf-8")
            html = html.replace("__LOCAL_TOKEN__", self.server.token)
            html = html.replace("__DEMO_AVAILABLE__", str(bool(self.server.options.demo_dir)).lower())
            return self.reply(200, html.encode(), "text/html; charset=utf-8")
        match = re.fullmatch(r"/download/([0-9a-f]{32})/(hq|excel|word)", path)
        if match:
            run = RUNS / match[1]
            if (run / "exported").is_file():
                data = json.loads((run / "report.json").read_text(encoding="utf-8"))
                name, suffix = OUTPUTS[match[2]]
                file = run / "bao-cao" / f"{name} - Tuần {data['week']}{suffix}"
                if file.is_file():
                    return self.reply(200, file.read_bytes(), "application/octet-stream", file.name)
        self.json(404, {"error": "Không tìm thấy tệp."})

    def do_POST(self):
        if not self.same_host() or not hmac.compare_digest(
            self.headers.get("X-Local-Token", ""), self.server.token
        ):
            return self.json(403, {"error": "Hãy mở lại giao diện từ địa chỉ của tool."})
        try:
            length = int(self.headers.get("Content-Length", "0"))
            if not 0 <= length <= MAX_UPLOAD:
                return self.json(413, {"error": "Tổng dung lượng tải lên tối đa là 100 MiB."})
            path = urlsplit(self.path).path
            if path == "/run":
                return self.json(200, self.run_report(length))
            if path == "/export":
                body = json.loads(self.rfile.read(length))
                return self.json(200, self.export_report(body.get("run", "")))
            self.json(404, {"error": "Không tìm thấy chức năng."})
        except (ValueError, OSError, KeyError, subprocess.TimeoutExpired) as error:
            self.json(400, {"error": str(error)})

    def run_report(self, length):
        content_type = self.headers.get("Content-Type", "")
        if not content_type.startswith("multipart/form-data;"):
            raise ValueError("Định dạng tệp gửi lên không hợp lệ.")
        message = BytesParser(policy=policy.default).parsebytes(
            f"Content-Type: {content_type}\r\n\r\n".encode() + self.rfile.read(length)
        )
        fields, files = {}, {}
        for part in message.iter_parts():
            name = part.get_param("name", header="content-disposition")
            payload = part.get_payload(decode=True)
            if part.get_filename() and name in FIELDS:
                suffix = Path(part.get_filename()).suffix.lower()
                if suffix not in (".xls", ".xlsx"):
                    raise ValueError(f"Tệp {name} phải là Excel .xls hoặc .xlsx.")
                files[name] = (suffix, payload)
            elif name in {"week", "bris_from", "bris_to", "demo"}:
                fields[name] = payload.decode("utf-8")
        week = int(fields.get("week", "0"))
        start = date.fromisoformat(fields.get("bris_from", ""))
        end = date.fromisoformat(fields.get("bris_to", ""))
        if not 1 <= week <= 53 or start > end:
            raise ValueError("Kiểm tra lại tuần báo cáo và khoảng ngày kích hoạt BRIS.")
        demo = fields.get("demo") == "true"
        options = self.server.options
        if demo and (not options.demo_dir or week != 46):
            raise ValueError("Bộ dữ liệu thử chỉ dành cho tuần 46.")
        run = RUNS / uuid.uuid4().hex
        upload = run / "nguon"
        upload.mkdir(parents=True)
        paths = {}
        for name, (suffix, payload) in files.items():
            paths[name] = upload / (name + suffix)
            paths[name].write_bytes(payload)
        if demo:
            paths.setdefault("ims", options.demo_dir / "RPVNP0173.xls")
            bris = sorted(path for path in options.demo_dir.rglob("hqkd_kenh_ct_ddtt*.xlsx") if path.is_file())
            if bris:
                paths.setdefault("bris", bris[0])
            paths.setdefault("reference", options.demo_reference)
            paths.setdefault("template", options.demo_template)
            paths.setdefault("td49", options.demo_reference)
            paths.setdefault("sim_tsd", options.demo_reference)
        required = ("ims", "bris", "reference", "template") if demo else (
            "ims", "bris", "smcs", "reference", "template"
        )
        for name in required:
            if not paths.get(name) or not paths[name].is_file():
                raise ValueError(f"Chưa có tệp đầu vào: {name}.")
        python = [sys.executable, "-B"]
        args = [
            *python, "scripts/tinh_hq.py", "--ims", str(paths["ims"]), "--bris", str(paths["bris"]),
            "--bris-from", str(start), "--bris-to", str(end), "--output", str(run / "hq"),
        ]
        if "smcs" in paths:
            args += ["--smcs", str(paths["smcs"])]
        else:
            args += ["--week46-catalog", str(paths["reference"])]
        for name in ("td49", "sim_tsd"):
            if paths.get(name):
                args += ["--" + name.replace("_", "-"), str(paths[name])]
        if demo:
            args += ["--compare", str(paths["reference"])]
        invoke(args)
        invoke([
            *python, "scripts/du_lieu_bao_cao_tuan.py", "--hq-dir", str(run / "hq"),
            "--reference", str(paths["reference"]), "--ims", str(paths["ims"]),
            "--excel-template", str(paths["template"]), "--week", str(week),
            "--output", str(run / "report.json"),
        ])
        meta = json.loads((run / "hq" / "run.json").read_text(encoding="utf-8"))
        data = json.loads((run / "report.json").read_text(encoding="utf-8"))
        if meta["catalog"] == "smcs":
            data["notes"][0] = ["SMCS", "Dùng tệp danh mục SMCS được chọn trên giao diện."]
        data["demo"] = demo
        if not demo:
            sample_only = {"Đối soát TD49", "Chiến dịch 120 ngày", "Tái sử dụng VinaKey1"}
            data["notes"] = [
                [name, description.replace("tuần 46", "tham chiếu đúng kỳ")]
                for name, description in data["notes"] if name not in sample_only
            ]
        for name, channel in (("td49", "VinaPro"), ("sim_tsd", "VinaKey1")):
            if not paths.get(name):
                data["notes"].append([f"Tái sử dụng {channel}", "Thiếu tệp bổ sung. Doanh thu đánh giá hiện chỉ tính PTM cho kênh này."])
        (run / "report.json").write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
        return {"run": run.name, "week": week, "hq": data["hq"], "bsc": data["bsc"], "notes": data["notes"]}

    def export_report(self, run_id):
        if not re.fullmatch(r"[0-9a-f]{32}", run_id):
            raise ValueError("Mã lần chạy không hợp lệ.")
        run = RUNS / run_id
        report = run / "report.json"
        if not report.is_file():
            raise ValueError("Hãy chạy tính HQ/BSC trước.")
        if not (run / "exported").exists():
            invoke([self.server.options.node, "scripts/xuat_bao_cao_tuan.mjs",
                    str(report), str(run / "bao-cao")])
            (run / "exported").touch()
        return {"files": {name: f"/download/{run_id}/{name}" for name in OUTPUTS},
                "folder": str(run / "bao-cao")}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--port", type=int, default=8765)
    parser.add_argument("--node", default="node")
    parser.add_argument("--demo-dir", type=Path)
    parser.add_argument("--demo-reference", type=Path)
    parser.add_argument("--demo-template", type=Path)
    options = parser.parse_args()
    if options.demo_dir and (not options.demo_reference or not options.demo_template):
        parser.error("Bộ thử cần --demo-reference và --demo-template.")
    with ThreadingHTTPServer(("127.0.0.1", options.port), App) as server:
        server.options = options
        server.token = secrets.token_hex(32)
        print(f"http://127.0.0.1:{server.server_port}", flush=True)
        server.serve_forever()


if __name__ == "__main__":
    main()
