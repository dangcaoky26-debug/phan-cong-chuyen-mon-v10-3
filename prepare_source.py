from pathlib import Path
import base64
import hashlib
import zipfile

EXPECTED_SIZE = 133928
EXPECTED_SHA256 = "621e8978f0bfbde952479853e079273f6f210482e473f91da3a45a135876ccae"

parts = sorted(Path("source_parts").glob("part*.b64"))
if len(parts) != 14:
    raise SystemExit(f"Need 14 parts, got {len(parts)}")

payload = b"".join(base64.b64decode(p.read_text(encoding="ascii").strip()) for p in parts)
if len(payload) != EXPECTED_SIZE:
    raise SystemExit(f"Size mismatch: {len(payload)} != {EXPECTED_SIZE}")
sha = hashlib.sha256(payload).hexdigest()
if sha != EXPECTED_SHA256:
    raise SystemExit(f"SHA256 mismatch: {sha}")

zip_path = Path("_source.zip")
zip_path.write_bytes(payload)
with zipfile.ZipFile(zip_path) as zf:
    zf.extractall(".")
zip_path.unlink()

# Chuyển server desktop sang dạng phù hợp với hosting web.
p = Path("phan_cong_v10_3_dragdrop.py")
s = p.read_text(encoding="utf-8")
s = s.replace(
    'HOST, PORT = "127.0.0.1", 8812',
    'HOST = os.environ.get("HOST", "0.0.0.0")\nPORT = int(os.environ.get("PORT", "8812"))'
)
s = s.replace(
    '    threading.Timer(0.8,lambda:webbrowser.open(url)).start()',
    '    if HOST in ("127.0.0.1", "localhost"):\n        threading.Timer(0.8, lambda: webbrowser.open(url)).start()'
)
p.write_text(s, encoding="utf-8")

Path("requirements.txt").write_text("PyMuPDF>=1.24,<2\n", encoding="utf-8")
Path("render.yaml").write_text(
    """services:\n  - type: web\n    name: phan-cong-chuyen-mon-v10-3\n    runtime: python\n    buildCommand: pip install -r requirements.txt\n    startCommand: python phan_cong_v10_3_dragdrop.py\n""",
    encoding="utf-8",
)
print(f"Source verified: {len(payload)} bytes, sha256={sha}")
