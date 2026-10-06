"""Render infographic tóm tắt buổi học (HTML → PNG) bằng Chrome headless.

    python build_infographic.py B1           # → 1_Sessions/B1_…/Handout/B1_Infographic_TomTat.png
                                             # (Mac: python3). Tìm Chrome/Edge tự động; đặt biến CHROME nếu cài chỗ khác.

Nguồn: _generator/infographic_b<n>.html (khung A4 dọc 1240 × 1754 px). Ảnh xuất ở 2x = 2480 × 3508 px (≈ 300 dpi khi in A4).
Muốn sửa nội dung thì sửa file HTML rồi chạy lại, không sửa ảnh PNG.

Số liệu: nếu có _generator/build_b<n>_meta.json thì HTML viết {{khóa.con|kiểu}} thay cho con số, script điền từ meta.
Kiểu: n (941.454.000, làm tròn), pct (34,2%), pct0 (34%), d2pct (0,96%), up (tỷ lệ 1,251 → 25,1%), d1/d2/d3 (số thập phân dấu phẩy), không ghi = giữ nguyên. Khóa thiếu trong meta → dừng, báo lỗi.
"""
import glob, json, os, pathlib, re, subprocess, sys, tempfile, time

ROOT = os.path.normpath(os.environ.get("MCI_ROOT") or os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))
GEN = os.path.dirname(os.path.abspath(__file__))
CHROME = next((p for p in [
    os.environ.get("CHROME"),
    "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
    os.path.join(os.environ.get("ProgramFiles", ""), r"Google\Chrome\Application\chrome.exe"),
    os.path.join(os.environ.get("ProgramFiles(x86)", ""), r"Google\Chrome\Application\chrome.exe"),
    os.path.join(os.environ.get("LOCALAPPDATA", ""), r"Google\Chrome\Application\chrome.exe"),
    os.path.join(os.environ.get("ProgramFiles(x86)", ""), r"Microsoft\Edge\Application\msedge.exe"),
] if p and os.path.exists(p)), None)
W, H, SCALE = 1240, 1754, 2


def fill(html, meta):
    def get(path):
        v = meta
        for k in path.split("."):
            v = v[k] if isinstance(v, dict) else v[int(k)]
        return v

    def sub(m):
        path, kind = m.group(1), m.group(2)
        try:
            v = get(path)
        except (KeyError, IndexError, ValueError):
            sys.exit(f"Meta không có khóa {path}")
        if kind == "n":
            return f"{round(v):,}".replace(",", ".")
        if kind == "pct":
            return f"{v * 100:.1f}%".replace(".", ",")
        if kind == "pct0":
            return f"{v * 100:.0f}%"
        if kind == "d2pct":
            return f"{v * 100:.2f}%".replace(".", ",")
        if kind == "up":                      # tỷ lệ 1,251 → "25,1%" (tăng so với mốc)
            return f"{(v - 1) * 100:.1f}%".replace(".", ",")
        if kind in ("d1", "d2", "d3"):
            return f"{v:.{kind[1]}f}".replace(".", ",")
        return str(v)
    return re.sub(r"\{\{\s*([\w. ]+?)\s*(?:\|\s*(\w+))?\s*\}\}", sub, html)


def build(session):
    n = session.upper().lstrip("B")
    src = os.path.join(GEN, f"infographic_b{n}.html")
    folders = glob.glob(os.path.join(ROOT, "1_Sessions", f"B{n}_*"))
    if not os.path.exists(src) or not folders:
        sys.exit(f"Không tìm thấy {src} hoặc thư mục buổi B{n}")
    out = os.path.join(folders[0], "Handout", f"B{n}_Infographic_TomTat.png")
    if os.path.exists(out):
        os.remove(out)
    meta_path = os.path.join(GEN, f"build_b{n}_meta.json")
    html = open(src, encoding="utf-8").read()
    tmp_src = None
    if not CHROME:
        sys.exit("Không tìm thấy Chrome hoặc Edge. Đặt biến môi trường CHROME = đường dẫn chrome.exe")
    if "{{" in html:
        if not os.path.exists(meta_path):
            sys.exit(f"{src} có {{{{…}}}} nhưng thiếu {meta_path}")
        tmp_src = os.path.join(GEN, f".infographic_b{n}_filled.html")
        open(tmp_src, "w", encoding="utf-8").write(fill(html, json.load(open(meta_path, encoding="utf-8"))))
        src = tmp_src
    # Chrome headless ghi xong ảnh nhưng có bản không tự thoát: chờ file ổn định rồi tắt tiến trình.
    with tempfile.TemporaryDirectory() as profile:
        p = subprocess.Popen([CHROME, "--headless=new", "--disable-gpu", "--hide-scrollbars", f"--user-data-dir={profile}",
                              f"--force-device-scale-factor={SCALE}", f"--window-size={W},{H}",
                              f"--screenshot={out}", pathlib.Path(src).resolve().as_uri()],
                             stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        size, t0 = -1, time.time()
        while time.time() - t0 < 60 and p.poll() is None:
            time.sleep(1)
            if os.path.exists(out) and os.path.getsize(out) == size > 0:
                break
            size = os.path.getsize(out) if os.path.exists(out) else -1
        if p.poll() is None:
            p.terminate()
            p.wait(10)
    if tmp_src:
        os.remove(tmp_src)
    if not os.path.exists(out):
        sys.exit("Chrome không xuất được ảnh")
    print("Đã dựng", os.path.relpath(out, ROOT))


if __name__ == "__main__":
    for s in sys.argv[1:] or ["B1"]:
        build(s)
