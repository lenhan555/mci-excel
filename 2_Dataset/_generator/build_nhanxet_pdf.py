"""Dựng PDF nhận xét bài tập về nhà từ <Tên>_NhanXet.md và danh sách ảnh cần chụp.

Chạy (Windows, cần Edge): python 2_Dataset/_generator/build_nhanxet_pdf.py [B1 B2 ...]
Không đối số: dựng mọi buổi có thư mục Homework/Cham_bai/.

Quy ước trong .md:
    ![Hình 2 · mô tả](Images/NX02a_ten.png)
    > **[CHỤP] NX02a_ten.png** · File: ... · Sheet: ... · Vùng: ... · Cần thấy: ...
- Có ảnh Images/NX02a_ten.png (cạnh file .md): chèn ảnh, bỏ ghi chú [CHỤP].
- Chưa có ảnh: vẽ khung nét đứt kèm ghi chú để người chụp biết chụp gì.
- Các ảnh cùng số hình ("Hình 2") xếp cạnh nhau (so sánh bài làm với đáp án).
Ngoài PDF, ghi Cham_bai/_DANH_SACH_ANH_NX.md (danh sách ảnh cần chụp, nhóm theo học viên).
"""
import glob
import pathlib
import re
import subprocess
import sys

import markdown

ROOT = pathlib.Path(__file__).resolve().parents[2]
SESSIONS = ROOT / "1_Sessions"
EDGE = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"

CSS = """@page{size:A4;margin:16mm 14mm}
body{font-family:'Segoe UI',Arial,sans-serif;font-size:10.5pt;line-height:1.5;color:#222}
h1{font-size:17pt;border-bottom:2px solid #1F4E79;padding-bottom:4px;color:#1F4E79}
h2{font-size:13pt;color:#1F4E79;margin-top:18px;border-bottom:1px solid #C9D6E5;padding-bottom:2px}
h3{font-size:11.5pt;color:#1F4E79;margin-top:14px;margin-bottom:4px}
table{border-collapse:collapse;width:100%;margin:8px 0;page-break-inside:avoid}
th,td{border:1px solid #999;padding:4px 8px;text-align:left;vertical-align:top}
th{background:#DCE6F1}
code{background:#f1f1f1;padding:0 3px;border-radius:2px;font-family:Consolas,monospace;font-size:9.5pt}
li{margin-bottom:4px}
blockquote{border-left:3px solid #C9D6E5;margin:8px 0;padding:2px 10px;color:#444}
.pair{display:flex;gap:8px;align-items:flex-start;margin:10px 0;page-break-inside:avoid}
.pair>.fig{flex:1;min-width:0}
.fig{margin:10px 0;page-break-inside:avoid}
.fig img{max-width:100%;border:1px solid #BBB}
.fig .cap{font-size:9pt;color:#555;margin-top:2px}
.frame{border:2px dashed #C0504D;background:#FDF3F2;padding:8px 10px;font-size:9pt;color:#7A2E2B}
.frame .tag{font-weight:bold;color:#C0504D;letter-spacing:.5px}
"""

FIG = re.compile(r"^!\[(?P<alt>[^\]]*)\]\((?P<src>[^)]+)\)\s*$")
NOTE = re.compile(r"^>\s*\*\*\[CHỤP\]\s*(?P<name>[^*]+?)\*\*\s*(?P<rest>.*)$")


def inline(text):
    return markdown.markdown(text, extensions=["tables"]).removeprefix("<p>").removesuffix("</p>")


def figure_html(alt, src, note, md_dir):
    exists = (md_dir / src).exists()
    cap = f'<div class="cap">{inline(alt)}</div>'
    if exists:
        return f'<div class="fig"><img src="{src}">{cap}</div>', None
    detail = inline(note["rest"].lstrip("· ").strip()) if note else ""
    html = (f'<div class="fig"><div class="frame"><span class="tag">CHỜ ẢNH · {note["name"] if note else src}</span>'
            f'<br>{detail}</div>{cap}</div>')
    return html, note


def convert(md_path):
    """Trả về (html, [ghi chú khung còn thiếu])."""
    md_dir = md_path.parent
    lines = md_path.read_text(encoding="utf-8").splitlines()
    out, missing, i = [], [], 0
    pending = []  # các hình cùng số "Hình N" liền nhau

    def flush():
        if pending:
            out.append("")
            out.append('<div class="pair">' + "".join(pending) + "</div>" if len(pending) > 1 else pending[0])
            out.append("")
            pending.clear()

    cur_no = None
    while i < len(lines):
        line = lines[i]
        m = FIG.match(line)
        if m:
            no = re.match(r"Hình\s+(\d+)", m["alt"])
            no = no.group(1) if no else m["src"]
            if pending and no != cur_no:
                flush()
            cur_no = no
            note = None
            j = i + 1
            if j < len(lines) and (n := NOTE.match(lines[j])):
                note = n.groupdict()
                note["name"] = note["name"].strip()
                i = j
            html, miss = figure_html(m["alt"], m["src"], note, md_dir)
            if miss:
                missing.append(miss)
            pending.append(html)
            i += 1
            # bỏ dòng trống giữa các hình cùng số
            while i < len(lines) and not lines[i].strip() and i + 1 < len(lines) and FIG.match(lines[i + 1]):
                i += 1
            continue
        flush()
        out.append(line)
        i += 1
    flush()
    txt = "\n".join(out)
    # markdown cần dòng trống trước danh sách; file .md viết liền
    txt = re.sub(r"(?m)^([^\s\-\d|#<>].*)\n(?=(?:- |\d+\. ))", r"\1\n\n", txt)
    body = markdown.markdown(txt, extensions=["tables", "md_in_html", "nl2br"])
    return body, missing


def to_pdf(md_path, body):
    h = md_path.with_suffix(".html")
    h.write_text(f"<!doctype html><html><head><meta charset='utf-8'><style>{CSS}</style></head><body>{body}</body></html>",
                 encoding="utf-8")
    pdf = md_path.with_suffix(".pdf")
    subprocess.run([EDGE, "--headless", "--disable-gpu", "--no-pdf-header-footer",
                    f"--print-to-pdf={pdf}", h.as_uri()], check=True, capture_output=True)
    h.unlink()
    return pdf


def main(buoi):
    for session in sorted(SESSIONS.glob("B*")):
        if buoi and session.name.split("_")[0] not in buoi:
            continue
        cham = session / "Homework" / "Cham_bai"
        mds = sorted(cham.glob("*/*_NhanXet.md"))
        if not mds:
            continue
        rows, total_missing = [], 0
        for md in mds:
            body, missing = convert(md)
            pdf = to_pdf(md, body)
            total_missing += len(missing)
            print(f"{session.name.split('_')[0]} {pdf.name}: còn {len(missing)} khung chờ ảnh")
            for n in missing:
                rows.append((md.parent.name, n["name"], n["rest"].lstrip("· ").strip()))
        lst = cham / "_DANH_SACH_ANH_NX.md"
        buoi_id = session.name.split("_")[0]
        head = [f"# Danh sách ảnh cần chụp — nhận xét bài tập về nhà {buoi_id}", "",
                "Mỗi ảnh lưu **PNG**, đúng tên, vào `Cham_bai/<Học viên>/Images/`. Chụp Excel 365 Windows, zoom 100%.",
                "Ảnh **\"bài của bạn\"** chụp từ file trong `Homework/Bai_nop/`; ảnh **\"đáp án\"** chụp từ file `_Answer` (hoặc `Practice/`, `Demo/` khi ghi rõ).",
                "Ô cần nhấn mạnh: tô viền đỏ trên **bản sao** của file, không sửa file gốc của học viên.",
                "Chạy lại `python 2_Dataset/_generator/build_nhanxet_pdf.py " + buoi_id + "` để ảnh tự vào PDF.", "",
                f"Còn thiếu: **{total_missing}** ảnh.", "",
                "| Học viên | Tên file | Chụp gì (file · sheet · vùng · cần thấy) |", "|---|---|---|"]
        body_rows = [f"| {s} | `{n}` | {r.replace('|', '/')} |" for s, n, r in rows]
        lst.write_text("\n".join(head + body_rows) + "\n", encoding="utf-8")
        print(f"→ {lst.relative_to(ROOT)} ({total_missing} ảnh)")


if __name__ == "__main__":
    main(set(sys.argv[1:]))
