"""작업 실행: 기준 + 대상 파일들 → 출력 폴더, 변경 요약, ZIP."""
import os
import shutil
import tempfile
import zipfile

from . import com_pptx
from .html_io import decode_bytes, html_to_model, parse_html, read_html_file
from .pptx_io import read_pptx
from .unify_html import build_html_from_html, build_html_from_model, template_from_pptx
from .unify_ppt import build_pptx

PPT_EXT = (".pptx", ".pptm")
HTML_EXT = (".html", ".htm")


def kind_of(path):
    e = os.path.splitext(path)[1].lower()
    if e in PPT_EXT:
        return "ppt"
    if e in HTML_EXT:
        return "html"
    return None


def _unique(path):
    if not os.path.exists(path):
        return path
    b, e = os.path.splitext(path)
    i = 2
    while os.path.exists("%s_%d%s" % (b, i, e)):
        i += 1
    return "%s_%d%s" % (b, i, e)


def run(to, base, targets, outdir, log=print):
    """to: 'ppt'|'html'. 반환: [{'file','out','ok','notes':[...]}]"""
    results = []
    os.makedirs(outdir, exist_ok=True)
    tmp = tempfile.mkdtemp(prefix="fu_")
    try:
        bk = kind_of(base)
        if bk is None:
            raise ValueError("기준 파일 형식을 알 수 없습니다: " + base)
        if to == "ppt" and bk != "ppt":
            raise ValueError("PPT로 출력하려면 기준 파일도 PPT여야 합니다. (HTML 기준 → PPT 출력은 아직 미지원)")
        # 기준 준비
        base_plain = None
        base_html_text = None
        if bk == "ppt":
            base_plain = com_pptx.ensure_plain(base, tmp)
            if to == "html":
                base_html_text = template_from_pptx(base_plain)
        else:
            with open(base, "rb") as f:
                base_html_text = decode_bytes(f.read())

        for t in targets:
            r = {"file": os.path.basename(t), "out": None, "ok": False, "notes": []}
            results.append(r)
            try:
                tk = kind_of(t)
                if tk is None:
                    raise ValueError("지원하지 않는 형식")
                stem = os.path.splitext(os.path.basename(t))[0]
                out = _unique(os.path.join(outdir, stem + "_unified." + ("pptx" if to == "ppt" else "html")))
                if to == "ppt":
                    title_slide = tk == "html"
                    if tk == "ppt":
                        src = com_pptx.ensure_plain(t, tmp)
                        doc = read_pptx(src)
                        r["notes"].append("PPT → PPT: 내용 이식")
                    else:
                        doc = html_to_model(read_html_file(t), os.path.basename(t), os.path.dirname(os.path.abspath(t)))
                        r["notes"].append("HTML → PPT: 제목(h1/h2) 단위로 슬라이드 분할")
                    r["notes"] += build_pptx(base_plain, doc, out, title_slide=title_slide)
                    r["notes"] += doc.warnings
                else:
                    if tk == "html":
                        with open(t, "rb") as f:
                            root = parse_html(decode_bytes(f.read()))
                        text, rep = build_html_from_html(base_html_text, root, os.path.basename(t))
                        r["notes"].append("HTML → HTML")
                    else:
                        src = com_pptx.ensure_plain(t, tmp)
                        doc = read_pptx(src)
                        text, rep = build_html_from_model(base_html_text, doc)
                        r["notes"].append("PPT → HTML")
                        rep += doc.warnings
                    with open(out, "w", encoding="utf-8") as f:
                        f.write(text)
                    r["notes"] += rep
                r["out"] = out
                r["ok"] = True
            except Exception as e:  # 파일 단위로 격리
                r["notes"].append("오류: %s" % e)
            log("%s %s" % ("OK " if r["ok"] else "ERR", r["file"]))
    finally:
        shutil.rmtree(tmp, ignore_errors=True)   # 복호화 임시 파일 삭제
    return results


def summary_text(to, base, results):
    lines = ["출력 형식: %s" % to.upper(), "기준 파일: %s" % os.path.basename(base), ""]
    for r in results:
        lines.append("[%s] %s → %s" % ("성공" if r["ok"] else "실패", r["file"], os.path.basename(r["out"]) if r["out"] else "-"))
        lines += ["    - " + n for n in r["notes"]]
    return "\n".join(lines) + "\n"


def make_zip(outdir, results, zip_path, summary):
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as z:
        for r in results:
            if r["out"]:
                z.write(r["out"], os.path.basename(r["out"]))
        z.writestr("변경요약.txt", summary)
    return zip_path
