"""pptx(zip+XML) 읽기. 표준 라이브러리만 사용."""
import posixpath
import re
import zipfile
import xml.etree.ElementTree as ET

from .model import Block, Document, Section

NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}
A = "{%s}" % NS["a"]
P = "{%s}" % NS["p"]
R = "{%s}" % NS["r"]
RELNS = "{http://schemas.openxmlformats.org/package/2006/relationships}"


def read_rels(z, part):
    """part 의 rels -> {rId: (type, absolute target path, mode)}"""
    d, n = posixpath.split(part)
    relpath = posixpath.join(d, "_rels", n + ".rels")
    out = {}
    if relpath not in z.namelist():
        return out
    root = ET.fromstring(z.read(relpath))
    for rel in root:
        tgt = rel.get("Target", "")
        if rel.get("TargetMode") == "External":
            out[rel.get("Id")] = (rel.get("Type"), tgt, "External")
        else:
            if tgt.startswith("/"):
                full = tgt.lstrip("/")
            else:
                full = posixpath.normpath(posixpath.join(d, tgt))
            out[rel.get("Id")] = (rel.get("Type"), full, "Internal")
    return out


def slide_paths(z):
    pres = ET.fromstring(z.read("ppt/presentation.xml"))
    rels = read_rels(z, "ppt/presentation.xml")
    out = []
    for sid in pres.iter(P + "sldId"):
        rid = sid.get(R + "id")
        if rid in rels:
            out.append(rels[rid][1])
    return out


def _para_text(p):
    parts = []
    for el in p:
        if el.tag in (A + "r", A + "fld"):
            t = el.find(A + "t")
            if t is not None and t.text:
                parts.append(t.text)
        elif el.tag == A + "br":
            parts.append(" ")
    return "".join(parts).strip()


def _ph(sp):
    nv = sp.find(P + "nvSpPr/" + P + "nvPr/" + P + "ph")
    return nv


def _ext_of(path):
    e = posixpath.splitext(path)[1].lower().lstrip(".")
    return "jpeg" if e == "jpg" else e


def _walk(el):
    """spTree 하위를 문서 순서로 순회 (그룹은 재귀)."""
    for ch in el:
        if ch.tag == P + "grpSp":
            yield from _walk(ch)
        elif ch.tag in (P + "sp", P + "pic", P + "graphicFrame"):
            yield ch


def read_pptx(path):
    doc = Document(title=posixpath.basename(path.replace("\\", "/")))
    with zipfile.ZipFile(path) as z:
        for i, sp_path in enumerate(slide_paths(z), 1):
            root = ET.fromstring(z.read(sp_path))
            rels = read_rels(z, sp_path)
            tree = root.find(P + "cSld/" + P + "spTree")
            sec = Section()
            cur_list = None
            if tree is None:
                continue
            for sh in _walk(tree):
                if sh.tag == P + "sp":
                    ph = _ph(sh)
                    ptype = ph.get("type") if ph is not None else None
                    if ptype in ("sldNum", "dt", "ftr", "hdr"):
                        continue
                    tx = sh.find(P + "txBody")
                    if tx is None:
                        continue
                    paras = [(pp, _para_text(pp)) for pp in tx.findall(A + "p")]
                    paras = [(pp, t) for pp, t in paras if t]
                    if not paras:
                        continue
                    if ptype in ("title", "ctrTitle"):
                        title = " ".join(t for _, t in paras)
                        sec.title = (sec.title + " " + title).strip() if sec.title else title
                        continue
                    in_body_ph = ph is not None and ptype in (None, "body", "obj")
                    for pp, t in paras:
                        ppr = pp.find(A + "pPr")
                        lvl = int(ppr.get("lvl", "0")) if ppr is not None else 0
                        has_bu = ppr is not None and (
                            ppr.find(A + "buChar") is not None or ppr.find(A + "buAutoNum") is not None)
                        no_bu = ppr is not None and ppr.find(A + "buNone") is not None
                        is_list = (has_bu or lvl > 0 or in_body_ph) and not no_bu
                        if is_list:
                            ordered = ppr is not None and ppr.find(A + "buAutoNum") is not None
                            if cur_list is None or cur_list.ordered != ordered:
                                cur_list = Block("list", ordered=ordered)
                                sec.blocks.append(cur_list)
                            cur_list.items.append((lvl, t))
                        else:
                            cur_list = None
                            sec.blocks.append(Block("para", text=t))
                    cur_list = None
                elif sh.tag == P + "graphicFrame":
                    tbl = sh.find(".//" + A + "tbl")
                    if tbl is None:
                        doc.warnings.append("슬라이드 %d: 표가 아닌 개체(차트/SmartArt 등)는 건너뜀" % i)
                        continue
                    rows = []
                    for tr in tbl.findall(A + "tr"):
                        row = []
                        for tc in tr.findall(A + "tc"):
                            row.append(" ".join(_para_text(pp) for pp in tc.iter(A + "p")).strip())
                        rows.append(row)
                    if rows:
                        sec.blocks.append(Block("table", rows=rows, header=True))
                    cur_list = None
                elif sh.tag == P + "pic":
                    blip = sh.find(".//" + A + "blip")
                    rid = blip.get(R + "embed") if blip is not None else None
                    if rid and rid in rels and rels[rid][2] == "Internal" and rels[rid][1] in z.namelist():
                        t = rels[rid][1]
                        ext = _ext_of(t)
                        if ext in ("png", "jpeg", "gif"):
                            cnv = sh.find(P + "nvPicPr/" + P + "cNvPr")
                            alt = cnv.get("descr", "") if cnv is not None else ""
                            sec.blocks.append(Block("image", data=z.read(t), ext=ext, alt=alt))
                        else:
                            doc.warnings.append("슬라이드 %d: 지원하지 않는 이미지 형식(%s) 건너뜀" % (i, ext))
                    cur_list = None
            if not sec.title:
                sec.title = "슬라이드 %d" % i
            doc.sections.append(sec)
    return doc


def read_theme(path):
    """기준 pptx 에서 슬라이드 크기·테마 폰트/색 추출 (HTML CSS 생성용)."""
    info = {"cx": 9144000, "cy": 6858000, "major": "", "minor": "", "colors": {}}
    with zipfile.ZipFile(path) as z:
        pres = z.read("ppt/presentation.xml").decode("utf-8", "replace")
        m = re.search(r'<p:sldSz[^>]*cx="(\d+)"[^>]*cy="(\d+)"', pres)
        if m:
            info["cx"], info["cy"] = int(m.group(1)), int(m.group(2))
        rels = read_rels(z, "ppt/presentation.xml")
        theme = next((v[1] for v in rels.values() if v[0].endswith("/theme")), None)
        if theme and theme in z.namelist():
            root = ET.fromstring(z.read(theme))
            for key, tag in (("major", "majorFont"), ("minor", "minorFont")):
                f = root.find(".//" + A + tag)
                if f is not None:
                    ea = f.find(A + "ea")
                    lat = f.find(A + "latin")
                    names = [x.get("typeface") for x in (lat, ea) if x is not None and x.get("typeface")]
                    info[key] = ", ".join(names)
            cs = root.find(".//" + A + "clrScheme")
            if cs is not None:
                for c in cs:
                    name = c.tag.replace(A, "")
                    for ch in c:
                        val = ch.get("val") or ch.get("lastClr")
                        if val:
                            info["colors"][name] = val
    return info
