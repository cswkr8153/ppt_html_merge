"""기준 pptx 의 마스터/레이아웃/테마를 유지하고 Document 내용만 새 슬라이드로 이식."""
import posixpath
import re
import struct
import zipfile
import xml.etree.ElementTree as ET
from xml.sax.saxutils import escape

from .model import Document
from .pptx_io import A, P, read_rels

CT_SLIDE = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"
REL_SLIDE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"
REL_LAYOUT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"
REL_IMAGE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
IMG_CT = {"png": "image/png", "jpeg": "image/jpeg", "gif": "image/gif"}
BAD = re.compile("[\x00-\x08\x0b\x0c\x0e-\x1f]")
MAX_ITEMS = 10      # 슬라이드당 본문 항목 상한
MAX_CHARS = 900     # 슬라이드당 본문 글자 상한


def esc(s):
    return escape(BAD.sub("", s))


def image_size(data):
    try:
        if data[:8] == b"\x89PNG\r\n\x1a\n":
            return struct.unpack(">II", data[16:24])
        if data[:6] in (b"GIF87a", b"GIF89a"):
            return struct.unpack("<HH", data[6:10])
        if data[:2] == b"\xff\xd8":
            i = 2
            while i < len(data):
                if data[i] != 0xFF:
                    i += 1
                    continue
                m = data[i + 1]
                if m in (0xC0, 0xC1, 0xC2):
                    h, w = struct.unpack(">HH", data[i + 5:i + 9])
                    return w, h
                i += 2 + struct.unpack(">H", data[i + 2:i + 4])[0]
    except Exception:
        pass
    return None


class Layout:
    def __init__(self, path, xml):
        self.path = path
        self.xml = xml
        root = ET.fromstring(xml)
        self.type = root.get("type") or ""
        self.title = None   # (type, idx, box|None)
        self.body = None
        self.sub = None
        for sp in root.iter(P + "sp"):
            ph = sp.find(P + "nvSpPr/" + P + "nvPr/" + P + "ph")
            if ph is None:
                continue
            t, idx = ph.get("type"), ph.get("idx")
            box = self._box(sp)
            if t in ("title", "ctrTitle"):
                self.title = (t, idx, box)
            elif t == "subTitle":
                self.sub = (t, idx, box)
            elif t in (None, "body", "obj") and self.body is None:
                self.body = (t, idx, box)

    @staticmethod
    def _box(sp):
        off = sp.find(P + "spPr/" + A + "xfrm/" + A + "off")
        ext = sp.find(P + "spPr/" + A + "xfrm/" + A + "ext")
        if off is None or ext is None:
            return None
        return (int(off.get("x")), int(off.get("y")), int(ext.get("cx")), int(ext.get("cy")))


def pick_layouts(z, master_body_box):
    layouts = []
    for n in sorted(z.namelist()):
        if re.match(r"ppt/slideLayouts/slideLayout\d+\.xml$", n):
            layouts.append(Layout(n, z.read(n)))
    content = None
    for want in ("obj", "tx"):
        content = next((l for l in layouts if l.type == want and l.title and l.body), None)
        if content:
            break
    if not content:
        content = next((l for l in layouts if l.title and l.body), None)
    title = next((l for l in layouts if l.type == "title" and l.title), None)
    if not content and layouts:
        content = layouts[0]
    # 본문 박스 보정: 레이아웃에 없으면 마스터 값
    return content, title, layouts


def master_boxes(z):
    boxes = {}
    for n in z.namelist():
        if re.match(r"ppt/slideMasters/slideMaster\d+\.xml$", n):
            root = ET.fromstring(z.read(n))
            for sp in root.iter(P + "sp"):
                ph = sp.find(P + "nvSpPr/" + P + "nvPr/" + P + "ph")
                if ph is None:
                    continue
                t = ph.get("type") or "body"
                b = Layout._box(sp)
                if b:
                    boxes.setdefault("title" if t in ("title", "ctrTitle") else t, b)
            break
    return boxes


def txbox(idn, name, box, paras_xml, ph=None):
    nv = '<p:nvPr>%s</p:nvPr>' % (ph or "")
    xfrm = ""
    if box:
        xfrm = '<a:xfrm><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></a:xfrm>' % box
    sppr = "<p:spPr>%s</p:spPr>" % xfrm if ph else '<p:spPr>%s<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>' % xfrm
    c = '<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>' if ph else '<p:cNvSpPr txBox="1"/>'
    return ('<p:sp><p:nvSpPr><p:cNvPr id="%d" name="%s"/>%s%s</p:nvSpPr>%s'
            '<p:txBody><a:bodyPr><a:normAutofit/></a:bodyPr><a:lstStyle/>%s</p:txBody></p:sp>'
            % (idn, esc(name), c, nv, sppr, paras_xml))


def p_xml(text, lvl=0, plain=False, ordered=False):
    ppr = ""
    if plain:
        ppr = '<a:pPr marL="0" indent="0"><a:buNone/></a:pPr>'
    elif ordered:
        ppr = '<a:pPr lvl="%d"><a:buFont typeface="+mj-lt"/><a:buAutoNum type="arabicPeriod"/></a:pPr>' % lvl
    elif lvl:
        ppr = '<a:pPr lvl="%d"/>' % lvl
    return "<a:p>%s<a:r><a:rPr lang=\"ko-KR\" dirty=\"0\"/><a:t>%s</a:t></a:r></a:p>" % (ppr, esc(text))


def table_xml(idn, box, rows, style_id):
    ncol = max(len(r) for r in rows)
    x, y, cx, cy = box
    colw = cx // ncol
    rowh = min(max(cy // max(len(rows), 1), 300000), 500000)
    sty = "<a:tableStyleId>%s</a:tableStyleId>" % style_id if style_id else ""
    out = ['<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="%d" name="Table %d"/>'
           '<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr>'
           '<p:xfrm><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></p:xfrm>'
           '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl>'
           '<a:tblPr firstRow="1" bandRow="1">%s</a:tblPr><a:tblGrid>'
           % (idn, idn, x, y, colw * ncol, rowh * len(rows), sty)]
    out.append("".join('<a:gridCol w="%d"/>' % colw for _ in range(ncol)))
    out.append("</a:tblGrid>")
    for r in rows:
        out.append('<a:tr h="%d">' % rowh)
        for ci in range(ncol):
            t = r[ci] if ci < len(r) else ""
            out.append('<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="ko-KR" sz="1400" dirty="0"/>'
                       '<a:t>%s</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>' % esc(t))
        out.append("</a:tr>")
    out.append("</a:tbl></a:graphicData></a:graphic></p:graphicFrame>")
    return "".join(out)


def pic_xml(idn, rid, box, data, alt):
    x, y, cx, cy = box
    sz = image_size(data)
    if sz and sz[0] and sz[1]:
        w, h = sz
        scale = min(cx / w, cy / h)
        pw, ph_ = int(w * scale), int(h * scale)
        x += (cx - pw) // 2
        y += (cy - ph_) // 2
        cx, cy = pw, ph_
    return ('<p:pic><p:nvPicPr><p:cNvPr id="%d" name="Picture %d" descr="%s"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>'
            '<p:blipFill><a:blip r:embed="%s"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>'
            '<p:spPr><a:xfrm><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>'
            % (idn, idn, esc(alt), rid, x, y, cx, cy))


def slide_xml(shapes):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
            'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
            '<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
            '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
            + "".join(shapes) +
            '</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>')


def ph_xml(info):
    t, idx, _ = info
    s = "<p:ph"
    if t:
        s += ' type="%s"' % t
    if idx:
        s += ' idx="%s"' % idx
    return s + "/>"


def split_text_blocks(blocks):
    """본문(para/list) 묶음을 슬라이드 단위 조각으로 분할. 표/이미지는 단독 조각."""
    pages, cur, n, chars = [], [], 0, 0

    def flush():
        nonlocal cur, n, chars
        if cur:
            pages.append(("text", cur))
        cur, n, chars = [], 0, 0

    for b in blocks:
        if b.kind in ("table", "image"):
            flush()
            pages.append((b.kind, b))
        elif b.kind == "para":
            if n + 1 > MAX_ITEMS or chars + len(b.text) > MAX_CHARS:
                flush()
            cur.append(("p", 0, b.text, False)); n += 1; chars += len(b.text)
        else:
            for lvl, t in b.items:
                if n + 1 > MAX_ITEMS or chars + len(t) > MAX_CHARS:
                    flush()
                cur.append(("l", lvl, t, b.ordered)); n += 1; chars += len(t)
    flush()
    return pages


def build_pptx(base_path, doc, out_path, title_slide=True):
    rep = []
    with zipfile.ZipFile(base_path) as zin:
        names = zin.namelist()
        pres = zin.read("ppt/presentation.xml").decode("utf-8")
        m = re.search(r'<p:sldSz[^>]*cx="(\d+)"[^>]*cy="(\d+)"', pres)
        sw, sh = (int(m.group(1)), int(m.group(2))) if m else (9144000, 6858000)
        mbox = master_boxes(zin)
        content, title_l, _ = pick_layouts(zin, mbox)
        if not content:
            raise RuntimeError("기준 파일에 슬라이드 레이아웃이 없습니다.")
        rep.append("슬라이드 크기 %.1f x %.1f cm 유지" % (sw / 360000, sh / 360000))
        rep.append("본문 레이아웃: %s" % posixpath.basename(content.path))

        def box_of(info, key):
            if info and info[2]:
                return info[2]
            return mbox.get(key)
        mg = int(sw * 0.06)
        t_box = box_of(content.title, "title") or (mg, int(sh * 0.05), sw - 2 * mg, int(sh * 0.15))
        b_box = box_of(content.body, "body") or (mg, int(sh * 0.24), sw - 2 * mg, int(sh * 0.66))
        style_id = None
        if "ppt/tableStyles.xml" in names:
            sm = re.search(r'def="(\{[^"]+\})"', zin.read("ppt/tableStyles.xml").decode("utf-8", "replace"))
            style_id = sm.group(1) if sm else None

        slides = []   # (xml, layout_path, [(rid, media_name)])
        media = {}    # name -> bytes

        def new_media(data, ext):
            name = "gen%d.%s" % (len(media) + 1, "jpeg" if ext == "jpg" else ext)
            media[name] = data
            return name

        def add_content_slide(title, kind, payload):
            shapes, rels = [], []
            tid = 2
            if content.title:
                shapes.append(txbox(tid, "Title", t_box if not content.title[2] else None,
                                    p_xml(title, plain=False).replace('<a:pPr/>', ''), ph_xml(content.title)))
            else:
                shapes.append(txbox(tid, "Title", t_box, p_xml(title, plain=True)))
            tid += 1
            if kind == "text":
                ps = "".join(p_xml(t, lvl, plain=(k == "p"), ordered=o) for k, lvl, t, o in payload)
                if content.body:
                    shapes.append(txbox(tid, "Content", None if content.body[2] else b_box, ps, ph_xml(content.body)))
                else:
                    shapes.append(txbox(tid, "Content", b_box, ps))
            elif kind == "table":
                shapes.append(table_xml(tid, b_box, payload.rows, style_id))
            elif kind == "image":
                nm = new_media(payload.data, payload.ext)
                rels.append(("rIdImg%d" % tid, nm))
                shapes.append(pic_xml(tid, "rIdImg%d" % tid, b_box, payload.data, payload.alt))
            slides.append((slide_xml(shapes), content.path, rels))

        if title_slide and title_l and len(doc.sections) > 1:
            shapes = [txbox(2, "Title", None, p_xml(doc.title.rsplit(".", 1)[0]), ph_xml(title_l.title))]
            slides.append((slide_xml(shapes), title_l.path, []))
        for sec in doc.sections:
            pages = split_text_blocks(sec.blocks) or [("text", [])]
            for i, (kind, payload) in enumerate(pages):
                add_content_slide(sec.title if i == 0 else sec.title + " (계속)", kind, payload)
        rep.append("슬라이드 %d장 생성 (원본 섹션 %d개)" % (len(slides), len(doc.sections)))

        # --- 패키지 재구성 ---
        old_slide_rels = {}
        for n in names:
            if re.match(r"ppt/slides/slide\d+\.xml$", n):
                old_slide_rels[n] = True
        drop = lambda n: (n.startswith("ppt/slides/") or n.startswith("ppt/notesSlides/")
                          or n.startswith("ppt/comments/") or n == "ppt/commentAuthors.xml")
        # presentation.xml.rels
        prels = zin.read("ppt/_rels/presentation.xml.rels").decode("utf-8")
        prels = re.sub(r'<Relationship [^>]*Type="[^"]*/(slide|commentAuthors)"[^>]*/>', "", prels)
        add = "".join('<Relationship Id="rIdNS%d" Type="%s" Target="slides/slide%d.xml"/>' % (i, REL_SLIDE, i)
                      for i in range(1, len(slides) + 1))
        prels = prels.replace("</Relationships>", add + "</Relationships>")
        # presentation.xml
        pres = re.sub(r"<p:sldIdLst>.*?</p:sldIdLst>|<p:sldIdLst/>", "", pres, flags=re.S)
        pres = re.sub(r"<p:custShowLst>.*?</p:custShowLst>", "", pres, flags=re.S)
        pres = re.sub(r'<p:ext uri="\{521415D9-36F7-43E2-AB2F-B90AF26B5E84\}">.*?</p:ext>', "", pres, flags=re.S)
        lst = "<p:sldIdLst>" + "".join('<p:sldId id="%d" r:id="rIdNS%d"/>' % (255 + i, i)
                                       for i in range(1, len(slides) + 1)) + "</p:sldIdLst>"
        pres = pres.replace("<p:sldSz", lst + "<p:sldSz", 1)
        # content types
        ct = zin.read("[Content_Types].xml").decode("utf-8")
        ct = re.sub(r'<Override [^>]*PartName="/ppt/(slides|notesSlides|comments)/[^"]*"[^>]*/>', "", ct)
        ct = re.sub(r'<Override [^>]*PartName="/ppt/commentAuthors.xml"[^>]*/>', "", ct)
        ovr = "".join('<Override PartName="/ppt/slides/slide%d.xml" ContentType="%s"/>' % (i, CT_SLIDE)
                      for i in range(1, len(slides) + 1))
        defaults = ""
        for ext, mt in IMG_CT.items():
            if 'Extension="%s"' % ext not in ct:
                defaults += '<Default Extension="%s" ContentType="%s"/>' % (ext, mt)
        ct = ct.replace("</Types>", defaults + ovr + "</Types>")

        # 유지되는 부분의 rels 가 참조하는 media 만 남김
        keep = [n for n in names if not drop(n)]
        used = set()
        for n in keep:
            if n.endswith(".rels"):
                base_dir = posixpath.dirname(posixpath.dirname(n))
                for t in re.findall(r'Target="([^"]+)"', zin.read(n).decode("utf-8", "replace")):
                    if "media/" in t:
                        used.add(posixpath.normpath(posixpath.join(base_dir, t)) if not t.startswith("/") else t.lstrip("/"))
        with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zout:
            zout.writestr("[Content_Types].xml", ct)
            for n in keep:
                if n == "[Content_Types].xml":
                    continue
                if n.startswith("ppt/media/") and n not in used:
                    continue
                if n == "ppt/presentation.xml":
                    zout.writestr(n, pres)
                elif n == "ppt/_rels/presentation.xml.rels":
                    zout.writestr(n, prels)
                else:
                    zout.writestr(n, zin.read(n))
            for i, (xml, lay, rels) in enumerate(slides, 1):
                zout.writestr("ppt/slides/slide%d.xml" % i, xml)
                rx = ('<Relationship Id="rId1" Type="%s" Target="../slideLayouts/%s"/>'
                      % (REL_LAYOUT, posixpath.basename(lay)))
                for rid, nm in rels:
                    rx += '<Relationship Id="%s" Type="%s" Target="../media/%s"/>' % (rid, REL_IMAGE, nm)
                zout.writestr("ppt/slides/_rels/slide%d.xml.rels" % i,
                              '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                              '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                              + rx + "</Relationships>")
            for k, data in media.items():
                zout.writestr("ppt/media/" + k, data)
    return rep
