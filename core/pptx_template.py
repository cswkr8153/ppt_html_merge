"""HTML 기준 → PPT 출력용: 기준 HTML 의 CSS(폰트·색)로 최소 pptx 템플릿(마스터/레이아웃/테마)을 생성."""
import os
import re
import zipfile
from xml.sax.saxutils import escape

NAMED = {"black": "000000", "white": "FFFFFF", "red": "FF0000", "green": "008000", "blue": "0000FF",
         "navy": "000080", "gray": "808080", "grey": "808080", "silver": "C0C0C0", "orange": "FFA500",
         "purple": "800080", "teal": "008080", "maroon": "800000", "yellow": "FFFF00", "darkblue": "00008B",
         "darkgreen": "006400", "darkred": "8B0000", "crimson": "DC143C", "tomato": "FF6347"}
GENERIC = {"sans-serif", "serif", "monospace", "system-ui", "cursive", "fantasy", "inherit", "initial"}
DEFAULT = {"font": "Malgun Gothic", "fg": "222222", "bg": "FFFFFF", "title": "1F3864", "accent": "4472C4"}
EMU_W, EMU_H = 12192000, 6858000


def parse_color(val):
    """CSS 색 문자열 → 'RRGGBB' 또는 None."""
    if not val:
        return None
    v = val.strip().lower()
    m = re.search(r"#([0-9a-f]{6})\b", v)
    if m:
        return m.group(1).upper()
    m = re.search(r"#([0-9a-f]{3})\b", v)
    if m:
        return "".join(ch * 2 for ch in m.group(1)).upper()
    m = re.search(r"rgba?\(\s*(\d+)[,\s]+(\d+)[,\s]+(\d+)", v)
    if m:
        return "%02X%02X%02X" % tuple(min(255, int(x)) for x in m.groups())
    for tok in re.findall(r"[a-z]+", v):
        if tok in NAMED:
            return NAMED[tok]
    return None


def _rules(css):
    css = re.sub(r"/\*.*?\*/", "", css, flags=re.S)
    rules = {}
    for sel, body in re.findall(r"([^{}@]+)\{([^{}]*)\}", css):
        props = {}
        for decl in body.split(";"):
            if ":" in decl:
                k, v = decl.split(":", 1)
                props[k.strip().lower()] = v.strip()
        for s in sel.split(","):
            rules.setdefault(s.strip().lower(), {}).update(props)
    return rules


def extract_css_theme(html_text, base_dir=None):
    """반환: (theme dict, 경고 list). CSS 가 없거나 못 읽으면 기본값 사용."""
    warn = []
    css = "\n".join(re.findall(r"<style[^>]*>(.*?)</style>", html_text, flags=re.S | re.I))
    for href in re.findall(r"<link[^>]*rel=[\"']?stylesheet[\"']?[^>]*href=[\"']([^\"']+)[\"']", html_text, flags=re.I):
        p = os.path.join(base_dir, href.split("?")[0]) if base_dir and not re.match(r"^[a-z]+://", href, re.I) else None
        if p and os.path.isfile(p):
            with open(p, "rb") as f:
                css += "\n" + f.read().decode("utf-8", "replace")
        else:
            warn.append("외부 CSS를 읽지 못함(기본값 사용): %s" % href[:60])
    rules = _rules(css)
    m = re.search(r"<body[^>]*style=[\"']([^\"']*)[\"']", html_text, flags=re.I)
    if m:
        inline = {}
        for decl in m.group(1).split(";"):
            if ":" in decl:
                k, v = decl.split(":", 1)
                inline[k.strip().lower()] = v.strip()
        rules.setdefault("body", {}).update(inline)

    def prop(sels, *names):
        for s in sels:
            r = rules.get(s)
            if r:
                for n in names:
                    if n in r:
                        return r[n]
        return None

    th = dict(DEFAULT)
    fam = prop(["body", "html", ":root"], "font-family")
    if fam:
        first = fam.split(",")[0].strip().strip("\"'")
        if first and first.lower() not in GENERIC and "var(" not in first:
            th["font"] = first
    th["fg"] = parse_color(prop(["body", "html", ":root"], "color")) or th["fg"]
    th["bg"] = parse_color(prop(["body", "html", ":root"], "background-color", "background")) or th["bg"]
    th["title"] = parse_color(prop(["h1", "h2"], "color")) or th["title"]
    acc = parse_color(prop(["a", ".btn", "th", "h2", "h1"], "color", "background-color", "background"))
    th["accent"] = acc or th["accent"]
    if not css.strip():
        warn.append("기준 HTML에서 CSS를 찾지 못해 기본 서식(맑은 고딕, 파란 계열)을 사용")
    return th, warn


# ---------------------------------------------------------------- 템플릿 XML
NS = ('xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
      'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"')
HDR = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
RT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"


def _rels(items):
    return (HDR + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            + "".join('<Relationship Id="%s" Type="%s%s" Target="%s"/>' % (i, RT, t, tg) for i, t, tg in items)
            + "</Relationships>")


def _ph_sp(idn, name, ph, box, text="", extra_body=""):
    x, y, cx, cy = box
    return ('<p:sp><p:nvSpPr><p:cNvPr id="%d" name="%s"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr>%s</p:nvPr></p:nvSpPr>'
            '<p:spPr><a:xfrm><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></a:xfrm></p:spPr>'
            '<p:txBody><a:bodyPr%s/><a:lstStyle/><a:p><a:r><a:rPr lang="ko-KR"/><a:t>%s</a:t></a:r></a:p></p:txBody></p:sp>'
            % (idn, name, ph, x, y, cx, cy, extra_body, escape(text)))


def _spTree(shapes):
    return ('<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
            '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
            + "".join(shapes) + "</p:spTree>")


def _lvl(n, sz, color, bullet):
    bu = '<a:buFont typeface="Arial"/><a:buChar char="%s"/>' % bullet if bullet else "<a:buNone/>"
    mar = 342900 * n
    return ('<a:lvl%dpPr marL="%d" indent="%d" algn="l"><a:spcBef><a:spcPts val="600"/></a:spcBef>%s'
            '<a:defRPr sz="%d" kern="1200"><a:solidFill>%s</a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/></a:defRPr></a:lvl%dpPr>'
            % (n, mar if bullet else 0, -285750 if bullet else 0, bu, sz, color, n))


def build_template(theme, out_path):
    fg, bg, ti, ac, font = theme["fg"], theme["bg"], theme["title"], theme["accent"], escape(theme["font"])
    clr = lambda h: '<a:srgbClr val="%s"/>' % h
    mx = int(EMU_W * 0.06)
    t_box = (mx, 365125, EMU_W - 2 * mx, 1143000)
    b_box = (mx, 1690688, EMU_W - 2 * mx, 4351338)
    ct_box = (1524000, 1122363, 9144000, 2387600)
    st_box = (1524000, 3602038, 9144000, 1655762)

    master = (HDR + '<p:sldMaster %s><p:cSld><p:bg><p:bgPr><a:solidFill>%s</a:solidFill><a:effectLst/></p:bgPr></p:bg>%s</p:cSld>'
              '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" '
              'accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>'
              '<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/><p:sldLayoutId id="2147483650" r:id="rId2"/></p:sldLayoutIdLst>'
              '<p:txStyles><p:titleStyle><a:lvl1pPr algn="l"><a:spcBef><a:spcPct val="0"/></a:spcBef><a:buNone/>'
              '<a:defRPr sz="4000" b="1" kern="1200"><a:solidFill>%s</a:solidFill><a:latin typeface="+mj-lt"/><a:ea typeface="+mj-ea"/></a:defRPr></a:lvl1pPr></p:titleStyle>'
              '<p:bodyStyle>%s%s%s</p:bodyStyle>'
              '<p:otherStyle><a:defPPr><a:defRPr lang="ko-KR"/></a:defPPr></p:otherStyle></p:txStyles></p:sldMaster>'
              % (NS, clr(bg),
                 _spTree([_ph_sp(2, "Title Placeholder 1", '<p:ph type="title"/>', t_box, "제목", ' anchor="ctr"'),
                          _ph_sp(3, "Text Placeholder 2", '<p:ph type="body" idx="1"/>', b_box, "본문")]),
                 clr(ti), _lvl(1, 2800, '<a:schemeClr val="tx1"/>', "•"), _lvl(2, 2400, '<a:schemeClr val="tx1"/>', "–"),
                 _lvl(3, 2000, '<a:schemeClr val="tx1"/>', "•")))
    lay1 = (HDR + '<p:sldLayout %s type="title" preserve="1"><p:cSld name="Title Slide">%s</p:cSld>'
            '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>'
            % (NS, _spTree([_ph_sp(2, "Title 1", '<p:ph type="ctrTitle"/>', ct_box, "제목", ' anchor="b"'),
                            _ph_sp(3, "Subtitle 2", '<p:ph type="subTitle" idx="1"/>', st_box, "부제목")])))
    lay2 = (HDR + '<p:sldLayout %s type="obj" preserve="1"><p:cSld name="Title and Content">%s</p:cSld>'
            '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>'
            % (NS, _spTree([_ph_sp(2, "Title 1", '<p:ph type="title"/>', t_box, "제목"),
                            _ph_sp(3, "Content Placeholder 2", '<p:ph idx="1"/>', b_box, "본문")])))
    fills = '<a:fillStyleLst>' + '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>' * 3 + '</a:fillStyleLst>'
    lns = '<a:lnStyleLst>' + '<a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>' * 3 + '</a:lnStyleLst>'
    eff = '<a:effectStyleLst>' + '<a:effectStyle><a:effectLst/></a:effectStyle>' * 3 + '</a:effectStyleLst>'
    bgf = '<a:bgFillStyleLst>' + '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>' * 3 + '</a:bgFillStyleLst>'
    others = ("accent2", "ED7D31"), ("accent3", "A5A5A5"), ("accent4", "FFC000"), ("accent5", "5B9BD5"), ("accent6", "70AD47")
    theme_xml = (HDR + '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Generated from HTML">'
                 '<a:themeElements><a:clrScheme name="HTML"><a:dk1>%s</a:dk1><a:lt1>%s</a:lt1><a:dk2>%s</a:dk2><a:lt2>%s</a:lt2>'
                 '<a:accent1>%s</a:accent1>%s<a:hlink>%s</a:hlink><a:folHlink>%s</a:folHlink></a:clrScheme>'
                 '<a:fontScheme name="HTML"><a:majorFont><a:latin typeface="%s"/><a:ea typeface="%s"/><a:cs typeface=""/></a:majorFont>'
                 '<a:minorFont><a:latin typeface="%s"/><a:ea typeface="%s"/><a:cs typeface=""/></a:minorFont></a:fontScheme>'
                 '<a:fmtScheme name="Office">%s%s%s%s</a:fmtScheme></a:themeElements></a:theme>'
                 % (clr(fg), clr(bg), clr(ti), clr("EEEEEE"), clr(ac),
                    "".join("<a:%s>%s</a:%s>" % (n, clr(c), n) for n, c in others), clr(ac), clr(ac),
                    font, font, font, font, fills, lns, eff, bgf))
    pres = (HDR + '<p:presentation %s><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>'
            '<p:sldSz cx="%d" cy="%d"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>' % (NS, EMU_W, EMU_H))
    ct = (HDR + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
          '<Default Extension="xml" ContentType="application/xml"/>'
          '<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>'
          '<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>'
          '<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>'
          '<Override PartName="/ppt/slideLayouts/slideLayout2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>'
          '<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>'
          '<Override PartName="/ppt/tableStyles.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml"/>'
          '</Types>')
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", ct)
        z.writestr("_rels/.rels", _rels([("rId1", "officeDocument", "ppt/presentation.xml")]))
        z.writestr("ppt/presentation.xml", pres)
        z.writestr("ppt/_rels/presentation.xml.rels", _rels([("rId1", "slideMaster", "slideMasters/slideMaster1.xml"),
                                                              ("rId2", "theme", "theme/theme1.xml"),
                                                              ("rId3", "tableStyles", "tableStyles.xml")]))
        z.writestr("ppt/slideMasters/slideMaster1.xml", master)
        z.writestr("ppt/slideMasters/_rels/slideMaster1.xml.rels", _rels([
            ("rId1", "slideLayout", "../slideLayouts/slideLayout1.xml"),
            ("rId2", "slideLayout", "../slideLayouts/slideLayout2.xml"), ("rId3", "theme", "../theme/theme1.xml")]))
        for i, lay in ((1, lay1), (2, lay2)):
            z.writestr("ppt/slideLayouts/slideLayout%d.xml" % i, lay)
            z.writestr("ppt/slideLayouts/_rels/slideLayout%d.xml.rels" % i,
                       _rels([("rId1", "slideMaster", "../slideMasters/slideMaster1.xml")]))
        z.writestr("ppt/theme/theme1.xml", theme_xml)
        z.writestr("ppt/tableStyles.xml", HDR + '<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
                   'def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"/>')
    return out_path


def template_from_html(html_text, out_path, base_dir=None):
    theme, warn = extract_css_theme(html_text, base_dir)
    build_template(theme, out_path)
    return theme, warn
