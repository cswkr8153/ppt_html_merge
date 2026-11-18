"""기준 HTML 의 head(CSS/메타)·헤더·푸터·레이아웃을 유지하고 대상 본문을 삽입."""
import html as htmlmod
import re

from .html_io import (Node, find_first, find_all, doc_title, html_to_model, model_to_nodes, parse_html, serialize)
from .pptx_io import read_theme

KEEP_TAGS = {"header", "nav", "footer", "aside", "script", "noscript", "style", "link"}
SLOT_TOKENS = ["content", "main-content", "maincontent", "main", "post", "article", "page-content", "container",
               "wrapper"]
PRESENTATIONAL = {"bgcolor", "align", "valign", "border", "cellpadding", "cellspacing", "color", "face", "size",
                  "width", "height", "background", "style", "hspace", "vspace", "nowrap", "bordercolor"}
REPLACE_TAG = {"b": "strong", "i": "em", "center": "div", "strike": "del", "u": None, "big": None, "small": None,
               "font": None, "tt": "code"}


def find_slot(root):
    body = find_first(root, "body") or root
    for pred in (lambda n: n.tag == "main", lambda n: n.get("role") == "main", lambda n: n.tag == "article"):
        for n in body.iter():
            if pred(n):
                return n
    for tok in SLOT_TOKENS:
        for n in body.iter():
            toks = ((n.get("id") or "") + " " + (n.get("class") or "")).lower().split()
            if tok in toks:
                return n
    return body


def fill_slot(slot, content_children):
    kept = [c for c in slot.children if isinstance(c, Node) and c.tag in KEEP_TAGS]
    before, after, seen_footer = [], [], False
    for k in kept:
        if k.tag == "footer" or (k.tag in ("script", "noscript") and seen_footer) or seen_footer:
            seen_footer = True
            after.append(k)
        else:
            before.append(k)
    slot.children = before + content_children + after
    for c in content_children:
        if isinstance(c, Node):
            c.parent = slot


def clean_content(nodes, stats):
    """프레젠테이션용 속성/구식 태그 정리 (구조 정리)."""
    out = []

    def walk(n):
        if isinstance(n, (str, tuple)):
            return [n]
        for k, _ in list(n.attrs):
            if k.lower() in PRESENTATIONAL or k.lower().startswith("on"):
                n.remove_attr(k)
                stats["attrs"] = stats.get("attrs", 0) + 1
        kids = []
        for c in n.children:
            kids.extend(walk(c))
        n.children = kids
        for c in kids:
            if isinstance(c, Node):
                c.parent = n
        if n.tag in REPLACE_TAG:
            new = REPLACE_TAG[n.tag]
            stats["tags"] = stats.get("tags", 0) + 1
            if new is None:
                return kids
            n.tag = new
        return [n]

    for n in nodes:
        out.extend(walk(n))
    # 빈 텍스트 연속 공백 정리는 직렬화 단계에서 수행
    return out


def extract_target_content(root, stats):
    for tag in ("script", "style", "link", "noscript", "meta", "title"):
        for n in find_all(root, tag):
            if n.parent is not None:
                n.parent.children = [c for c in n.parent.children if c is not n]
                stats["removed_" + tag] = stats.get("removed_" + tag, 0) + 1
    body = find_first(root, "body") or root
    # 대상의 header/nav/footer 는 템플릿 것을 쓰므로 제거
    for tag in ("header", "nav", "footer"):
        for n in find_all(body, tag):
            if n.parent is not None:
                n.parent.children = [c for c in n.parent.children if c is not n]
                stats["removed_" + tag] = stats.get("removed_" + tag, 0) + 1
    main = find_first(body, "main") or find_first(body, "article") or body
    return list(main.children)


def apply_to_template(template_root, content_children, title):
    t = find_first(template_root, "title")
    if t is not None and title:
        t.children = [title]
    slot = find_slot(template_root)
    fill_slot(slot, content_children)


def build_html_from_html(base_root_text, target_root, target_name):
    """같은 형식: 대상 본문의 서식 속성을 정리해 기준 템플릿 슬롯에 삽입."""
    stats = {}
    rep = []
    tpl = parse_html(base_root_text)
    title = doc_title(target_root, target_name)
    content = extract_target_content(target_root, stats)
    content = clean_content(content, stats)
    apply_to_template(tpl, content, title)
    rep.append("기준 head(CSS/메타)·헤더·푸터 레이아웃 적용")
    if stats.get("attrs"):
        rep.append("서식 속성/인라인 스타일 %d개 제거" % stats["attrs"])
    if stats.get("tags"):
        rep.append("구식 태그 %d개 정리(b→strong 등)" % stats["tags"])
    for k, lab in (("removed_style", "<style>"), ("removed_link", "<link>"), ("removed_script", "<script>"),
                   ("removed_header", "<header>"), ("removed_nav", "<nav>"), ("removed_footer", "<footer>")):
        if stats.get(k):
            rep.append("대상 파일의 %s %d개 제거(기준 것으로 대체)" % (lab, stats[k]))
    return serialize(tpl), rep


def build_html_from_model(base_html_text, doc):
    """PPT 등에서 온 Document 를 기준 HTML 템플릿에 삽입."""
    tpl = parse_html(base_html_text)
    frags = model_to_nodes(doc)
    nodes = []
    for f in frags:
        nodes.extend(parse_html(f).children)
    apply_to_template(tpl, nodes, doc.title.rsplit(".", 1)[0])
    rep = ["기준 head(CSS/메타)·헤더·푸터 레이아웃 적용", "슬라이드 %d개 → <section> 변환" % len(doc.sections)]
    return serialize(tpl), rep


def template_from_pptx(pptx_path):
    """기준이 PPT 이고 출력이 HTML 일 때: 슬라이드 비율/테마 폰트·색으로 CSS 템플릿 생성."""
    th = read_theme(pptx_path)
    c = th["colors"]
    col = lambda k, d: "#" + c.get(k, d)
    font = th["minor"] or "sans-serif"
    hfont = th["major"] or font
    css = """
body {{ font-family: {font}, sans-serif; color: {fg}; background: {bg}; margin: 0; }}
main {{ max-width: 960px; margin: 0 auto; padding: 24px; }}
section {{ aspect-ratio: {w} / {h}; box-sizing: border-box; padding: 32px 48px; margin: 24px 0;
  background: {bg}; border: 1px solid {acc}33; border-radius: 6px; overflow: auto; }}
h1, h2 {{ font-family: {hfont}, sans-serif; color: {acc}; margin-top: 0; }}
table {{ border-collapse: collapse; width: 100%; }}
th {{ background: {acc}; color: {bg}; }}
th, td {{ border: 1px solid {acc}66; padding: 6px 10px; text-align: left; }}
img {{ max-width: 100%; height: auto; }}
""".format(font=font, hfont=hfont, fg=col("dk1", "000000"), bg=col("lt1", "FFFFFF"), acc=col("accent1", "4472C4"),
           w=th["cx"], h=th["cy"])
    return ("<!DOCTYPE html>\n<html lang=\"ko\">\n<head>\n<meta charset=\"utf-8\">\n"
            "<meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">\n<title></title>\n"
            "<style>%s</style>\n</head>\n<body>\n<main></main>\n</body>\n</html>\n" % css)
