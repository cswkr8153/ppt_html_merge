"""HTML 파싱(표준 라이브러리 HTMLParser) → 트리 → 직렬화 / 공통 모델 변환."""
import base64
import html as htmlmod
import os
import re
from html.parser import HTMLParser

from .model import Block, Document, Section

VOID = {"area", "base", "br", "col", "embed", "hr", "img", "input", "link", "meta", "source", "track", "wbr", "param"}
RAW = {"script", "style"}
PRE = {"pre", "textarea"}
BLOCK = {"html", "head", "body", "div", "section", "article", "main", "header", "footer", "nav", "aside", "p", "ul",
         "ol", "li", "table", "thead", "tbody", "tfoot", "tr", "td", "th", "h1", "h2", "h3", "h4", "h5", "h6",
         "blockquote", "pre", "form", "figure", "figcaption", "hr", "title", "meta", "link", "style", "script",
         "dl", "dt", "dd", "details", "summary", "fieldset", "br_"}
# 자동 닫힘 규칙 (새 태그가 열릴 때 닫히는 열린 태그)
AUTOCLOSE = {"li": {"li"}, "p": {"p"}, "td": {"td", "th"}, "th": {"td", "th"}, "tr": {"tr", "td", "th"},
             "dt": {"dt", "dd"}, "dd": {"dt", "dd"}, "option": {"option"}}
P_CLOSERS = {"div", "ul", "ol", "table", "h1", "h2", "h3", "h4", "h5", "h6", "section", "article", "header", "footer",
             "pre", "blockquote", "form", "hr", "nav", "main", "aside", "dl"}


class Node:
    def __init__(self, tag, attrs=None, parent=None):
        self.tag = tag
        self.attrs = attrs if attrs is not None else []   # [(name, value|None)]
        self.children = []   # Node | str(text) | ("comment", s) | ("raw", s) | ("decl", s)
        self.parent = parent

    def get(self, name, default=None):
        for k, v in self.attrs:
            if k == name:
                return v if v is not None else ""
        return default

    def set(self, name, value):
        for i, (k, _) in enumerate(self.attrs):
            if k == name:
                self.attrs[i] = (name, value)
                return
        self.attrs.append((name, value))

    def remove_attr(self, name):
        self.attrs = [(k, v) for k, v in self.attrs if k != name]

    def elements(self):
        return [c for c in self.children if isinstance(c, Node)]

    def iter(self):
        for c in self.children:
            if isinstance(c, Node):
                yield c
                yield from c.iter()

    def text(self):
        out = []
        for c in self.children:
            if isinstance(c, Node):
                if c.tag in RAW:
                    continue
                out.append(c.text())
            elif isinstance(c, str):
                out.append(c)
        return "".join(out)


class TreeBuilder(HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.root = Node("#root")
        self.cur = self.root

    def _open_has(self, tags):
        n = self.cur
        while n is not None and n.tag != "#root":
            if n.tag in tags:
                return n
            if n.tag in ("table", "ul", "ol", "body", "html", "div", "section"):
                return None
            n = n.parent
        return None

    def handle_decl(self, decl):
        self.cur.children.append(("decl", decl))

    def handle_comment(self, data):
        self.cur.children.append(("comment", data))

    def handle_starttag(self, tag, attrs):
        if tag in AUTOCLOSE:
            n = self._open_has(AUTOCLOSE[tag])
            if n is not None:
                self.cur = n.parent
        if tag in P_CLOSERS:
            n = self._open_has({"p"})
            if n is not None:
                self.cur = n.parent
        node = Node(tag, [(k, v) for k, v in attrs], self.cur)
        self.cur.children.append(node)
        if tag not in VOID:
            self.cur = node

    def handle_startendtag(self, tag, attrs):
        node = Node(tag, [(k, v) for k, v in attrs], self.cur)
        self.cur.children.append(node)
        if tag not in VOID and tag in RAW:
            pass

    def handle_endtag(self, tag):
        if tag in VOID:
            return
        n = self.cur
        while n is not None and n.tag != "#root":
            if n.tag == tag:
                self.cur = n.parent
                return
            n = n.parent

    def handle_data(self, data):
        if self.cur.tag in RAW:
            self.cur.children.append(("raw", data))
        else:
            self.cur.children.append(data)


def parse_html(text):
    b = TreeBuilder()
    b.feed(text)
    b.close()
    return b.root


def decode_bytes(raw):
    for enc in ("utf-8-sig", "cp949", "euc-kr", "latin-1"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", "replace")


def read_html_file(path):
    with open(path, "rb") as f:
        return parse_html(decode_bytes(f.read()))


# ---------------------------------------------------------------- 직렬화
def _attrs(node):
    out = ""
    for k, v in node.attrs:
        out += " " + k if v is None else ' %s="%s"' % (k, htmlmod.escape(v, quote=True))
    return out


def _inline(node, pre=False):
    """들여쓰기 없이 직렬화."""
    if isinstance(node, str):
        return htmlmod.escape(node, quote=False) if pre else htmlmod.escape(re.sub(r"\s+", " ", node), quote=False)
    if isinstance(node, tuple):
        kind, s = node
        return "<!--%s-->" % s if kind == "comment" else ("<!%s>" % s if kind == "decl" else s)
    s = "<%s%s>" % (node.tag, _attrs(node))
    if node.tag in VOID:
        return s
    ip = pre or node.tag in PRE
    s += "".join(_inline(c, ip) for c in node.children)
    return s + "</%s>" % node.tag


def serialize(node, level=0):
    ind = "  " * level
    if isinstance(node, str):
        t = re.sub(r"\s+", " ", node).strip()
        return ind + htmlmod.escape(t, quote=False) if t else ""
    if isinstance(node, tuple):
        return ind + _inline(node)
    if node.tag == "#root":
        return "\n".join(x for x in (serialize(c, 0) for c in node.children) if x) + "\n"
    if node.tag in VOID:
        return ind + _inline(node)
    if node.tag in RAW:
        body = "".join(c[1] for c in node.children if isinstance(c, tuple))
        if "\n" in body.strip():
            inner = "\n".join(("  " * (level + 1) + l.rstrip()) if l.strip() else "" for l in body.strip("\n").split("\n"))
            return "%s<%s%s>\n%s\n%s</%s>" % (ind, node.tag, _attrs(node), inner, ind, node.tag)
        return "%s<%s%s>%s</%s>" % (ind, node.tag, _attrs(node), body.strip(), node.tag)
    if node.tag in PRE:
        return ind + _inline(node)
    kids = [c for c in node.children if not (isinstance(c, str) and not c.strip())]
    has_block = any(isinstance(c, Node) and c.tag in BLOCK for c in kids)
    if node.tag in BLOCK and has_block:
        lines = [serialize(c, level + 1) for c in kids]
        lines = [l for l in lines if l]
        return "%s<%s%s>\n%s\n%s</%s>" % (ind, node.tag, _attrs(node), "\n".join(lines), ind, node.tag)
    return ind + _inline(Node_strip(node))


def Node_strip(node):
    """인라인 직렬화 전, 맨 앞/뒤 공백 텍스트 정리."""
    c = list(node.children)
    while c and isinstance(c[0], str) and not c[0].strip():
        c.pop(0)
    while c and isinstance(c[-1], str) and not c[-1].strip():
        c.pop()
    if c and isinstance(c[0], str):
        c[0] = c[0].lstrip()
    if c and isinstance(c[-1], str):
        c[-1] = c[-1].rstrip()
    n = Node(node.tag, node.attrs)
    n.children = c
    return n


# ---------------------------------------------------------------- 탐색 도우미
def find_first(node, tag):
    for n in node.iter():
        if n.tag == tag:
            return n
    return None


def find_all(node, tag):
    return [n for n in node.iter() if n.tag == tag]


def doc_title(root, fallback=""):
    t = find_first(root, "title")
    s = t.text().strip() if t else ""
    return s or fallback


# ---------------------------------------------------------------- HTML → 공통 모델
def _list_items(node, level, out):
    for li in node.elements():
        if li.tag != "li":
            continue
        txt_parts, subs = [], []
        for c in li.children:
            if isinstance(c, Node) and c.tag in ("ul", "ol"):
                subs.append(c)
            elif isinstance(c, Node):
                txt_parts.append(c.text())
            elif isinstance(c, str):
                txt_parts.append(c)
        t = re.sub(r"\s+", " ", "".join(txt_parts)).strip()
        if t:
            out.append((level, t))
        for s in subs:
            _list_items(s, level + 1, out)


def _load_image(src, base_dir, doc):
    if src.startswith("data:"):
        m = re.match(r"data:image/(png|jpe?g|gif);base64,(.*)", src, re.S)
        if m:
            ext = "jpeg" if m.group(1).startswith("jp") else m.group(1)
            try:
                return base64.b64decode(m.group(2)), ext
            except Exception:
                return None
        return None
    if re.match(r"^[a-z]+://", src, re.I) or not base_dir:
        doc.warnings.append("외부/원격 이미지는 가져오지 못함: %s" % src[:60])
        return None
    p = os.path.normpath(os.path.join(base_dir, src.split("?")[0].replace("%20", " ")))
    ext = os.path.splitext(p)[1].lower().lstrip(".")
    ext = "jpeg" if ext == "jpg" else ext
    if ext in ("png", "jpeg", "gif") and os.path.isfile(p):
        with open(p, "rb") as f:
            return f.read(), ext
    doc.warnings.append("이미지 파일을 찾을 수 없거나 형식 미지원: %s" % src[:60])
    return None


def html_to_model(root, name, base_dir=None):
    doc = Document(title=doc_title(root, name))
    body = find_first(root, "body") or root
    main = find_first(body, "main") or find_first(body, "article") or body
    cur = Section(title=doc.title)
    started = [False]

    def new_section(title):
        nonlocal cur
        if cur.blocks or started[0]:
            doc.sections.append(cur)
        cur = Section(title=title)
        started[0] = True

    def walk(n):
        for c in n.children:
            if isinstance(c, str):
                t = re.sub(r"\s+", " ", c).strip()
                if t:
                    cur.blocks.append(Block("para", text=t))
                continue
            if not isinstance(c, Node):
                continue
            t = c.tag
            if t in ("script", "style", "link", "meta", "noscript", "nav", "header", "footer", "head", "title"):
                continue
            if t in ("h1", "h2"):
                new_section(re.sub(r"\s+", " ", c.text()).strip() or doc.title)
            elif t in ("h3", "h4", "h5", "h6"):
                s = re.sub(r"\s+", " ", c.text()).strip()
                if s:
                    cur.blocks.append(Block("para", text=s))
            elif t in ("ul", "ol"):
                items = []
                _list_items(c, 0, items)
                if items:
                    cur.blocks.append(Block("list", items=items, ordered=(t == "ol")))
            elif t == "table":
                rows, hdr = [], False
                for tr in find_all(c, "tr"):
                    cells = [x for x in tr.elements() if x.tag in ("td", "th")]
                    if cells:
                        if not rows and any(x.tag == "th" for x in cells):
                            hdr = True
                        rows.append([re.sub(r"\s+", " ", x.text()).strip() for x in cells])
                if rows:
                    cur.blocks.append(Block("table", rows=rows, header=hdr or True))
            elif t == "img":
                src = c.get("src", "")
                got = _load_image(src, base_dir, doc) if src else None
                if got:
                    cur.blocks.append(Block("image", data=got[0], ext=got[1], alt=c.get("alt", "")))
            elif t in ("p", "pre", "blockquote", "figcaption", "dd", "dt", "summary"):
                imgs = [x for x in c.iter() if x.tag == "img"]
                s = re.sub(r"\s+", " ", c.text()).strip()
                if s:
                    cur.blocks.append(Block("para", text=s))
                for im in imgs:
                    got = _load_image(im.get("src", ""), base_dir, doc)
                    if got:
                        cur.blocks.append(Block("image", data=got[0], ext=got[1], alt=im.get("alt", "")))
            elif t in ("div", "section", "article", "main", "aside", "form", "figure", "details", "dl", "tbody"):
                walk(c)
            else:   # 인라인 요소(a, span, b ...)
                s = re.sub(r"\s+", " ", c.text()).strip()
                if s:
                    cur.blocks.append(Block("para", text=s))

    walk(main)
    if cur.blocks or not doc.sections:
        doc.sections.append(cur)
    return doc


# ---------------------------------------------------------------- 공통 모델 → HTML 조각
def model_to_nodes(doc):
    """Document → HTML 문자열 (섹션별 <section>). 템플릿 삽입 시 raw 로 붙인다."""
    parts = []
    for sec in doc.sections:
        h = ['<section>', '<h2>%s</h2>' % htmlmod.escape(sec.title)]
        for b in sec.blocks:
            if b.kind == "para":
                h.append("<p>%s</p>" % htmlmod.escape(b.text))
            elif b.kind == "list":
                tag = "ol" if b.ordered else "ul"
                h.append(_list_html(b.items, tag))
            elif b.kind == "table":
                rows = []
                for i, r in enumerate(b.rows):
                    cell = "th" if (i == 0 and b.header) else "td"
                    rows.append("<tr>%s</tr>" % "".join("<%s>%s</%s>" % (cell, htmlmod.escape(c), cell) for c in r))
                h.append("<table>%s</table>" % "".join(rows))
            elif b.kind == "image":
                mt = "jpeg" if b.ext == "jpg" else b.ext
                h.append('<img src="data:image/%s;base64,%s" alt="%s">'
                         % (mt, base64.b64encode(b.data).decode("ascii"), htmlmod.escape(b.alt, quote=True)))
        h.append("</section>")
        parts.append("".join(h))
    return parts


def _list_html(items, tag):
    out, depth = [], -1
    for lvl, t in items:
        lvl = min(lvl, depth + 1)
        if lvl > depth:
            out.append("<%s>" % tag)
            depth += 1
        else:
            out.append("</li>")
            while depth > lvl:
                out.append("</%s></li>" % tag)
                depth -= 1
        out.append("<li>%s" % htmlmod.escape(t))
    out.append("</li>")
    while depth > 0:
        out.append("</%s></li>" % tag)
        depth -= 1
    out.append("</%s>" % tag)
    return "".join(out)
