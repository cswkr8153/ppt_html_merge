"""테스트용 샘플 생성 (python-pptx 는 테스트 샘플 생성에만 사용. 본 프로그램은 불필요)."""
import os
import sys

from pptx import Presentation
from pptx.util import Inches
from pptx.dml.color import RGBColor


def make(d):
    os.makedirs(d, exist_ok=True)
    # 기준 pptx: 16:9 + 제목 색 지정
    base = Presentation()
    base.slide_width, base.slide_height = Inches(13.333), Inches(7.5)
    s = base.slides.add_slide(base.slide_layouts[1])
    s.shapes.title.text = "기준 제목"
    s.placeholders[1].text = "기준 본문"
    base.save(os.path.join(d, "base.pptx"))

    # 대상 pptx: 4:3, 제목+목록+표
    t = Presentation()
    s = t.slides.add_slide(t.slide_layouts[1])
    s.shapes.title.text = "분기 실적"
    tf = s.placeholders[1].text_frame
    tf.text = "매출 증가"
    p = tf.add_paragraph(); p.text = "신규 고객 확대"; p.level = 1
    s2 = t.slides.add_slide(t.slide_layouts[5])
    s2.shapes.title.text = "표 슬라이드"
    tb = s2.shapes.add_table(2, 2, Inches(1), Inches(2), Inches(6), Inches(1.5)).table
    for r in range(2):
        for c in range(2):
            tb.cell(r, c).text = "r%dc%d" % (r, c)
    t.save(os.path.join(d, "target.pptx"))

    open(os.path.join(d, "base.html"), "w", encoding="utf-8").write(
        '<!DOCTYPE html><html><head><meta charset="utf-8"><title>BASE</title>'
        '<style>body{font-family:Malgun Gothic;color:#223}h2{color:#c00}</style></head>'
        '<body><header><h1>회사 포털</h1></header><div id="content"><p>placeholder</p></div>'
        '<footer>© 회사</footer><script>var x=1;</script></body></html>')
    open(os.path.join(d, "target.html"), "w", encoding="utf-8").write(
        '<html><head><title>보고서</title><style>h1{color:blue}</style></head><body bgcolor="#eee">'
        '<center><font face="Arial" size="4"><b>굵은 글</b></font></center>'
        '<h1 style="color:red">개요</h1><p align="center">첫 문단<p>둘째 문단'
        '<ul><li>하나<ul><li>하나-1</li></ul><li>둘</ul>'
        '<h2>데이터</h2><table border="1"><tr><th>이름</th><th>값</th></tr><tr><td>a</td><td>1</td></tr></table>'
        '<footer>대상 푸터</footer></body></html>')


if __name__ == "__main__":
    make(sys.argv[1])
