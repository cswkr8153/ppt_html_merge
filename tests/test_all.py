import os, sys, tempfile, zipfile, subprocess, unittest
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import make_samples
from core import engine, com_pptx
from core.pptx_io import read_pptx, read_theme


class T(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.d = tempfile.mkdtemp()
        make_samples.make(cls.d)
        j = lambda n: os.path.join(cls.d, n)
        cls.j = staticmethod(j)

    def run_pair(self, to, base, tgt):
        out = os.path.join(self.d, "out_" + to + "_" + base.split(".")[1] + tgt.split(".")[1])
        res = engine.run(to, self.j(base), [self.j(tgt)], out, log=lambda s: None)
        self.assertTrue(res[0]["ok"], res[0]["notes"])
        return res[0]["out"], res[0]

    def test_ppt_to_ppt(self):
        out, r = self.run_pair("ppt", "base.pptx", "target.pptx")
        self.assertTrue(com_pptx.is_plain_pptx(out))
        self.assertEqual(read_theme(out)["cx"], 12191695 if False else read_theme(self.j("base.pptx"))["cx"])
        doc = read_pptx(out)
        titles = [s.title for s in doc.sections]
        self.assertIn("분기 실적", titles)
        pp = __import__("pptx").Presentation(out)   # 유효성 (테스트 한정)
        self.assertEqual(len(pp.slides), 2)

    def test_html_to_ppt(self):
        out, r = self.run_pair("ppt", "base.pptx", "target.html")
        pp = __import__("pptx").Presentation(out)
        self.assertGreaterEqual(len(pp.slides), 2)

    def test_html_to_html(self):
        out, r = self.run_pair("html", "base.html", "target.html")
        t = open(out, encoding="utf-8").read()
        self.assertIn("Malgun Gothic", t)
        self.assertIn("회사 포털", t)
        self.assertIn("© 회사", t)
        self.assertNotIn("대상 푸터", t)
        self.assertNotIn("color:red", t)
        self.assertIn("<strong>", t)
        self.assertIn("하나-1", t)
        self.assertIn("<title>보고서</title>", t)

    def test_ppt_to_html(self):
        out, r = self.run_pair("html", "base.html", "target.pptx")
        t = open(out, encoding="utf-8").read()
        self.assertIn("분기 실적", t)
        self.assertIn("<table>", t)

    def test_ppt_base_html_out(self):
        out, r = self.run_pair("html", "base.pptx", "target.html")
        self.assertIn("aspect-ratio", open(out, encoding="utf-8").read())

    def test_html_base_ppt(self):
        import zipfile as zf
        for tgt in ("target.pptx", "target.html"):
            out, r = self.run_pair("ppt", "base.html", tgt)
            pp = __import__("pptx").Presentation(out)   # 유효성 (테스트 한정)
            self.assertGreaterEqual(len(pp.slides), 2)
            th = zf.ZipFile(out).read("ppt/theme/theme1.xml").decode()
            self.assertIn("Malgun Gothic", th)           # base.html 의 font-family
            self.assertIn("CC0000", th.upper())          # h2 color #c00 -> 제목색
            self.assertTrue(any("PPT 템플릿 생성" in n for n in r["notes"]))

    def test_css_theme_fallback(self):
        from core.pptx_template import extract_css_theme
        th, warn = extract_css_theme("<html><body>x</body></html>")
        self.assertEqual(th["bg"], "FFFFFF")
        self.assertTrue(warn)
        th, _ = extract_css_theme("<style>body{font-family:'Noto Sans KR',sans-serif;color:rgb(1,2,3);background:#fff}</style>")
        self.assertEqual((th["font"], th["fg"], th["bg"]), ("Noto Sans KR", "010203", "FFFFFF"))

    def test_encrypted_non_windows(self):
        f = self.j("enc.pptx")
        open(f, "wb").write(b"\xd0\xcf\x11\xe0not a zip")
        res = engine.run("ppt", self.j("base.pptx"), [f], self.d + "/e", log=lambda s: None)
        self.assertFalse(res[0]["ok"])


if __name__ == "__main__":
    unittest.main()
