#!/usr/bin/env python3
"""PPT/HTML 포맷 일치화 도구.

GUI:  python format_unifier.py
CLI:  python format_unifier.py --to ppt --base 기준.pptx 파일1.pptx 파일2.html --out 결과폴더
"""
import argparse
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))


def main():
    ap = argparse.ArgumentParser(description="PPT/HTML 포맷 일치화")
    ap.add_argument("--to", choices=["ppt", "html"])
    ap.add_argument("--base")
    ap.add_argument("--out", default="unified_out")
    ap.add_argument("files", nargs="*")
    a = ap.parse_args()
    if not (a.to and a.base and a.files):
        try:
            from core.gui import launch
        except ImportError as e:
            print("GUI를 띄울 수 없습니다(tkinter 없음?): %s\nCLI 사용법: --to ppt|html --base 기준 파일... --out 폴더" % e)
            return 1
        launch()
        return 0
    from core import engine
    res = engine.run(a.to, a.base, a.files, a.out)
    s = engine.summary_text(a.to, a.base, res)
    print(s)
    z = engine.make_zip(a.out, res, os.path.join(a.out, "unified.zip"), s)
    print("ZIP:", z)
    return 0 if all(r["ok"] for r in res) else 2


if __name__ == "__main__":
    sys.exit(main())
