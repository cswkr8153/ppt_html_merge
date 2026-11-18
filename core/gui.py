"""Tkinter GUI: ① 출력 형식 → ② 기준 파일 → ③ 대상 파일들 → ④ 실행 → ⑤ 요약/ZIP."""
import os
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk

from . import com_pptx, engine

FT_ALL = [("PPT / HTML", "*.pptx *.pptm *.html *.htm"), ("PPT", "*.pptx *.pptm"), ("HTML", "*.html *.htm"), ("모든 파일", "*.*")]


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("PPT / HTML 포맷 일치화")
        self.geometry("760x640")
        self.to = tk.StringVar(value="")
        self.base = tk.StringVar()
        self.outdir = tk.StringVar(value=os.path.join(os.path.expanduser("~"), "unified_out"))
        self.targets = []
        self.results = None
        self._build()

    def _build(self):
        pad = {"padx": 10, "pady": 6}
        f1 = ttk.LabelFrame(self, text="① 출력 형식 선택 (먼저 선택하세요)")
        f1.pack(fill="x", **pad)
        ttk.Radiobutton(f1, text="PPT로 만들기", variable=self.to, value="ppt", command=self._refresh).pack(side="left", padx=10, pady=6)
        ttk.Radiobutton(f1, text="HTML로 만들기", variable=self.to, value="html", command=self._refresh).pack(side="left", padx=10)
        ttk.Label(f1, text="COM(암호화 PPT): %s" % ("사용 가능" if com_pptx.com_available() else "Windows 아님 - 불가")).pack(side="right", padx=10)

        self.f2 = ttk.LabelFrame(self, text="② 기준 파일 (1개)")
        self.f2.pack(fill="x", **pad)
        ttk.Entry(self.f2, textvariable=self.base, state="readonly").pack(side="left", fill="x", expand=True, padx=6, pady=6)
        self.b_base = ttk.Button(self.f2, text="기준 파일 가져오기", command=self.pick_base)
        self.b_base.pack(side="right", padx=6)

        self.f3 = ttk.LabelFrame(self, text="③ 일치화할 파일들 (PPT/HTML 혼합 가능)")
        self.f3.pack(fill="both", expand=True, **pad)
        self.lb = tk.Listbox(self.f3, selectmode="extended", height=8)
        self.lb.pack(side="left", fill="both", expand=True, padx=6, pady=6)
        bf = ttk.Frame(self.f3)
        bf.pack(side="right", padx=6)
        self.b_add = ttk.Button(bf, text="파일 추가", command=self.add_targets)
        self.b_add.pack(fill="x", pady=2)
        ttk.Button(bf, text="선택 제거", command=self.remove_targets).pack(fill="x", pady=2)
        ttk.Button(bf, text="모두 제거", command=self.clear_targets).pack(fill="x", pady=2)

        f4 = ttk.Frame(self)
        f4.pack(fill="x", **pad)
        ttk.Label(f4, text="저장 폴더:").pack(side="left")
        ttk.Entry(f4, textvariable=self.outdir).pack(side="left", fill="x", expand=True, padx=6)
        ttk.Button(f4, text="찾기", command=self.pick_out).pack(side="left")
        self.b_run = ttk.Button(self, text="④ 실행", command=self.run)
        self.b_run.pack(fill="x", **pad)

        self.log = scrolledtext.ScrolledText(self, height=12, state="disabled")
        self.log.pack(fill="both", expand=True, **pad)
        self.b_zip = ttk.Button(self, text="⑤ ZIP으로 저장", command=self.save_zip, state="disabled")
        self.b_zip.pack(fill="x", **pad)
        self._refresh()

    def _refresh(self):
        on = bool(self.to.get())
        for w in (self.b_base, self.b_add, self.b_run):
            w.state(["!disabled"] if on else ["disabled"])

    def pick_base(self):
        p = filedialog.askopenfilename(title="기준 파일", filetypes=FT_ALL)
        if p:
            if engine.kind_of(p) is None:
                return messagebox.showerror("오류", "PPT/HTML 파일이 아닙니다.")
            if self.to.get() == "ppt" and engine.kind_of(p) != "ppt":
                return messagebox.showwarning("안내", "PPT로 출력하려면 기준 파일도 PPT여야 합니다.")
            self.base.set(p)

    def add_targets(self):
        for p in filedialog.askopenfilenames(title="대상 파일", filetypes=FT_ALL):
            if engine.kind_of(p) and p not in self.targets and p != self.base.get():
                self.targets.append(p)
                self.lb.insert("end", p)

    def remove_targets(self):
        for i in reversed(self.lb.curselection()):
            self.lb.delete(i)
            del self.targets[i]

    def clear_targets(self):
        self.lb.delete(0, "end")
        self.targets.clear()

    def pick_out(self):
        d = filedialog.askdirectory()
        if d:
            self.outdir.set(d)

    def _write(self, s):
        self.log.configure(state="normal")
        self.log.insert("end", s)
        self.log.see("end")
        self.log.configure(state="disabled")
        self.update_idletasks()

    def run(self):
        if not self.base.get():
            return messagebox.showwarning("안내", "기준 파일을 먼저 가져오세요.")
        if not self.targets:
            return messagebox.showwarning("안내", "일치화할 파일을 추가하세요.")
        self.log.configure(state="normal"); self.log.delete("1.0", "end"); self.log.configure(state="disabled")
        self.b_run.state(["disabled"])
        try:
            self.results = engine.run(self.to.get(), self.base.get(), self.targets, self.outdir.get(), log=lambda s: self._write(s + "\n"))
            self.summary = engine.summary_text(self.to.get(), self.base.get(), self.results)
            self._write("\n" + self.summary)
            self.b_zip.state(["!disabled"])
        except Exception as e:
            messagebox.showerror("오류", str(e))
        finally:
            self.b_run.state(["!disabled"])

    def save_zip(self):
        p = filedialog.asksaveasfilename(defaultextension=".zip", initialfile="unified.zip", filetypes=[("ZIP", "*.zip")])
        if p:
            engine.make_zip(self.outdir.get(), self.results, p, self.summary)
            messagebox.showinfo("완료", "저장했습니다:\n" + p)


def launch():
    App().mainloop()
