"""사내 DRM 암호화 pptx 대응: PowerPoint COM으로 열어 일반 pptx 복사본을 만든다.

- Windows + PowerPoint 설치 + (사내 DRM 플러그인 자동 복호화) 환경에서만 동작.
- pywin32 가 있으면 win32com 사용, 없으면 PowerShell(New-Object -ComObject) 사용.
"""
import os
import subprocess
import sys
import tempfile
import zipfile

PS_SCRIPT = r"""
param([string]$Src, [string]$Dst)
$ErrorActionPreference = 'Stop'
$app = New-Object -ComObject PowerPoint.Application
try {
  $pres = $app.Presentations.Open($Src, $true, $false, $false)
  $pres.SaveCopyAs($Dst, 24)
  $pres.Close()
} finally { $app.Quit() }
"""


def is_plain_pptx(path):
    """zip 구조이고 ppt/presentation.xml 이 있으면 일반 pptx."""
    try:
        if not zipfile.is_zipfile(path):
            return False
        with zipfile.ZipFile(path) as z:
            return "ppt/presentation.xml" in z.namelist()
    except Exception:
        return False


def com_available():
    return sys.platform.startswith("win")


def decrypt_copy(path, tmpdir):
    """COM으로 열어 평문 pptx 복사본 경로를 반환. 실패 시 RuntimeError."""
    if not com_available():
        raise RuntimeError("암호화된 PPT는 Windows + PowerPoint(COM)에서만 열 수 있습니다: " + os.path.basename(path))
    dst = os.path.join(tmpdir, "dec_%d.pptx" % abs(hash(path)))
    src = os.path.abspath(path)
    try:
        import win32com.client  # type: ignore
        app = win32com.client.Dispatch("PowerPoint.Application")
        try:
            pres = app.Presentations.Open(src, True, False, False)
            pres.SaveCopyAs(dst, 24)
            pres.Close()
        finally:
            app.Quit()
    except ImportError:
        ps = os.path.join(tmpdir, "dec.ps1")
        with open(ps, "w", encoding="utf-8-sig") as f:
            f.write(PS_SCRIPT)
        r = subprocess.run(
            ["powershell", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass",
             "-File", ps, "-Src", src, "-Dst", dst],
            capture_output=True, text=True)
        if r.returncode != 0:
            raise RuntimeError("PowerPoint COM 실패: " + (r.stderr or r.stdout).strip()[:300])
    if not is_plain_pptx(dst):
        raise RuntimeError("COM 복호화 결과가 올바른 pptx가 아닙니다: " + os.path.basename(path))
    return dst


def ensure_plain(path, tmpdir):
    """일반 pptx면 그대로, 아니면 COM으로 복호화한 복사본 경로 반환."""
    if is_plain_pptx(path):
        return path
    return decrypt_copy(path, tmpdir)
