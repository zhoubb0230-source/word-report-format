# -*- coding: utf-8 -*-
"""
msword.py — Microsoft Word (Windows) converter via COM automation.

When running on Windows with Microsoft Office / Word installed, we can convert
.doc <-> .docx by driving Word through COM. This is the natural fallback where
LibreOffice (soffice) is absent — a very common setup on user Windows machines.

Self-contained: driven through PowerShell (always present on Windows), so no
pywin32 / pip dependency is required. Detection is registry-based (checks that
the ``Word.Application`` COM ProgID is registered) so we never spawn Word just
to answer "is it available?".
"""

import os
import subprocess
import tempfile

# Word SaveAs format codes (WdSaveFormat).
_WD_FORMAT = {"docx": 16, "doc": 0}  # 16=wdFormatDocumentDefault, 0=wdFormatDocument


def _is_windows():
    return os.name == "nt"


def _powershell():
    """Return a runnable PowerShell command name, or None."""
    if not _is_windows():
        return None
    import shutil
    return shutil.which("powershell") or shutil.which("pwsh")


def msword_available():
    """True only on Windows where Word's COM automation is registered."""
    ps = _powershell()
    if not ps:
        return False
    try:
        res = subprocess.run(
            [ps, "-NoProfile", "-NonInteractive", "-Command",
             "if ([Type]::GetTypeFromProgID('Word.Application')) "
             "{ Write-Output 'yes' } else { Write-Output 'no' }"],
            capture_output=True, timeout=60,
        )
    except Exception:
        return False
    return b"yes" in (res.stdout or b"")


_PS_CONVERT = r"""
param([string]$Src, [string]$Dst, [int]$Fmt)
$ErrorActionPreference = 'Stop'
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
    # Open read-only (never touches the original), no add-to-recent.
    $doc = $word.Documents.Open($Src, $false, $true)
    $doc.SaveAs([ref]$Dst, [ref]$Fmt)
    $doc.Close($false)
} finally {
    $word.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
}
"""


def convert(input_path, out_format, out_dir):
    """
    Convert ``input_path`` to ``out_format`` ('docx' or 'doc') into ``out_dir``
    using Microsoft Word via COM. Returns the output path, or raises on failure.
    """
    ps = _powershell()
    if not ps:
        raise RuntimeError("PowerShell not available; cannot drive Microsoft Word")
    if out_format not in _WD_FORMAT:
        raise RuntimeError("unsupported target format for Word: %s" % out_format)
    os.makedirs(out_dir, exist_ok=True)

    src = os.path.abspath(input_path)
    stem = os.path.splitext(os.path.basename(input_path))[0]
    out = os.path.abspath(os.path.join(out_dir, stem + "." + out_format))

    script_fd, script_path = tempfile.mkstemp(suffix=".ps1", prefix="msword_conv_")
    try:
        with os.fdopen(script_fd, "w", encoding="utf-8") as f:
            f.write(_PS_CONVERT)
        res = subprocess.run(
            [ps, "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass",
             "-File", script_path, "-Src", src, "-Dst", out,
             "-Fmt", str(_WD_FORMAT[out_format])],
            capture_output=True, timeout=600,
        )
    finally:
        try:
            os.remove(script_path)
        except OSError:
            pass

    if not os.path.exists(out):
        raise RuntimeError(
            "Word conversion to %s failed: %s"
            % (out_format, (res.stderr or b"").decode("utf-8", "replace")[:500])
        )
    return out


if __name__ == "__main__":
    import json
    print(json.dumps({"msword_available": msword_available()}))
