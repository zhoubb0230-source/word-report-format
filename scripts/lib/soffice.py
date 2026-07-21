# -*- coding: utf-8 -*-
"""
soffice.py — vendored LibreOffice runner.

Runs `soffice` headless with a private temporary user profile, which is required
in non-root sandboxes (otherwise soffice aborts with "User installation could
not be completed"). Applies an AF_UNIX shim only if the sandbox blocks unix
sockets. Self-contained; used for .doc<->.docx conversion.
"""

import contextlib
import os
import shutil
import socket
import subprocess
import tempfile
from pathlib import Path


def soffice_available():
    return shutil.which("soffice") is not None or shutil.which("libreoffice") is not None


def _bin():
    return shutil.which("soffice") or shutil.which("libreoffice")


_SHIM_SO = Path(tempfile.gettempdir()) / "lo_socket_shim_fr.so"


def _needs_shim():
    try:
        s = socket.socket(socket.AF_UNIX, socket.SOCK_STREAM)
        s.close()
        return False
    except OSError:
        return True


def _ensure_shim():
    if _SHIM_SO.exists():
        return _SHIM_SO
    c_src = _SHIM_SO.with_suffix(".c")
    c_src.write_text(
        "#define _GNU_SOURCE\n#include <sys/socket.h>\n#include <errno.h>\n"
        "int socket(int d,int t,int p){\n"
        "  extern int __real_socket(int,int,int);\n"
        "  if(d==AF_UNIX){errno=EAFNOSUPPORT;return -1;}\n"
        "  return syscall(41,d,t,p);\n}\n"
    )
    # Best-effort compile; if no compiler, just skip the shim.
    try:
        subprocess.run(["cc", "-shared", "-fPIC", "-ldl", str(c_src), "-o", str(_SHIM_SO)],
                       check=True, capture_output=True)
    except Exception:
        return None
    return _SHIM_SO if _SHIM_SO.exists() else None


def _env():
    env = os.environ.copy()
    env["SAL_USE_VCLPLUGIN"] = "svp"
    if _needs_shim():
        shim = _ensure_shim()
        if shim:
            env["LD_PRELOAD"] = str(shim)
    return env


def run(args, timeout=600):
    binp = _bin()
    if not binp:
        raise RuntimeError("soffice/libreoffice not found on PATH")
    with contextlib.ExitStack() as stack:
        profile = stack.enter_context(
            tempfile.TemporaryDirectory(prefix="lo_profile_fr_", ignore_cleanup_errors=True)
        )
        full = [binp, "-env:UserInstallation=%s" % Path(profile).as_uri()] + list(args)
        return subprocess.run(full, env=_env(), capture_output=True, timeout=timeout)


def convert(input_path, out_format, out_dir):
    """
    Convert input_path to out_format ('docx' or 'doc') into out_dir.
    Returns the output path, or raises on failure.
    """
    os.makedirs(out_dir, exist_ok=True)
    filt = {"docx": "MS Word 2007 XML", "doc": "MS Word 97"}.get(out_format, out_format)
    res = run(["--headless", "--convert-to", "%s:%s" % (out_format, filt)
               if filt != out_format else out_format,
               "--outdir", out_dir, input_path])
    stem = Path(input_path).stem
    out = Path(out_dir) / (stem + "." + out_format)
    if not out.exists():
        raise RuntimeError(
            "conversion to %s failed: %s"
            % (out_format, (res.stderr or b"").decode("utf-8", "replace")[:500])
        )
    return str(out)


if __name__ == "__main__":
    import json
    import sys
    print(json.dumps({"soffice_available": soffice_available(), "bin": _bin()}))
