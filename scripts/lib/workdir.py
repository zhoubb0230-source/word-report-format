# -*- coding: utf-8 -*-
"""Shared helpers for locating and validating the pipeline's working directory.

All intermediate artifacts (working.docx / structure.json / shards/ / fixes.json
/ out_pkg/ / formatted.docx ...) live under ONE per-run ``<workdir>``. To keep
those files from scattering, and to keep concurrent sessions from clobbering
each other, every run gets its OWN uniquely-named ``<workdir>`` created atomically
under a single fixed temp base beneath the current working directory:

    <cwd>/.word_report_work/run_<YYYYMMDD_HHMMSS>_<unique>/

Both 05_new_workdir.py (create) and 59_cleanup.py (remove) import this module so
they agree on where the base is and what is safe to delete.
"""
import os

# Component name of the fixed temp base directory. Anything to be cleaned up
# MUST live under a directory with exactly this name — the safety guard in
# is_inside_base() keeps 59_cleanup.py from ever rmtree-ing an arbitrary path.
BASE_NAME = ".word_report_work"

# Env var lets a caller relocate the base (e.g. onto a scratch volume) without
# touching the scripts. When unset we anchor to the current working directory,
# i.e. the "工作目录下的某个临时目录" the user expects.
BASE_ENV = "WORD_REPORT_WORK_BASE"


def base_dir():
    """Absolute path of the fixed temp base that holds every run's workdir."""
    override = os.environ.get(BASE_ENV)
    if override:
        return os.path.abspath(override)
    return os.path.abspath(os.path.join(os.getcwd(), BASE_NAME))


def is_inside_base(path):
    """True if ``path`` sits under a directory named BASE_NAME (guards deletes).

    We do not require the *current* base_dir() specifically — a caller may have
    set/unset WORD_REPORT_WORK_BASE between create and cleanup — only that the
    path clearly belongs to a word_report_work base and is not the base itself.
    """
    real = os.path.realpath(path)
    parts = real.split(os.sep)
    if BASE_NAME not in parts:
        return False
    # Refuse the base directory itself; only individual run dirs may be removed.
    if parts[-1] == BASE_NAME:
        return False
    return True
