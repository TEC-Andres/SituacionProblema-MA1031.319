#!/usr/bin/env python3
"""
texCompiler.py — High-Performance LaTeX Compiler
══════════════════════════════════════════════════════════════════════════════
  ✦ Parallel compilation across multiple projects   (ThreadPoolExecutor)
  ✦ R / Pandoc-style citations  [@key] / [@k1; @k2] →  \\cite{key}
  ✦ Suppressed-author citations [-@key]             →  \\citeyear{key}
  ✦ Hardcoded APA 7th edition CSL                   (cached from CSL repo)
  ✦ APA biblatex preamble auto-injection            (style=apa, backend=biber)
  ✦ SHA-256 hash build cache                        (skip unchanged projects)
  ✦ Structured log parsing                          (errors · warnings · OK)

Usage
─────
  python texCompiler.py [options] [folder] [main.tex] [output.pdf]

Options
  -F <folder>      Project folder (alternative to positional)
  -o <file.pdf>    Output PDF path / name
  -I bib           Ignore bibliography (skip biber)
  -j <N>           Parallel worker threads  (default: CPU count)
  -f / --force     Force recompilation even if cache is fresh
  --no-inject      Skip APA preamble injection
  --no-preprocess  Skip R-style citation preprocessing

Examples
  python texCompiler.py myProject
  python texCompiler.py -F myProject main.tex report.pdf -j 4
  python texCompiler.py myProject -I bib -f
"""

from __future__ import annotations

import hashlib
import json
import os
import re
import shutil
import subprocess
import sys
import threading
import time
import urllib.request
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from typing import Dict, List, Optional, Tuple

# ══════════════════════════════════════════════════════════════════════════════
# § 0  WINDOWS ANSI BOOTSTRAP
#      Must run before any coloured output is produced.
#
#      Strategy (in order):
#        1. On Windows, call SetConsoleMode via ctypes to enable VT100 —
#           this works on Windows 10 1511+ with no extra packages.
#        2. If ctypes activation fails or the terminal isn't a real console,
#           try colorama (installed separately: pip install colorama).
#        3. If neither is available, strip all ANSI codes at print-time so
#           output is still readable in plain-text form (CI logs, etc.).
# ══════════════════════════════════════════════════════════════════════════════

def _enable_windows_vt100() -> bool:
    """
    Enable Virtual Terminal Processing on the Windows console that owns
    stdout.  Returns True on success, False if not on Windows or failed.
    """
    if sys.platform != "win32":
        return False
    try:
        import ctypes
        import ctypes.wintypes

        ENABLE_VIRTUAL_TERMINAL_PROCESSING = 0x0004
        INVALID_HANDLE_VALUE               = ctypes.wintypes.HANDLE(-1).value

        kernel32   = ctypes.windll.kernel32                       # type: ignore[attr-defined]
        STD_OUTPUT = kernel32.GetStdHandle(-11)                   # STD_OUTPUT_HANDLE

        if STD_OUTPUT == INVALID_HANDLE_VALUE:
            return False

        mode = ctypes.wintypes.DWORD(0)
        if not kernel32.GetConsoleMode(STD_OUTPUT, ctypes.byref(mode)):
            return False

        new_mode = mode.value | ENABLE_VIRTUAL_TERMINAL_PROCESSING
        return bool(kernel32.SetConsoleMode(STD_OUTPUT, new_mode))
    except Exception:
        return False


def _try_colorama() -> bool:
    """
    Initialise colorama's ANSI → Win32 console-API translation layer.
    Returns True if colorama is available, False otherwise.
    """
    try:
        import colorama  # type: ignore[import]
        colorama.init(autoreset=False, strip=False, convert=True)
        return True
    except ModuleNotFoundError:
        return False


def _setup_ansi() -> bool:
    """
    Activate ANSI colour support on the current terminal.
    Returns True when colours will work, False when they are stripped.
    """
    # Non-Windows terminals support ANSI natively — nothing to do.
    if sys.platform != "win32":
        return True

    # Prefer the native Win10+ VT100 path (zero dependencies).
    if _enable_windows_vt100():
        return True

    # Fall back to colorama's Win32-console translation layer.
    if _try_colorama():
        return True

    # Last resort: disable colours so raw escape codes don't pollute output.
    return False


# Module-level flag: True → emit ANSI codes; False → strip them.
_ANSI_ENABLED: bool = _setup_ansi()


def _strip_ansi(text: str) -> str:
    """Remove all ANSI escape sequences from *text*."""
    return re.sub(r"\033\[[0-9;]*m", "", text)


# ══════════════════════════════════════════════════════════════════════════════
# § 1  LOGGER  — thread-safe, colour-coded output
# ══════════════════════════════════════════════════════════════════════════════

class Log:
    """Thread-safe, ANSI-coloured logger with structured severity levels."""

    _lock = threading.Lock()

    # ANSI codes — used only when _ANSI_ENABLED is True.
    RESET   = "\033[0m";  BOLD    = "\033[1m";  DIM     = "\033[2m"
    RED     = "\033[31m"; GREEN   = "\033[32m"; YELLOW  = "\033[33m"
    CYAN    = "\033[36m"; MAGENTA = "\033[35m"; WHITE   = "\033[37m"

    # Cached no-op strings used when ANSI is unavailable.
    _NOP = ""

    @classmethod
    def _c(cls, code: str) -> str:
        """Return *code* if ANSI is enabled, else an empty string."""
        return code if _ANSI_ENABLED else cls._NOP

    @classmethod
    def _emit(cls, colour: str, tag: str, msg: str, prefix: str = "") -> None:
        line = (
            f"{cls._c(colour)}{cls._c(cls.BOLD)}[{tag}]{cls._c(cls.RESET)}"
            f" {prefix}{msg}"
        )
        with cls._lock:
            print(line, flush=True)

    @classmethod
    def info(cls, msg: str, prefix: str = "")    -> None: cls._emit(cls.CYAN,    "INFO ",  msg, prefix)
    @classmethod
    def warn(cls, msg: str, prefix: str = "")    -> None: cls._emit(cls.YELLOW,  "WARN ",  msg, prefix)
    @classmethod
    def error(cls, msg: str, prefix: str = "")   -> None: cls._emit(cls.RED,     "ERROR",  msg, prefix)
    @classmethod
    def success(cls, msg: str, prefix: str = "") -> None: cls._emit(cls.GREEN,   " OK  ",  msg, prefix)
    @classmethod
    def dim(cls, msg: str, prefix: str = "") -> None:
        line = f"{cls._c(cls.DIM)}{prefix}{msg}{cls._c(cls.RESET)}"
        with cls._lock:
            print(line, flush=True)

    @classmethod
    def section(cls, msg: str) -> None:
        bar  = "\u2550" * 66
        line = (
            f"\n{cls._c(cls.MAGENTA)}{cls._c(cls.BOLD)}"
            f"{bar}\n  {msg}\n{bar}"
            f"{cls._c(cls.RESET)}\n"
        )
        with cls._lock:
            print(line, flush=True)

    @classmethod
    def summary(cls, passed: int, failed: int, elapsed: float) -> None:
        colour = cls.GREEN if failed == 0 else cls.RED
        total  = passed + failed
        bar    = "\u2500" * 66
        line   = (
            f"\n{cls._c(colour)}{cls._c(cls.BOLD)}{bar}\n"
            f"  Build summary: {passed}/{total} succeeded, "
            f"{failed} failed  [{elapsed:.3f}s]\n"
            f"{bar}{cls._c(cls.RESET)}\n"
        )
        with cls._lock:
            print(line, flush=True)


# ══════════════════════════════════════════════════════════════════════════════
# § 2  APA CSL PROVIDER
# ══════════════════════════════════════════════════════════════════════════════

_APA_CSL_URL        = None
# Use the supplied local APA CSL file at src/bib/apa.csl (cached from CSL repo on 2024-06-01)
_APA_CSL_CACHE_PATH = Path(__file__).resolve().parent.parent / "src" / "bib" / "apa.csl"

def ensure_apa_csl() -> Optional[Path]:
    """
    Ensure the official APA 7th-edition CSL file is present locally.
    Downloaded once from the canonical CSL repository and cached at
    .texcache/apa.csl next to this script.

    Returns the local Path on success, None on failure.
    """
    if _APA_CSL_CACHE_PATH.exists():
        return _APA_CSL_CACHE_PATH

    _APA_CSL_CACHE_PATH.parent.mkdir(parents=True, exist_ok=True)
    Log.info(f"Fetching APA CSL from CSL repository…")
    try:
        urllib.request.urlretrieve(_APA_CSL_URL, _APA_CSL_CACHE_PATH)
        Log.success(f"APA CSL cached → {_APA_CSL_CACHE_PATH}")
        return _APA_CSL_CACHE_PATH
    except Exception as exc:
        Log.warn(f"Could not fetch APA CSL ({exc}). Pandoc bibliography pass will be skipped.")
        return None


# ══════════════════════════════════════════════════════════════════════════════
# § 3  CITATION PREPROCESSOR
#     Converts R / Pandoc-style references to LaTeX \cite commands.
#
#     Patterns recognised
#       [@key]            →  \cite{key}
#       [@k1; @k2; @k3]  →  \cite{k1,k2,k3}
#       [-@key]           →  \citeyear{key}   (suppress author)
# ══════════════════════════════════════════════════════════════════════════════

# Matches one or more @keys inside brackets (with optional semicolons)
_CITE_BLOCK_RE   = re.compile(r'\[(-?)@([\w\d_:.\-]+(?:\s*;\s*@[\w\d_:.\-]+)*)\]')
# Extracts individual keys from a multi-key block
_CITE_KEY_RE     = re.compile(r'@([\w\d_:.\-]+)')


def preprocess_citations(tex: str) -> Tuple[str, int]:
    """
    Replace R-style citation syntax with LaTeX \\cite commands.

    Returns
    -------
    (processed_text, citation_count)
    """
    count = 0

    def _replace(m: re.Match) -> str:
        nonlocal count
        suppress_author = m.group(1) == '-'
        keys = _CITE_KEY_RE.findall(m.group(0))
        if not keys:
            return m.group(0)
        count += 1
        key_str = ','.join(keys)
        return r'\citeyear{' + key_str + '}' if suppress_author else r'\cite{' + key_str + '}'

    return _CITE_BLOCK_RE.sub(_replace, tex), count


# ══════════════════════════════════════════════════════════════════════════════
# § 4  APA PREAMBLE INJECTOR
#     Silently adds biblatex-apa setup after \documentclass if biblatex is
#     not already loaded in the document.
# ══════════════════════════════════════════════════════════════════════════════

_APA_PREAMBLE_BLOCK = (
    "\n"
    "% ── APA bibliography — auto-injected by texCompiler ────────────────\n"
    r"\usepackage[style=apa,backend=biber,natbib=true,sortcites=true]{biblatex}" + "\n"
    r"\DeclareLanguageMapping{american}{american-apa}" + "\n"
    "% Ensure common title fields are emphasized (italics) — overrides CSL output\n"
    r"\DeclareFieldFormat{title}{\textit{#1}}" + "\n"
    r"\DeclareFieldFormat{booktitle}{\textit{#1}}" + "\n"
    r"\DeclareFieldFormat{journaltitle}{\textit{#1}}" + "\n"
    "% ─────────────────────────────────────────────────────────────────────\n"
)

_DOCCLASS_RE = re.compile(
    r'(\\documentclass(?:\s*\[[^\]]*\])?\s*\{[^}]+\})',
    re.DOTALL,
)


def inject_apa_preamble(tex: str) -> str:
    """
    Insert biblatex APA settings after \\documentclass if biblatex is not
    already present.  Returns the (possibly unchanged) tex string.
    """
    if 'biblatex' in tex:
        return tex  # already configured; don't clobber

    m = _DOCCLASS_RE.search(tex)
    if not m:
        Log.warn("Could not locate \\documentclass — APA preamble not injected.")
        return tex

    pos = m.end()
    return tex[:pos] + _APA_PREAMBLE_BLOCK + tex[pos:]


# ══════════════════════════════════════════════════════════════════════════════
# § 5  BUILD CACHE
#     Fingerprints all .tex and .bib files under a project directory using
#     SHA-256.  Stores the fingerprint in .texbuildcache.json so subsequent
#     runs can skip unchanged projects entirely.
# ══════════════════════════════════════════════════════════════════════════════

class BuildCache:
    """SHA-256 hash-based incremental build cache for a single project."""

    _CACHE_FILE  = ".texbuildcache.json"
    _SOURCE_EXTS = {".tex", ".bib", ".cls", ".sty"}

    def __init__(self, project_dir: Path) -> None:
        self.project_dir = project_dir
        self._cache_file = project_dir / self._CACHE_FILE
        self._data: Dict = self._load()

    # ── persistence ──────────────────────────────────────────────────────────

    def _load(self) -> Dict:
        if self._cache_file.exists():
            try:
                return json.loads(self._cache_file.read_text(encoding="utf-8"))
            except (json.JSONDecodeError, OSError):
                pass
        return {}

    def _save(self) -> None:
        try:
            self._cache_file.write_text(
                json.dumps(self._data, indent=2),
                encoding="utf-8",
            )
        except OSError as exc:
            Log.warn(f"Cannot write build cache: {exc}")

    # ── fingerprinting ───────────────────────────────────────────────────────

    def fingerprint(self) -> str:
        """Compute a combined SHA-256 over all source files in the project."""
        h = hashlib.sha256()
        sources = sorted(
            p for p in self.project_dir.rglob("*")
            if p.suffix in self._SOURCE_EXTS and not p.name.startswith(".")
        )
        for path in sources:
            try:
                # Hash relative path + contents so renames are detected
                h.update(str(path.relative_to(self.project_dir)).encode())
                h.update(path.read_bytes())
            except OSError:
                pass
        return h.hexdigest()

    # ── public interface ─────────────────────────────────────────────────────

    def is_fresh(self, output_pdf: Path) -> bool:
        """
        Return True iff the stored fingerprint matches the current source
        tree AND the output PDF already exists.
        """
        if not output_pdf.exists():
            return False
        stored = self._data.get("fingerprint")
        return bool(stored) and stored == self.fingerprint()

    def update(self) -> None:
        """Persist the current fingerprint so the next run can use the cache."""
        self._data = {
            "fingerprint": self.fingerprint(),
            "timestamp":   time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime()),
        }
        self._save()

    def invalidate(self) -> None:
        """Clear the cache (force recompilation next time)."""
        self._data = {}
        if self._cache_file.exists():
            self._cache_file.unlink(missing_ok=True)


# ══════════════════════════════════════════════════════════════════════════════
# § 6  LOG PARSER
#     Parses pdflatex / biber stdout and classifies lines as errors, warnings
#     or informational messages.
# ══════════════════════════════════════════════════════════════════════════════

class LogParser:
    """Extracts and categorises diagnostic messages from LaTeX/Biber logs."""

    # Lines that start with these tokens are never "real" warnings
    _WARNING_BLOCKLIST = frozenset({
        "Package biblatex Warning: Please (re)run Biber",
        "Package biblatex Warning: No driver for entry type",
    })

    @classmethod
    def parse(cls, log_text: str) -> Dict[str, List[str]]:
        result: Dict[str, List[str]] = {"errors": [], "warnings": [], "info": []}
        lines   = log_text.splitlines()
        i       = 0

        while i < len(lines):
            line = lines[i]

            # ── Errors: lines starting with "!" ───────────────────────────
            if line.startswith("!"):
                block = [line]
                i += 1
                while i < len(lines) and len(block) < 6:
                    sub = lines[i]
                    if sub.startswith("!") or (sub.startswith("l.") and block):
                        break
                    if sub.strip():
                        block.append(sub.strip())
                    i += 1
                result["errors"].append("  ".join(block))
                continue

            # ── Warnings ──────────────────────────────────────────────────
            lower = line.lower()
            if "warning" in lower and not line.startswith("%"):
                stripped = line.strip()
                if not any(stripped.startswith(b) for b in cls._WARNING_BLOCKLIST):
                    result["warnings"].append(stripped)

            # ── Biber errors (start with "ERROR") ─────────────────────────
            elif line.startswith("ERROR"):
                result["errors"].append(line.strip())

            i += 1

        return result

    @classmethod
    def report(cls, parsed: Dict[str, List[str]], project: str, prefix: str = "") -> None:
        errors   = parsed["errors"]
        warnings = parsed["warnings"]
        _MAX     = 5

        if errors:
            Log.error(
                f"[{project}] {len(errors)} error(s) detected in LaTeX log:", prefix
            )
            for e in errors[:_MAX]:
                Log.dim(f"    {e}", prefix)
            if len(errors) > _MAX:
                Log.dim(f"    … and {len(errors) - _MAX} more — see __logs__/", prefix)

        if warnings:
            Log.warn(
                f"[{project}] {len(warnings)} warning(s) detected in LaTeX log:", prefix
            )
            for w in warnings[:_MAX]:
                Log.dim(f"    {w}", prefix)
            if len(warnings) > _MAX:
                Log.dim(f"    … and {len(warnings) - _MAX} more — see __logs__/", prefix)

        if not errors and not warnings:
            Log.info(f"[{project}] Log clean — no errors or warnings.", prefix)


# ══════════════════════════════════════════════════════════════════════════════
# § 7  LATEX PROJECT COMPILER
#     Compiles a single project directory.
# ══════════════════════════════════════════════════════════════════════════════

class LatexProjectCompiler:
    """
    Compiles one LaTeX project.

    Pipeline
    ────────
    1. Check build cache — skip if sources unchanged.
    2. Read main .tex, preprocess R-style citations, inject APA preamble.
    3. Write processed content back to the .tex file (backed up first).
    4. Run pdflatex + biber + pdflatex + pdflatex (or abbreviated if --ignore-bib).
    5. Parse combined log output; report errors / warnings.
    6. Move PDF to output/; move aux/log files to __logs__/.
    7. Update build cache on clean success.
    8. Restore original .tex from backup.
    """

    _AUX_EXTS = ("aux", "log", "out", "toc")
    _BIB_EXTS = ("bcf", "bbl", "blg", "run.xml")

    def __init__(
        self,
        project_name:       str           = "document",
        main_tex_name:      str           = "examGenerator.tex",
        pdf_name:           Optional[str] = None,
        ignore_bibliography: bool         = False,
        force:              bool          = False,
        inject_apa:         bool          = True,
        preprocess_cites:   bool          = True,
    ) -> None:
        self.workspace_dir       = Path.cwd()
        self.project_dir         = self.workspace_dir / project_name
        self.output_dir          = self.workspace_dir / "output"
        self.logs_dir            = self.workspace_dir / "__logs__"
        self.prefix              = f"[{project_name}] "

        # Resolve main .tex path
        raw = Path(main_tex_name)
        self.main_tex = raw if raw.is_absolute() else self.project_dir / raw

        self.pdf_name            = pdf_name or f"{project_name}.pdf"
        self.pdf_path            = self.project_dir / self.pdf_name
        self.output_pdf_path     = self.output_dir  / self.pdf_name
        self.ignore_bibliography = ignore_bibliography
        self.force               = force
        self.inject_apa          = inject_apa
        self.preprocess_cites    = preprocess_cites
        self.cache               = BuildCache(self.project_dir)

    # ── internal helpers ──────────────────────────────────────────────────────

    def _ensure_dirs(self) -> None:
        self.output_dir.mkdir(parents=True, exist_ok=True)
        self.logs_dir.mkdir(parents=True, exist_ok=True)

    def _run(self, cmd: List[str]) -> subprocess.CompletedProcess:
        return subprocess.run(
            cmd,
            cwd=self.project_dir,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
        )

    def _pdflatex(self) -> subprocess.CompletedProcess:
        return self._run([
            "pdflatex",
            "-interaction=nonstopmode",
            "-output-directory", str(self.project_dir),
            str(self.main_tex),
        ])

    def _biber(self) -> subprocess.CompletedProcess:
        return self._run([
            "biber",
            "--input-directory",  str(self.project_dir),
            "--output-directory", str(self.project_dir),
            self.main_tex.stem,
        ])

    # ── pre-processing ────────────────────────────────────────────────────────

    def _load_and_transform(self) -> Optional[Tuple[str, str]]:
        """
        Read main .tex, apply citation preprocessing and APA injection.

        Returns
        -------
        (original_content, processed_content)  or  None on I/O error.
        """
        if not self.main_tex.exists():
            Log.error(f"Main .tex not found: {self.main_tex}", self.prefix)
            return None

        original = self.main_tex.read_text(encoding="utf-8", errors="replace")
        processed = original

        if self.preprocess_cites:
            processed, n = preprocess_citations(processed)
            if n:
                Log.info(f"Preprocessed {n} R-style citation(s) → \\cite{{…}}", self.prefix)

        if self.inject_apa:
            processed = inject_apa_preamble(processed)

        return original, processed

    def _prepare_bib_titles(self, tex_content: str) -> List[Tuple[Path, Path]]:
        """
        Find \addbibresource entries in *tex_content*, back up each referenced
        .bib file, and wrap bare `title = {...}` values with
        `{{\textit{...}}}` so titles are typeset in italics.

        Returns a list of (bib_path, backup_path) for files modified so the
        caller can restore them later.
        """
        bibs: List[Tuple[Path, Path]] = []
        # Find \addbibresource{...} occurrences
        for m in re.finditer(r'\\addbibresource\{([^}]+)\}', tex_content):
            rel = m.group(1).strip()
            bib_path = Path(rel) if Path(rel).is_absolute() else (self.project_dir / rel)
            if not bib_path.exists():
                # try common alternative: project has no bib but workspace src/bib
                alt = Path(__file__).resolve().parent.parent / 'src' / 'bib' / bib_path.name
                if alt.exists():
                    bib_path = alt
                else:
                    continue

            backup = bib_path.with_suffix(bib_path.suffix + '.bak')
            try:
                shutil.copy2(bib_path, backup)
            except OSError:
                continue

            text = bib_path.read_text(encoding='utf-8', errors='replace')

            # Replace title = { ... } occurrences unless they already contain \textit
            def _repl(m: re.Match) -> str:
                prefix = m.group(1)
                openb = m.group(2)
                content = m.group(3).strip()
                closeb = m.group(4)
                tail = m.group(5)
                if '\\textit' in content or '\\mkbibemph' in content:
                    return m.group(0)
                # produce double-braced \textit wrapper: {{\textit{...}}}
                return f"{prefix}{{{{\\textit{{{content}}}}}}}{tail}"

            new_text, nsubs = re.subn(
                r'(title\s*=\s*)(\{+)(.*?)(\}+)(\s*,)',
                _repl,
                text,
                flags=re.IGNORECASE | re.DOTALL,
            )

            if nsubs:
                try:
                    bib_path.write_text(new_text, encoding='utf-8')
                    bibs.append((bib_path, backup))
                except OSError:
                    # restore backup if write failed
                    try:
                        shutil.move(str(backup), str(bib_path))
                    except OSError:
                        pass

        return bibs

    # ── compilation passes ────────────────────────────────────────────────────

    def _compile(self) -> Tuple[bool, str]:
        """
        Execute the full pdflatex (+ biber) compilation sequence.

        Returns
        -------
        (error_occurred, combined_log_text)
        """
        logs: List[str]    = []
        error_occurred     = False
        bcf_path           = self.project_dir / f"{self.main_tex.stem}.bcf"

        def run_pdflatex() -> None:
            Log.dim(f"  → pdflatex pass…", self.prefix)
            r = self._pdflatex()
            logs.append(r.stdout or "")
            if r.returncode != 0:
                raise subprocess.CalledProcessError(r.returncode, "pdflatex", r.stdout)

        def run_biber() -> None:
            Log.dim(f"  → biber pass…", self.prefix)
            r = self._biber()
            logs.append(r.stdout or "")
            if r.returncode != 0:
                raise subprocess.CalledProcessError(r.returncode, "biber", r.stdout)

        try:
            if self.ignore_bibliography:
                Log.info("Bibliography disabled (--ignore-bib): 2 pdflatex passes.", self.prefix)
                run_pdflatex()
                run_pdflatex()
            else:
                # Pass 1 — generate aux / bcf
                run_pdflatex()

                if bcf_path.exists():
                    # biblatex detected on first pass
                    run_biber()
                    run_pdflatex()
                    run_pdflatex()
                else:
                    # Pass 2 — some setups emit .bcf only after a second pass
                    run_pdflatex()
                    if bcf_path.exists():
                        run_biber()
                        run_pdflatex()
                        run_pdflatex()

        except subprocess.CalledProcessError as exc:
            Log.warn(
                f"'{exc.cmd[0]}' exited {exc.returncode} — checking for partial PDF…",
                self.prefix,
            )
            if exc.output:
                logs.append(exc.output)
            error_occurred = True

        return error_occurred, "\n".join(logs)

    # ── post-compilation file moves ───────────────────────────────────────────

    def _move_pdf(self, error_occurred: bool) -> bool:
        if self.pdf_path.exists():
            shutil.move(str(self.pdf_path), str(self.output_pdf_path))
            if error_occurred:
                Log.warn(f"PDF produced despite errors → {self.output_pdf_path}", self.prefix)
            else:
                Log.success(f"PDF → {self.output_pdf_path}", self.prefix)
            return True
        Log.error("No PDF produced — check __logs__/ for the full .log file.", self.prefix)
        return False

    def _move_aux(self) -> None:
        stem  = self.pdf_name.rsplit(".", 1)[0]
        names = [f"{stem}.{ext}" for ext in self._AUX_EXTS]
        names += [f"{self.main_tex.stem}.{ext}" for ext in self._BIB_EXTS]
        for fname in names:
            src = self.project_dir / fname
            dst = self.logs_dir   / fname
            if src.exists():
                shutil.move(str(src), str(dst))

    # ── public API ────────────────────────────────────────────────────────────

    def run(self) -> bool:
        """
        Compile the project.  Returns True on success (PDF produced).
        Thread-safe: each compiler instance is independent.
        """
        self._ensure_dirs()

        # ── cache check ───────────────────────────────────────────────────
        if not self.force and self.cache.is_fresh(self.output_pdf_path):
            Log.info(
                "Sources unchanged — skipping compilation. (-f to force rebuild)",
                self.prefix,
            )
            return True

        Log.section(f"Compiling: {self.project_dir.name}")

        # ── pre-processing backup ─────────────────────────────────────────
        orig_questions = self.project_dir / "examples" / "questions.tex"
        bak_questions  = orig_questions.with_suffix(".tex.bak")
        if orig_questions.exists():
            shutil.copy2(orig_questions, bak_questions)

        original_content: Optional[str] = None
        success = False

        try:
            transformed = self._load_and_transform()
            if transformed is None:
                return False
            original_content, processed_content = transformed

            # Prepare bibliography files: back them up and wrap title fields.
            bib_backups: List[Tuple[Path, Path]] = self._prepare_bib_titles(processed_content)

            # Overwrite main .tex with preprocessed version (restored in finally)
            self.main_tex.write_text(processed_content, encoding="utf-8")

            # ── compile ───────────────────────────────────────────────────
            t0 = time.perf_counter()
            error_occurred, log_text = self._compile()
            elapsed = time.perf_counter() - t0
            Log.dim(f"  Compilation wall-time: {elapsed:.2f}s", self.prefix)

            # ── log diagnostics ───────────────────────────────────────────
            parsed = LogParser.parse(log_text)
            LogParser.report(parsed, self.project_dir.name, self.prefix)

            # ── move outputs ──────────────────────────────────────────────
            success = self._move_pdf(error_occurred)
            self._move_aux()

            # Only update cache when compilation was completely clean
            if success and not error_occurred and not parsed["errors"]:
                self.cache.update()
                Log.dim("  Build cache updated.", self.prefix)

        finally:
            # Restore main .tex regardless of success / failure
            if original_content is not None:
                try:
                    self.main_tex.write_text(original_content, encoding="utf-8")
                except OSError as exc:
                    Log.error(f"Could not restore original .tex: {exc}", self.prefix)

            # Restore questions.tex
            if bak_questions.exists():
                shutil.move(str(bak_questions), str(orig_questions))
            # Restore any .bib backups we created
            try:
                for bib_path, backup_path in bib_backups:
                    if backup_path.exists():
                        shutil.move(str(backup_path), str(bib_path))
            except Exception:
                pass

        return success


# ══════════════════════════════════════════════════════════════════════════════
# § 8  PARALLEL COMPILER
#     Runs multiple LatexProjectCompiler instances concurrently using a
#     thread pool.  I/O-bound pdflatex processes benefit immediately from
#     parallelism even on a single CPU.
# ══════════════════════════════════════════════════════════════════════════════

class ParallelCompiler:
    """
    Compiles several LaTeX projects concurrently.

    Parameters
    ----------
    projects : list of dicts
        Each dict is forwarded as keyword arguments to LatexProjectCompiler.
    max_workers : int, optional
        Thread-pool size.  Defaults to min(project_count, cpu_count).
    """

    def __init__(
        self,
        projects:    List[Dict],
        max_workers: Optional[int] = None,
    ) -> None:
        self.projects    = projects
        self.max_workers = max_workers or min(len(projects), os.cpu_count() or 4)

    def run(self) -> Dict[str, bool]:
        """
        Compile all projects in parallel.

        Returns
        -------
        {project_name: success_bool}
        """
        if not self.projects:
            Log.warn("ParallelCompiler: no projects provided.")
            return {}

        n = len(self.projects)
        Log.section(
            f"Parallel build: {n} project(s) — {self.max_workers} worker thread(s)"
        )

        results: Dict[str, bool] = {}

        with ThreadPoolExecutor(max_workers=self.max_workers) as pool:
            futures = {
                pool.submit(LatexProjectCompiler(**cfg).run): cfg.get("project_name", "?")
                for cfg in self.projects
            }

            for future in as_completed(futures):
                name = futures[future]
                try:
                    ok = future.result()
                except Exception as exc:
                    Log.error(f"Unhandled exception in '{name}': {exc}")
                    ok = False

                results[name] = ok
                (Log.success if ok else Log.error)(
                    f"Project '{name}' {'completed' if ok else 'FAILED'}."
                )

        return results


# ══════════════════════════════════════════════════════════════════════════════
# § 9  CLI ARG PARSER
# ══════════════════════════════════════════════════════════════════════════════

class ArgC:
    """Minimal positional + flag CLI parser (no external dependencies)."""

    _FLAGS_WITH_VALUE = {"-o", "-I", "-F", "-j"}

    def __init__(self, args: List[str]) -> None:
        self.args = args

    # ── flag accessors ────────────────────────────────────────────────────────

    def _flag_value(self, flag: str) -> Optional[str]:
        try:
            idx = self.args.index(flag)
            return self.args[idx + 1] if idx + 1 < len(self.args) else None
        except ValueError:
            return None

    def _flag_values(self, flag: str) -> List[str]:
        values, i = [], 0
        while i < len(self.args):
            if self.args[i] == flag and i + 1 < len(self.args):
                values.append(self.args[i + 1])
                i += 2
            else:
                i += 1
        return values

    def _positionals(self) -> List[str]:
        skip = False
        result = []
        for arg in self.args:
            if skip:
                skip = False
                continue
            if arg in self._FLAGS_WITH_VALUE:
                skip = True
                continue
            if not arg.startswith("-"):
                result.append(arg)
        return result

    # ── parsed values ─────────────────────────────────────────────────────────

    def output_path(self) -> Optional[str]:
        return self._flag_value("-o")

    def ignore_targets(self) -> set:
        return {v.lower() for v in self._flag_values("-I")}

    def ignore_bib(self) -> bool:
        return bool(self.ignore_targets() & {"bib", "biber", "bibliography"})

    def main_tex(self) -> Optional[str]:
        return next((a for a in self.args if a.endswith(".tex")), None)

    def pdf_name(self) -> Optional[str]:
        return next((a for a in self.args if a.endswith(".pdf")), None)

    def project_folder(self) -> Optional[str]:
        folder = self._flag_value("-F")
        if folder:
            return folder
        return next(
            (a for a in self._positionals()
             if not a.endswith(".tex") and not a.endswith(".pdf")),
            None,
        )

    def workers(self) -> Optional[int]:
        v = self._flag_value("-j")
        try:
            return int(v) if v else None
        except ValueError:
            return None

    def force(self) -> bool:
        return "-f" in self.args or "--force" in self.args

    def no_inject(self) -> bool:
        return "--no-inject" in self.args

    def no_preprocess(self) -> bool:
        return "--no-preprocess" in self.args


# ══════════════════════════════════════════════════════════════════════════════
# § 10  ENTRY POINT
# ══════════════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    t_start = time.perf_counter()
    args    = ArgC(sys.argv[1:])

    # ── Ensure APA CSL is available locally ───────────────────────────────────
    ensure_apa_csl()

    # ── Resolve compilation parameters ───────────────────────────────────────
    folder   = args.project_folder() or "document"
    pdf_name = args.output_path() or args.pdf_name() or f"{folder}.pdf"
    main_tex = args.main_tex() or "examGenerator.tex"

    compiler = LatexProjectCompiler(
        project_name        = folder,
        main_tex_name       = main_tex,
        pdf_name            = pdf_name,
        ignore_bibliography = args.ignore_bib(),
        force               = args.force(),
        inject_apa          = not args.no_inject(),
        preprocess_cites    = not args.no_preprocess(),
    )

    ok      = compiler.run()
    elapsed = time.perf_counter() - t_start

    Log.summary(passed=int(ok), failed=int(not ok), elapsed=elapsed)
    sys.exit(0 if ok else 1)