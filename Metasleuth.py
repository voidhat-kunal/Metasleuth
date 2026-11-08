from __future__ import annotations
import base64
import collections
import concurrent.futures
import csv
import datetime
import hashlib
import html
import importlib
import importlib.util
import io
import json
import math
import mmap
import os
import queue
import re
import subprocess
import sys
import threading
import webbrowser
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

APP_NAME = "MetaSleuth"
APP_VERSION = "1.0"
FROZEN = bool(getattr(sys, "frozen", False))
CONFIG_DIR = Path(os.environ.get("APPDATA") or Path.home()) / (
    "MetaSleuth" if os.name == "nt" else ".metasleuth"
)


# ============================================================================
# DEPENDENCY BOOTSTRAP (only when running as a plain .py, never inside the .exe)
# ============================================================================
def _fatal(message):
    try:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(APP_NAME, message)
        root.destroy()
    except Exception:
        print(message)
    sys.exit(1)


def _ensure_dependencies():
    if FROZEN:
        return
    needed = {
        "customtkinter": "customtkinter",
        "PIL": "pillow",
        "pypdf": "pypdf",
        "mutagen": "mutagen",
    }
    missing = [pkg for mod, pkg in needed.items() if importlib.util.find_spec(mod) is None]
    if missing:
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", *missing])
            importlib.invalidate_caches()
        except Exception as exc:
            _fatal(
                "MetaSleuth needs these Python packages:\n    "
                + " ".join(missing)
                + "\n\nInstall them with:\n    python -m pip install "
                + " ".join(missing)
                + f"\n\n({exc})"
            )
    # optional: drag & drop support (tried once, failure is fine)
    if importlib.util.find_spec("tkinterdnd2") is None:
        marker = CONFIG_DIR / ".dnd_tried"
        if not marker.exists():
            try:
                subprocess.call(
                    [sys.executable, "-m", "pip", "install", "tkinterdnd2"],
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL,
                )
                importlib.invalidate_caches()
            except Exception:
                pass
            try:
                CONFIG_DIR.mkdir(parents=True, exist_ok=True)
                marker.write_text("1")
            except Exception:
                pass


_ensure_dependencies()

import customtkinter as ctk  # noqa: E402
from PIL import ExifTags, Image, ImageChops, ImageOps, ImageStat, ImageTk  # noqa: E402

try:
    from PIL import ImageCms
except Exception:  # pragma: no cover
    ImageCms = None
try:
    import pypdf
except Exception:  # pragma: no cover
    pypdf = None
try:
    import mutagen
except Exception:  # pragma: no cover
    mutagen = None
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD

    HAS_DND = True
except Exception:  # pragma: no cover
    HAS_DND = False


# ============================================================================
# THEME / CONSTANTS
# ============================================================================
C = {  # (light, dark)
    "bg": ("#F1F5F9", "#0B1220"),
    "panel": ("#E5ECF6", "#0E1830"),
    "card": ("#FFFFFF", "#14213D"),
    "card_alt": ("#EEF2F7", "#1B2B4D"),
    "text": ("#0F172A", "#E6EDF7"),
    "muted": ("#64748B", "#93A4C0"),
    "accent": ("#1F6FEB", "#2F8BFF"),
    "accent_hover": ("#1A5FCC", "#1E6FE0"),
    "neutral": ("#CBD5E1", "#2A3B63"),
    "neutral_hover": ("#B6C2D4", "#375083"),
    "teal": ("#0D9488", "#0F9F92"),
    "teal_hover": ("#0F766E", "#0B7C72"),
}
SEV = {"HIGH": "#EF4444", "MEDIUM": "#F59E0B", "LOW": "#3B82F6", "INFO": "#64748B"}
SEV_ORDER = {"HIGH": 3, "MEDIUM": 2, "LOW": 1, "INFO": 0}
RISK_COLOR = {
    "HIGH": "#EF4444",
    "MEDIUM": "#F59E0B",
    "LOW": "#3B82F6",
    "CLEAN": "#10B981",
    "ERROR": "#9CA3AF",
    "...": "#9CA3AF",
}
MONO = "Consolas" if os.name == "nt" else ("Menlo" if sys.platform == "darwin" else "DejaVu Sans Mono")
UI_FONT = "Segoe UI" if os.name == "nt" else "Helvetica"
HEX_PAGE = 4096

IMAGE_EXT = {".jpg", ".jpeg", ".jpe", ".jfif", ".png", ".gif", ".bmp", ".tif", ".tiff", ".webp", ".heic", ".heif", ".ico", ".avif"}
MEDIA_EXT = {".mp3", ".mp4", ".m4a", ".m4v", ".wav", ".flac", ".ogg", ".oga", ".opus", ".avi", ".mkv", ".mov", ".wma", ".wmv", ".aac", ".webm", ".3gp"}
OFFICE_EXT = {".docx", ".docm", ".xlsx", ".xlsm", ".pptx", ".pptm", ".dotx", ".xltx", ".potx", ".odt", ".ods", ".odp"}
EDITOR_HINTS = (
    "photoshop", "gimp", "lightroom", "paint.net", "paintshop", "pixlr", "canva", "snapseed",
    "picsart", "affinity", "illustrator", "capture one", "darktable", "luminar", "facetune",
    "adobe", "corel", "fotor", "photoscape", "acdsee", "irfanview", "imagemagick", "krita",
)
AI_HINTS = (
    "stable diffusion", "midjourney", "dall-e", "dall·e", "firefly", "comfyui", "automatic1111",
    "novelai", "sdxl", "leonardo.ai", "openai", "imagen", "gemini", "invokeai", "dreamstudio",
)
COMPRESSED_HINTS = (
    "jpeg", "png", "gif", "webp", "zip", "rar", "7-zip", "gzip", "mp3", "mp4", "matroska", "ogg",
    "flac", "pdf", "heic", "avif", "cabinet", "quicktime", "m4a", "3gp",
)


# ============================================================================
# SETTINGS + AUDIT LOG
# ============================================================================
class Settings:
    DEFAULTS = {
        "appearance": "Dark",
        "scaling": 100,
        "recursive": True,
        "min_string_len": 6,
        "ela_quality": 90,
        "case": {"case_id": "", "examiner": "", "organization": "", "notes": ""},
    }

    def __init__(self):
        self.path = CONFIG_DIR / "settings.json"
        self.data = json.loads(json.dumps(self.DEFAULTS))
        try:
            if self.path.exists():
                loaded = json.loads(self.path.read_text(encoding="utf-8"))
                for k, v in loaded.items():
                    if k == "case" and isinstance(v, dict):
                        self.data["case"].update(v)
                    elif k in self.data:
                        self.data[k] = v
        except Exception:
            pass

    def __getitem__(self, key):
        return self.data[key]

    def __setitem__(self, key, value):
        self.data[key] = value

    def save(self):
        try:
            CONFIG_DIR.mkdir(parents=True, exist_ok=True)
            self.path.write_text(json.dumps(self.data, indent=2), encoding="utf-8")
        except Exception:
            pass


class AuditLog:
    def __init__(self):
        self.entries = []
        self.path = CONFIG_DIR / "audit.log"
        self.add("SESSION", f"{APP_NAME} {APP_VERSION} started")

    def add(self, action, detail=""):
        ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.entries.append((ts, action, detail))
        try:
            CONFIG_DIR.mkdir(parents=True, exist_ok=True)
            with open(self.path, "a", encoding="utf-8") as fh:
                fh.write(f"{ts}\t{action}\t{detail}\n")
        except Exception:
            pass

    def text(self):
        return "\n".join(f"[{ts}]  {a:<14} {d}" for ts, a, d in self.entries)


# ============================================================================
# SMALL HELPERS
# ============================================================================
def fmt_size(n):
    try:
        n = float(n)
    except Exception:
        return "Unknown"
    for unit in ("B", "KB", "MB", "GB", "TB"):
        if n < 1024 or unit == "TB":
            return f"{int(n)} B" if unit == "B" else f"{n:.2f} {unit}"
        n /= 1024.0


def fmt_dt(dt):
    return dt.strftime("%Y-%m-%d %H:%M:%S")


def safe_str(v, limit=400):
    """Turns any metadata value (bytes, rationals, tuples...) into a clean single-line string."""
    try:
        if isinstance(v, (bytes, bytearray)):
            b = bytes(v)
            if not b:
                return ""
            try:
                t = b.decode("utf-8").rstrip("\x00")
                if not all(ch in "\r\n\t" or ch.isprintable() for ch in t):
                    raise ValueError
                s = t
            except Exception:
                tail = "…" if len(b) > 16 else ""
                s = f"<{len(b)} bytes: {b[:16].hex(' ').upper()}{tail}>"
        elif isinstance(v, (tuple, list)):
            s = "(" + ", ".join(safe_str(x, 80) for x in v) + ")"
        elif hasattr(v, "numerator") and hasattr(v, "denominator") and not isinstance(v, (int, bool)):
            try:
                f = float(v)
                s = str(int(f)) if f == int(f) else f"{f:.6g}"
            except Exception:
                s = str(v)
        else:
            s = str(v)
    except Exception:
        s = repr(v)
    s = s.replace("\r\n", " ⏎ ").replace("\n", " ⏎ ").replace("\r", " ").replace("\x00", "")
    return s if len(s) <= limit else s[: limit - 1] + "…"


_PDF_DATE = re.compile(r"^D?:?(\d{4})(\d{2})?(\d{2})?(\d{2})?(\d{2})?(\d{2})?")
_STD_DATE = re.compile(r"^(\d{4})[:\-/](\d{2})[:\-/](\d{2})(?:[T\s](\d{2}):(\d{2})(?::(\d{2}))?)?")


def parse_dt(value):
    """Parses EXIF / ISO / PDF date strings into naive datetimes (timezone is ignored)."""
    if value is None:
        return None
    s = str(value).strip().strip("\x00")
    if not s:
        return None
    try:
        if s.startswith("D:"):
            m = _PDF_DATE.match(s)
            parts = [int(x) if x else d for x, d in zip(m.groups(), (1970, 1, 1, 0, 0, 0))]
            return datetime.datetime(*parts)
        m = _STD_DATE.match(s)
        if m:
            return datetime.datetime(*[int(x) if x else 0 for x in m.groups()])
    except (ValueError, TypeError, AttributeError):
        return None
    return None


def compute_hashes(path):
    md5, sha1, sha256, sha512 = hashlib.md5(), hashlib.sha1(), hashlib.sha256(), hashlib.sha512()
    with open(path, "rb") as fh:
        while True:
            chunk = fh.read(1 << 20)
            if not chunk:
                break
            md5.update(chunk)
            sha1.update(chunk)
            sha256.update(chunk)
            sha512.update(chunk)
    return {
        "MD5": md5.hexdigest().upper(),
        "SHA-1": sha1.hexdigest().upper(),
        "SHA-256": sha256.hexdigest().upper(),
        "SHA-512": sha512.hexdigest().upper(),
    }


def shannon_entropy(path, limit=16 * 1024 * 1024):
    counts = collections.Counter()
    total = 0
    with open(path, "rb") as fh:
        while total < limit:
            chunk = fh.read(min(1 << 20, limit - total))
            if not chunk:
                break
            counts.update(chunk)
            total += len(chunk)
    if not total:
        return 0.0
    return -sum((c / total) * math.log2(c / total) for c in counts.values())


def entropy_label(e):
    if e < 1.0:
        return "Very low (mostly uniform data)"
    if e < 5.0:
        return "Low-medium (text / structured data)"
    if e < 7.2:
        return "High (typical for media / documents)"
    return "Very high (compressed or encrypted)"


def dms_to_dd(dms, ref):
    d, m, s = (float(x) for x in dms)
    dd = d + m / 60.0 + s / 3600.0
    ref = ref.decode(errors="ignore") if isinstance(ref, bytes) else str(ref)
    if ref.strip().upper().startswith(("S", "W")):
        dd = -dd
    return round(dd, 6)


def _is_bmp(head):
    return head[:2] == b"BM" and len(head) >= 15 and head[6:10] == b"\x00\x00\x00\x00" and head[14] in (12, 40, 56, 108, 124)


def _is_pe(head):
    if head[:2] != b"MZ" or len(head) < 64:
        return False
    return int.from_bytes(head[0x3C:0x40], "little") <= 0x10000


SIGNATURES = [
    (b"\xff\xd8\xff", "JPEG image", {".jpg", ".jpeg", ".jpe", ".jfif"}),
    (b"\x89PNG\r\n\x1a\n", "PNG image", {".png"}),
    (b"GIF87a", "GIF image", {".gif"}),
    (b"GIF89a", "GIF image", {".gif"}),
    (b"II*\x00", "TIFF image", {".tif", ".tiff"}),
    (b"MM\x00*", "TIFF image", {".tif", ".tiff"}),
    (b"%PDF-", "PDF document", {".pdf"}),
    (b"PK\x03\x04", "ZIP / OOXML / ODF container", {".zip", ".jar", ".apk", ".epub", ".xpi", ".vsdx", ".whl", ".kmz", ".nupkg"} | OFFICE_EXT),
    (b"PK\x05\x06", "ZIP (empty)", {".zip"}),
    (b"Rar!\x1a\x07", "RAR archive", {".rar"}),
    (b"7z\xbc\xaf\x27\x1c", "7-Zip archive", {".7z"}),
    (b"\x1f\x8b", "GZIP archive", {".gz", ".tgz"}),
    (b"\x7fELF", "ELF executable", {".elf", ".so", ".o", ".bin", ".out", ""}),
    (b"ID3", "MP3 audio (ID3)", {".mp3"}),
    (b"OggS", "Ogg container (audio/video)", {".ogg", ".oga", ".ogv", ".opus"}),
    (b"fLaC", "FLAC audio", {".flac"}),
    (b"\x1aE\xdf\xa3", "Matroska / WebM video", {".mkv", ".webm", ".mka"}),
    (b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1", "OLE2 compound (legacy Office / MSI)", {".doc", ".xls", ".ppt", ".msi", ".msg", ".dot", ".xlt", ".pps", ".pub", ".vsd"}),
    (b"SQLite format 3\x00", "SQLite database", {".db", ".sqlite", ".sqlite3", ".db3"}),
    (b"{\\rtf", "RTF document", {".rtf"}),
    (b"\x00\x00\x01\x00", "Windows icon (image)", {".ico"}),
    (b"MSCF", "Cabinet archive", {".cab"}),
    (b"L\x00\x00\x00\x01\x14\x02\x00", "Windows shortcut", {".lnk"}),
]


def detect_signature(head):
    """Returns (label, set_of_expected_extensions) from the file's magic bytes, or (None, None)."""
    if _is_bmp(head):
        return "BMP image", {".bmp", ".dib"}
    if _is_pe(head):
        return "Windows executable (PE)", {".exe", ".dll", ".sys", ".scr", ".ocx", ".cpl", ".drv", ".efi", ".mui"}
    for magic, label, exts in SIGNATURES:
        if head.startswith(magic):
            return label, exts
    if head[:4] == b"RIFF" and len(head) >= 12:
        tag = head[8:12]
        if tag == b"WEBP":
            return "WebP image", {".webp"}
        if tag == b"WAVE":
            return "WAV audio", {".wav", ".wave"}
        if tag == b"AVI ":
            return "AVI video", {".avi"}
        return "RIFF container", {".wav", ".avi", ".webp", ".ani", ".rmi"}
    if len(head) >= 12 and head[4:8] == b"ftyp":
        brand = head[8:12]
        if brand in (b"heic", b"heix", b"hevc", b"hevx", b"mif1", b"msf1"):
            return "HEIC image", {".heic", ".heif", ".hif"}
        if brand in (b"avif", b"avis"):
            return "AVIF image", {".avif"}
        if brand == b"M4A ":
            return "M4A audio", {".m4a", ".mp4", ".m4b"}
        if brand == b"qt  ":
            return "QuickTime movie", {".mov", ".qt"}
        if brand.startswith(b"3g"):
            return "3GP video", {".3gp", ".3g2"}
        return "MP4 video (ISO media)", {".mp4", ".m4v", ".mov", ".m4a", ".3gp", ".f4v"}
    return None, None


def classify_kind(label, ext):
    if label:
        low = label.lower()
        if "image" in low or "icon" in low:
            return "image"
        if "pdf" in low:
            return "pdf"
        if "zip" in low:
            return "office" if ext in OFFICE_EXT else "archive"
        if any(w in low for w in ("audio", "video", "movie", "ogg", "matroska")):
            return "media"
        if "ole2" in low:
            return "ole"
    if ext in IMAGE_EXT:
        return "image"
    if ext == ".pdf":
        return "pdf"
    if ext in OFFICE_EXT:
        return "office"
    if ext in MEDIA_EXT:
        return "media"
    return "other"


# ============================================================================
# ANALYSIS RECORD
# ============================================================================
class Record:
    """Everything MetaSleuth learned about one file."""

    def __init__(self, path):
        self.path = str(path)
        self.name = Path(path).name
        self.sections = collections.OrderedDict()
        self.facts = collections.OrderedDict()
        self.flags = []
        self.events = []  # (datetime, label)
        self.hashes = {}
        self.gps = None
        self.kind = "other"
        self.detected_type = "Unknown"
        self.size = 0
        self.error = None
        self.when = {}

    def add(self, section, key, value):
        self.sections.setdefault(section, collections.OrderedDict())[str(key)] = safe_str(value)

    def fact(self, label, value):
        if value not in (None, "", "N/A"):
            self.facts[label] = safe_str(value, 200)

    def flag(self, severity, title, detail=""):
        self.flags.append({"severity": severity, "title": title, "detail": detail})

    def event(self, when, label):
        dt = when if isinstance(when, datetime.datetime) else parse_dt(when)
        if dt:
            self.events.append((dt, label))
        return dt

    @property
    def risk(self):
        if self.error and not self.sections:
            return "ERROR"
        worst = max((SEV_ORDER[f["severity"]] for f in self.flags), default=0)
        return {3: "HIGH", 2: "MEDIUM", 1: "LOW"}.get(worst, "CLEAN")

    def flat(self):
        out = collections.OrderedDict()
        for sec, vals in self.sections.items():
            for k, v in vals.items():
                out[f"{sec} › {k}"] = v
        return out

    def to_dict(self):
        return {
            "file": self.name,
            "path": self.path,
            "size_bytes": self.size,
            "detected_type": self.detected_type,
            "risk": self.risk,
            "hashes": self.hashes,
            "gps": {"latitude": self.gps[0], "longitude": self.gps[1]} if self.gps else None,
            "flags": self.flags,
            "timeline": [{"time": fmt_dt(d), "event": s} for d, s in sorted(self.events)],
            "metadata": {sec: dict(vals) for sec, vals in self.sections.items()},
            "error": self.error,
        }


# ============================================================================
# ANALYSIS ENGINE
# ============================================================================
def _jpeg_end(data):
    """Offset just past the real EOI marker of the first JPEG image, or None."""
    n, i = len(data), 2
    while i < n - 1:
        i = data.find(b"\xff", i)
        if i < 0 or i + 1 >= n:
            return None
        m = data[i + 1]
        if m == 0xFF:
            i += 1
            continue
        if m == 0xD9:
            return i + 2
        if m in (0x00, 0x01) or 0xD0 <= m <= 0xD8:
            i += 2
            continue
        if i + 3 >= n:
            return None
        i += 2 + ((data[i + 2] << 8) | data[i + 3])
    return None


def _png_end(data):
    pos, n, chunks = 8, len(data), []
    while pos + 8 <= n:
        length = int.from_bytes(data[pos:pos + 4], "big")
        ctype = data[pos + 4:pos + 8]
        chunks.append(ctype.decode("latin-1"))
        pos += 12 + length
        if ctype == b"IEND":
            return pos, chunks
    return None, chunks


class Analyzer:
    @classmethod
    def analyze(cls, path):
        rec = Record(path)
        p = Path(path)
        try:
            st = os.stat(path)
            with open(path, "rb") as fh:
                head = fh.read(64)
        except Exception as exc:
            rec.error = f"Cannot read file: {exc}"
            rec.flag("HIGH", "File could not be read", str(exc))
            return rec

        rec.size = st.st_size
        ext = p.suffix.lower()
        cls._filesystem(rec, p, st)

        label, exts = detect_signature(head)
        if label:
            rec.detected_type = label
        elif not head:
            rec.detected_type = "Empty file"
        elif b"\x00" not in head and all(b in (9, 10, 13) or 32 <= b < 127 or b >= 128 for b in head):
            rec.detected_type = "Plain text (no binary signature)"
        else:
            rec.detected_type = "Unknown binary"
        sec = "Signature & Entropy"
        rec.add(sec, "Detected type", rec.detected_type)
        rec.add(sec, "Magic bytes (hex)", head[:8].hex(" ").upper())
        if label:
            ok = ext in exts
            rec.add(sec, "Extension matches content", "Yes" if ok else f"NO - extension '{ext or '(none)'}' but content is {label}")
            if not ok:
                rec.flag(
                    "HIGH",
                    "File extension does not match its content",
                    f"The name ends in '{ext or '(none)'}' but the file signature identifies it as: {label}. "
                    "Files disguised with a wrong extension are a common concealment technique.",
                )
        rec.kind = classify_kind(label, ext)

        rec.fact("Detected type", rec.detected_type)
        rec.fact("Size", f"{fmt_size(st.st_size)} ({st.st_size:,} bytes)")

        try:
            rec.hashes = compute_hashes(path)
            for k, v in rec.hashes.items():
                rec.add("Hashes", k, v)
            rec.fact("SHA-256", rec.hashes["SHA-256"])
        except Exception as exc:
            rec.add("Hashes", "Error", exc)

        try:
            ent = shannon_entropy(path)
            rec.add(sec, "Shannon entropy", f"{ent:.4f} bits/byte - {entropy_label(ent)}")
            rec.fact("Entropy", f"{ent:.2f} / 8.00")
            known_compressed = label and any(h in label.lower() for h in COMPRESSED_HINTS)
            if ent > 7.5 and not known_compressed and st.st_size > 2048:
                rec.flag(
                    "MEDIUM",
                    "Very high entropy in a non-compressed file type",
                    f"Entropy {ent:.2f}/8. Content looks encrypted, packed or compressed even though its type "
                    "does not normally behave that way.",
                )
        except Exception as exc:
            rec.add(sec, "Entropy error", exc)

        handlers = {"image": cls._image, "pdf": cls._pdf, "office": cls._office, "media": cls._media, "archive": cls._archive}
        handler = handlers.get(rec.kind)
        if handler:
            try:
                handler(rec, path)
            except Exception as exc:
                rec.add("Analysis Notes", "Parser warning", f"{type(exc).__name__}: {exc}")

        cls._rules(rec)
        rec.flags.sort(key=lambda f: -SEV_ORDER[f["severity"]])
        return rec

    # ------------------------------------------------------------------ file system
    @classmethod
    def _filesystem(cls, rec, p, st):
        sec = "File System"
        birth = getattr(st, "st_birthtime", None)
        created = datetime.datetime.fromtimestamp(birth if birth else st.st_ctime)
        modified = datetime.datetime.fromtimestamp(st.st_mtime)
        accessed = datetime.datetime.fromtimestamp(st.st_atime)
        has_birth = bool(birth) or os.name == "nt"
        c_label = "Created" if has_birth else "Metadata changed (ctime)"
        rec.add(sec, "File name", p.name)
        rec.add(sec, "Full path", str(p.resolve()))
        rec.add(sec, "Folder", str(p.resolve().parent))
        rec.add(sec, "Extension", p.suffix or "(none)")
        rec.add(sec, "Size", f"{fmt_size(st.st_size)} ({st.st_size:,} bytes)")
        rec.add(sec, c_label, fmt_dt(created))
        rec.add(sec, "Modified", fmt_dt(modified))
        rec.add(sec, "Accessed", fmt_dt(accessed))
        rec.event(created, f"Filesystem - {c_label}")
        rec.event(modified, "Filesystem - Modified")
        rec.event(accessed, "Filesystem - Accessed")
        rec.when["fs_modified"] = modified
        if has_birth:
            rec.when["fs_created"] = created
        rec.fact(c_label, fmt_dt(created))
        rec.fact("Modified", fmt_dt(modified))
        attrs = getattr(st, "st_file_attributes", None)
        if attrs is not None:
            names = {0x1: "Read-only", 0x2: "Hidden", 0x4: "System", 0x20: "Archive", 0x100: "Temporary",
                     0x400: "Reparse point", 0x800: "Compressed", 0x4000: "Encrypted"}
            rec.add(sec, "Windows attributes", ", ".join(n for b, n in names.items() if attrs & b) or "Normal")
            if attrs & 0x2:
                rec.flag("INFO", "Hidden file", "The Windows 'Hidden' attribute is set.")
        else:
            rec.add(sec, "Permissions", oct(st.st_mode & 0o777))
        if os.path.islink(p):
            rec.add(sec, "Symbolic link target", os.readlink(p))
        # Mark of the Web (NTFS alternate data stream)
        if os.name == "nt":
            try:
                with open(str(p) + ":Zone.Identifier", "r", encoding="utf-8", errors="replace") as fh:
                    zone = fh.read()
                zid = {"0": "My Computer", "1": "Local Intranet", "2": "Trusted Sites", "3": "Internet", "4": "Restricted Sites"}
                for line in zone.splitlines():
                    if "=" in line:
                        k, v = line.split("=", 1)
                        if k.strip() == "ZoneId":
                            v = f"{v.strip()} ({zid.get(v.strip(), 'Unknown')})"
                        rec.add("Mark of the Web (Zone.Identifier)", k.strip(), v.strip())
                host = rec.sections.get("Mark of the Web (Zone.Identifier)", {}).get("HostUrl", "")
                rec.flag("INFO", "File was downloaded from the Internet (Mark of the Web)", host or "Zone.Identifier stream present.")
                rec.fact("Downloaded from", host)
            except Exception:
                pass

    # ------------------------------------------------------------------ images
    @classmethod
    def _image(cls, rec, path):
        try:
            img = Image.open(path)
        except Exception as exc:
            rec.add("Image", "Decode error", exc)
            rec.flag("LOW", "Image could not be decoded", f"{type(exc).__name__}: {exc}")
            return
        software, texts = [], []
        fmt = None
        with img:
            fmt = img.format
            w, h = img.size
            rec.add("Image", "Format", img.format)
            rec.add("Image", "Color mode", img.mode)
            rec.add("Image", "Dimensions", f"{w} x {h} px")
            rec.add("Image", "Megapixels", f"{w * h / 1e6:.2f} MP")
            rec.fact("Dimensions", f"{w} × {h} px ({w * h / 1e6:.2f} MP)")
            if img.info.get("dpi"):
                rec.add("Image", "DPI", img.info["dpi"])
            frames = getattr(img, "n_frames", 1)
            if frames and frames > 1:
                rec.add("Image", "Frames / pages", frames)

            icc = img.info.get("icc_profile")
            if icc:
                desc = ""
                if ImageCms:
                    try:
                        desc = ImageCms.getOpenProfile(io.BytesIO(icc)).profile.profile_description
                    except Exception:
                        desc = ""
                rec.add("Image", "ICC color profile", f"{(desc or '').strip()} ({len(icc)} bytes)".strip())

            exif = img.getexif()
            gps_ifd, exif_ifd = {}, {}
            try:
                exif_ifd = exif.get_ifd(0x8769) or {}
                gps_ifd = exif.get_ifd(0x8825) or {}
            except Exception:
                pass

            def tag_name(i):
                return ExifTags.TAGS.get(i, f"Tag 0x{i:04X}")

            for tid, val in exif.items():
                if tid in (0x8769, 0x8825, 0xA005):
                    continue
                name = tag_name(tid)
                rec.add("EXIF - Main (IFD0)", name, "<binary data>" if name == "MakerNote" else val)
                if name in ("Software", "HostComputer"):
                    software.append(safe_str(val))
            for tid, val in exif_ifd.items():
                name = tag_name(tid)
                if name == "MakerNote":
                    rec.add("EXIF - Camera (Exif IFD)", name, f"<{len(val) if hasattr(val, '__len__') else '?'} bytes of vendor data>")
                else:
                    rec.add("EXIF - Camera (Exif IFD)", name, val)
                if name in ("ImageDescription", "UserComment"):
                    texts.append(safe_str(val))
            main = rec.sections.get("EXIF - Main (IFD0)", {})
            camera = rec.sections.get("EXIF - Camera (Exif IFD)", {})
            make, model = main.get("Make", ""), main.get("Model", "")
            if make or model:
                rec.fact("Camera", f"{make} {model}".strip())
            if main.get("Software"):
                rec.fact("Software", main["Software"])
            if camera.get("LensModel"):
                rec.fact("Lens", camera["LensModel"])
            if camera.get("BodySerialNumber"):
                rec.flag("INFO", "Camera serial number embedded", f"Serial {camera['BodySerialNumber']} can link this photo to a specific device.")

            for key, label in (("DateTimeOriginal", "EXIF - Captured"), ("DateTimeDigitized", "EXIF - Digitized")):
                dt = rec.event(camera.get(key), label)
                if key == "DateTimeOriginal" and dt:
                    rec.when["exif_original"] = dt
                    rec.fact("Captured", fmt_dt(dt))
            dt = rec.event(main.get("DateTime"), "EXIF - Last modified")
            if dt:
                rec.when["exif_modified"] = dt

            if gps_ifd:
                for k, v in gps_ifd.items():
                    rec.add("GPS", ExifTags.GPSTAGS.get(k, f"GPS {k}"), v)
                try:
                    g = {ExifTags.GPSTAGS.get(k, k): v for k, v in gps_ifd.items()}
                    lat = dms_to_dd(g["GPSLatitude"], g["GPSLatitudeRef"])
                    lon = dms_to_dd(g["GPSLongitude"], g["GPSLongitudeRef"])
                    if -90 <= lat <= 90 and -180 <= lon <= 180:
                        rec.gps = (lat, lon)
                        rec.add("GPS", "Decimal latitude", lat)
                        rec.add("GPS", "Decimal longitude", lon)
                        rec.add("GPS", "Google Maps", f"https://www.google.com/maps?q={lat},{lon}")
                        rec.add("GPS", "OpenStreetMap", f"https://www.openstreetmap.org/?mlat={lat}&mlon={lon}#map=17/{lat}/{lon}")
                        rec.fact("GPS position", f"{lat}, {lon}")
                        rec.flag("INFO", "GPS coordinates embedded", f"Position {lat}, {lon} - privacy-sensitive location data.")
                    if "GPSAltitude" in g:
                        alt = float(g["GPSAltitude"])
                        below = g.get("GPSAltitudeRef") in (1, b"\x01")
                        rec.add("GPS", "Altitude", f"{'-' if below else ''}{alt:.1f} m")
                    if "GPSDateStamp" in g and "GPSTimeStamp" in g:
                        hh, mm, ss = (int(float(x)) for x in g["GPSTimeStamp"])
                        rec.event(f"{g['GPSDateStamp']} {hh:02d}:{mm:02d}:{ss:02d}", "GPS - satellite time (UTC)")
                except Exception as exc:
                    rec.add("GPS", "Parse note", exc)

            # container-level info (PNG text chunks, JFIF, etc.)
            skip = {"exif", "icc_profile", "xmp", "XML:com.adobe.xmp", "photoshop", "dpi"}
            for i, (k, v) in enumerate(img.info.items()):
                if k in skip or i > 60:
                    continue
                rec.add("Embedded Info (container)", k, v)
                if isinstance(v, str):
                    texts.append(v)
                    if str(k).lower() in ("software", "comment", "parameters", "prompt", "workflow"):
                        software.append(v[:200])

            xmp = img.info.get("xmp") or img.info.get("XML:com.adobe.xmp")
            if xmp:
                xt = xmp.decode("utf-8", "ignore") if isinstance(xmp, bytes) else str(xmp)
                rec.add("XMP", "XMP packet", f"present ({len(xt)} chars)")
                pat = re.compile(r"(?:xmp:|xmpMM:|photoshop:)?(CreatorTool|CreateDate|ModifyDate|MetadataDate|DocumentID|OriginalDocumentID|InstanceID)(?:=\"([^\"]*)\"|>([^<]*)<)")
                seen = set()
                for m in pat.finditer(xt):
                    key, val = m.group(1), (m.group(2) or m.group(3) or "").strip()
                    if val and (key, val) not in seen:
                        seen.add((key, val))
                        rec.add("XMP", key, val)
                        if key == "CreatorTool":
                            software.append(val)
                        if key in ("CreateDate", "ModifyDate"):
                            rec.event(val, f"XMP - {key}")
                texts.append(xt[:20000])

        # ---- deep structural checks on raw bytes
        data = None
        if rec.size <= 100 * 1024 * 1024:
            try:
                with open(path, "rb") as fh:
                    data = fh.read()
            except Exception:
                data = None
        if data:
            end = None
            if fmt == "JPEG" and data[:2] == b"\xff\xd8":
                end = _jpeg_end(data)
            elif fmt == "PNG" and data[:8] == b"\x89PNG\r\n\x1a\n":
                end, chunks = _png_end(data)
                counts = collections.Counter(chunks)
                rec.add("Image", "PNG chunks", ", ".join(f"{k}x{v}" if v > 1 else k for k, v in counts.items()))
            if end is not None and end < len(data):
                tail = data[end:]
                if not tail.strip(b"\x00"):
                    rec.flag("INFO", "Zero padding after end of image", f"{len(tail):,} null bytes follow the image data.")
                else:
                    sig, _ = detect_signature(tail[:64])
                    rec.add("Image", "Data after image end", f"{len(tail):,} bytes" + (f" - looks like {sig}" if sig else ""))
                    if sig:
                        rec.flag("HIGH", "Embedded file appended after the image", f"{len(tail):,} bytes of {sig} data follow the end of the image. This is a classic way of hiding payloads inside pictures.")
                    else:
                        rec.flag("MEDIUM", "Unexplained data appended after the image", f"{len(tail):,} bytes follow the image end marker. Can be a hidden payload, or legitimate (Motion Photo / multi-picture files).")
            if b"c2pa" in data[:8_000_000] or b"jumb" in data[:8_000_000]:
                rec.add("Image", "Content Credentials (C2PA)", "manifest marker found")
                rec.flag("INFO", "Content Credentials (C2PA) marker present", "The file carries provenance data that may describe its origin or AI generation.")

        blob = " ".join(software + texts).lower()
        hits = sorted({s for s in software if any(h in s.lower() for h in EDITOR_HINTS)})
        if hits:
            rec.flag("MEDIUM", "Image touched by editing software", "Software tags: " + "; ".join(hits)[:300])
        ai = sorted({h for h in AI_HINTS if h in blob})
        if ai:
            rec.flag("MEDIUM", "Possible AI-generated / AI-edited image metadata", "Metadata mentions: " + ", ".join(ai) + ". Indicator only - not proof.")
        if fmt == "JPEG" and not rec.sections.get("EXIF - Main (IFD0)") and not rec.sections.get("EXIF - Camera (Exif IFD)"):
            rec.flag("INFO", "JPEG contains no EXIF data", "Metadata may have been stripped (social media, messaging apps, screenshot tools, editors).")

    # ------------------------------------------------------------------ PDF
    @classmethod
    def _pdf(cls, rec, path):
        sec = "PDF Document"
        if pypdf is None:
            rec.add(sec, "Notice", "pypdf is not installed")
            return
        try:
            reader = pypdf.PdfReader(path)
            if reader.is_encrypted:
                try:
                    reader.decrypt("")
                except Exception:
                    pass
                rec.add(sec, "Encrypted", "Yes")
                rec.flag("LOW", "PDF is encrypted", "Some metadata may be unavailable without the password.")
            rec.add(sec, "PDF version", getattr(reader, "pdf_header", "?"))
            try:
                rec.add(sec, "Page count", len(reader.pages))
                box = reader.pages[0].mediabox
                rec.add(sec, "First page size", f"{float(box.width):.0f} x {float(box.height):.0f} pt")
                rec.fact("Pages", len(reader.pages))
            except Exception:
                pass
            info = reader.metadata
            if info:
                for k, v in info.items():
                    name = k[1:] if str(k).startswith("/") else str(k)
                    rec.add(sec, name, v)
                    if name in ("Producer", "Creator"):
                        rec.fact(name, v)
                    if name == "Author":
                        rec.fact("Author", v)
                c = rec.event(str(info.get("/CreationDate", "")), "PDF - Created")
                m = rec.event(str(info.get("/ModDate", "")), "PDF - Modified")
                if c:
                    rec.when["doc_created"] = c
                if m:
                    rec.when["doc_modified"] = m
                if c and m and (m - c).total_seconds() > 60:
                    rec.flag("INFO", "PDF modified after creation", f"Created {fmt_dt(c)}, modified {fmt_dt(m)}.")
                prod = " ".join(str(info.get(k, "")) for k in ("/Producer", "/Creator")).lower()
                if any(h in prod for h in EDITOR_HINTS) and "acrobat" not in prod and "distiller" not in prod:
                    rec.flag("LOW", "PDF produced by an editing tool", prod.strip()[:200])
            else:
                rec.add(sec, "Metadata", "No document information dictionary")
            try:
                ids = reader.trailer.get("/ID")
                if ids and len(ids) == 2:
                    a, b = bytes(ids[0]).hex().upper(), bytes(ids[1]).hex().upper()
                    rec.add(sec, "Document ID (original)", a)
                    rec.add(sec, "Document ID (current)", b)
                    if a != b:
                        rec.flag("INFO", "PDF document ID changed", "The file was saved again after it was first created.")
            except Exception:
                pass
        except Exception as exc:
            rec.add(sec, "Parser error", f"{type(exc).__name__}: {exc}")

        if rec.size <= 100 * 1024 * 1024:
            with open(path, "rb") as fh:
                raw = fh.read()
            counts = {
                "JavaScript": len(re.findall(rb"/JavaScript\b|/JS\b", raw)),
                "Launch actions": len(re.findall(rb"/Launch\b", raw)),
                "OpenAction": len(re.findall(rb"/OpenAction\b", raw)),
                "Additional actions (/AA)": len(re.findall(rb"/AA\b", raw)),
                "Embedded files": len(re.findall(rb"/EmbeddedFile\b", raw)),
                "URI links": len(re.findall(rb"/URI\b", raw)),
                "Forms (AcroForm/XFA)": len(re.findall(rb"/AcroForm\b|/XFA\b", raw)),
                "Revisions (%%EOF markers)": len(re.findall(rb"%%EOF", raw)),
            }
            for k, v in counts.items():
                rec.add("PDF Structure (raw scan)", k, v)
            if counts["JavaScript"]:
                rec.flag("HIGH", "PDF contains JavaScript", f"{counts['JavaScript']} JavaScript reference(s) - frequently abused by malicious PDFs.")
            if counts["Launch actions"]:
                rec.flag("HIGH", "PDF contains Launch actions", "Can start external programs when opened.")
            if counts["OpenAction"] or counts["Additional actions (/AA)"]:
                rec.flag("MEDIUM", "PDF runs actions automatically", "OpenAction / AA entries execute when the file is opened or interacted with.")
            if counts["Embedded files"]:
                rec.flag("MEDIUM", "PDF contains embedded files", f"{counts['Embedded files']} embedded file stream(s).")
            if counts["Revisions (%%EOF markers)"] > 1:
                rec.flag("LOW", "PDF has multiple revisions", f"{counts['Revisions (%%EOF markers)']} incremental saves - earlier versions of the content may still be recoverable.")

    # ------------------------------------------------------------------ Office / ODF
    @classmethod
    def _office(cls, rec, path):
        sec = "Office / ODF Document"
        try:
            zf = zipfile.ZipFile(path)
        except Exception as exc:
            rec.add(sec, "Error", f"Not a valid ZIP container: {exc}")
            rec.flag("MEDIUM", "Document container is corrupt", str(exc))
            return

        def local(tag):
            return tag.split("}")[-1]

        def read_xml(name):
            try:
                with zf.open(name) as fh:
                    return ET.fromstring(fh.read(2_000_000))
            except Exception:
                return None

        with zf:
            names = zf.namelist()
            rec.add(sec, "Container entries", len(names))
            core = read_xml("docProps/core.xml")
            labels = {
                "creator": "Author (creator)", "lastModifiedBy": "Last modified by", "created": "Created",
                "modified": "Last modified", "revision": "Revision", "title": "Title", "subject": "Subject",
                "keywords": "Keywords", "description": "Description", "category": "Category", "lastPrinted": "Last printed",
            }
            if core is not None:
                for el in core:
                    key = local(el.tag)
                    val = (el.text or "").strip()
                    if val:
                        rec.add(sec, labels.get(key, key), val)
                        if key == "creator":
                            rec.fact("Author", val)
                        if key == "lastModifiedBy":
                            rec.fact("Last modified by", val)
                        if key in ("created", "modified", "lastPrinted"):
                            dt = rec.event(val, f"Document - {labels[key]}")
                            if dt:
                                rec.when["doc_created" if key == "created" else "doc_" + key] = dt
                c, m = rec.when.get("doc_created"), rec.when.get("doc_modified")
                if c:
                    rec.fact("Document created", fmt_dt(c))
                if m:
                    rec.fact("Document modified", fmt_dt(m))
            app = read_xml("docProps/app.xml")
            if app is not None:
                for el in app:
                    key = local(el.tag)
                    val = (el.text or "").strip()
                    if val and key in ("Application", "AppVersion", "Company", "Template", "TotalTime", "Pages", "Words", "Characters", "Slides", "Manager", "Lines", "Paragraphs"):
                        rec.add(sec, key, val)
                        if key == "Application":
                            rec.fact("Application", val)
                        if key == "Company":
                            rec.fact("Company", val)
            odf = read_xml("meta.xml")
            if odf is not None:
                omap = {"initial-creator": "Original author", "creator": "Last modified by", "creation-date": "Created", "date": "Last modified",
                        "generator": "Generator", "title": "Title", "subject": "Subject", "editing-cycles": "Editing cycles", "editing-duration": "Editing duration"}
                for el in odf.iter():
                    key = local(el.tag)
                    if key in omap and (el.text or "").strip():
                        rec.add(sec, omap[key], el.text.strip())
                        if key == "initial-creator":
                            rec.fact("Author", el.text.strip())
                        if key == "creation-date":
                            rec.when["doc_created"] = rec.event(el.text, "Document - Created")
                        if key == "date":
                            rec.when["doc_modified"] = rec.event(el.text, "Document - Modified")

            lowered = [n.lower() for n in names]
            if any(n.endswith("vbaproject.bin") for n in lowered):
                rec.flag("HIGH", "Document contains VBA macros", "A vbaProject.bin part is present. Macros are a primary malware delivery vector.")
            if any("activex" in n for n in lowered):
                rec.flag("MEDIUM", "Document contains ActiveX controls", "ActiveX parts found inside the container.")
            emb = [n for n in names if "/embeddings/" in n.lower() or n.lower().startswith("embeddings/")]
            if emb:
                rec.add(sec, "Embedded objects", ", ".join(emb[:10]))
                rec.flag("MEDIUM", "Document contains embedded objects", f"{len(emb)} embedded object(s).")
            externals = []
            for n in names:
                if n.endswith(".rels"):
                    root = read_xml(n)
                    if root is None:
                        continue
                    for rel in root:
                        if rel.get("TargetMode") == "External":
                            rtype = (rel.get("Type") or "").rsplit("/", 1)[-1]
                            if rtype != "hyperlink":
                                externals.append((rtype, rel.get("Target", "")))
            for i, (rtype, target) in enumerate(externals[:20]):
                rec.add("External References", f"{rtype} #{i + 1}", target)
            if externals:
                risky = [e for e in externals if e[0] in ("attachedTemplate", "oleObject", "frame")]
                if risky:
                    rec.flag("MEDIUM", "Document loads remote content", "External " + ", ".join(sorted({r[0] for r in risky})) + " reference (remote template / OLE injection technique).")
                else:
                    rec.flag("LOW", "Document references external resources", f"{len(externals)} external relationship(s) (e.g. remote images).")

    # ------------------------------------------------------------------ other archives
    @classmethod
    def _archive(cls, rec, path):
        sec = "Archive"
        with zipfile.ZipFile(path) as zf:
            infos = zf.infolist()
            total = sum(i.file_size for i in infos)
            comp = sum(i.compress_size for i in infos)
            rec.add(sec, "Entries", len(infos))
            rec.add(sec, "Uncompressed size", fmt_size(total))
            rec.add(sec, "Compressed size", fmt_size(comp))
            enc = sum(1 for i in infos if i.flag_bits & 0x1)
            if enc:
                rec.add(sec, "Encrypted entries", enc)
                rec.flag("LOW", "Archive contains encrypted entries", f"{enc} password-protected item(s).")
            if zf.comment:
                rec.add(sec, "Archive comment", zf.comment)
            for i, info in enumerate(infos[:25]):
                rec.add("Archive Contents (first 25)", info.filename, f"{fmt_size(info.file_size)}  {fmt_dt(datetime.datetime(*info.date_time))}")
            risky = [i.filename for i in infos if i.filename.lower().endswith((".exe", ".scr", ".js", ".vbs", ".bat", ".cmd", ".ps1", ".lnk", ".jar"))]
            if risky:
                rec.flag("MEDIUM", "Archive contains executable / script files", ", ".join(risky[:8]))
            if total and comp and total / max(comp, 1) > 1000:
                rec.flag("MEDIUM", "Extreme compression ratio", "Possible decompression bomb.")

    # ------------------------------------------------------------------ audio / video
    @classmethod
    def _media(cls, rec, path):
        sec = "Audio / Video"
        if mutagen is None:
            rec.add(sec, "Notice", "mutagen is not installed")
            return
        audio = mutagen.File(path)
        if audio is None:
            rec.add(sec, "Notice", "Container not recognised by mutagen (tags unavailable)")
            return
        rec.add(sec, "Parser", type(audio).__name__)
        info = getattr(audio, "info", None)
        if info is not None:
            for attr in ("length", "bitrate", "channels", "sample_rate", "bits_per_sample", "codec", "codec_description", "mode", "version"):
                if hasattr(info, attr):
                    val = getattr(info, attr)
                    if attr == "length":
                        val = f"{int(val // 60)}m {int(val % 60)}s ({val:.2f} s)"
                        rec.fact("Duration", val)
                    elif attr == "bitrate" and isinstance(val, (int, float)):
                        val = f"{val / 1000:.1f} kbps"
                    elif attr == "sample_rate":
                        val = f"{val} Hz"
                    rec.add(sec, attr.replace("_", " ").title(), val)
        if getattr(audio, "tags", None):
            for i, (k, v) in enumerate(audio.tags.items()):
                if i > 120:
                    break
                key = str(k)
                rec.add("Tags", key, "<embedded picture>" if key.startswith(("APIC", "covr", "PIC")) else v)
            tags = rec.sections.get("Tags", {})
            for k in ("TPE1", "artist", "\xa9ART"):
                if k in tags:
                    rec.fact("Artist", tags[k])
            for k in ("TIT2", "title", "\xa9nam"):
                if k in tags:
                    rec.fact("Title", tags[k])
            for k in ("TDRC", "date", "\xa9day"):
                if k in tags:
                    rec.event(tags[k], "Media - Release / recording date")

    # ------------------------------------------------------------------ rules
    @classmethod
    def _rules(cls, rec):
        now = datetime.datetime.now()
        w = rec.when
        o, m = w.get("exif_original"), w.get("exif_modified")
        if o and m and (m - o).total_seconds() > 60:
            rec.flag("MEDIUM", "Image modified after capture", f"EXIF capture time {fmt_dt(o)}; last EXIF modification {fmt_dt(m)}.")
        fsm = w.get("fs_modified")
        if o and fsm and (o - fsm).total_seconds() > 86400:
            rec.flag("MEDIUM", "File modified time predates capture time", f"Captured {fmt_dt(o)} but the file claims to be modified {fmt_dt(fsm)}. Suggests clock manipulation or timestamp tampering.")
        dc, dm = w.get("doc_created"), w.get("doc_modified")
        if dc and dm and (dc - dm).total_seconds() > 60:
            rec.flag("MEDIUM", "Document modified date is earlier than created date", f"Created {fmt_dt(dc)}, modified {fmt_dt(dm)} - inconsistent or tampered metadata.")
        fc = w.get("fs_created")
        if fc and fsm and (fc - fsm).total_seconds() > 1:
            rec.flag("INFO", "File modified earlier than it was created", "Typical of a file that was copied or extracted from another location.")
        future = [(d, s) for d, s in rec.events if d > now + datetime.timedelta(days=1)]
        if future:
            d, s = future[0]
            rec.flag("MEDIUM", "Timestamp lies in the future", f"{s}: {fmt_dt(d)}")


# ============================================================================
# DEEP TOOLS (run on demand)
# ============================================================================
def read_hex_page(path, offset, length=HEX_PAGE):
    with open(path, "rb") as fh:
        fh.seek(offset)
        data = fh.read(length)
    lines = []
    for i in range(0, len(data), 16):
        chunk = data[i:i + 16]
        left = " ".join(f"{b:02X}" for b in chunk[:8])
        right = " ".join(f"{b:02X}" for b in chunk[8:])
        txt = "".join(chr(b) if 32 <= b < 127 else "." for b in chunk)
        lines.append(f"{offset + i:08X}  {(left + '  ' + right).ljust(48)}  |{txt}|")
    return "\n".join(lines)


def hex_find(path, needle, start=0):
    if not needle:
        return -1
    with open(path, "rb") as fh:
        try:
            mm = mmap.mmap(fh.fileno(), 0, access=mmap.ACCESS_READ)
        except (ValueError, OSError):
            return -1
        try:
            return mm.find(needle, max(0, start))
        finally:
            mm.close()


def extract_strings(path, min_len=6, max_bytes=64 * 1024 * 1024, limit=60000):
    with open(path, "rb") as fh:
        data = fh.read(max_bytes)
    min_len = max(3, min(int(min_len), 200))
    ascii_re = re.compile(rb"[\x20-\x7e]{%d,}" % min_len)
    wide_re = re.compile(rb"(?:[\x20-\x7e]\x00){%d,}" % min_len)
    found = [(m.start(), "A", m.group().decode("ascii")) for m in ascii_re.finditer(data)]
    found += [(m.start(), "U", m.group().decode("utf-16le", "ignore")) for m in wide_re.finditer(data)]
    found.sort()
    truncated = len(found) > limit
    found = found[:limit]
    lines = [f"{off:08X}  [{enc}]  {txt}" for off, enc, txt in found]
    joined = "\n".join(t for _, _, t in found)

    def uniq(rx, flags=0, validate=None, cap=5000):
        seen, out = set(), []
        for m in re.finditer(rx, joined, flags):
            v = m.group().rstrip(".,;:)")
            if v not in seen and (validate is None or validate(v)):
                seen.add(v)
                out.append(v)
                if len(out) >= cap:
                    break
        return out

    def ip_ok(v):
        return all(int(p) <= 255 for p in v.split("."))

    return {
        "All strings": lines,
        "URLs": uniq(r"https?://[^\s\"'<>\\]{4,}"),
        "Emails": uniq(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}"),
        "IP addresses": uniq(r"\b(?:\d{1,3}\.){3}\d{1,3}\b", validate=ip_ok),
        "File paths": uniq(r"[A-Za-z]:\\(?:[^\\/:*?\"<>|\r\n]+\\)*[^\\/:*?\"<>|\r\n\s]*"),
        "_meta": {"truncated": truncated, "scanned_bytes": len(data), "total": len(found)},
    }


def make_thumbnail(path, box):
    with Image.open(path) as im:
        try:
            im.draft("RGB", (box[0] * 2, box[1] * 2))
        except Exception:
            pass
        im.load()
        try:
            im = ImageOps.exif_transpose(im)
        except Exception:
            pass
        out = im.copy()
    out.thumbnail(box, Image.LANCZOS)
    return out if out.mode in ("RGB", "RGBA") else out.convert("RGBA")


def compute_ela(path, quality=90, box=(480, 480)):
    with Image.open(path) as im:
        im.load()
        orig = im.convert("RGB")
    if orig.width * orig.height > 16_000_000:
        orig.thumbnail((4000, 4000))
    buf = io.BytesIO()
    orig.save(buf, "JPEG", quality=int(quality))
    buf.seek(0)
    with Image.open(buf) as rs:
        resaved = rs.convert("RGB")
    diff = ImageChops.difference(orig, resaved)
    mx = max(e[1] for e in diff.getextrema()) or 1
    scale = 255.0 / mx
    ela = diff.point(lambda p: min(255, int(p * scale)))
    mean = sum(ImageStat.Stat(diff).mean) / 3.0
    o_view, e_view = orig.copy(), ela.copy()
    o_view.thumbnail(box, Image.LANCZOS)
    e_view.thumbnail(box, Image.LANCZOS)
    return o_view, e_view, {"max": mx, "mean": mean, "quality": int(quality)}


def strip_metadata(src, dst):
    with Image.open(src) as im:
        im.load()
        im2 = ImageOps.exif_transpose(im)
        fmt = Image.registered_extensions().get(Path(dst).suffix.lower(), im.format or "PNG")
        if fmt == "JPEG" and im2.mode not in ("RGB", "L", "CMYK"):
            im2 = im2.convert("RGB")
        clean = Image.new(im2.mode, im2.size)
        if im2.mode == "P" and im2.getpalette():
            clean.putpalette(im2.getpalette())
        clean.paste(im2)
        kwargs = {"quality": 95} if fmt in ("JPEG", "WEBP") else {}
        clean.save(dst, format=fmt, **kwargs)
    return os.path.getsize(src), os.path.getsize(dst)


# ============================================================================
# REPORT / EXPORT BUILDERS
# ============================================================================
def build_json(recs, case):
    return json.dumps(
        {
            "tool": f"{APP_NAME} {APP_VERSION}",
            "generated": datetime.datetime.now().isoformat(timespec="seconds"),
            "case": case,
            "files": [r.to_dict() for r in recs],
        },
        indent=2,
        ensure_ascii=False,
        default=str,
    )


def build_csv_rows(recs):
    rows = [["File", "Section", "Property", "Value"]]
    for r in recs:
        for f in r.flags:
            rows.append([r.name, "Forensic Indicators", f["severity"] + " - " + f["title"], f["detail"]])
        for sec, vals in r.sections.items():
            for k, v in vals.items():
                rows.append([r.name, sec, k, v])
    return rows


def build_text_report(recs, case):
    out = ["=" * 78, f"{APP_NAME} {APP_VERSION} - METADATA REPORT", "=" * 78]
    out.append(f"Generated : {datetime.datetime.now():%Y-%m-%d %H:%M:%S}")
    for k, label in (("case_id", "Case ID"), ("examiner", "Examiner"), ("organization", "Organization")):
        if case.get(k):
            out.append(f"{label:<10}: {case[k]}")
    if case.get("notes"):
        out.append(f"Notes     : {case['notes']}")
    for r in recs:
        out += ["", "#" * 78, f"FILE: {r.name}", f"PATH: {r.path}", f"RISK: {r.risk}", "#" * 78]
        if r.flags:
            out.append("\n[FORENSIC INDICATORS]")
            for f in r.flags:
                out.append(f"  [{f['severity']}] {f['title']}\n         {f['detail']}")
        for sec, vals in r.sections.items():
            out.append(f"\n[{sec.upper()}]")
            out.append("-" * 50)
            for k, v in vals.items():
                out.append(f"{k:<34}: {v}")
        if r.events:
            out.append("\n[TIMELINE]")
            for d, s in sorted(r.events):
                out.append(f"{fmt_dt(d)}  {s}")
    return "\n".join(out)


def build_timeline_rows(recs):
    rows = [["Time", "Event", "File", "Path"]]
    ev = [(d, s, r.name, r.path) for r in recs for d, s in r.events]
    for d, s, n, p in sorted(ev, key=lambda e: e[0]):
        rows.append([fmt_dt(d), s, n, p])
    return rows


def build_html_report(recs, case, logo_b64, audit_text=""):
    e = html.escape
    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    css = """
    body{font-family:"Segoe UI",Roboto,Helvetica,Arial,sans-serif;background:#eef2f7;color:#0f172a;margin:0}
    .wrap{max-width:1100px;margin:0 auto;padding:28px}
    header{display:flex;align-items:center;gap:18px;background:linear-gradient(135deg,#0b1220,#16305f);color:#fff;padding:22px 28px;border-radius:14px}
    header img{width:72px;height:72px;border-radius:14px} header h1{margin:0;font-size:26px} header p{margin:4px 0 0;color:#a9bddc;font-size:13px}
    .card{background:#fff;border:1px solid #dbe3ee;border-radius:12px;margin:22px 0;overflow:hidden}
    .card h2,.card h3{margin:0;padding:12px 18px;background:#f1f5fb;border-bottom:1px solid #dbe3ee;font-size:15px;color:#13407f}
    table{width:100%;border-collapse:collapse;font-size:13px} th,td{text-align:left;padding:8px 18px;border-bottom:1px solid #edf1f6;vertical-align:top;word-break:break-word}
    th{width:30%;color:#475569;font-weight:600} tr:nth-child(even) td,tr:nth-child(even) th{background:#fafcff}
    .badge{display:inline-block;padding:2px 10px;border-radius:99px;color:#fff;font-size:11px;font-weight:700}
    .HIGH{background:#ef4444}.MEDIUM{background:#f59e0b}.LOW{background:#3b82f6}.INFO{background:#64748b}.CLEAN{background:#10b981}.ERROR{background:#9ca3af}
    .hash{font-family:Consolas,monospace;font-size:12px}
    footer{text-align:center;color:#64748b;font-size:12px;margin:30px 0 10px}
    @media print{body{background:#fff}.card{break-inside:avoid}}
    """
    p = [f"<!DOCTYPE html><html lang='en'><head><meta charset='utf-8'><title>{APP_NAME} Report</title><style>{css}</style></head><body><div class='wrap'>"]
    p.append(f"<header><img alt='logo' src='data:image/png;base64,{logo_b64}'><div><h1>{APP_NAME} Evidence Examination Report</h1><p>Generated {now} · {APP_NAME} {APP_VERSION} · {len(recs)} file(s)</p></div></header>")
    case_rows = [(l, case.get(k, "")) for k, l in (("case_id", "Case ID"), ("examiner", "Examiner"), ("organization", "Organization"), ("notes", "Notes")) if case.get(k)]
    if case_rows:
        p.append("<div class='card'><h2>Case Information</h2><table>" + "".join(f"<tr><th>{e(l)}</th><td>{e(v)}</td></tr>" for l, v in case_rows) + "</table></div>")
    if len(recs) > 1:
        p.append("<div class='card'><h2>Evidence Summary</h2><table><tr><th>File</th><th>Type</th><th>Size</th><th>Risk</th><th>SHA-256</th></tr>")
        for r in recs:
            p.append(f"<tr><td>{e(r.name)}</td><td>{e(r.detected_type)}</td><td>{e(fmt_size(r.size))}</td><td><span class='badge {r.risk}'>{r.risk}</span></td><td class='hash'>{e(r.hashes.get('SHA-256', ''))}</td></tr>")
        p.append("</table></div>")
    for r in recs:
        p.append(f"<div class='card'><h2>{e(r.name)} &nbsp;<span class='badge {r.risk}'>{r.risk}</span></h2><table>")
        p.append(f"<tr><th>Path</th><td>{e(r.path)}</td></tr><tr><th>Detected type</th><td>{e(r.detected_type)}</td></tr><tr><th>Size</th><td>{e(fmt_size(r.size))} ({r.size:,} bytes)</td></tr>")
        for k, v in r.hashes.items():
            p.append(f"<tr><th>{e(k)}</th><td class='hash'>{e(v)}</td></tr>")
        p.append("</table></div>")
        if r.flags:
            p.append("<div class='card'><h3>Forensic Indicators</h3><table>")
            for f in r.flags:
                p.append(f"<tr><th><span class='badge {f['severity']}'>{f['severity']}</span> {e(f['title'])}</th><td>{e(f['detail'])}</td></tr>")
            p.append("</table></div>")
        for sec, vals in r.sections.items():
            if sec == "Hashes" or not vals:
                continue
            p.append(f"<div class='card'><h3>{e(sec)}</h3><table>")
            for k, v in vals.items():
                if v.startswith(("http://", "https://")) and k in ("Google Maps", "OpenStreetMap"):
                    cell = f"<a href='{e(v)}'>{e(v)}</a>"
                else:
                    cell = e(v)
                p.append(f"<tr><th>{e(k)}</th><td>{cell}</td></tr>")
            p.append("</table></div>")
        if r.events:
            p.append("<div class='card'><h3>Timeline</h3><table>")
            for d, s in sorted(r.events):
                p.append(f"<tr><th>{fmt_dt(d)}</th><td>{e(s)}</td></tr>")
            p.append("</table></div>")
    if audit_text:
        p.append(f"<div class='card'><h3>Examiner Action Log</h3><pre style='padding:12px 18px;font-size:12px;white-space:pre-wrap'>{e(audit_text)}</pre></div>")
    p.append("<footer>Timestamps are shown as recorded in the file (time zones are not normalised). Automated indicators are leads for an examiner, not proof of tampering.</footer></div></body></html>")
    return "".join(p)


# ============================================================================
# EMBEDDED APP ICON (MetaSleuth logo, 256x256 PNG, base64)
# ============================================================================
APP_ICON_B64 = """
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
14cOH2Hq5DwatckWj/vnNCuIOqPYQYaz/p6JRsdx1jWdf0SGBzxluG4ZPipDMLmzgC2CTERdCUz7X5q7Aw1BErw2pulQs1yJN+Qu
oQbFhvpVGlrxMJqW5Uocz4913PmSO/mD35VmLSJVDcggy7jThRfwqEc9mo995APEnZVUEwH/yFGUkC1Oce75F/ChD72HODbkuWIk
QsWV+k1nfel3++17mZs5iUnaPlUXiKII2xsAlvPvdAEPvP995LK7XcaZZ55Bu91i7tQct952u/7oJ7/k05/7Mkf23UzUnfQBnUZw
eZ9LL7sLV9znriwu9njYg3+nBJCKnX3Hjh08//kvoN3pYq3Sarf5/Je+Jjdcd51GcQub9nCux84zz+VRj3w497/vfeSsXWeyetUq
Wq2EPMuZnpnl5ltu5Qtf/Aof/PCnmDlxjHhkTKzNVHyoY6A/iGaLh1m5ZgtPfvKfyCMe9kDOOH0nKyYmiJKYNE05cvSY/OznV/KR
j3xMv/G1rwBtos5IuFVTtm7ZjDGGLM181WMIKUrHfOCey4FIXL+PY5HzL7wrj/n9R3HPe96NnTt2sGrVJEkSk+c5J05McdXVV/Pp
z3yej3/yK5KlGVErRm3qD7c8p92J5Tl/9mLWrl1Nv98POirDW972fo4dPYKJ4hrCr2iesnr1Kp79zD+TpOVVgUnc4tChg7zxjW9m
x+ln8eQn/gGLi6e4z33u4zf7mm/CGTt38KIX/1WoWvzm+94PfJKDt0+xfdumagMP5+C+/ftJ+ycYGT+NR/3uU3jAA67gjDN3MT42
xtTUDFf+6mo+/OFP8qtf/gLTXVlO7pY3/6bBypQliHYNezNN+Y11/4+c4OUJQKGfONU/V96RpfKMU5bnzljz92qjR3zMctpTrWz+
fcu6B1hZfS8rK+5iZfIiKxMXWJk4zzJ+jpWxXVZGz7SMnG4Z2WHpbrfS3ZZLd0tOZ3NO+7Rc2hvD14Zc2uut6W60RCvst7/zfauq
ttcf2EGa2TTN7SDNbJbldpDmdrE/sKpqf/yTn9koGbcSr7CSrMyltSqX1uo86q7PiVbk6zZss9ffcJNVVdvvD6y11uZ5bvPc2jTL
bT/N7UKvb1XVvuktb7cgNumssiZZaePuWgvYbTt25R/44MfzfnjOZb5yVc0PHzlqn/7MF1iYsHF3vU3C33/gAx9u/H6aZTbPrc2t
tVmWLfuYD3/kYyzEedwaz9etOy3/99e+IZ89dSof/j2b5zYfeoybbt5t73aP+1tIbNxdb6W9ypr2GhuPbLTQso9+9GPtbbftbj5O
llub5Utexxe++BW744w7WRNP2lZ3jQXs3/zNK/y16fXtYFBch5/aKOpYicdz05rMJRq1m7eebt/17vfbfvh8l7522/jfP/zRz+y2
nXeyyKiN2qtt1F5loWN3nX2nJX+b5ZndtvOCHMbzqL0ml2R1TrI6jzrrcojsFfd70JK/uemW2yxgn/OcP21ejzS1WZbbPMttnufL
vtbz73QPC9hXveofraraQX9g0zS1qmpf+n/+xt71srvba6+97jfdHzbNcvuqf3qDJd5s49Ft1oxstdLdZqW73TKyw8roTsvoGVbG
zrIytssyfo5l/Hwrk3eysvJSK2vuaWX9fa1sfriVrY+3ctd/sNF7Fq28yVp5fTaQ96kmP+i/Kazd+I5rAQpWVKmuqx/1hWZ6OK2E
ErwpOf461Ns1aJLFo4q6LJM169axa9cuqmGY3wlbcUyaO6zznOQ0y7nbZXfhQQ/6Hb78xU8Td1aLdc5TfO2ATtvIJz7+Uc45+ywW
ewOiKCKzjnYSY20h23Rlp3z1r39dvp4oTsh6Mzz6957AW9/6Blm3dg02t/T7/aqY1uA/byLJnWPD+nW8422vZ9369fzTq/6GuLMC
pM3mLZvJsowsy4mTJExutAQR0zQlsw7nHEkc0+v1uP6GPSCjqCof+MAHeOAD74e1ThYXez7AIjIY8UxJVR/dKIGie9aZO/jSFz/J
Ffd7KNdefR1RqwUC+eI0z3ruc3nrm14HQK/XKwG5KEnAWtLBoORnCsrDH/Ygtm3byt3udhmD1IOWO4sxWIF5APv3H8DaAUm7Q547
Jse7fP2rX+bss8/COUev1wdVojgiigw29+83iqPgxaHc4+534Quf+zCXX35fZk/1iJMOFmXT5m04B2k2QK3Sbre4+urrOHRwrxDH
WiDvITpQLMqWLZtwzjEYDECh1W5x0403AbB923ayLCNNMzrtVtliFtyONE1xzpFbz/K74Yab2LP7JqDFOeecW0NjDdY57n63u/Oi
F72INatW0O/3w/0qtUmYksQRf/Oy59NPc/7xFa8iGluB2lq0njbdmqWcEA67O9do9lbL8Z/4L//n37mjQUDqG4CrWIFLuP26VPpT
F5+htf9bCvwaYyAfcMbOraxbu5rBIPczB+dvnKuuuZ5er0cU+aiHNPNl+gte8KeI6QQteISIEZvOyJvf8p/c5953oz9IieMohPPm
/OLX14Zu0oX0GP9R3HDjDSWPLusd44+f9jQ+/akPs27tGnq9AYPM0up06HQ6dLtdRka6dDodklZEZCIGaU6WZfzj37+M33nww8n7
U6xctYqzzzmXJEloj3SpdOjB9i2OaLVajHY7jI+O0Gm3mJub4/ixI5hWG5v1+MGPfgzAwmKfkZEu3W6HVsu727babdqdDkkS+9Yz
illY7LFq5Qr+4/X/iok8i9P2Z7jgwvP5j9e+BuccWZYRRzHtdpskSVicn2dhYbF8vE6nXV6ej3/is/QWe2Ue/bbt22t8B//P23bv
KYtWExmZm5vmuuuux1rLYm9At9uhO9Kl1WoRRTGtdotOt0OURJgA0C0s9rjg/HP56795BS4fEMUJYDl95xYf6qmedGMiw+EjR8kG
ixpFpmSZajkHM+zYuSNgN/76RlHEwUOHALjokotJkoTR0ZGa6rIAdf316HQ6jI2O0koSjh0/zsLcFHGcsCNMP1z4/dw6fvcRD2Vy
fIzBIPf3RqdNp93CqWKMIYnjkBOb87K/fA47z96F7S14vosMc2Uqd6emxF+GrNjqwqaSUvBbGAOmaXDQLXYfo3Xj3yb9dWj2KUMS
T7SZUiGNjSIQsjPOPe88osiQhx3SOg9v/ttrX8u3v/lNWrH/4NVEzC5m3O9+9+Pelz8AO1ig1eqS92d4xav+maf+0ROZWRjgiOgN
LJ1WzOtf93quvfoq4sj4U1eVJI6Zmp7ltt23gySkvWl+54EP5p1vfxNZbslyi4kiRrotpqdmeNe73s3Tnv4Mnvq0P+Fd734vaebK
PjQLpJFXvfLlRPEYIyOj3Hj9dfzwhz/ilpt3V+GRQBwLp6Zn+PnPfs5Pf/pTfvSjH3P11dfwpS//l87NHlVDDqbNO97xXqanZ2Vi
fIRbd9/OG9/0Fh7/B0/gnve+F5dfcT/+7uV/z9TUDJiYPFfipEWW51x+r7tx/oXnkQ/mUc154QufT7vdIss8Uy1pJXzrW9/hQQ9+
KJdceh8uvvPdeOADH8w73vEuBoMBnU6HH/7oJ7zmn/9BTdRVm8Po2Eq2bt3iF4Fq2Xfu3bunYvMJWKv827+/jiiKGBvt8uurruWf
Xv0afu/3H8c973Vv7n//B/Da176OPLXl5+Y3K3jsY36XyVUb/elNzs4dW/19oJAHCvXeffsA58ewqkO6fdi2eUu5nIrWeM/uWwE4
cuQIV/7yl1x33fVNQ0UR5hfm+clP/PX42c9+xtVXX8NX/uu/AGVyxQir166pat1QEKdpRpLEZNmAH/zwx3zzW99h3/4DjHbbDRbJ
YmoZHenw8Ic9EOyiP/AKF8slMs8a8Wt4ulWkmFYsz2B57+SOxwAO9c/h7YOUPzlpee7JzLzC5tHDPpFz2tNy2fx7Vtbe38qqe1pW
XGqZuJOVifMt4+f6/n/sLMuo7//F9/+WzlYrnU057dNyWhty2htyaW+w0t5gk9FNFtr27e9+v1VVOz3ft/O9zM4upjZVtRdcdGf7
0Ic83KqqnesN7Klebk/M9W2uaj/26S9YMBaMfeZz/tz/zmJqT/UyOzXne9APfeQTNopb9prrbrJ9q3Z2MbWnFn3//Ourr7btkRVW
JLHr12+0e/cesP1M7dziwC70fK/33e/90J5xxtkegq19PeIRj7bzCz0738/sfD+zC33/mPe4z0MtjNhWe8QC9t///fXhdfXt/KLv
nd/3wY9ZwEZx24ppWRN1cohzTDeXeCKP2mtyaOV/+3evyv/p1f9iV6xYNfT8sQXsFVc8wM4v9OziILf9QWYXev7xn/SHT7GAXbFq
g71t916bq5Y/O378hJ1csTZAyqMW2uXj3vNe97Pf/8GP7CV3vswCedKeyGHEnnHmOXZmZsamudr5XmrnFv1n85CHPcICedxe4XGA
ZIVFRuzb3v5u+4I//wvbanXCY0vj9T/zmc+21vpr1R/kNsutHQwye+4Fdynf20c/+gmrqnah1y/xmpe89K9zIE86qzOJJ3PiyVzi
yTxqrbQQ269+9eslhrQY3u8Tn/SHFrBJ4l/LQx780PA7afmZfPXr37CANSa2IpE1UWLFRBawZ551jj01N2+z3Np+WmBR/lq/+z3v
s7t2nWtF/O+uXrXGvvnNb7Oqauf7qZ0L96qq2ne+7+MWVtt4bEcuI9tyGdmRM3J6iQEwdraV8XOtTF5gZcUlVlbd1cqae1mz4X5W
tvyulW1PtHLZq61526KVN1gr/54NeJdq9N3ef/xPMID//xXAK4qNIFVs5dKDw9tzDSH/VRJNM3xS6hrv4e+WbsJ+B8ytRdqTXHj+
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
ylpzIEfVYSKDc0K+mNFes4aztu+UjRvXMTrS5uw731t2HztJemJK6SY+sMOGjaPG7qSoSkxcK19rYa6NAYpP5i0XuobY7kZ14EIi
bhGn7crqTMOESPCAcRWXXmBKNS8J0aEw7Xq7UeDhtRyK0jNEhoxpioWvS48QkTLLunmyh4qg9PCXGgPeVBtA2EClKPfDYvfSXlNW
ARXq78d9FIvftBFpw9g65MJt6FzoFnJi7dos35B/MYzr/0eL/3+4AVRPEG1OP+PWJX9j90QxxqKLiuw4F274CZg+YvpocQNiQGI/
OilXWz3QzFt3UUuiqayEiqPb1Uu1Ydv0asfVWnu6TOlXq/WGHFfrW4uIeCGmlvSicqOoX/yS/eWXpwm9Xe3GoD7aIZKwcWixKVDb
7ZUobAaJ/17oBf3Noc0RZQCdRB3iUtQOwPURSRXJgFzU+arLGHD9ReLRlfLiv3o6z//j32V8YhyX9ogi4ec//jSz06d454e+yGv+
4730po4SjXS80Krgohfvobwxww1bCFZKbCOYo6lFXNmKqKoTj67XNwIFdaJqEVdb6MW9ENoY//uu8vwPceABWxKtYrWrY0G0EVAz
7FUntcqQWtVQktXqe5nWtpe6PZcsx3itNSUiqo0NNFRxUn1PinvE1DZX0wpfiWfZmg4Sj4AdhwvOQNdOooesEoljNEpkPP2Fbv/y
1aEn/i1vACKOl6vJtsi1yabBD93K0St0cS5joR/J+q2wdit6bBaVFpAGFDkKPZEOBYqqnxaUJ7PWxjNS7/qLBq/WvzXhPNXh2WY1
qpGiVNfqNKhvJtJY3IVxSTnHaZpC1p1ahvp7qe30SHHCl4tEK1pnJMUuL6aFFpRPAv0z8hwDf/MEhN9Ic+MsSiSXgut5imC+6Of3
rg9u4OcmJsGlKWs3nc6n3/9q7n7JWbz5PZ/n7e/4KM9/3pPZvGkD/+dv/5MnPuHBvORPnyi//4j78fAn/qXuveUmopFuSEcO4pni
pow6HpE2Le9jECVVO1Ocqs6h5KEKcKC5istB8xCnbUtuiLhMUOv5Ec5XLH40WuvzXWhrwqbiF224V9RV4TL1KqFx8mtz8TdwCa1V
pZW9XVmJihRQYoUoSljtpUJPho4jj+iKhDauaKFKPKCu9KtOf4n8qe8/5xaYThj/jUO6Fr3HGZ6+4hScONaBbNLPqTzO8m2Na6SH
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
"""


# ============================================================================
# USER INTERFACE
# ============================================================================
def _mode_color(pair):
    return pair[1] if ctk.get_appearance_mode() == "Dark" else pair[0]


def _set_text(box, text):
    box.configure(state="normal")
    box.delete("1.0", "end")
    box.insert("1.0", text)
    box.configure(state="disabled")


TAB_NAMES = ["Overview", "Metadata", "Hashes", "Hex View", "Strings & IOCs", "Image Forensics", "Timeline", "Compare", "Audit Log"]
_BASES = (ctk.CTk,) + ((TkinterDnD.DnDWrapper,) if HAS_DND else ())


class MetaSleuthApp(*_BASES):
    def __init__(self):
        settings = Settings()
        ctk.set_appearance_mode(settings["appearance"])
        ctk.set_default_color_theme("blue")
        try:
            ctk.set_widget_scaling(float(settings["scaling"]) / 100.0)
        except Exception:
            pass
        super().__init__()
        self.settings = settings
        self.audit = AuditLog()
        self.title(f"{APP_NAME} - Forensic Metadata Analysis Tool")
        self.geometry("1380x880")
        self.minsize(1120, 720)

        self.dnd_ok = False
        if HAS_DND:
            try:
                self.TkdndVersion = TkinterDnD._require(self)
                self.dnd_ok = True
            except Exception:
                self.dnd_ok = False

        self.items = collections.OrderedDict()  # uid -> {"path", "rec", "strings"}
        self._uid = 0
        self.selected_uid = None
        self.epoch = 0
        self.total = 0
        self.done = 0
        self.ui_q = queue.Queue()
        self.pool = concurrent.futures.ThreadPoolExecutor(max_workers=4)
        self._suppress_select = False
        self._refs = {}
        self._hex_hit = -1
        self.hex_offset = 0
        self.compare_map = {}

        self.mono = ctk.CTkFont(family=MONO, size=12)
        self._setup_icon()
        self._apply_tree_style()
        self._build_ui()
        self._bind_keys()
        self.report_callback_exception = self._on_tk_error
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.after(60, self._poll_queue)
        self._show_current()

    # ------------------------------------------------------------------ icon / style
    def _setup_icon(self):
        try:
            if sys.platform.startswith("win"):
                import ctypes

                ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("MetaSleuth.Forensics.3")
        except Exception:
            pass
        self.icon_pil = None
        try:
            self.icon_pil = load_app_icon()
            self.icon_photo = ImageTk.PhotoImage(self.icon_pil)
            self.iconphoto(True, self.icon_photo)
            self.after(300, lambda: self.iconphoto(True, self.icon_photo))
            if sys.platform.startswith("win"):
                import tempfile

                ico = os.path.join(tempfile.gettempdir(), "metasleuth_icon.ico")
                self.icon_pil.save(ico, format="ICO", sizes=[(16, 16), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)])
                self.after(300, lambda: self.iconbitmap(ico))
        except Exception:
            pass
        size = (48, 48)
        self.logo_img = ctk.CTkImage(light_image=self.icon_pil, dark_image=self.icon_pil, size=size) if self.icon_pil else None
        self.blank_img = ctk.CTkImage(light_image=Image.new("RGBA", (1, 1), (0, 0, 0, 0)), size=(1, 1))

    def _apply_tree_style(self):
        dark = ctk.get_appearance_mode() == "Dark"
        pal = (
            dict(bg="#14213D", fg="#E6EDF7", head="#1B2B4D", sel="#1D4ED8", selfg="#FFFFFF")
            if dark
            else dict(bg="#FFFFFF", fg="#0F172A", head="#E2E8F0", sel="#BFDBFE", selfg="#0F172A")
        )
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except Exception:
            pass
        try:
            dpi = self.winfo_fpixels("1i") / 96.0
        except Exception:
            dpi = 1.0
        scale = float(self.settings["scaling"]) / 100.0
        size = max(8, round(10 * scale))
        style.configure("MS.Treeview", background=pal["bg"], fieldbackground=pal["bg"], foreground=pal["fg"],
                        rowheight=int(26 * dpi * scale), borderwidth=0, font=(UI_FONT, size))
        style.configure("MS.Treeview.Heading", background=pal["head"], foreground=pal["fg"], relief="flat",
                        font=(UI_FONT, size, "bold"))
        style.map("MS.Treeview", background=[("selected", pal["sel"])], foreground=[("selected", pal["selfg"])])
        style.map("MS.Treeview.Heading", background=[("active", pal["head"])])
        self._tree_font_size = size

    def _style_tree_tags(self, tree):
        for name, color in RISK_COLOR.items():
            tree.tag_configure(name, foreground=color)
        tree.tag_configure("section", font=(UI_FONT, self._tree_font_size, "bold"), foreground="#38BDF8" if ctk.get_appearance_mode() == "Dark" else "#1D4ED8")
        tree.tag_configure("diff", foreground="#F59E0B")

    def _make_tree(self, parent, columns, headings, widths, show="headings", height=10):
        frame = ctk.CTkFrame(parent, fg_color="transparent")
        frame.grid_rowconfigure(0, weight=1)
        frame.grid_columnconfigure(0, weight=1)
        tree = ttk.Treeview(frame, columns=columns, show=show, style="MS.Treeview", selectmode="browse", height=height)
        for col, head, width in zip(columns, headings, widths):
            tree.heading(col, text=head, anchor="w")
            tree.column(col, width=width, minwidth=40, anchor="w", stretch=(col == columns[-1]))
        sb = ctk.CTkScrollbar(frame, command=tree.yview)
        tree.configure(yscrollcommand=sb.set)
        tree.grid(row=0, column=0, sticky="nsew")
        sb.grid(row=0, column=1, sticky="ns")
        self._style_tree_tags(tree)
        return frame, tree

    # ------------------------------------------------------------------ layout
    def _build_ui(self):
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)
        self._build_sidebar()
        self._build_main()
        self._build_statusbar()
        if self.dnd_ok:
            for w in (self, self.main_frame, self.ev_tree):
                try:
                    w.drop_target_register(DND_FILES)
                    w.dnd_bind("<<Drop>>", self._on_drop)
                except Exception:
                    pass

    def _build_sidebar(self):
        sb = ctk.CTkFrame(self, width=350, corner_radius=0, fg_color=C["panel"])
        sb.grid(row=0, column=0, sticky="nsew")
        sb.grid_propagate(False)
        sb.grid_columnconfigure(0, weight=1)
        sb.grid_rowconfigure(5, weight=1)

        brand = ctk.CTkFrame(sb, fg_color="transparent")
        brand.grid(row=0, column=0, sticky="ew", padx=18, pady=(18, 10))
        ctk.CTkLabel(brand, text="", image=self.logo_img).pack(side="left")
        tb = ctk.CTkFrame(brand, fg_color="transparent")
        tb.pack(side="left", padx=10)
        ctk.CTkLabel(tb, text="MetaSleuth", font=ctk.CTkFont(size=23, weight="bold")).pack(anchor="w")
        ctk.CTkLabel(tb, text=f"Forensic Metadata Suite · v{APP_VERSION}", font=ctk.CTkFont(size=11), text_color=C["muted"]).pack(anchor="w")

        ctk.CTkButton(sb, text="＋  Add Files", height=40, font=ctk.CTkFont(size=14, weight="bold"),
                      fg_color=C["accent"], hover_color=C["accent_hover"], command=self.add_files).grid(row=1, column=0, padx=18, pady=(6, 4), sticky="ew")
        ctk.CTkButton(sb, text="📁  Add Folder", height=36, fg_color=C["neutral"], hover_color=C["neutral_hover"],
                      text_color=C["text"], command=self.add_folder).grid(row=2, column=0, padx=18, pady=4, sticky="ew")
        row = ctk.CTkFrame(sb, fg_color="transparent")
        row.grid(row=3, column=0, padx=18, pady=4, sticky="ew")
        row.grid_columnconfigure((0, 1), weight=1)
        ctk.CTkButton(row, text="Case Info", height=32, fg_color=C["neutral"], hover_color=C["neutral_hover"], text_color=C["text"],
                      command=self.open_case_dialog).grid(row=0, column=0, padx=(0, 3), sticky="ew")
        ctk.CTkButton(row, text="Settings", height=32, fg_color=C["neutral"], hover_color=C["neutral_hover"], text_color=C["text"],
                      command=self.open_settings).grid(row=0, column=1, padx=(3, 0), sticky="ew")

        self.filter_var = tk.StringVar()
        self.filter_var.trace_add("write", lambda *_: self._refresh_list())
        self.filter_entry = ctk.CTkEntry(sb, textvariable=self.filter_var, placeholder_text="Filter evidence…  (Ctrl+F)", height=32)
        self.filter_entry.grid(row=4, column=0, padx=18, pady=(10, 4), sticky="ew")

        frame, self.ev_tree = self._make_tree(sb, ("risk", "name", "size"), ("Risk", "Evidence file", "Size"), (68, 170, 70), height=14)
        frame.grid(row=5, column=0, padx=(18, 12), pady=4, sticky="nsew")
        self.ev_tree.column("size", anchor="e")
        self.ev_tree.bind("<<TreeviewSelect>>", self._on_select)

        r2 = ctk.CTkFrame(sb, fg_color="transparent")
        r2.grid(row=6, column=0, padx=18, pady=4, sticky="ew")
        r2.grid_columnconfigure((0, 1), weight=1)
        ctk.CTkButton(r2, text="Remove", height=30, fg_color=C["neutral"], hover_color=C["neutral_hover"], text_color=C["text"],
                      command=self.remove_selected).grid(row=0, column=0, padx=(0, 3), sticky="ew")
        ctk.CTkButton(r2, text="Clear all", height=30, fg_color=C["neutral"], hover_color=C["neutral_hover"], text_color=C["text"],
                      command=self.clear_all).grid(row=0, column=1, padx=(3, 0), sticky="ew")

        self.btn_export = ctk.CTkButton(sb, text="Export  ▾", height=36, fg_color=C["neutral"], hover_color=C["neutral_hover"], text_color=C["text"])
        self.btn_export.grid(row=7, column=0, padx=18, pady=(6, 4), sticky="ew")
        self.btn_export.configure(command=self._popup_export_menu)
        ctk.CTkButton(sb, text="📑  Generate Forensic Report", height=40, font=ctk.CTkFont(size=13, weight="bold"),
                      fg_color=C["teal"], hover_color=C["teal_hover"], command=self.generate_report).grid(row=8, column=0, padx=18, pady=(4, 18), sticky="ew")

        self.export_menu = tk.Menu(self, tearoff=0)
        for scope, title in (("current", "Current file"), ("all", "All files")):
            sub = tk.Menu(self.export_menu, tearoff=0)
            for fmt, label in (("json", "JSON"), ("csv", "CSV"), ("txt", "Text report"), ("html", "HTML report")):
                sub.add_command(label=label, command=lambda s=scope, f=fmt: self.export(s, f))
            self.export_menu.add_cascade(label=title, menu=sub)
        self.export_menu.add_separator()
        self.export_menu.add_command(label="Timeline (CSV, all files)", command=self.export_timeline)

    def _popup_export_menu(self):
        try:
            x = self.btn_export.winfo_rootx()
            y = self.btn_export.winfo_rooty() + self.btn_export.winfo_height()
            self.export_menu.tk_popup(x, y)
        finally:
            self.export_menu.grab_release()

    def _build_main(self):
        self.main_frame = ctk.CTkFrame(self, fg_color=C["bg"], corner_radius=0)
        self.main_frame.grid(row=0, column=1, sticky="nsew")
        self.main_frame.grid_rowconfigure(1, weight=1)
        self.main_frame.grid_columnconfigure(0, weight=1)

        hdr = ctk.CTkFrame(self.main_frame, fg_color=C["card"], corner_radius=14)
        hdr.grid(row=0, column=0, padx=16, pady=(16, 8), sticky="ew")
        hdr.grid_columnconfigure(0, weight=1)
        top = ctk.CTkFrame(hdr, fg_color="transparent")
        top.grid(row=0, column=0, padx=18, pady=(14, 2), sticky="ew")
        self.hdr_title = ctk.CTkLabel(top, text="", font=ctk.CTkFont(size=18, weight="bold"), anchor="w")
        self.hdr_title.pack(side="left")
        self.risk_badge = ctk.CTkLabel(top, text="", corner_radius=10, font=ctk.CTkFont(size=11, weight="bold"), text_color="#FFFFFF", fg_color="#64748B", width=70, height=22)
        self.risk_badge.pack(side="left", padx=12)
        self.hdr_sub = ctk.CTkLabel(hdr, text="", font=ctk.CTkFont(size=12), text_color=C["muted"], anchor="w")
        self.hdr_sub.grid(row=1, column=0, padx=18, pady=(0, 12), sticky="w")
        btns = ctk.CTkFrame(hdr, fg_color="transparent")
        btns.grid(row=0, column=1, rowspan=2, padx=14, pady=10, sticky="e")
        self.btn_maps = ctk.CTkButton(btns, text="📍 Map", width=84, state="disabled", fg_color="#EF4444", hover_color="#DC2626", command=self.open_maps)
        self.btn_maps.pack(side="left", padx=3)
        self.btn_copy_hash = ctk.CTkButton(btns, text="Copy SHA-256", width=108, fg_color=C["neutral"], hover_color=C["neutral_hover"], text_color=C["text"], command=self.copy_sha256)
        self.btn_copy_hash.pack(side="left", padx=3)
        self.btn_reveal = ctk.CTkButton(btns, text="Show in folder", width=108, fg_color=C["neutral"], hover_color=C["neutral_hover"], text_color=C["text"], command=self.reveal_in_folder)
        self.btn_reveal.pack(side="left", padx=3)

        self.tabs = ctk.CTkTabview(self.main_frame, corner_radius=12, command=self._on_tab_change)
        self.tabs.grid(row=1, column=0, padx=16, pady=(0, 10), sticky="nsew")
        self.tab = {name: self.tabs.add(name) for name in TAB_NAMES}
        self._build_overview(self.tab["Overview"])
        self._build_metadata(self.tab["Metadata"])
        self._build_hashes(self.tab["Hashes"])
        self._build_hex(self.tab["Hex View"])
        self._build_strings(self.tab["Strings & IOCs"])
        self._build_imgforensics(self.tab["Image Forensics"])
        self._build_timeline(self.tab["Timeline"])
        self._build_compare(self.tab["Compare"])
        self._build_audit(self.tab["Audit Log"])

    def _build_statusbar(self):
        bar = ctk.CTkFrame(self, height=28, corner_radius=0, fg_color=C["panel"])
        bar.grid(row=1, column=0, columnspan=2, sticky="ew")
        bar.grid_columnconfigure(0, weight=1)
        self.status = ctk.CTkLabel(bar, text="Ready - add files, or drop them onto the window." if self.dnd_ok else "Ready - add files to begin.", font=ctk.CTkFont(size=11), text_color=C["muted"], anchor="w")
        self.status.grid(row=0, column=0, padx=14, pady=3, sticky="w")
        self.progress = ctk.CTkProgressBar(bar, width=220, height=8)
        self.progress.set(0)
        self.progress.grid(row=0, column=1, padx=14, pady=3, sticky="e")

    # ------------------------------------------------------------------ tabs: Overview
    def _build_overview(self, tab):
        tab.grid_columnconfigure(1, weight=1)
        tab.grid_rowconfigure(0, weight=1)
        left = ctk.CTkFrame(tab, width=330, fg_color=C["card"], corner_radius=12)
        left.grid(row=0, column=0, sticky="ns", padx=(0, 10), pady=4)
        left.grid_propagate(False)
        self.preview_label = ctk.CTkLabel(left, text="No preview", text_color=C["muted"], image=self.blank_img, compound="center")
        self.preview_label.pack(padx=10, pady=12)
        right = ctk.CTkScrollableFrame(tab, fg_color="transparent")
        right.grid(row=0, column=1, sticky="nsew")
        self.flags_card = ctk.CTkFrame(right, fg_color=C["card"], corner_radius=12)
        self.flags_card.pack(fill="x", pady=(4, 8))
        self.facts_card = ctk.CTkFrame(right, fg_color=C["card"], corner_radius=12)
        self.facts_card.pack(fill="x", pady=4)

    def _render_overview(self, rec):
        for card in (self.flags_card, self.facts_card):
            for w in card.winfo_children():
                w.destroy()
        ctk.CTkLabel(self.flags_card, text="Forensic Indicators", font=ctk.CTkFont(size=15, weight="bold")).pack(anchor="w", padx=16, pady=(12, 6))
        if rec is None:
            ctk.CTkLabel(self.flags_card, text="Add evidence files to see automated findings here.", text_color=C["muted"]).pack(anchor="w", padx=16, pady=(0, 14))
        elif not rec.flags:
            ctk.CTkLabel(self.flags_card, text="✔  No anomalies detected by the automated checks.", text_color="#10B981").pack(anchor="w", padx=16, pady=(0, 14))
        else:
            for f in rec.flags:
                row = ctk.CTkFrame(self.flags_card, fg_color="transparent")
                row.pack(fill="x", padx=14, pady=4)
                ctk.CTkLabel(row, text=f["severity"], width=72, height=22, corner_radius=8, fg_color=SEV[f["severity"]], text_color="#FFFFFF", font=ctk.CTkFont(size=10, weight="bold")).pack(side="left", anchor="n", pady=2)
                col = ctk.CTkFrame(row, fg_color="transparent")
                col.pack(side="left", fill="x", expand=True, padx=10)
                ctk.CTkLabel(col, text=f["title"], font=ctk.CTkFont(size=13, weight="bold"), anchor="w", justify="left", wraplength=620).pack(anchor="w")
                if f["detail"]:
                    ctk.CTkLabel(col, text=f["detail"], font=ctk.CTkFont(size=11), text_color=C["muted"], anchor="w", justify="left", wraplength=620).pack(anchor="w")
            ctk.CTkFrame(self.flags_card, height=6, fg_color="transparent").pack()
        ctk.CTkLabel(self.facts_card, text="Key Facts", font=ctk.CTkFont(size=15, weight="bold")).pack(anchor="w", padx=16, pady=(12, 6))
        grid = ctk.CTkFrame(self.facts_card, fg_color="transparent")
        grid.pack(fill="x", padx=16, pady=(0, 14))
        grid.grid_columnconfigure(1, weight=1)
        if rec is not None:
            for i, (k, v) in enumerate(rec.facts.items()):
                ctk.CTkLabel(grid, text=k, text_color=C["muted"], anchor="w", font=ctk.CTkFont(size=12)).grid(row=i, column=0, sticky="nw", padx=(0, 18), pady=3)
                ctk.CTkLabel(grid, text=v, anchor="w", justify="left", wraplength=560, font=ctk.CTkFont(size=12, weight="bold")).grid(row=i, column=1, sticky="w", pady=3)

    # ------------------------------------------------------------------ tabs: Metadata
    def _build_metadata(self, tab):
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(1, weight=1)
        bar = ctk.CTkFrame(tab, fg_color="transparent")
        bar.grid(row=0, column=0, sticky="ew", pady=(2, 6))
        bar.grid_columnconfigure(0, weight=1)
        self.meta_filter = tk.StringVar()
        self.meta_filter.trace_add("write", lambda *_: self._render_metadata(self._cur_rec()))
        ctk.CTkEntry(bar, textvariable=self.meta_filter, placeholder_text="Search tags and values…", height=32).grid(row=0, column=0, sticky="ew", padx=(0, 6))
        ctk.CTkButton(bar, text="Copy value", width=96, height=32, command=self.copy_meta_value).grid(row=0, column=1, padx=3)
        ctk.CTkButton(bar, text="Collapse all", width=96, height=32, fg_color=C["neutral"], hover_color=C["neutral_hover"], text_color=C["text"], command=lambda: self._meta_toggle(False)).grid(row=0, column=2, padx=3)
        ctk.CTkButton(bar, text="Expand all", width=96, height=32, fg_color=C["neutral"], hover_color=C["neutral_hover"], text_color=C["text"], command=lambda: self._meta_toggle(True)).grid(row=0, column=3, padx=3)
        frame, self.meta_tree = self._make_tree(tab, ("value",), ("Value",), (700,), show="tree headings")
        self.meta_tree.heading("#0", text="Section / Tag", anchor="w")
        self.meta_tree.column("#0", width=300, minwidth=160, stretch=False)
        frame.grid(row=1, column=0, sticky="nsew")
        self.meta_tree.bind("<<TreeviewSelect>>", self._on_meta_select)
        self.meta_tree.bind("<Button-3>", lambda e: self.copy_meta_value())
        self.meta_detail = ctk.CTkTextbox(tab, height=84, font=self.mono, wrap="word")
        self.meta_detail.grid(row=2, column=0, sticky="ew", pady=(6, 0))
        _set_text(self.meta_detail, "Select a row to see the full value. Right-click a row to copy its value.")

    def _render_metadata(self, rec):
        tree = self.meta_tree
        tree.delete(*tree.get_children())
        if rec is None:
            return
        needle = self.meta_filter.get().strip().lower()
        for si, (sec, vals) in enumerate(rec.sections.items()):
            rows = [(k, v) for k, v in vals.items() if not needle or needle in k.lower() or needle in v.lower()]
            if not rows:
                continue
            parent = tree.insert("", "end", iid=f"s{si}", text=f"{sec}  ({len(rows)})", values=("",), open=True, tags=("section",))
            for ri, (k, v) in enumerate(rows):
                tree.insert(parent, "end", iid=f"s{si}r{ri}", text=k, values=(v,))

    def _meta_toggle(self, opened):
        for iid in self.meta_tree.get_children():
            self.meta_tree.item(iid, open=opened)

    def _on_meta_select(self, _e=None):
        sel = self.meta_tree.selection()
        if not sel:
            return
        item = self.meta_tree.item(sel[0])
        val = item["values"][0] if item["values"] else ""
        _set_text(self.meta_detail, f"{item['text']}\n{val}" if val != "" else item["text"])

    def copy_meta_value(self):
        sel = self.meta_tree.selection()
        if sel:
            item = self.meta_tree.item(sel[0])
            if item["values"]:
                self._copy(str(item["values"][0]))

    # ------------------------------------------------------------------ tabs: Hashes
    def _build_hashes(self, tab):
        tab.grid_columnconfigure(0, weight=1)
        card = ctk.CTkFrame(tab, fg_color=C["card"], corner_radius=12)
        card.grid(row=0, column=0, sticky="ew", pady=4)
        card.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(card, text="Cryptographic hashes", font=ctk.CTkFont(size=15, weight="bold")).grid(row=0, column=0, columnspan=3, padx=16, pady=(12, 8), sticky="w")
        self.hash_entries = {}
        for i, algo in enumerate(("MD5", "SHA-1", "SHA-256", "SHA-512")):
            ctk.CTkLabel(card, text=algo, width=70, anchor="w", font=ctk.CTkFont(size=12, weight="bold")).grid(row=i + 1, column=0, padx=(16, 6), pady=5)
            ent = ctk.CTkEntry(card, font=self.mono, height=32)
            ent.grid(row=i + 1, column=1, sticky="ew", pady=5)
            ctk.CTkButton(card, text="Copy", width=64, height=30, command=lambda a=algo: self.copy_hash(a)).grid(row=i + 1, column=2, padx=(6, 16), pady=5)
            self.hash_entries[algo] = ent
        ctk.CTkFrame(card, height=8, fg_color="transparent").grid(row=6, column=0)

        v = ctk.CTkFrame(tab, fg_color=C["card"], corner_radius=12)
        v.grid(row=1, column=0, sticky="ew", pady=8)
        v.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(v, text="Verify against a known hash", font=ctk.CTkFont(size=15, weight="bold")).grid(row=0, column=0, padx=16, pady=(12, 2), sticky="w")
        ctk.CTkLabel(v, text="Paste an expected MD5 / SHA-1 / SHA-256 / SHA-512 value. The algorithm is detected from its length.", text_color=C["muted"], font=ctk.CTkFont(size=11)).grid(row=1, column=0, padx=16, sticky="w")
        row = ctk.CTkFrame(v, fg_color="transparent")
        row.grid(row=2, column=0, sticky="ew", padx=16, pady=10)
        row.grid_columnconfigure(0, weight=1)
        self.verify_entry = ctk.CTkEntry(row, font=self.mono, height=34, placeholder_text="Expected hash…")
        self.verify_entry.grid(row=0, column=0, sticky="ew", padx=(0, 8))
        ctk.CTkButton(row, text="Verify", width=90, height=34, command=self.verify_hash).grid(row=0, column=1)
        self.verify_result = ctk.CTkLabel(v, text="", font=ctk.CTkFont(size=13, weight="bold"))
        self.verify_result.grid(row=3, column=0, padx=16, pady=(0, 14), sticky="w")

    def _fill_hashes(self, rec):
        for algo, ent in self.hash_entries.items():
            ent.configure(state="normal")
            ent.delete(0, "end")
            if rec is not None:
                ent.insert(0, rec.hashes.get(algo, ""))
            ent.configure(state="readonly")
        self.verify_result.configure(text="")

    def copy_hash(self, algo):
        rec = self._cur_rec()
        if rec and rec.hashes.get(algo):
            self._copy(rec.hashes[algo])
            self._status(f"{algo} copied to clipboard.")

    def copy_sha256(self):
        self.copy_hash("SHA-256")

    def verify_hash(self):
        rec = self._cur_rec()
        if not rec:
            messagebox.showwarning(APP_NAME, "Select an analysed file first.")
            return
        want = re.sub(r"[^0-9A-Fa-f]", "", self.verify_entry.get()).upper()
        algo = {32: "MD5", 40: "SHA-1", 64: "SHA-256", 128: "SHA-512"}.get(len(want))
        if not algo:
            self.verify_result.configure(text="✖  Not a valid MD5 / SHA-1 / SHA-256 / SHA-512 length.", text_color="#F59E0B")
        elif rec.hashes.get(algo) == want:
            self.verify_result.configure(text=f"✔  {algo} MATCH - file integrity confirmed.", text_color="#10B981")
        else:
            self.verify_result.configure(text=f"✖  {algo} MISMATCH - the file differs from the expected hash.", text_color="#EF4444")
        self.audit.add("HASH VERIFY", f"{rec.name}: {algo or 'invalid'}")

    # ------------------------------------------------------------------ tabs: Hex
    def _build_hex(self, tab):
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(1, weight=1)
        bar = ctk.CTkFrame(tab, fg_color="transparent")
        bar.grid(row=0, column=0, sticky="ew", pady=(2, 6))
        ctk.CTkButton(bar, text="◀ Prev", width=70, height=30, command=lambda: self._hex_move(-HEX_PAGE)).pack(side="left", padx=2)
        ctk.CTkButton(bar, text="Next ▶", width=70, height=30, command=lambda: self._hex_move(HEX_PAGE)).pack(side="left", padx=2)
        self.hex_goto = ctk.CTkEntry(bar, width=110, height=30, placeholder_text="Offset (0x…)")
        self.hex_goto.pack(side="left", padx=(12, 2))
        ctk.CTkButton(bar, text="Go", width=44, height=30, command=self._hex_go).pack(side="left", padx=2)
        self.hex_mode = ctk.CTkSegmentedButton(bar, values=["Text", "Hex"])
        self.hex_mode.set("Text")
        self.hex_mode.pack(side="left", padx=(16, 4))
        self.hex_find_entry = ctk.CTkEntry(bar, width=200, height=30, placeholder_text="Find in file…")
        self.hex_find_entry.pack(side="left", padx=2)
        self.hex_find_entry.bind("<Return>", lambda e: self._hex_find())
        ctk.CTkButton(bar, text="Find next", width=80, height=30, command=self._hex_find).pack(side="left", padx=2)
        self.hex_info = ctk.CTkLabel(bar, text="", text_color=C["muted"], font=ctk.CTkFont(size=11))
        self.hex_info.pack(side="right", padx=6)
        self.hex_text = ctk.CTkTextbox(tab, font=self.mono, wrap="none", state="disabled")
        self.hex_text.grid(row=1, column=0, sticky="nsew")

    def _hex_load(self, offset=None):
        rec = self._cur_rec()
        if rec is None or rec.error:
            _set_text(self.hex_text, "")
            self.hex_info.configure(text="")
            return
        if offset is not None:
            self.hex_offset = max(0, min(int(offset), max(0, rec.size - 1)))
        self.hex_offset -= self.hex_offset % 16
        try:
            _set_text(self.hex_text, read_hex_page(rec.path, self.hex_offset))
            self.hex_info.configure(text=f"0x{self.hex_offset:08X}  ·  {self.hex_offset:,} / {rec.size:,} bytes")
        except Exception as exc:
            _set_text(self.hex_text, f"Cannot read file: {exc}")

    def _hex_move(self, delta):
        self._hex_load(self.hex_offset + delta)

    def _hex_go(self):
        txt = self.hex_goto.get().strip()
        try:
            self._hex_load(int(txt, 16) if txt.lower().startswith("0x") else int(txt))
        except ValueError:
            self._status("Enter an offset such as 0x1F40 or 8000.")

    def _hex_find(self):
        rec = self._cur_rec()
        if rec is None:
            return
        raw = self.hex_find_entry.get()
        try:
            if self.hex_mode.get() == "Hex":
                clean = re.sub(r"[^0-9A-Fa-f]", "", raw)
                if len(clean) % 2 or not clean:
                    raise ValueError
                needle = bytes.fromhex(clean)
            else:
                needle = raw.encode("utf-8")
        except ValueError:
            self._status("Invalid search pattern.")
            return
        start = self._hex_hit + 1 if self._hex_hit >= 0 else self.hex_offset
        pos = hex_find(rec.path, needle, start)
        if pos < 0 and start > 0:
            pos = hex_find(rec.path, needle, 0)
        if pos < 0:
            self._hex_hit = -1
            self._status("Pattern not found.")
        else:
            self._hex_hit = pos
            self._hex_load(pos)
            self._status(f"Found at offset 0x{pos:X} ({pos:,}).")

    # ------------------------------------------------------------------ tabs: Strings & IOCs
    def _build_strings(self, tab):
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(1, weight=1)
        bar = ctk.CTkFrame(tab, fg_color="transparent")
        bar.grid(row=0, column=0, sticky="ew", pady=(2, 6))
        ctk.CTkLabel(bar, text="Min length").pack(side="left", padx=(0, 4))
        self.str_min = ctk.CTkEntry(bar, width=50, height=30)
        self.str_min.insert(0, str(self.settings["min_string_len"]))
        self.str_min.pack(side="left")
        ctk.CTkButton(bar, text="Extract", width=80, height=30, command=self._strings_run).pack(side="left", padx=8)
        self.str_cat = ctk.CTkSegmentedButton(bar, values=["All strings", "URLs", "Emails", "IP addresses", "File paths"], command=lambda v: self._strings_render())
        self.str_cat.set("All strings")
        self.str_cat.pack(side="left", padx=8)
        self.str_filter = tk.StringVar()
        self.str_filter.trace_add("write", lambda *_: self._strings_render())
        ctk.CTkEntry(bar, textvariable=self.str_filter, width=180, height=30, placeholder_text="Filter…").pack(side="left", padx=4)
        self.str_status = ctk.CTkLabel(bar, text="", text_color=C["muted"], font=ctk.CTkFont(size=11))
        self.str_status.pack(side="right", padx=6)
        self.str_text = ctk.CTkTextbox(tab, font=self.mono, wrap="none", state="disabled")
        self.str_text.grid(row=1, column=0, sticky="nsew")

    def _strings_run(self):
        rec = self._cur_rec()
        if rec is None or rec.error:
            return
        try:
            n = int(self.str_min.get())
        except ValueError:
            n = self.settings["min_string_len"]
        uid = self.selected_uid
        self.str_status.configure(text="Extracting…")
        self.run_bg(lambda: extract_strings(rec.path, n), lambda res, err: self._strings_done(uid, res, err))

    def _strings_done(self, uid, res, err):
        if err:
            self.str_status.configure(text=f"Error: {err}")
            return
        if uid in self.items:
            self.items[uid]["strings"] = res
            self.audit.add("STRINGS", f"{self.items[uid]['path']} ({res['_meta']['total']} strings)")
        if uid == self.selected_uid:
            self._strings_render()

    def _strings_render(self):
        item = self.items.get(self.selected_uid)
        data = item.get("strings") if item else None
        if not data:
            _set_text(self.str_text, "Press 'Extract' to scan this file for printable ASCII / UTF-16 strings and indicators (URLs, e-mails, IPs, paths).")
            self.str_status.configure(text="")
            return
        lines = data.get(self.str_cat.get(), [])
        needle = self.str_filter.get().strip().lower()
        if needle:
            lines = [l for l in lines if needle in l.lower()]
        shown = lines[:4000]
        note = f"\n\n… {len(lines) - len(shown):,} more (refine the filter or export)" if len(lines) > len(shown) else ""
        _set_text(self.str_text, ("\n".join(shown) or "(nothing found)") + note)
        meta = data["_meta"]
        extra = " · truncated" if meta["truncated"] else ""
        self.str_status.configure(text=f"{len(lines):,} result(s) · scanned {fmt_size(meta['scanned_bytes'])}{extra}")

    # ------------------------------------------------------------------ tabs: Image Forensics
    def _build_imgforensics(self, tab):
        tab.grid_columnconfigure((0, 1), weight=1)
        tab.grid_rowconfigure(2, weight=1)
        bar = ctk.CTkFrame(tab, fg_color="transparent")
        bar.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(2, 6))
        self.btn_ela = ctk.CTkButton(bar, text="Run ELA", width=90, height=32, command=self.run_ela)
        self.btn_ela.pack(side="left", padx=2)
        self.ela_q_label = ctk.CTkLabel(bar, text=f"JPEG quality {self.settings['ela_quality']}", width=110)
        self.ela_q_label.pack(side="left", padx=(10, 2))
        self.ela_slider = ctk.CTkSlider(bar, from_=60, to=98, number_of_steps=38, width=150,
                                        command=lambda v: self.ela_q_label.configure(text=f"JPEG quality {int(v)}"))
        self.ela_slider.set(self.settings["ela_quality"])
        self.ela_slider.pack(side="left")
        self.btn_strip = ctk.CTkButton(bar, text="🧹 Save metadata-free copy", width=190, height=32, fg_color=C["teal"], hover_color=C["teal_hover"], command=self.strip_current)
        self.btn_strip.pack(side="left", padx=18)
        self.btn_osm = ctk.CTkButton(bar, text="OpenStreetMap", width=110, height=32, fg_color=C["neutral"], hover_color=C["neutral_hover"], text_color=C["text"], state="disabled", command=lambda: self.open_maps("osm"))
        self.btn_osm.pack(side="right", padx=2)
        ctk.CTkLabel(tab, text="Original", font=ctk.CTkFont(size=13, weight="bold")).grid(row=1, column=0, sticky="w", padx=6)
        ctk.CTkLabel(tab, text="Error Level Analysis (ELA)", font=ctk.CTkFont(size=13, weight="bold")).grid(row=1, column=1, sticky="w", padx=6)
        self.if_left = ctk.CTkLabel(tab, text="", image=self.blank_img, compound="center", fg_color=C["card"], corner_radius=12)
        self.if_left.grid(row=2, column=0, sticky="nsew", padx=4, pady=4)
        self.if_right = ctk.CTkLabel(tab, text="Press 'Run ELA'", image=self.blank_img, compound="center", fg_color=C["card"], corner_radius=12, text_color=C["muted"])
        self.if_right.grid(row=2, column=1, sticky="nsew", padx=4, pady=4)
        self.if_info = ctk.CTkLabel(tab, text="", text_color=C["muted"], font=ctk.CTkFont(size=11), wraplength=900, justify="left")
        self.if_info.grid(row=3, column=0, columnspan=2, sticky="w", padx=6, pady=(2, 0))

    def _setup_imgforensics(self, rec):
        is_img = rec is not None and rec.kind == "image" and not rec.error
        state = "normal" if is_img else "disabled"
        self.btn_ela.configure(state=state)
        self.btn_strip.configure(state=state)
        self.btn_osm.configure(state="normal" if (rec and rec.gps) else "disabled")
        self.if_left.configure(image=self.blank_img, text="" if is_img else "Select an image file to use these tools.")
        self.if_right.configure(image=self.blank_img, text="Press 'Run ELA'" if is_img else "")
        self.if_info.configure(text="ELA re-saves the image as JPEG and amplifies the difference. Uniform noise is normal; sharply brighter patches can mark pasted, retouched or re-compressed regions. It is an indicator, not proof." if is_img else "")
        if is_img:
            uid = self.selected_uid
            self.run_bg(lambda: make_thumbnail(rec.path, (470, 470)), lambda res, err: self._show_pil(uid, self.if_left, res, err, (470, 470)))

    def _show_pil(self, uid, label, pil, err, box):
        if err or uid != self.selected_uid or pil is None:
            return
        img = ctk.CTkImage(light_image=pil, dark_image=pil, size=pil.size)
        self._refs[id(label)] = img
        label.configure(image=img, text="")

    def run_ela(self):
        rec = self._cur_rec()
        if rec is None:
            return
        q = int(self.ela_slider.get())
        uid = self.selected_uid
        self.if_right.configure(image=self.blank_img, text="Computing…")
        self.btn_ela.configure(state="disabled")

        def done(res, err):
            self.btn_ela.configure(state="normal")
            if err or res is None:
                self.if_right.configure(text=f"ELA failed: {err}")
                return
            orig, ela, stats = res
            if uid == self.selected_uid:
                self._show_pil(uid, self.if_left, orig, None, orig.size)
                self._show_pil(uid, self.if_right, ela, None, ela.size)
                self.if_info.configure(text=f"ELA at JPEG quality {stats['quality']} · max error level {stats['max']} · mean {stats['mean']:.2f}. Bright, isolated regions deserve a closer look; uniform brightness is normal.")
            self.audit.add("ELA", f"{rec.name} (quality {q})")

        self.run_bg(lambda: compute_ela(rec.path, q), done)

    def strip_current(self):
        rec = self._cur_rec()
        if rec is None:
            return
        src = Path(rec.path)
        dst = filedialog.asksaveasfilename(title="Save metadata-free copy", initialfile=f"{src.stem}_clean{src.suffix}", defaultextension=src.suffix)
        if not dst:
            return
        if os.path.abspath(dst) == os.path.abspath(rec.path):
            messagebox.showerror(APP_NAME, "Choose a different file name - the original evidence is never overwritten.")
            return

        def done(res, err):
            if err:
                messagebox.showerror(APP_NAME, f"Could not create clean copy:\n{err}")
            else:
                messagebox.showinfo(APP_NAME, f"Clean copy saved:\n{dst}\n\n{fmt_size(res[0])}  →  {fmt_size(res[1])}\nEXIF, GPS, XMP and text chunks were not copied.")
                self.audit.add("STRIP METADATA", f"{rec.path} -> {dst}")

        self.run_bg(lambda: strip_metadata(rec.path, dst), done)

    # ------------------------------------------------------------------ tabs: Timeline
    def _build_timeline(self, tab):
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(1, weight=1)
        bar = ctk.CTkFrame(tab, fg_color="transparent")
        bar.grid(row=0, column=0, sticky="ew", pady=(2, 6))
        self.tl_scope = ctk.CTkSegmentedButton(bar, values=["All files", "Current file"], command=lambda v: self._render_timeline())
        self.tl_scope.set("All files")
        self.tl_scope.pack(side="left")
        ctk.CTkButton(bar, text="Export CSV", width=100, height=30, command=self.export_timeline).pack(side="right")
        self.tl_info = ctk.CTkLabel(bar, text="", text_color=C["muted"], font=ctk.CTkFont(size=11))
        self.tl_info.pack(side="right", padx=12)
        frame, self.tl_tree = self._make_tree(tab, ("time", "event", "file"), ("Time", "Event", "File"), (160, 260, 400))
        frame.grid(row=1, column=0, sticky="nsew")

    def _render_timeline(self):
        self.tl_tree.delete(*self.tl_tree.get_children())
        if self.tl_scope.get() == "Current file":
            recs = [r for r in [self._cur_rec()] if r]
        else:
            recs = self._all_recs()
        events = sorted(((d, s, r.name) for r in recs for d, s in r.events), key=lambda e: e[0])
        for d, s, n in events[:20000]:
            self.tl_tree.insert("", "end", values=(fmt_dt(d), s, n))
        self.tl_info.configure(text=f"{len(events):,} event(s) · times shown as recorded, time zones not normalised")

    # ------------------------------------------------------------------ tabs: Compare
    def _build_compare(self, tab):
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(1, weight=1)
        bar = ctk.CTkFrame(tab, fg_color="transparent")
        bar.grid(row=0, column=0, sticky="ew", pady=(2, 6))
        self.cmp_a = ctk.CTkOptionMenu(bar, values=["-"], width=250, dynamic_resizing=False)
        self.cmp_b = ctk.CTkOptionMenu(bar, values=["-"], width=250, dynamic_resizing=False)
        ctk.CTkLabel(bar, text="A").pack(side="left", padx=(0, 4))
        self.cmp_a.pack(side="left")
        ctk.CTkLabel(bar, text="B").pack(side="left", padx=(12, 4))
        self.cmp_b.pack(side="left")
        self.cmp_diff = ctk.CTkCheckBox(bar, text="Only differences", command=self._run_compare)
        self.cmp_diff.select()
        self.cmp_diff.pack(side="left", padx=14)
        ctk.CTkButton(bar, text="Compare", width=90, height=30, command=self._run_compare).pack(side="left")
        self.cmp_info = ctk.CTkLabel(bar, text="", text_color=C["muted"], font=ctk.CTkFont(size=11))
        self.cmp_info.pack(side="right", padx=6)
        frame, self.cmp_tree = self._make_tree(tab, ("a", "b"), ("File A", "File B"), (380, 380), show="tree headings")
        self.cmp_tree.heading("#0", text="Property", anchor="w")
        self.cmp_tree.column("#0", width=300, stretch=False)
        frame.grid(row=1, column=0, sticky="nsew")

    def _refresh_compare_choices(self):
        self.compare_map = {}
        labels = []
        for i, (uid, it) in enumerate(self.items.items(), 1):
            if it["rec"] is not None:
                label = f"{i}. {it['rec'].name}"[:44]
                self.compare_map[label] = uid
                labels.append(label)
        if not labels:
            labels = ["-"]
        self.cmp_a.configure(values=labels)
        self.cmp_b.configure(values=labels)
        if self.cmp_a.get() not in labels:
            self.cmp_a.set(labels[0])
        if self.cmp_b.get() not in labels:
            self.cmp_b.set(labels[1] if len(labels) > 1 else labels[0])

    def _run_compare(self):
        self.cmp_tree.delete(*self.cmp_tree.get_children())
        ua, ub = self.compare_map.get(self.cmp_a.get()), self.compare_map.get(self.cmp_b.get())
        if not ua or not ub:
            self.cmp_info.configure(text="Load at least two analysed files.")
            return
        ra, rb = self.items[ua]["rec"], self.items[ub]["rec"]
        fa, fb = ra.flat(), rb.flat()
        keys = list(fa.keys()) + [k for k in fb if k not in fa]
        only_diff = bool(self.cmp_diff.get())
        shown = diffs = 0
        for i, key in enumerate(keys):
            va, vb = fa.get(key, "—"), fb.get(key, "—")
            differs = va != vb
            diffs += differs
            if only_diff and not differs:
                continue
            shown += 1
            self.cmp_tree.insert("", "end", text=key, values=(va, vb), tags=("diff",) if differs else ())
        self.cmp_info.configure(text=f"{diffs} difference(s) in {len(keys)} properties")

    # ------------------------------------------------------------------ tabs: Audit
    def _build_audit(self, tab):
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(1, weight=1)
        bar = ctk.CTkFrame(tab, fg_color="transparent")
        bar.grid(row=0, column=0, sticky="ew", pady=(2, 6))
        ctk.CTkLabel(bar, text="Chain-of-actions log of this session (also saved to the MetaSleuth config folder).", text_color=C["muted"], font=ctk.CTkFont(size=11)).pack(side="left")
        ctk.CTkButton(bar, text="Export log", width=96, height=30, command=self.export_audit).pack(side="right")
        self.audit_text = ctk.CTkTextbox(tab, font=self.mono, wrap="none", state="disabled")
        self.audit_text.grid(row=1, column=0, sticky="nsew")

    def export_audit(self):
        path = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Text", "*.txt")], initialfile="metasleuth_audit_log.txt")
        if path:
            Path(path).write_text(self.audit.text(), encoding="utf-8")

    # ------------------------------------------------------------------ evidence management
    def _bind_keys(self):
        self.bind("<Control-o>", lambda e: self.add_files())
        self.bind("<Control-O>", lambda e: self.add_folder())
        self.bind("<Control-f>", lambda e: self.filter_entry.focus_set())
        self.bind("<Delete>", lambda e: self.remove_selected() if self.focus_get() is self.ev_tree else None)
        self.bind("<F5>", lambda e: self._reanalyze())

    def _all_recs(self):
        return [it["rec"] for it in self.items.values() if it["rec"] is not None]

    def _cur_rec(self):
        it = self.items.get(self.selected_uid)
        return it["rec"] if it else None

    def add_files(self):
        paths = filedialog.askopenfilenames(title="Select evidence files")
        if paths:
            self.load_paths(list(paths))

    def add_folder(self):
        folder = filedialog.askdirectory(title="Select evidence folder")
        if folder:
            self.load_paths([folder])

    def _on_drop(self, event):
        try:
            self.load_paths(list(self.tk.splitlist(event.data)))
        except Exception as exc:
            self._status(f"Drop failed: {exc}")

    def load_paths(self, paths):
        files = []
        recursive = bool(self.settings["recursive"])
        for p in paths:
            if os.path.isdir(p):
                if recursive:
                    for root, _, names in os.walk(p):
                        files += [os.path.join(root, n) for n in sorted(names)]
                else:
                    files += [os.path.join(p, n) for n in sorted(os.listdir(p)) if os.path.isfile(os.path.join(p, n))]
            elif os.path.isfile(p):
                files.append(p)
        existing = {os.path.normcase(os.path.abspath(it["path"])) for it in self.items.values()}
        new = []
        for f in files:
            key = os.path.normcase(os.path.abspath(f))
            if key not in existing:
                existing.add(key)
                new.append(os.path.abspath(f))
        if not new:
            self._status("No new files to add.")
            return
        for f in new:
            self._uid += 1
            uid = f"u{self._uid}"
            self.items[uid] = {"path": f, "rec": None, "strings": None}
            self.pool.submit(self._worker, self.epoch, uid, f)
        self.total += len(new)
        self.audit.add("LOAD", f"{len(new)} file(s) added, first: {new[0]}")
        self._refresh_list()
        self._status(f"Analysing {len(new)} file(s)…")

    def _worker(self, epoch, uid, path):
        if epoch != self.epoch:
            return
        try:
            rec = Analyzer.analyze(path)
        except Exception as exc:
            rec = Record(path)
            rec.error = f"{type(exc).__name__}: {exc}"
            rec.flag("HIGH", "Analysis failed", rec.error)
        self.ui_q.put(("result", epoch, uid, rec))

    def run_bg(self, func, callback):
        def job():
            try:
                self.ui_q.put(("cb", callback, func(), None))
            except Exception as exc:
                self.ui_q.put(("cb", callback, None, f"{type(exc).__name__}: {exc}"))

        self.pool.submit(job)

    def _poll_queue(self):
        try:
            for _ in range(80):
                msg = self.ui_q.get_nowait()
                try:
                    if msg[0] == "result":
                        self._on_result(*msg[1:])
                    elif msg[0] == "cb":
                        msg[1](msg[2], msg[3])
                except Exception as exc:
                    self._log_error(exc)
        except queue.Empty:
            pass
        finally:
            self.after(60, self._poll_queue)

    def _on_result(self, epoch, uid, rec):
        if epoch != self.epoch or uid not in self.items:
            return
        self.items[uid]["rec"] = rec
        self.audit.add("ANALYSED", f"{rec.name} | risk {rec.risk} | SHA-256 {rec.hashes.get('SHA-256', 'n/a')}")
        self.done += 1
        self._update_row(uid)
        if self.selected_uid is None:
            self._select(uid)
        elif uid == self.selected_uid:
            self._show_current()
        if self.done >= self.total:
            self.progress.set(1)
            high = sum(1 for r in self._all_recs() if r.risk == "HIGH")
            self._status(f"Done - {self.done} file(s) analysed" + (f", {high} with HIGH-risk indicators." if high else "."))
            self._refresh_compare_choices()
            if self.tabs.get() == "Timeline":
                self._render_timeline()
        else:
            self.progress.set(self.done / max(1, self.total))
            self._status(f"Analysing… {self.done} / {self.total}")

    def _reanalyze(self):
        uid = self.selected_uid
        if uid in self.items:
            self.items[uid]["rec"] = None
            self.total += 1
            self._update_row(uid)
            self.pool.submit(self._worker, self.epoch, uid, self.items[uid]["path"])

    def remove_selected(self):
        uid = self.selected_uid
        if uid in self.items:
            if self.items[uid]["rec"] is None:
                self.total = max(0, self.total - 1)
            else:
                self.done = max(0, self.done - 1)
                self.total = max(0, self.total - 1)
            self.audit.add("REMOVE", self.items[uid]["path"])
            del self.items[uid]
            self.selected_uid = None
            self._refresh_list()
            first = next(iter(self.items), None)
            if first:
                self._select(first)
            else:
                self._show_current()
            self._refresh_compare_choices()

    def clear_all(self):
        if self.items and not messagebox.askyesno(APP_NAME, "Remove all evidence from the workspace?\n(The original files are never touched.)"):
            return
        self.epoch += 1
        self.items.clear()
        self.selected_uid = None
        self.total = self.done = 0
        self.progress.set(0)
        self.audit.add("CLEAR", "Workspace cleared")
        self._refresh_list()
        self._show_current()
        self._refresh_compare_choices()
        self._status("Workspace cleared.")

    def _row_values(self, uid):
        it = self.items[uid]
        rec = it["rec"]
        name = Path(it["path"]).name
        if rec is None:
            return ("...", name, "")
        return (rec.risk, name, fmt_size(rec.size))

    def _refresh_list(self):
        needle = self.filter_var.get().strip().lower()
        self._suppress_select = True
        try:
            self.ev_tree.delete(*self.ev_tree.get_children())
            for uid in self.items:
                vals = self._row_values(uid)
                if needle and needle not in vals[1].lower() and needle not in vals[0].lower():
                    continue
                self.ev_tree.insert("", "end", iid=uid, values=vals, tags=(vals[0],))
            if self.selected_uid and self.ev_tree.exists(self.selected_uid):
                self.ev_tree.selection_set(self.selected_uid)
        finally:
            self._suppress_select = False

    def _update_row(self, uid):
        if self.ev_tree.exists(uid):
            vals = self._row_values(uid)
            self.ev_tree.item(uid, values=vals, tags=(vals[0],))

    def _select(self, uid):
        self.selected_uid = uid
        self._suppress_select = True
        try:
            if self.ev_tree.exists(uid):
                self.ev_tree.selection_set(uid)
                self.ev_tree.see(uid)
        finally:
            self._suppress_select = False
        self._show_current()

    def _on_select(self, _e=None):
        if self._suppress_select:
            return
        sel = self.ev_tree.selection()
        if sel and sel[0] != self.selected_uid:
            self.selected_uid = sel[0]
            self._show_current()

    # ------------------------------------------------------------------ rendering the selected record
    def _show_current(self):
        rec = self._cur_rec()
        pending = self.selected_uid in self.items and rec is None
        if rec is None:
            self.hdr_title.configure(text="Analysing…" if pending else "No evidence loaded")
            self.hdr_sub.configure(text="" if pending else "Add files or a folder (Ctrl+O) to begin" + (" - or drag & drop them here." if self.dnd_ok else "."))
            self.risk_badge.configure(text="", fg_color="transparent")
            self.btn_maps.configure(state="disabled", text="📍 Map")
            self.preview_label.configure(image=self.blank_img, text="No preview")
        else:
            self.hdr_title.configure(text=rec.name)
            self.hdr_sub.configure(text=f"{rec.detected_type}  ·  {fmt_size(rec.size)}  ·  {rec.path}")
            self.risk_badge.configure(text=rec.risk, fg_color=RISK_COLOR[rec.risk])
            self.btn_maps.configure(state="normal" if rec.gps else "disabled", text="📍 Map" if rec.gps else "📍 No GPS")
            self.preview_label.configure(image=self.blank_img, text="Loading preview…" if rec.kind == "image" else f"{rec.detected_type}\n\n(no visual preview)")
            if rec.kind == "image" and not rec.error:
                uid = self.selected_uid
                self.run_bg(lambda: make_thumbnail(rec.path, (300, 300)), lambda res, err: self._preview_done(uid, res, err))
        self._render_overview(rec)
        self._render_metadata(rec)
        self._fill_hashes(rec)
        self._hex_hit = -1
        self._hex_load(0)
        self._strings_render()
        self._setup_imgforensics(rec)
        _set_text(self.meta_detail, "Select a row to see the full value. Right-click a row to copy its value.")
        if self.tabs.get() == "Timeline" and self.tl_scope.get() == "Current file":
            self._render_timeline()

    def _preview_done(self, uid, pil, err):
        if uid != self.selected_uid:
            return
        if err or pil is None:
            self.preview_label.configure(text="Preview unavailable")
            return
        img = ctk.CTkImage(light_image=pil, dark_image=pil, size=pil.size)
        self._refs["preview"] = img
        self.preview_label.configure(image=img, text="")

    def _on_tab_change(self):
        name = self.tabs.get()
        if name == "Timeline":
            self._render_timeline()
        elif name == "Compare":
            self._refresh_compare_choices()
        elif name == "Audit Log":
            _set_text(self.audit_text, self.audit.text())

    # ------------------------------------------------------------------ actions
    def _copy(self, text):
        self.clipboard_clear()
        self.clipboard_append(text)

    def _status(self, text):
        self.status.configure(text=text)

    def open_maps(self, which="google"):
        rec = self._cur_rec()
        if rec and rec.gps:
            lat, lon = rec.gps
            url = f"https://www.openstreetmap.org/?mlat={lat}&mlon={lon}#map=17/{lat}/{lon}" if which == "osm" else f"https://www.google.com/maps?q={lat},{lon}"
            webbrowser.open_new_tab(url)

    def reveal_in_folder(self):
        rec = self._cur_rec()
        if not rec:
            return
        try:
            if sys.platform.startswith("win"):
                subprocess.Popen(["explorer", "/select,", os.path.normpath(rec.path)])
            elif sys.platform == "darwin":
                subprocess.Popen(["open", "-R", rec.path])
            else:
                subprocess.Popen(["xdg-open", os.path.dirname(rec.path)])
        except Exception as exc:
            messagebox.showerror(APP_NAME, str(exc))

    # ------------------------------------------------------------------ export / report
    def _export_recs(self, scope):
        recs = [self._cur_rec()] if scope == "current" else self._all_recs()
        recs = [r for r in recs if r is not None]
        if not recs:
            messagebox.showwarning(APP_NAME, "Nothing to export yet - load and analyse a file first.")
        return recs

    def export(self, scope, fmt):
        recs = self._export_recs(scope)
        if not recs:
            return
        stem = Path(recs[0].name).stem if len(recs) == 1 else "metasleuth_all_files"
        ext = {"json": ".json", "csv": ".csv", "txt": ".txt", "html": ".html"}[fmt]
        default = f"{stem}_metadata{ext}" if fmt != "html" else f"FORENSIC_REPORT_{stem}.html"
        types = {".json": "JSON", ".csv": "CSV", ".txt": "Text", ".html": "HTML report"}
        path = filedialog.asksaveasfilename(defaultextension=ext, filetypes=[(types[ext], "*" + ext)], initialfile=default)
        if not path:
            return
        try:
            case = self.settings["case"]
            if fmt == "json":
                Path(path).write_text(build_json(recs, case), encoding="utf-8")
            elif fmt == "csv":
                with open(path, "w", newline="", encoding="utf-8-sig") as fh:
                    csv.writer(fh).writerows(build_csv_rows(recs))
            elif fmt == "txt":
                Path(path).write_text(build_text_report(recs, case), encoding="utf-8")
            else:
                Path(path).write_text(build_html_report(recs, case, self._logo_b64(), self.audit.text()), encoding="utf-8")
                webbrowser.open_new_tab(Path(path).resolve().as_uri())
            self.audit.add("EXPORT", f"{fmt.upper()} ({len(recs)} file(s)) -> {path}")
            self._status(f"Saved {path}")
            if fmt != "html":
                messagebox.showinfo(APP_NAME, f"Saved:\n{path}")
        except Exception as exc:
            messagebox.showerror(APP_NAME, f"Export failed:\n{exc}")

    def generate_report(self):
        n = len(self._all_recs())
        if n == 0:
            messagebox.showwarning(APP_NAME, "Load a file before generating a report.")
            return
        scope = "current"
        if n > 1:
            ans = messagebox.askyesnocancel("Forensic report", f"Include all {n} analysed files?\n\nYes = all files\nNo = current file only")
            if ans is None:
                return
            scope = "all" if ans else "current"
        self.export(scope, "html")

    def export_timeline(self):
        recs = self._all_recs()
        if not recs:
            messagebox.showwarning(APP_NAME, "No analysed files yet.")
            return
        path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV", "*.csv")], initialfile="metasleuth_timeline.csv")
        if path:
            with open(path, "w", newline="", encoding="utf-8-sig") as fh:
                csv.writer(fh).writerows(build_timeline_rows(recs))
            self.audit.add("EXPORT", f"Timeline CSV -> {path}")
            self._status(f"Timeline saved to {path}")

    def _logo_b64(self):
        return "".join(APP_ICON_B64.split())

    # ------------------------------------------------------------------ dialogs
    def _dialog(self, title, w, h):
        top = ctk.CTkToplevel(self)
        top.title(title)
        x = self.winfo_rootx() + max(0, (self.winfo_width() - w) // 2)
        y = self.winfo_rooty() + max(0, (self.winfo_height() - h) // 3)
        top.geometry(f"{w}x{h}+{x}+{y}")
        top.resizable(False, False)
        try:
            top.transient(self)
            top.after(200, top.grab_set)
            if self.icon_pil:
                top.after(300, lambda: top.iconphoto(False, self.icon_photo))
        except Exception:
            pass
        return top

    def open_case_dialog(self):
        top = self._dialog("Case Information", 480, 470)
        case = self.settings["case"]
        entries = {}
        for i, (key, label) in enumerate((("case_id", "Case / reference number"), ("examiner", "Examiner"), ("organization", "Organization"))):
            ctk.CTkLabel(top, text=label, anchor="w").grid(row=i * 2, column=0, padx=24, pady=(14 if i == 0 else 6, 0), sticky="w")
            ent = ctk.CTkEntry(top, width=432, height=34)
            ent.insert(0, case.get(key, ""))
            ent.grid(row=i * 2 + 1, column=0, padx=24)
            entries[key] = ent
        ctk.CTkLabel(top, text="Notes", anchor="w").grid(row=6, column=0, padx=24, pady=(6, 0), sticky="w")
        notes = ctk.CTkTextbox(top, width=432, height=130)
        notes.insert("1.0", case.get("notes", ""))
        notes.grid(row=7, column=0, padx=24)

        def save():
            for k, e in entries.items():
                case[k] = e.get().strip()
            case["notes"] = notes.get("1.0", "end").strip()
            self.settings.save()
            self.audit.add("CASE INFO", f"Updated (case {case.get('case_id') or '-'}, examiner {case.get('examiner') or '-'})")
            top.destroy()

        ctk.CTkButton(top, text="Save", height=36, command=save).grid(row=8, column=0, padx=24, pady=18, sticky="e")

    def open_settings(self):
        top = self._dialog("Settings", 460, 460)
        st = self.settings
        ctk.CTkLabel(top, text="Appearance", anchor="w").grid(row=0, column=0, padx=24, pady=(18, 2), sticky="w")
        mode = ctk.CTkOptionMenu(top, values=["Dark", "Light", "System"], width=200)
        mode.set(st["appearance"])
        mode.grid(row=1, column=0, padx=24, sticky="w")
        sc_lbl = ctk.CTkLabel(top, text=f"Interface scale: {st['scaling']}%", anchor="w")
        sc_lbl.grid(row=2, column=0, padx=24, pady=(14, 2), sticky="w")
        scale = ctk.CTkSlider(top, from_=80, to=150, number_of_steps=14, width=300, command=lambda v: sc_lbl.configure(text=f"Interface scale: {int(v)}%"))
        scale.set(st["scaling"])
        scale.grid(row=3, column=0, padx=24, sticky="w")
        rec_var = tk.BooleanVar(value=bool(st["recursive"]))
        ctk.CTkCheckBox(top, text="Include sub-folders when adding a folder", variable=rec_var).grid(row=4, column=0, padx=24, pady=16, sticky="w")
        ctk.CTkLabel(top, text="Default minimum string length", anchor="w").grid(row=5, column=0, padx=24, sticky="w")
        msl = ctk.CTkEntry(top, width=80)
        msl.insert(0, str(st["min_string_len"]))
        msl.grid(row=6, column=0, padx=24, pady=(2, 10), sticky="w")
        ctk.CTkLabel(top, text=f"Settings folder:\n{CONFIG_DIR}", anchor="w", justify="left", text_color=C["muted"], font=ctk.CTkFont(size=11)).grid(row=7, column=0, padx=24, pady=8, sticky="w")

        def save():
            st["appearance"] = mode.get()
            st["scaling"] = int(scale.get())
            st["recursive"] = bool(rec_var.get())
            try:
                st["min_string_len"] = max(3, int(msl.get()))
            except ValueError:
                pass
            st.save()
            ctk.set_appearance_mode(st["appearance"])
            ctk.set_widget_scaling(st["scaling"] / 100.0)
            self._apply_tree_style()
            for t in (self.ev_tree, self.meta_tree, self.tl_tree, self.cmp_tree):
                self._style_tree_tags(t)
            self.audit.add("SETTINGS", f"appearance={st['appearance']} scale={st['scaling']}%")
            top.destroy()

        row = ctk.CTkFrame(top, fg_color="transparent")
        row.grid(row=8, column=0, padx=24, pady=10, sticky="ew")
        ctk.CTkButton(row, text="About", width=90, fg_color=C["neutral"], hover_color=C["neutral_hover"], text_color=C["text"],
                      command=lambda: messagebox.showinfo(APP_NAME, f"{APP_NAME} {APP_VERSION}\nForensic metadata analysis suite.\n\nIndicators are investigative leads, not proof of tampering.")).pack(side="left")
        ctk.CTkButton(row, text="Save", width=110, command=save).pack(side="right")

    # ------------------------------------------------------------------ errors / shutdown
    def _log_error(self, exc):
        import traceback

        try:
            CONFIG_DIR.mkdir(parents=True, exist_ok=True)
            with open(CONFIG_DIR / "error.log", "a", encoding="utf-8") as fh:
                fh.write(f"{datetime.datetime.now():%Y-%m-%d %H:%M:%S}\n{''.join(traceback.format_exception(type(exc), exc, exc.__traceback__))}\n")
        except Exception:
            pass

    def _on_tk_error(self, exc, val, tb):
        self._log_error(val)
        try:
            self._status(f"Internal error: {val}  (details in {CONFIG_DIR / 'error.log'})")
        except Exception:
            pass

    def _on_close(self):
        self.epoch += 1
        self.audit.add("SESSION", "Closed")
        try:
            self.settings["ela_quality"] = int(self.ela_slider.get())
            self.settings.save()
        except Exception:
            pass
        try:
            self.pool.shutdown(wait=False, cancel_futures=True)
        except TypeError:
            self.pool.shutdown(wait=False)
        self.destroy()


# ============================================================================
# ICON LOADER + .EXE BUILDER + ENTRY POINT
# ============================================================================
def load_app_icon():
    data = base64.b64decode("".join(APP_ICON_B64.split()))
    return Image.open(io.BytesIO(data)).convert("RGBA")


def build_exe():
    """python Metasleuth.py --build-exe   ->   dist/MetaSleuth.exe (icon included)."""
    script = Path(__file__).resolve()
    out = script.parent
    py = sys.executable
    subprocess.check_call([py, "-m", "pip", "install", "--upgrade", "pyinstaller", "customtkinter", "pillow", "pypdf", "mutagen"])
    has_dnd = subprocess.call([py, "-m", "pip", "install", "tkinterdnd2"]) == 0
    ico = out / "metasleuth.ico"
    load_app_icon().save(ico, format="ICO", sizes=[(16, 16), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)])
    cmd = [
        py, "-m", "PyInstaller", "--noconfirm", "--clean", "--onefile", "--windowed",
        "--name", "MetaSleuth", "--icon", str(ico),
        "--collect-all", "customtkinter",
        "--collect-submodules", "mutagen",
        "--collect-submodules", "pypdf",
        "--distpath", str(out / "dist"), "--workpath", str(out / "build"), "--specpath", str(out),
    ]
    if has_dnd:
        cmd += ["--collect-all", "tkinterdnd2"]
    cmd.append(str(script))
    subprocess.check_call(cmd)
    print("\nDONE ->", out / "dist" / "MetaSleuth.exe")


def main():
    if "--build-exe" in sys.argv:
        build_exe()
        return
    app = MetaSleuthApp()
    app.mainloop()


if __name__ == "__main__":
    main()
