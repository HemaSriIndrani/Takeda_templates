"""
takeda_template_entyvio.py — ENTYVIO template (self-contained) + chrome.
Part of the SPLIT (Option 4b) build: each brand template lives in its OWN file so a
project loads only the brand(s) it needs. Template ships as a disk attachment (takeda_template_entyvio.pptx); chrome stays inline.

This module is self-contained: the ENTYVIO template carries its own masters, theme,
layouts, and chrome — it does NOT require the corporate template at runtime.

USE — at the top of a generation script, import ONLY the brand module(s) you need:
      import sys; sys.path.insert(0, '/mnt/project'); import takeda_template_entyvio
  Then: Deck(brand="ENTYVIO") auto-loads this template.

The four template modules (corporate + 3 brands) share an identical API:
  template_path(brand) / chrome_path(name) / CHROME_GEOMETRY
Importing any one registers that brand. Import several to use several. `takeda.py`
discovers whichever modules are present.
"""
import base64 as _b64, gzip as _gz, os as _os

_CACHE_DIR = "/tmp/takeda_templates"
_os.makedirs(_CACHE_DIR, exist_ok=True)
BRAND = "ENTYVIO"

CHROME_GEOMETRY = {
    "FRUQ_BANNER": {
        "L": 0.0,
        "T": 0.0,
        "W": 13.333,
        "H": 7.5,
        "role": "full-bleed gradient (banner is top portion; thin line ~0.5in)"
    },
    "FRUQ_LOGO": {
        "L": 10.7,
        "T": 6.44,
        "W": 2.24,
        "H": 0.9,
        "role": "bottom-right logo, whitespace-contingent"
    },
    "ENTYVIO_BANNER": {
        "L": 0.0,
        "T": 7.15,
        "W": 13.333,
        "H": 0.35,
        "role": "bottom full-width gradient strip"
    },
    "ICLUSIG_LOGO": {
        "L": 11.4,
        "T": 6.59,
        "W": 1.56,
        "H": 0.67,
        "role": "bottom-right logo, whitespace-contingent"
    }
}

CHROME_ENTYVIO_BANNER_B64GZ = (
    "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")

def _decode(s): return _gz.decompress(_b64.b64decode(s))

_TEMPLATE_NAME = "takeda_template_entyvio.pptx"
_SEARCH_DIRS = ["/mnt/project", "/home/claude", "/mnt/user-data/uploads", _os.getcwd()]

def _find_brand_pptx():
    for d in _SEARCH_DIRS:
        p = _os.path.join(d, _TEMPLATE_NAME)
        if _os.path.exists(p):
            return p
    try:
        import takeda_remote as _remote
        return _remote.get(_TEMPLATE_NAME)   # download from GitHub
    except Exception:
        return None

def _seed_template_from_disk():
    """Copy the attached brand .pptx into the /tmp cache path takeda.py expects."""
    src = _find_brand_pptx()
    dst = _os.path.join(_CACHE_DIR, "takeda_entyvio.pptx")
    if src and not _os.path.exists(dst):
        import shutil as _sh; _sh.copyfile(src, dst)

def template_path(brand="ENTYVIO"):
    """Path to the ENTYVIO template .pptx (disk attachment: takeda_template_entyvio.pptx)."""
    p = _find_brand_pptx()
    if not p:
        raise FileNotFoundError(
            "ENTYVIO template not found. Attach 'takeda_template_entyvio.pptx' to the project "
            "or place it in one of: " + repr(_SEARCH_DIRS))
    return p

# Register into a shared module-level registry that takeda.py reads. Using a global
# dict on a well-known module name lets multiple brand modules coexist.
import builtins as _bi
if not hasattr(_bi, "_TAKEDA_TEMPLATE_REGISTRY"):
    _bi._TAKEDA_TEMPLATE_REGISTRY = {}
if not hasattr(_bi, "_TAKEDA_CHROME_REGISTRY"):
    _bi._TAKEDA_CHROME_REGISTRY = {}
if not hasattr(_bi, "_TAKEDA_CHROME_GEOMETRY"):
    _bi._TAKEDA_CHROME_GEOMETRY = {}

_bi._TAKEDA_TEMPLATE_REGISTRY["ENTYVIO"] = "__DISK__"  # template is a disk attachment now
try:
    _seed_template_from_disk()
except Exception:
    pass
_bi._TAKEDA_CHROME_GEOMETRY.update(CHROME_GEOMETRY)
_bi._TAKEDA_CHROME_REGISTRY["ENTYVIO_BANNER"] = CHROME_ENTYVIO_BANNER_B64GZ

_EXT = {'ENTYVIO_BANNER': 'jpg'}


def chrome_path(name):
    """Path to a decoded chrome image (gunzipped to /tmp on first call)."""
    reg = _bi._TAKEDA_CHROME_REGISTRY
    if name not in reg:
        raise KeyError(f"Chrome {name!r} not loaded. Import the owning brand module. "
                       f"Loaded: {sorted(reg)}")
    ext = _EXT.get(name, "png")
    path = _os.path.join(_CACHE_DIR, f"chrome_{name.lower()}.{ext}")
    if not _os.path.exists(path):
        with open(path, "wb") as f:
            f.write(_decode(reg[name]))
    return path
