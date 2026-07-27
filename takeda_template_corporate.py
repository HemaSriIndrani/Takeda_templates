"""
takeda_template_corporate.py — CORPORATE template loader (disk-backed).

The 2026 corporate template now ships as a real binary attachment,
`Takeda_Slide_Template_EN.potx`, instead of a gzip+base64 blob embedded here.
Attach that .potx to the project (or drop it in one of the search dirs below).

USE — at the top of a generation script:
      import sys; sys.path.insert(0, '/mnt/project'); import takeda_template_corporate
  Then: Deck(brand="CORPORATE") auto-loads the template.

API is unchanged: template_path(brand) / chrome_path(name) / CHROME_GEOMETRY.
`takeda.py` already prefers a .potx found on disk, so importing this module is
optional for corporate decks when the .potx is present.
"""
import base64 as _b64, gzip as _gz, os as _os, glob as _glob
import takeda_remote as _remote

_CACHE_DIR = "/tmp/takeda_templates"
_os.makedirs(_CACHE_DIR, exist_ok=True)
BRAND = "CORPORATE"

# Where we look for the attached corporate template, in priority order.
_SEARCH_DIRS = ["/mnt/project", "/home/claude", "/mnt/user-data/uploads", _os.getcwd()]
_TEMPLATE_NAMES = ["Takeda_Slide_Template_EN.potx", "Takeda_Slide_Template_EN.pptx"]

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

def _find_potx():
    for d in _SEARCH_DIRS:
        for n in _TEMPLATE_NAMES:
            p = _os.path.join(d, n)
            if _os.path.exists(p):
                return p
        hits = sorted(_glob.glob(_os.path.join(d, "Takeda*.pot*"))) if _os.path.isdir(d) else []
        if hits:
            return hits[0]
    return None

def template_path(brand="CORPORATE"):
    """Path to the corporate .potx on disk. Raises if the attachment is missing."""
    p = _find_potx()
    if p:
        return p
    return _remote.get("Takeda_Slide_Template_EN.potx")  # download from GitHub

def _decode(s):  # kept for API parity (chrome blobs, if any brand registers them)
    return _gz.decompress(_b64.b64decode(s))

import builtins as _bi
if not hasattr(_bi, "_TAKEDA_TEMPLATE_REGISTRY"): _bi._TAKEDA_TEMPLATE_REGISTRY = {}
if not hasattr(_bi, "_TAKEDA_CHROME_REGISTRY"):   _bi._TAKEDA_CHROME_REGISTRY   = {}
if not hasattr(_bi, "_TAKEDA_CHROME_GEOMETRY"):   _bi._TAKEDA_CHROME_GEOMETRY   = {}
_bi._TAKEDA_CHROME_GEOMETRY.update(CHROME_GEOMETRY)

# Corporate is disk-backed: no base64 blob goes in the template registry.
# takeda.py resolves CORPORATE via disk first, then falls back to template_path().

_EXT = {'ENTYVIO_BANNER': 'jpg'}

def chrome_path(name):
    """Path to a decoded chrome image (for brands that register chrome blobs)."""
    reg = _bi._TAKEDA_CHROME_REGISTRY
    if name not in reg:
        raise KeyError("Chrome %r not loaded. Import the owning brand module. Loaded: %s"
                       % (name, sorted(reg)))
    ext = _EXT.get(name, "png")
    path = _os.path.join(_CACHE_DIR, "chrome_%s.%s" % (name.lower(), ext))
    if not _os.path.exists(path):
        with open(path, "wb") as f:
            f.write(_decode(reg[name]))
    return path

_LEGACY_CACHE = "/tmp/takeda_template_cache.potx"
_p = _find_potx()
if _p and not _os.path.exists(_LEGACY_CACHE):
    try:
        import shutil as _sh; _sh.copyfile(_p, _LEGACY_CACHE)
    except Exception as _e:
        import warnings; warnings.warn("corporate template cache copy failed: %s" % _e)
