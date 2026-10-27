# core package — install the startup accelerator before any submodule can
# import PyMuPDF, so the heavy AI-analysis stack never loads (see fast_pymupdf).
from . import fast_pymupdf as _fast_pymupdf

_fast_pymupdf.install()
