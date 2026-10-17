"""Helpers for drawing figures with Thai text in the thesis font.

    import sys; sys.path.insert(0, r"<skill>/scripts")
    from thai_fonts import setup_matplotlib, pil_font
    setup_matplotlib()                 # matplotlib now uses TH Sarabun PSK
    font = pil_font(28)                # ImageFont for PIL mock screenshots

TH SarabunPSK is often installed per user (%LOCALAPPDATA%\\Microsoft\\Windows\\
Fonts), where matplotlib does not look, and its family name inside
matplotlib is "TH Sarabun PSK" (with a space). PIL without libraqm cannot
shape Thai, so upper vowels / tone marks may sit slightly off - keep Thai
in PIL mock-ups short, or draw the mock-up with matplotlib instead.
"""
import glob, os

_DIRS = [os.path.join(os.environ.get("LOCALAPPDATA", ""), "Microsoft", "Windows", "Fonts"),
         os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts")]


def font_file(bold=False):
    for d in _DIRS:
        for f in glob.glob(os.path.join(d, "THSarabun*")):
            name = os.path.basename(f).lower()
            if ("bold" in name) == bold and "italic" not in name:
                return f
    raise FileNotFoundError("TH SarabunPSK is not installed")


def setup_matplotlib(size=16):
    import matplotlib
    from matplotlib import font_manager
    for b in (False, True):
        try:
            font_manager.fontManager.addfont(font_file(b))
        except FileNotFoundError:
            pass
    matplotlib.rcParams.update({"font.family": "TH Sarabun PSK", "font.size": size,
                                "axes.unicode_minus": False, "savefig.dpi": 200})


def pil_font(size, bold=False):
    from PIL import ImageFont
    return ImageFont.truetype(font_file(bold), size)
