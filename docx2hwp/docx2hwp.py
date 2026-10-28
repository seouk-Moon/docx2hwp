#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
docx2hwp — Word(.docx) 문서를 한글(.hwpx / .hwp) 문서로 변환

  * 순수 파이썬(표준 라이브러리만) → Windows / macOS / Linux 어디서나 .hwpx 생성
  * Windows 에 한컴오피스 한글이 설치돼 있으면 .hwp(바이너리)로도 저장
  * Word 수식(OMML)은 한글 수식(편집 가능한 수식 개체)으로 변환

사용법
  python docx2hwp.py 문서.docx                 → 문서.hwpx
  python docx2hwp.py 문서.docx -o 결과.hwp      → (한글 설치 시) 결과.hwp
  python docx2hwp.py 폴더/ --format both       → 폴더 안 모든 docx 를 hwpx+hwp 로
  python docx2hwp.py --gui                     → 파일 선택 창
"""
from __future__ import print_function

import argparse
import datetime
import glob
import math
import os
import posixpath
import random
import re
import struct
import sys
import zipfile
import xml.etree.ElementTree as ET

try:
    from omml2hwpeq import OmmlConverter
except ImportError:  # 같은 폴더가 아닌 곳에서 실행될 때
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    from omml2hwpeq import OmmlConverter
import omml2hwpeq as eqlearn

__version__ = "1.0.0"

# 한컴 공개 문서 이용 조건에 따른 표기
HANCOM_NOTICE = "본 제품은 한글과컴퓨터의 한글 문서 파일(.hwp) 공개 문서를 참고하여 개발하였습니다."

# ---------------------------------------------------------------------------
# 네임스페이스
# ---------------------------------------------------------------------------
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
WPS = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
V = "urn:schemas-microsoft-com:vml"
PR = "http://schemas.openxmlformats.org/package/2006/relationships"
O = "urn:schemas-microsoft-com:office:office"


def w(tag):
    return "{%s}%s" % (W, tag)


def local(el):
    t = el.tag
    return t.split("}", 1)[1] if isinstance(t, str) and "}" in t else t


def nsof(el):
    t = el.tag
    return t[1:].split("}", 1)[0] if isinstance(t, str) and t.startswith("{") else ""


def wval(el, name=None, default=None):
    """<w:name w:val=".."/> 값 읽기. name 이 None 이면 el 자체."""
    if el is None:
        return default
    if name is not None:
        el = el.find(w(name))
        if el is None:
            return default
    v = el.get(w("val"))
    return default if v is None else v


def wbool(el, name):
    """토글 속성(<w:b/>, <w:b w:val="0"/>) → True/False/None"""
    if el is None:
        return None
    c = el.find(w(name))
    if c is None:
        return None
    v = c.get(w("val"))
    return v is None or v not in ("0", "false", "off", "none")


def to_int(v, default=0):
    try:
        return int(round(float(v)))
    except (TypeError, ValueError):
        return default


_XML_BAD = re.compile("[\x00-\x08\x0b\x0c\x0e-\x1f￾￿]")


def esc(s):
    s = _XML_BAD.sub("", s)
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def attr(s):
    return esc(str(s)).replace('"', "&quot;")


# 단위: HWPUNIT = 1/7200 inch,  twip = 1/1440 inch,  EMU = 1/914400 inch
def tw(v):            # twip → HWPUNIT
    return int(round(to_int(v) * 5))


def emu(v):           # EMU → HWPUNIT
    return int(round(to_int(v) / 127.0))


def pt(v):            # point → HWPUNIT
    return int(round(float(v) * 100))


HP_NS_DECL = (
    'xmlns:ha="http://www.hancom.co.kr/hwpml/2011/app" '
    'xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph" '
    'xmlns:hp10="http://www.hancom.co.kr/hwpml/2016/paragraph" '
    'xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section" '
    'xmlns:hc="http://www.hancom.co.kr/hwpml/2011/core" '
    'xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head" '
    'xmlns:hhs="http://www.hancom.co.kr/hwpml/2011/history" '
    'xmlns:hm="http://www.hancom.co.kr/hwpml/2011/master-page" '
    'xmlns:hpf="http://www.hancom.co.kr/schema/2011/hpf" '
    'xmlns:dc="http://purl.org/dc/elements/1.1/" '
    'xmlns:opf="http://www.idpf.org/2007/opf/" '
    'xmlns:ooxmlchart="http://www.hancom.co.kr/hwpml/2016/ooxmlchart" '
    'xmlns:hwpunitchar="http://www.hancom.co.kr/hwpml/2016/HwpUnitChar" '
    'xmlns:epub="http://www.idpf.org/2007/ops" '
    'xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0"'
)
XML_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>'
HUC = "http://www.hancom.co.kr/hwpml/2016/HwpUnitChar"

LANGS = ("hangul", "latin", "hanja", "japanese", "other", "symbol", "user")

HIGHLIGHT = {
    "yellow": "#FFFF00", "green": "#00FF00", "cyan": "#00FFFF",
    "magenta": "#FF00FF", "blue": "#0000FF", "red": "#FF0000",
    "darkBlue": "#000080", "darkCyan": "#008080", "darkGreen": "#008000",
    "darkMagenta": "#800080", "darkRed": "#800000", "darkYellow": "#808000",
    "darkGray": "#808080", "lightGray": "#C0C0C0", "black": "#000000",
    "white": "#FFFFFF",
}

LINE_TYPES = {
    "single": "SOLID", "thick": "SOLID", "double": "DOUBLE_SLIM",
    "dotted": "DOT", "dashed": "DASH", "dashSmallGap": "DASH",
    "dotDash": "DASH_DOT", "dotDotDash": "DASH_DOT_DOT",
    "triple": "TRIPLE_SLIM" if False else "SLIM_THICK_SLIM",
    "thinThickSmallGap": "SLIM_THICK", "thickThinSmallGap": "THICK_SLIM",
    "thinThickThinSmallGap": "SLIM_THICK_SLIM", "wave": "WAVE",
    "doubleWave": "DOUBLE_WAVE", "dashDotStroked": "DASH_DOT",
    "threeDEmboss": "SOLID", "threeDEngrave": "SOLID", "outset": "SOLID",
    "inset": "SOLID",
}

UNDERLINE = {
    "single": "SOLID", "words": "SOLID", "double": "DOUBLE_SLIM",
    "thick": "SOLID", "dotted": "DOT", "dottedHeavy": "DOT", "dash": "DASH",
    "dashedHeavy": "DASH", "dashLong": "LONG_DASH", "dashLongHeavy": "LONG_DASH",
    "dotDash": "DASH_DOT", "dashDotHeavy": "DASH_DOT", "dotDotDash": "DASH_DOT_DOT",
    "dashDotDotHeavy": "DASH_DOT_DOT", "wave": "WAVE", "wavyHeavy": "WAVE",
    "wavyDouble": "DOUBLE_WAVE",
}

BORDER_WIDTHS = [(0.1, "0.1 mm"), (0.12, "0.12 mm"), (0.15, "0.15 mm"),
                 (0.2, "0.2 mm"), (0.25, "0.25 mm"), (0.3, "0.3 mm"),
                 (0.4, "0.4 mm"), (0.5, "0.5 mm"), (0.6, "0.6 mm"),
                 (0.7, "0.7 mm"), (1.0, "1.0 mm"), (1.5, "1.5 mm"),
                 (2.0, "2.0 mm"), (3.0, "3.0 mm"), (4.0, "4.0 mm"),
                 (5.0, "5.0 mm")]


def border_width(sz_eighths):
    mm = to_int(sz_eighths, 4) / 8.0 * 0.3528
    best = min(BORDER_WIDTHS, key=lambda x: abs(x[0] - mm))
    return best[1]


def norm_color(c, default="#000000"):
    if not c or c == "auto":
        return default
    c = c.strip().lstrip("#")
    if re.match(r"^[0-9A-Fa-f]{6}$", c):
        return "#" + c.upper()
    return default


# Symbol / Wingdings 글꼴 문자 → 유니코드 (자주 쓰이는 것만)
SYMBOL_FONT = {
    0x61: "α", 0x62: "β", 0x63: "χ", 0x64: "δ", 0x65: "ε", 0x66: "φ",
    0x67: "γ", 0x68: "η", 0x69: "ι", 0x6A: "ϕ", 0x6B: "κ", 0x6C: "λ",
    0x6D: "μ", 0x6E: "ν", 0x6F: "ο", 0x70: "π", 0x71: "θ", 0x72: "ρ",
    0x73: "σ", 0x74: "τ", 0x75: "υ", 0x76: "ϖ", 0x77: "ω", 0x78: "ξ",
    0x79: "ψ", 0x7A: "ζ", 0x41: "Α", 0x42: "Β", 0x43: "Χ", 0x44: "Δ",
    0x45: "Ε", 0x46: "Φ", 0x47: "Γ", 0x48: "Η", 0x49: "Ι", 0x4B: "Κ",
    0x4C: "Λ", 0x4D: "Μ", 0x4E: "Ν", 0x4F: "Ο", 0x50: "Π", 0x51: "Θ",
    0x52: "Ρ", 0x53: "Σ", 0x54: "Τ", 0x55: "Υ", 0x57: "Ω", 0x58: "Ξ",
    0x59: "Ψ", 0x5A: "Ζ", 0xB7: "•", 0xB0: "°", 0xB1: "±", 0xB4: "×",
    0xB8: "÷", 0xB9: "≠", 0xBA: "≡", 0xBB: "≈", 0xA3: "≤", 0xB3: "≥",
    0xA5: "∞", 0xB6: "∂", 0xD6: "√", 0xE5: "∑", 0xF2: "∫", 0xAE: "→",
    0xAC: "←", 0xDE: "⇒", 0xDB: "⇔", 0xCE: "∈", 0xC7: "∩", 0xC8: "∪",
    0xC6: "∅", 0xA2: "′", 0xBC: "…", 0xD7: "·", 0x2D: "−", 0x3D: "=",
}
WINGDINGS = {
    0xA7: "▪", 0xA8: "□", 0x6C: "●", 0x6E: "■", 0x71: "❑", 0x75: "◆",
    0x76: "❖", 0xD8: "➢", 0xFC: "✓", 0xFB: "✗", 0xFE: "☑", 0x9F: "•",
    0xA1: "○", 0xE0: "➔", 0xE8: "➔", 0x77: "◆", 0x70: "□", 0x6F: "□",
    0x4A: "☺", 0x46: "☞", 0xF0: "⇨",
}


def map_symbol_char(code, font):
    font = (font or "").lower()
    if code >= 0xF000:
        code -= 0xF000
    if "wingding" in font:
        return WINGDINGS.get(code, "•")
    if "symbol" in font:
        return SYMBOL_FONT.get(code, chr(code) if code >= 0x20 else "")
    return chr(code)


def fix_pua(text, font=None):
    """Symbol 계열 글꼴의 사설영역(U+F0xx) 문자를 일반 유니코드로."""
    if not text or not any(0xF000 <= ord(c) <= 0xF0FF for c in text):
        return text
    return "".join(map_symbol_char(ord(c), font or "symbol")
                   if 0xF000 <= ord(c) <= 0xF0FF else c for c in text)


# ---------------------------------------------------------------------------
# 이미지 크기 읽기 (PNG / JPEG / GIF / BMP)
# ---------------------------------------------------------------------------
def image_pixel_size(data):
    try:
        if data[:8] == b"\x89PNG\r\n\x1a\n":
            wpx, hpx = struct.unpack(">II", data[16:24])
            return wpx, hpx
        if data[:6] in (b"GIF87a", b"GIF89a"):
            wpx, hpx = struct.unpack("<HH", data[6:10])
            return wpx, hpx
        if data[:2] == b"BM":
            wpx, hpx = struct.unpack("<ii", data[18:26])
            return wpx, abs(hpx)
        if data[:2] == b"\xff\xd8":
            i = 2
            while i < len(data) - 9:
                if data[i] != 0xFF:
                    i += 1
                    continue
                marker = data[i + 1]
                if marker in (0xC0, 0xC1, 0xC2, 0xC3, 0xC5, 0xC6, 0xC7,
                              0xC9, 0xCA, 0xCB, 0xCD, 0xCE, 0xCF):
                    hpx, wpx = struct.unpack(">HH", data[i + 5:i + 9])
                    return wpx, hpx
                seglen = struct.unpack(">H", data[i + 2:i + 4])[0]
                i += 2 + seglen
    except Exception:
        pass
    return None


MEDIA_TYPES = {
    "png": "image/png", "jpg": "image/jpg", "jpeg": "image/jpg",
    "gif": "image/gif", "bmp": "image/bmp", "wmf": "image/x-wmf",
    "emf": "image/x-emf", "tif": "image/tif", "tiff": "image/tif",
    "svg": "image/svg+xml",
}


# ---------------------------------------------------------------------------
# DOCX 패키지 읽기
# ---------------------------------------------------------------------------
class Docx(object):
    def __init__(self, path):
        self.path = path
        self.zip = zipfile.ZipFile(path)
        self.names = set(self.zip.namelist())
        self.main = self._main_part()
        self.doc = self.xml(self.main)
        self.rels = self.load_rels(self.main)
        self.styles = {}
        self.default_pstyle = None
        self.default_cstyle = None
        self.default_tstyle = None
        self.doc_rpr = {}
        self.doc_ppr = {}
        self.theme = {}
        self.abstract = {}
        self.nums = {}
        self.default_tab = 720
        self.even_odd = False
        self.footnotes = {}
        self.endnotes = {}
        self._load_theme()
        self._load_styles()
        self._load_numbering()
        self._load_settings()
        self._load_notes()

    def _main_part(self):
        try:
            root = ET.fromstring(self.zip.read("_rels/.rels"))
            for r in root:
                if r.get("Type", "").endswith("/officeDocument"):
                    return r.get("Target").lstrip("/")
        except KeyError:
            pass
        return "word/document.xml"

    def has(self, name):
        return name in self.names

    def read(self, name):
        return self.zip.read(name)

    def xml(self, name):
        if name not in self.names:
            return None
        return ET.fromstring(self.zip.read(name))

    def load_rels(self, part):
        d = posixpath.dirname(part)
        relp = posixpath.join(d, "_rels", posixpath.basename(part) + ".rels")
        out = {}
        root = self.xml(relp)
        if root is None:
            return out
        for r in root:
            tgt = r.get("Target", "")
            mode = r.get("TargetMode")
            if mode != "External":
                if tgt.startswith("/"):
                    tgt = tgt.lstrip("/")
                else:
                    tgt = posixpath.normpath(posixpath.join(d, tgt))
            out[r.get("Id")] = (r.get("Type", "").rsplit("/", 1)[-1], tgt, mode)
        return out

    def part_for_type(self, rtype):
        for _rid, (t, tgt, _m) in self.rels.items():
            if t == rtype:
                return tgt
        return None

    # -- 테마 글꼴 ------------------------------------------------------------
    def _load_theme(self):
        part = self.part_for_type("theme")
        root = self.xml(part) if part else None
        if root is None:
            return
        for kind in ("majorFont", "minorFont"):
            f = root.find(".//{%s}%s" % (A, kind))
            if f is None:
                continue
            latin = f.find("{%s}latin" % A)
            ea = f.find("{%s}ea" % A)
            hang = None
            for fnt in f.findall("{%s}font" % A):
                if fnt.get("script") == "Hang":
                    hang = fnt.get("typeface")
            key = "major" if kind == "majorFont" else "minor"
            self.theme[key + "Latin"] = latin.get("typeface") if latin is not None else ""
            eav = ea.get("typeface") if ea is not None else ""
            self.theme[key + "EA"] = eav or hang or ""

    def theme_font(self, name):
        if not name:
            return None
        major = name.startswith("major")
        key = ("major" if major else "minor") + ("EA" if "EastAsia" in name else "Latin")
        v = self.theme.get(key)
        if not v and key.endswith("Latin"):
            v = None
        return v or None

    # -- 스타일 ---------------------------------------------------------------
    def _load_styles(self):
        part = self.part_for_type("styles") or "word/styles.xml"
        root = self.xml(part)
        if root is None:
            return
        dd = root.find(w("docDefaults"))
        if dd is not None:
            rpr = dd.find("%s/%s" % (w("rPrDefault"), w("rPr")))
            ppr = dd.find("%s/%s" % (w("pPrDefault"), w("pPr")))
            self.doc_rpr = self.parse_rpr(rpr)
            self.doc_ppr = parse_ppr(ppr)
        for st in root.findall(w("style")):
            sid = st.get(w("styleId"))
            typ = st.get(w("type"))
            info = {
                "type": typ,
                "name": wval(st, "name", sid),
                "basedOn": wval(st, "basedOn"),
                "rpr": self.parse_rpr(st.find(w("rPr"))),
                "ppr": parse_ppr(st.find(w("pPr"))),
                "tblPr": st.find(w("tblPr")),
                "tcPr": st.find(w("tcPr")),
            }
            self.styles[sid] = info
            if st.get(w("default")) in ("1", "true"):
                if typ == "paragraph":
                    self.default_pstyle = sid
                elif typ == "character":
                    self.default_cstyle = sid
                elif typ == "table":
                    self.default_tstyle = sid

    def style_chain(self, sid):
        chain = []
        seen = set()
        while sid and sid in self.styles and sid not in seen:
            seen.add(sid)
            chain.append(self.styles[sid])
            sid = self.styles[sid]["basedOn"]
        return list(reversed(chain))

    def style_rpr(self, sid):
        out = {}
        for s in self.style_chain(sid):
            out.update(s["rpr"])
        return out

    def style_ppr(self, sid):
        out = {}
        for s in self.style_chain(sid):
            merge_ppr(out, s["ppr"])
        return out

    def parse_rpr(self, rpr):
        d = {}
        if rpr is None:
            return d
        f = rpr.find(w("rFonts"))
        if f is not None:
            for key, attr_names in (("ascii", ("ascii", "asciiTheme")),
                                    ("eastAsia", ("eastAsia", "eastAsiaTheme")),
                                    ("hAnsi", ("hAnsi", "hAnsiTheme"))):
                v = f.get(w(attr_names[0]))
                th = f.get(w(attr_names[1]))
                if th:
                    tv = self.theme_font(th)
                    if tv:
                        v = tv
                if v:
                    d[key] = v
        sz = wval(rpr, "sz")
        if sz is None:
            sz = wval(rpr, "szCs") if rpr.find(w("sz")) is None and False else None
        if sz is not None:
            d["sz"] = to_int(sz, 20)
        for name in ("b", "i", "strike", "dstrike", "caps", "smallCaps",
                     "vanish", "emboss", "imprint", "outline", "shadow"):
            v = wbool(rpr, name)
            if v is not None:
                d[name] = v
        u = rpr.find(w("u"))
        if u is not None:
            d["u"] = u.get(w("val"), "single")
            d["u_color"] = u.get(w("color"))
        c = wval(rpr, "color")
        if c is not None:
            d["color"] = c
        hl = wval(rpr, "highlight")
        if hl is not None:
            d["highlight"] = hl
        shd = rpr.find(w("shd"))
        if shd is not None:
            fill = shd.get(w("fill"))
            if fill and fill != "auto":
                d["shd"] = fill
        va = wval(rpr, "vertAlign")
        if va is not None:
            d["vert"] = va
        sp = wval(rpr, "spacing")
        if sp is not None:
            d["spacing"] = to_int(sp)
        sc = wval(rpr, "w")
        if sc is not None:
            d["scale"] = to_int(sc, 100)
        pos = wval(rpr, "position")
        if pos is not None:
            d["position"] = to_int(pos)
        rs = wval(rpr, "rStyle")
        if rs is not None:
            d["rStyle"] = rs
        return d

    # -- 번호 매기기 ----------------------------------------------------------
    def _load_numbering(self):
        part = self.part_for_type("numbering")
        root = self.xml(part) if part else None
        if root is None:
            return
        for an in root.findall(w("abstractNum")):
            aid = an.get(w("abstractNumId"))
            lvls = {}
            link = wval(an, "numStyleLink")
            for lvl in an.findall(w("lvl")):
                il = to_int(lvl.get(w("ilvl")), 0)
                lvls[il] = self._parse_lvl(lvl)
            self.abstract[aid] = {"lvls": lvls, "link": link}
        for num in root.findall(w("num")):
            nid = num.get(w("numId"))
            aid = wval(num, "abstractNumId")
            overrides = {}
            for ov in num.findall(w("lvlOverride")):
                il = to_int(ov.get(w("ilvl")), 0)
                o = {}
                so = wval(ov, "startOverride")
                if so is not None:
                    o["start"] = to_int(so, 1)
                lvl = ov.find(w("lvl"))
                if lvl is not None:
                    o["lvl"] = self._parse_lvl(lvl)
                overrides[il] = o
            self.nums[nid] = {"abstract": aid, "overrides": overrides}

    def _parse_lvl(self, lvl):
        rpr = lvl.find(w("rPr"))
        rfonts = rpr.find(w("rFonts")) if rpr is not None else None
        sym_font = rfonts.get(w("ascii")) if rfonts is not None else None
        return {
            "start": to_int(wval(lvl, "start", "1"), 1),
            "fmt": wval(lvl, "numFmt", "decimal"),
            "text": wval(lvl, "lvlText", ""),
            "restart": wval(lvl, "lvlRestart"),
            "suff": wval(lvl, "suff", "tab"),
            "isLgl": wbool(lvl, "isLgl"),
            "ppr": parse_ppr(lvl.find(w("pPr"))),
            "rpr": self.parse_rpr(rpr),
            "font": sym_font,
            "jc": wval(lvl, "lvlJc", "left"),
        }

    def num_level(self, num_id, ilvl):
        num = self.nums.get(num_id)
        if num is None:
            return None, None
        aid = num["abstract"]
        an = self.abstract.get(aid)
        # numStyleLink → 스타일의 numPr 로 연결된 정의
        if an is not None and not an["lvls"] and an.get("link"):
            st = self.styles.get(an["link"])
            if st:
                nid2 = st["ppr"].get("numId")
                if nid2 and nid2 != num_id and nid2 in self.nums:
                    an = self.abstract.get(self.nums[nid2]["abstract"])
        if an is None:
            return None, None
        lvl = an["lvls"].get(ilvl)
        ov = num["overrides"].get(ilvl, {})
        if "lvl" in ov:
            lvl = ov["lvl"]
        return aid, (lvl, ov.get("start"))

    # -- 설정 / 각주 ----------------------------------------------------------
    def _load_settings(self):
        part = self.part_for_type("settings")
        root = self.xml(part) if part else None
        if root is None:
            return
        self.default_tab = to_int(wval(root, "defaultTabStop", "720"), 720) or 720
        self.even_odd = bool(wbool(root, "evenAndOddHeaders"))

    def _load_notes(self):
        for rtype, store, tag in (("footnotes", self.footnotes, "footnote"),
                                  ("endnotes", self.endnotes, "endnote")):
            part = self.part_for_type(rtype)
            root = self.xml(part) if part else None
            if root is None:
                continue
            for fn in root.findall(w(tag)):
                if fn.get(w("type")) in ("separator", "continuationSeparator",
                                         "continuationNotice"):
                    continue
                store[fn.get(w("id"))] = fn


def parse_ppr(ppr):
    d = {}
    if ppr is None:
        return d
    jc = wval(ppr, "jc")
    if jc is not None:
        d["jc"] = jc
    ind = ppr.find(w("ind"))
    if ind is not None:
        for key, names in (("left", ("start", "left")), ("right", ("end", "right")),
                           ("firstLine", ("firstLine",)), ("hanging", ("hanging",))):
            for n in names:
                v = ind.get(w(n))
                if v is not None:
                    d["ind_" + key] = to_int(v)
                    break
        if ind.get(w("firstLine")) is not None:
            d.pop("ind_hanging", None)
            d["ind_hanging_set"] = True
        if ind.get(w("hanging")) is not None:
            d.pop("ind_firstLine", None)
            d["ind_hanging_set"] = True
    sp = ppr.find(w("spacing"))
    if sp is not None:
        for key in ("before", "after", "line"):
            v = sp.get(w(key))
            if v is not None:
                d[key] = to_int(v)
        lr = sp.get(w("lineRule"))
        if lr is not None:
            d["lineRule"] = lr
        if sp.get(w("beforeAutospacing")) in ("1", "true", "on"):
            d["before"] = 280
        if sp.get(w("afterAutospacing")) in ("1", "true", "on"):
            d["after"] = 280
    for name in ("keepNext", "keepLines", "pageBreakBefore", "widowControl",
                 "contextualSpacing", "wordWrap", "bidi", "suppressAutoHyphens"):
        v = wbool(ppr, name)
        if v is not None:
            d[name] = v
    np_ = ppr.find(w("numPr"))
    if np_ is not None:
        nid = wval(np_, "numId")
        il = wval(np_, "ilvl")
        if nid is not None:
            d["numId"] = nid
        if il is not None:
            d["ilvl"] = to_int(il, 0)
    ol = wval(ppr, "outlineLvl")
    if ol is not None:
        d["outline"] = to_int(ol, 9)
    ps = wval(ppr, "pStyle")
    if ps is not None:
        d["pStyle"] = ps
    tabs = ppr.find(w("tabs"))
    if tabs is not None:
        lst = []
        for t in tabs.findall(w("tab")):
            lst.append((t.get(w("val"), "left"), to_int(t.get(w("pos")), 0),
                        t.get(w("leader"), "none")))
        d["tabs"] = lst
    shd = ppr.find(w("shd"))
    if shd is not None:
        fill = shd.get(w("fill"))
        if fill and fill != "auto":
            d["shd"] = fill
    bdr = ppr.find(w("pBdr"))
    if bdr is not None:
        b = {}
        for side in ("top", "left", "bottom", "right"):
            e = bdr.find(w(side))
            if e is not None:
                b[side] = (e.get(w("val"), "single"), e.get(w("sz"), "4"),
                           e.get(w("color"), "auto"))
        d["bdr"] = b
    return d


def merge_ppr(base, over):
    """pPr 병합: 들여쓰기 firstLine/hanging 은 서로 배타."""
    if "ind_hanging_set" in over:
        base.pop("ind_firstLine", None)
        base.pop("ind_hanging", None)
    if "tabs" in over and "tabs" in base:
        merged = {pos: (val, pos, ld) for (val, pos, ld) in base["tabs"]}
        for (val, pos, ld) in over["tabs"]:
            if val == "clear":
                merged.pop(pos, None)
            else:
                merged[pos] = (val, pos, ld)
        over = dict(over)
        over["tabs"] = sorted(merged.values(), key=lambda x: x[1])
    base.update(over)
    return base


# ---------------------------------------------------------------------------
# 번호 문자열
# ---------------------------------------------------------------------------
def roman(n):
    if n <= 0:
        return str(n)
    vals = [(1000, "M"), (900, "CM"), (500, "D"), (400, "CD"), (100, "C"),
            (90, "XC"), (50, "L"), (40, "XL"), (10, "X"), (9, "IX"),
            (5, "V"), (4, "IV"), (1, "I")]
    out = ""
    for v, s in vals:
        while n >= v:
            out += s
            n -= v
    return out


def letters(n, upper):
    if n <= 0:
        return str(n)
    base = "ABCDEFGHIJKLMNOPQRSTUVWXYZ" if upper else "abcdefghijklmnopqrstuvwxyz"
    rep = (n - 1) // 26 + 1
    return base[(n - 1) % 26] * rep


GANADA = "가나다라마바사아자차카타파하"
CHOSUNG = "ㄱㄴㄷㄹㅁㅂㅅㅇㅈㅊㅋㅌㅍㅎ"
KDIGIT = "영일이삼사오육칠팔구"
KCOUNT = ["", "하나", "둘", "셋", "넷", "다섯", "여섯", "일곱", "여덟", "아홉", "열"]
CIRCLED = "①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳"


def format_number(n, fmt):
    if fmt in ("decimal", "decimalHalfWidth"):
        return str(n)
    if fmt == "decimalZero":
        return "%02d" % n
    if fmt == "upperRoman":
        return roman(n)
    if fmt == "lowerRoman":
        return roman(n).lower()
    if fmt == "upperLetter":
        return letters(n, True)
    if fmt == "lowerLetter":
        return letters(n, False)
    if fmt == "ordinal":
        suf = "th" if 10 <= n % 100 <= 20 else {1: "st", 2: "nd", 3: "rd"}.get(n % 10, "th")
        return "%d%s" % (n, suf)
    if fmt in ("ganada", "koreanLegal"):
        return GANADA[(n - 1) % len(GANADA)] if n > 0 else str(n)
    if fmt == "chosung":
        return CHOSUNG[(n - 1) % len(CHOSUNG)] if n > 0 else str(n)
    if fmt in ("koreanDigital", "koreanDigital2"):
        return "".join(KDIGIT[int(c)] for c in str(n))
    if fmt == "koreanCounting":
        return KCOUNT[n] if 0 < n <= 10 else str(n)
    if fmt in ("decimalEnclosedCircle", "decimalEnclosedCircleChinese"):
        return CIRCLED[n - 1] if 0 < n <= 20 else "(%d)" % n
    if fmt == "decimalEnclosedParen":
        return "(%d)" % n
    if fmt == "decimalEnclosedFullstop":
        return "%d." % n
    if fmt == "decimalFullWidth":
        return "".join(chr(0xFF10 + int(c)) for c in str(n))
    if fmt == "none":
        return ""
    return str(n)


# ---------------------------------------------------------------------------
# 한글 문서 속성 레지스트리 (header.xml)
# ---------------------------------------------------------------------------
class HeaderRegistry(object):
    def __init__(self):
        self.fonts = []
        self.font_idx = {}
        self.charprs = []
        self.charpr_idx = {}
        self.paraprs = []
        self.parapr_idx = {}
        self.bfs = []
        self.bf_idx = {}
        self.tabprs = []
        self.tabpr_idx = {}
        self.styles = []
        self.font("함초롬바탕")
        self.font("함초롬돋움")
        # borderFill 1: 테두리 없음, 2: 테두리 없음 + 빈 채우기 (한글 기본과 같음)
        self.borderfill(None, None, None, None, None, _bf_id_force=1)
        self.borderfill(None, None, None, None, "none", _bf_id_force=2)
        self.tabpr(())                              # 0
        self.tabpr((), auto_left=True)              # 1

    # 글꼴
    def font(self, name):
        name = (name or "").strip() or "함초롬바탕"
        if name not in self.font_idx:
            self.font_idx[name] = len(self.fonts)
            self.fonts.append(name)
        return self.font_idx[name]

    # 테두리/배경
    def borderfill(self, left, right, top, bottom, fill, diag=None, _bf_id_force=None):
        key = (left, right, top, bottom, fill, diag)
        if key in self.bf_idx:
            return self.bf_idx[key]
        bid = len(self.bfs) + 1
        self.bfs.append(key)
        self.bf_idx[key] = bid
        return bid

    def tabpr(self, items, auto_left=False):
        key = (tuple(items), auto_left)
        if key not in self.tabpr_idx:
            self.tabpr_idx[key] = len(self.tabprs)
            self.tabprs.append(key)
        return self.tabpr_idx[key]

    def charpr(self, cp):
        key = tuple(sorted(cp.items()))
        if key not in self.charpr_idx:
            self.charpr_idx[key] = len(self.charprs)
            self.charprs.append(dict(cp))
        return self.charpr_idx[key]

    def parapr(self, pp):
        key = tuple(sorted(pp.items()))
        if key not in self.parapr_idx:
            self.parapr_idx[key] = len(self.paraprs)
            self.paraprs.append(dict(pp))
        return self.parapr_idx[key]

    # -- 직렬화 ---------------------------------------------------------------
    def xml(self, sec_cnt):
        out = [XML_HEAD, '<hh:head %s version="1.4" secCnt="%d">' % (HP_NS_DECL, sec_cnt),
               '<hh:beginNum page="1" footnote="1" endnote="1" pic="1" tbl="1" equation="1"/>',
               "<hh:refList>"]
        # fontfaces
        out.append('<hh:fontfaces itemCnt="7">')
        for lang in ("HANGUL", "LATIN", "HANJA", "JAPANESE", "OTHER", "SYMBOL", "USER"):
            out.append('<hh:fontface lang="%s" fontCnt="%d">' % (lang, len(self.fonts)))
            for i, f in enumerate(self.fonts):
                out.append('<hh:font id="%d" face="%s" type="TTF" isEmbedded="0">'
                           '<hh:typeInfo familyType="FCAT_GOTHIC" weight="6" proportion="0" '
                           'contrast="0" strokeVariation="1" armStyle="1" letterform="1" '
                           'midline="1" xHeight="1"/></hh:font>' % (i, attr(f)))
            out.append("</hh:fontface>")
        out.append("</hh:fontfaces>")
        # borderFills
        out.append('<hh:borderFills itemCnt="%d">' % len(self.bfs))
        for i, (l, r, t, b, fill, diag) in enumerate(self.bfs):
            out.append(self._bf_xml(i + 1, l, r, t, b, fill))
        out.append("</hh:borderFills>")
        # charProperties
        out.append('<hh:charProperties itemCnt="%d">' % len(self.charprs))
        for i, cp in enumerate(self.charprs):
            out.append(self._charpr_xml(i, cp))
        out.append("</hh:charProperties>")
        # tabProperties
        out.append('<hh:tabProperties itemCnt="%d">' % len(self.tabprs))
        for i, (items, auto_left) in enumerate(self.tabprs):
            out.append('<hh:tabPr id="%d" autoTabLeft="%d" autoTabRight="0">' % (i, 1 if auto_left else 0))
            for (pos, typ, leader) in items:
                out.append('<hp:switch><hp:case hp:required-namespace="%s">'
                           '<hh:tabItem pos="%d" type="%s" leader="%s" unit="HWPUNIT"/></hp:case>'
                           '<hp:default><hh:tabItem pos="%d" type="%s" leader="%s"/></hp:default>'
                           '</hp:switch>' % (HUC, pos, typ, leader, pos * 2, typ, leader))
            out.append("</hh:tabPr>")
        out.append("</hh:tabProperties>")
        # numberings (개요 번호 – 한글 기본값)
        out.append('<hh:numberings itemCnt="1"><hh:numbering id="1" start="0">')
        heads = [("DIGIT", "^1."), ("HANGUL_SYLLABLE", "^2."), ("DIGIT", "^3)"),
                 ("HANGUL_SYLLABLE", "^4)"), ("DIGIT", "(^5)"),
                 ("HANGUL_SYLLABLE", "(^6)"), ("CIRCLED_DIGIT", "^7")]
        for lv, (fmt, txt) in enumerate(heads):
            out.append('<hh:paraHead start="1" level="%d" align="LEFT" useInstWidth="1" '
                       'autoIndent="1" widthAdjust="0" textOffsetType="PERCENT" textOffset="50" '
                       'numFormat="%s" charPrIDRef="4294967295" checkable="%d">%s</hh:paraHead>'
                       % (lv + 1, fmt, 1 if lv == 6 else 0, esc(txt)))
        out.append("</hh:numbering></hh:numberings>")
        # paraProperties
        out.append('<hh:paraProperties itemCnt="%d">' % len(self.paraprs))
        for i, pp in enumerate(self.paraprs):
            out.append(self._parapr_xml(i, pp))
        out.append("</hh:paraProperties>")
        # styles
        out.append('<hh:styles itemCnt="%d">' % len(self.styles))
        for i, st in enumerate(self.styles):
            out.append('<hh:style id="%d" type="%s" name="%s" engName="%s" paraPrIDRef="%d" '
                       'charPrIDRef="%d" nextStyleIDRef="%d" langID="1042" lockForm="0"/>'
                       % (i, st["type"], attr(st["name"]), attr(st.get("eng", "")),
                          st["para"], st["char"], st.get("next", i)))
        out.append("</hh:styles>")
        out.append("</hh:refList>")
        out.append('<hh:compatibleDocument targetProgram="HWP201X"><hh:layoutCompatibility/>'
                   '</hh:compatibleDocument>')
        out.append('<hh:docOption><hh:linkinfo path="" pageInherit="1" footnoteInherit="0"/>'
                   '</hh:docOption>')
        out.append('<hh:trackchageConfig flags="56"/>')
        out.append("</hh:head>")
        return "".join(out)

    @staticmethod
    def _line_xml(tag, spec):
        if spec is None:
            return '<hh:%s type="NONE" width="0.1 mm" color="#000000"/>' % tag
        typ, width, color = spec
        return '<hh:%s type="%s" width="%s" color="%s"/>' % (tag, typ, width, color)

    def _bf_xml(self, bid, l, r, t, b, fill):
        s = ('<hh:borderFill id="%d" threeD="0" shadow="0" centerLine="NONE" '
             'breakCellSeparateLine="0">'
             '<hh:slash type="NONE" Crooked="0" isCounter="0"/>'
             '<hh:backSlash type="NONE" Crooked="0" isCounter="0"/>' % bid)
        s += self._line_xml("leftBorder", l) + self._line_xml("rightBorder", r)
        s += self._line_xml("topBorder", t) + self._line_xml("bottomBorder", b)
        s += '<hh:diagonal type="SOLID" width="0.1 mm" color="#000000"/>'
        if fill == "none":
            s += ('<hc:fillBrush><hc:winBrush faceColor="none" hatchColor="#FF000000" '
                  'alpha="0"/></hc:fillBrush>')
        elif fill:
            s += ('<hc:fillBrush><hc:winBrush faceColor="%s" hatchColor="#000000" '
                  'alpha="0"/></hc:fillBrush>' % fill)
        s += "</hh:borderFill>"
        return s

    def _charpr_xml(self, i, cp):
        fid = cp.get("font_hangul", 0)
        lid = cp.get("font_latin", fid)
        refs = {"hangul": fid, "latin": lid, "hanja": fid, "japanese": fid,
                "other": lid, "symbol": lid, "user": lid}
        ratio = cp.get("ratio", 100)
        spacing = cp.get("spacing", 0)
        s = ('<hh:charPr id="%d" height="%d" textColor="%s" shadeColor="%s" useFontSpace="0" '
             'useKerning="0" symMark="NONE" borderFillIDRef="2">'
             % (i, cp.get("height", 1000), cp.get("color", "#000000"), cp.get("shade", "none")))
        s += "<hh:fontRef %s/>" % " ".join('%s="%d"' % (k, refs[k]) for k in LANGS)
        s += "<hh:ratio %s/>" % " ".join('%s="%d"' % (k, ratio) for k in LANGS)
        s += "<hh:spacing %s/>" % " ".join('%s="%d"' % (k, spacing) for k in LANGS)
        s += "<hh:relSz %s/>" % " ".join('%s="100"' % k for k in LANGS)
        s += "<hh:offset %s/>" % " ".join('%s="%d"' % (k, cp.get("offset", 0)) for k in LANGS)
        if cp.get("bold"):
            s += "<hh:bold/>"
        if cp.get("italic"):
            s += "<hh:italic/>"
        if cp.get("underline"):
            s += '<hh:underline type="BOTTOM" shape="%s" color="%s"/>' % (
                cp["underline"], cp.get("ucolor", cp.get("color", "#000000")))
        else:
            s += '<hh:underline type="NONE" shape="SOLID" color="#000000"/>'
        s += '<hh:strikeout shape="%s" color="%s"/>' % (
            cp.get("strike", "NONE"), cp.get("color", "#000000"))
        s += '<hh:outline type="%s"/>' % ("SOLID" if cp.get("outline") else "NONE")
        s += '<hh:shadow type="%s" color="#B2B2B2" offsetX="10" offsetY="10"/>' % (
            "DROP" if cp.get("shadow") else "NONE")
        if cp.get("emboss"):
            s += "<hh:emboss/>"
        if cp.get("engrave"):
            s += "<hh:engrave/>"
        if cp.get("vert") == "sup":
            s += "<hh:supscript/>"
        elif cp.get("vert") == "sub":
            s += "<hh:subscript/>"
        s += "</hh:charPr>"
        return s

    def _parapr_xml(self, i, pp):
        def margin_block(k):
            ls_type = pp.get("ls_type", "PERCENT")
            ls_val = pp.get("ls_val", 160)
            if ls_type != "PERCENT":
                ls_val = ls_val * k
            return ('<hh:margin><hc:intent value="%d" unit="HWPUNIT"/>'
                    '<hc:left value="%d" unit="HWPUNIT"/><hc:right value="%d" unit="HWPUNIT"/>'
                    '<hc:prev value="%d" unit="HWPUNIT"/><hc:next value="%d" unit="HWPUNIT"/>'
                    '</hh:margin><hh:lineSpacing type="%s" value="%d" unit="HWPUNIT"/>'
                    % (pp.get("intent", 0) * k, pp.get("left", 0) * k, pp.get("right", 0) * k,
                       pp.get("prev", 0) * k, pp.get("next", 0) * k, ls_type, ls_val))

        head = pp.get("heading")
        s = ('<hh:paraPr id="%d" tabPrIDRef="%d" condense="0" fontLineHeight="0" '
             'snapToGrid="%d" suppressLineNumbers="0" checked="0">'
             % (i, pp.get("tab", 0), pp.get("snap", 1)))
        s += '<hh:align horizontal="%s" vertical="BASELINE"/>' % pp.get("align", "JUSTIFY")
        if head:
            s += '<hh:heading type="OUTLINE" idRef="1" level="%d"/>' % head
        else:
            s += '<hh:heading type="NONE" idRef="0" level="0"/>'
        s += ('<hh:breakSetting breakLatinWord="KEEP_WORD" breakNonLatinWord="%s" '
              'widowOrphan="%d" keepWithNext="%d" keepLines="%d" pageBreakBefore="%d" '
              'lineWrap="BREAK"/>'
              % (pp.get("nonlatin", "KEEP_WORD"), pp.get("widow", 0), pp.get("keepnext", 0),
                 pp.get("keeplines", 0), pp.get("pbb", 0)))
        s += '<hh:autoSpacing eAsianEng="0" eAsianNum="0"/>'
        s += ('<hp:switch><hp:case hp:required-namespace="%s">%s</hp:case>'
              '<hp:default>%s</hp:default></hp:switch>' % (HUC, margin_block(1), margin_block(2)))
        s += ('<hh:border borderFillIDRef="%d" offsetLeft="0" offsetRight="0" offsetTop="0" '
              'offsetBottom="0" connect="0" ignoreMargin="0"/>' % pp.get("bf", 2))
        s += "</hh:paraPr>"
        return s


# ---------------------------------------------------------------------------
# 변환기 본체
# ---------------------------------------------------------------------------
class Para(object):
    """한글 문단 하나를 만드는 동안 쓰는 버퍼."""

    def __init__(self, pp_id, style_id, page_break=False, col_break=False):
        self.pp = pp_id
        self.style = style_id
        self.page_break = page_break
        self.col_break = col_break
        self.items = []          # (kind, charpr, payload)
        self.mark_cp = 0
        self.width_units = 0.0   # 글자 폭 합 (HWPUNIT)
        self.max_font = 1000
        self.obj_heights = []    # 개체(표/그림/수식) 높이
        self.text_chars = 0
        self.prefix = ""         # 구역 정의 등 맨 앞에 들어갈 XML

    def add_text(self, cp, text, font_h):
        if not text:
            return
        self.items.append(("t", cp, text))
        for ch in text:
            o = ord(ch)
            self.width_units += font_h * (1.0 if o >= 0x1100 else 0.52)
        self.text_chars += len(text)
        self.max_font = max(self.max_font, font_h)

    def add(self, kind, cp, payload, height=0, width=0, font_h=1000):
        self.items.append((kind, cp, payload))
        if height:
            self.obj_heights.append(height)
        self.width_units += width
        self.max_font = max(self.max_font, font_h)

    def is_empty(self):
        return not self.items

    def xml(self):
        out = ['<hp:p id="0" paraPrIDRef="%d" styleIDRef="%d" pageBreak="%d" columnBreak="%d" merged="0">'
               % (self.pp, self.style, 1 if self.page_break else 0, 1 if self.col_break else 0)]
        first_cp = self.items[0][1] if self.items else self.mark_cp
        if self.prefix:
            out.append('<hp:run charPrIDRef="%d">%s</hp:run>' % (first_cp, self.prefix))
        if not self.items:
            if not self.prefix:
                out.append('<hp:run charPrIDRef="%d"/>' % self.mark_cp)
        else:
            i = 0
            n = len(self.items)
            while i < n:
                kind, cp, payload = self.items[i]
                if kind in ("t", "tab", "br"):
                    buf = []
                    while i < n and self.items[i][0] in ("t", "tab", "br") and self.items[i][1] == cp:
                        k, _c, p = self.items[i]
                        if k == "t":
                            buf.append(esc(p))
                        elif k == "tab":
                            buf.append('<hp:tab width="%d" leader="0" type="1"/>' % p)
                        else:
                            buf.append("<hp:lineBreak/>")
                        i += 1
                    out.append('<hp:run charPrIDRef="%d"><hp:t>%s</hp:t></hp:run>' % (cp, "".join(buf)))
                    continue
                # 개체 / 컨트롤
                out.append('<hp:run charPrIDRef="%d">%s<hp:t/></hp:run>' % (cp, payload))
                i += 1
        out.append("</hp:p>")
        return "".join(out)


class Ctx(object):
    """문단을 변환할 때의 환경(사용 가능한 폭, 표 스타일 등)."""

    def __init__(self, width, tbl_rpr=None, tbl_ppr=None, in_note=False, in_header=False):
        self.width = width
        self.tbl_rpr = tbl_rpr or {}
        self.tbl_ppr = tbl_ppr or {}
        self.in_note = in_note
        self.in_header = in_header

    def sub(self, **kw):
        c = Ctx(self.width, self.tbl_rpr, self.tbl_ppr, self.in_note, self.in_header)
        for k, v in kw.items():
            setattr(c, k, v)
        return c


class Converter(object):
    def __init__(self, docx_path, log=None, eq_rules=None):
        self.eq_rules = eq_rules or []
        self.dx = Docx(docx_path)
        self.reg = HeaderRegistry()
        self.log = log or (lambda *a: None)
        self.warnings = []
        self.bins = []            # (id, zip_path, data, media_type)
        self.bin_by_part = {}
        self.obj_id = 1000000
        self.zorder = 0
        self.counters = {}        # (abstract_id) → [count per level]
        self.last_num_seen = {}
        self.footnote_no = 0
        self.endnote_no = 0
        self.eq_count = 0
        self.eq_scripts = []
        self.plain_text = []
        self.rels_stack = [self.dx.rels]
        self.part_stack = [self.dx.main]
        self.style_map = {}
        self._init_default_fonts()
        self._init_styles()

    # -- 기본값 --------------------------------------------------------------
    def warn(self, msg):
        if msg not in self.warnings:
            self.warnings.append(msg)

    def next_id(self):
        self.obj_id += 1
        return self.obj_id + random.randint(0, 999) * 1000000

    def _init_default_fonts(self):
        th = self.dx.theme
        self.default_ea = th.get("minorEA") or "맑은 고딕"
        self.default_latin = th.get("minorLatin") or self.default_ea
        d = self.dx.doc_rpr
        if "eastAsia" in d:
            self.default_ea = d["eastAsia"]
        if "ascii" in d:
            self.default_latin = d["ascii"]

    def _init_styles(self):
        base_cp = self.charpr_id(self.base_rpr(None))
        base_pp = self.parapr_id(self.base_ppr(None), Ctx(42520))
        st = self.reg.styles
        st.append({"type": "PARA", "name": "바탕글", "eng": "Normal", "para": base_pp, "char": base_cp})
        # 한글 기본 스타일 몇 가지 (개요/머리말/각주 등) – 편집할 때 편의를 위해
        for lv in range(1, 8):
            pp = dict(self.reg.paraprs[base_pp])
            pp["heading"] = lv
            pp["left"] = 1000 * lv
            pid = self.reg.parapr(pp)
            st.append({"type": "PARA", "name": "개요 %d" % lv, "eng": "Outline %d" % lv,
                       "para": pid, "char": base_cp})
        st.append({"type": "CHAR", "name": "쪽 번호", "eng": "Page Number", "para": 0, "char": base_cp})
        st.append({"type": "PARA", "name": "머리말", "eng": "Header", "para": base_pp, "char": base_cp})
        st.append({"type": "PARA", "name": "각주", "eng": "Footnote", "para": base_pp, "char": base_cp})
        st.append({"type": "PARA", "name": "미주", "eng": "Endnote", "para": base_pp, "char": base_cp})
        self.STYLE_HEADER = 9
        self.STYLE_FOOTNOTE = 10
        self.STYLE_ENDNOTE = 11

    def style_id_for(self, word_sid):
        """Word 문단 스타일 → 한글 스타일 id (처음 쓰일 때 등록)"""
        if not word_sid or word_sid not in self.dx.styles:
            return 0
        if word_sid == self.dx.default_pstyle:
            return 0
        if word_sid in self.style_map:
            return self.style_map[word_sid]
        info = self.dx.styles[word_sid]
        rpr = self.base_rpr(word_sid)
        ppr = self.base_ppr(word_sid)
        cp = self.charpr_id(rpr)
        pp = self.parapr_id(ppr, Ctx(42520))
        sid = len(self.reg.styles)
        name = info["name"] or word_sid
        names = {s["name"] for s in self.reg.styles}
        if name in names:
            name = name + " (Word)"
        self.reg.styles.append({"type": "PARA", "name": name, "eng": word_sid,
                                "para": pp, "char": cp})
        self.style_map[word_sid] = sid
        return sid

    # -- 속성 계산 ------------------------------------------------------------
    def base_rpr(self, pstyle, ctx=None):
        d = dict(self.dx.doc_rpr)
        if ctx is not None and ctx.tbl_rpr:
            d.update(ctx.tbl_rpr)
        sid = pstyle or self.dx.default_pstyle
        if sid:
            d.update(self.dx.style_rpr(sid))
        return d

    def base_ppr(self, pstyle, ctx=None):
        d = {}
        merge_ppr(d, self.dx.doc_ppr)
        if ctx is not None and ctx.tbl_ppr:
            merge_ppr(d, ctx.tbl_ppr)
        sid = pstyle or self.dx.default_pstyle
        if sid:
            merge_ppr(d, self.dx.style_ppr(sid))
        return d

    def run_rpr(self, base, rpr_el):
        d = dict(base)
        direct = self.dx.parse_rpr(rpr_el) if rpr_el is not None else {}
        cs = direct.get("rStyle")
        if cs:
            d.update(self.dx.style_rpr(cs))
        d.update(direct)
        return d

    def charpr_id(self, r):
        ea = r.get("eastAsia") or self.default_ea
        latin = r.get("ascii") or r.get("hAnsi") or self.default_latin or ea
        cp = {
            "height": max(100, int(r.get("sz", 20) * 50)),
            "font_hangul": self.reg.font(ea),
            "font_latin": self.reg.font(latin),
            "color": norm_color(r.get("color")),
        }
        if r.get("b"):
            cp["bold"] = True
        if r.get("i"):
            cp["italic"] = True
        u = r.get("u")
        if u and u != "none":
            cp["underline"] = UNDERLINE.get(u, "SOLID")
            if r.get("u_color") and r["u_color"] != "auto":
                cp["ucolor"] = norm_color(r["u_color"])
        if r.get("dstrike"):
            cp["strike"] = "DOUBLE_SLIM"
        elif r.get("strike"):
            cp["strike"] = "SOLID"
        hl = r.get("highlight")
        if hl and hl != "none":
            cp["shade"] = HIGHLIGHT.get(hl, norm_color(hl, "#FFFF00"))
        elif r.get("shd"):
            cp["shade"] = norm_color(r["shd"], "none")
        if r.get("vert") in ("superscript", "subscript"):
            cp["vert"] = "sup" if r["vert"] == "superscript" else "sub"
        if r.get("spacing"):
            sz = r.get("sz", 20)
            cp["spacing"] = max(-50, min(50, int(round(r["spacing"] * 10.0 / sz))))
        if r.get("scale") and r["scale"] != 100:
            cp["ratio"] = max(50, min(200, r["scale"]))
        if r.get("emboss"):
            cp["emboss"] = True
        if r.get("imprint"):
            cp["engrave"] = True
        if r.get("outline"):
            cp["outline"] = True
        if r.get("shadow"):
            cp["shadow"] = True
        if r.get("position"):
            sz = r.get("sz", 20)
            cp["offset"] = max(-100, min(100, -int(round(r["position"] * 100.0 / sz))))
        return self.reg.charpr(cp)

    def parapr_id(self, p, ctx, list_auto_tab=False):
        jc = p.get("jc", "left")
        align = {"left": "LEFT", "start": "LEFT", "center": "CENTER", "right": "RIGHT",
                 "end": "RIGHT", "both": "JUSTIFY", "distribute": "DISTRIBUTE",
                 "lowKashida": "JUSTIFY", "mediumKashida": "JUSTIFY",
                 "highKashida": "JUSTIFY", "thaiDistribute": "DISTRIBUTE"}.get(jc, "LEFT")
        left = tw(p.get("ind_left", 0))
        right = tw(p.get("ind_right", 0))
        intent = 0
        if p.get("ind_hanging"):
            h = tw(p["ind_hanging"])
            left -= h
            intent = -h
        elif p.get("ind_firstLine"):
            intent = tw(p["ind_firstLine"])
        line = p.get("line")
        rule = p.get("lineRule", "auto")
        if line is None:
            ls_type, ls_val = "PERCENT", self._percent_for(240)
        elif rule == "exact":
            ls_type, ls_val = "FIXED", max(100, tw(line))
        elif rule == "atLeast":
            ls_type, ls_val = "AT_LEAST", max(100, tw(line))
        else:
            ls_type, ls_val = "PERCENT", self._percent_for(line)
        tabs = []
        for (val, pos, leader) in p.get("tabs", []):
            if val in ("clear", "bar", "num"):
                continue
            typ = {"left": "LEFT", "start": "LEFT", "right": "RIGHT", "end": "RIGHT",
                   "center": "CENTER", "decimal": "DECIMAL"}.get(val, "LEFT")
            ld = {"dot": "DOT", "hyphen": "DASH", "underscore": "SOLID",
                  "middleDot": "DOT", "heavy": "SOLID"}.get(leader, "NONE")
            tabs.append((tw(pos), typ, ld))
        auto_left = list_auto_tab or bool(p.get("ind_hanging"))
        tab_id = self.reg.tabpr(tabs, auto_left) if (tabs or auto_left) else 0
        bf = 2
        if p.get("shd") or p.get("bdr"):
            sides = {}
            for side, spec in (p.get("bdr") or {}).items():
                typ = LINE_TYPES.get(spec[0])
                if typ:
                    sides[side] = (typ, border_width(spec[1]), norm_color(spec[2]))
            fill = norm_color(p.get("shd"), None) if p.get("shd") else "none"
            bf = self.reg.borderfill(sides.get("left"), sides.get("right"), sides.get("top"),
                                     sides.get("bottom"), fill)
        pp = {
            "align": align, "left": max(-20000, left), "right": right, "intent": intent,
            "prev": tw(p.get("before", 0)), "next": tw(p.get("after", 0)),
            "ls_type": ls_type, "ls_val": ls_val, "tab": tab_id, "bf": bf,
            "keepnext": 1 if p.get("keepNext") else 0,
            "keeplines": 1 if p.get("keepLines") else 0,
            "pbb": 0,
            "widow": 1 if p.get("widowControl", True) else 0,
            "nonlatin": "KEEP_WORD" if p.get("wordWrap") is False else "BREAK_WORD",
        }
        return self.reg.parapr(pp)

    def _percent_for(self, line):
        """Word '배수' 줄 간격 → 한글 % (한글 100% = 글자 크기, Word 1줄 ≈ 글꼴 높이의 1.2~1.33배)"""
        k = 1.3 if any(s in (self.default_ea or "") for s in ("맑은", "Malgun", "malgun")) else 1.2
        return int(round(line / 240.0 * 100 * k))

    # -- 이미지 --------------------------------------------------------------
    def add_bin(self, part):
        if part in self.bin_by_part:
            return self.bin_by_part[part]
        if not self.dx.has(part):
            return None
        data = self.dx.read(part)
        ext = part.rsplit(".", 1)[-1].lower() if "." in part else "png"
        if ext == "jpeg":
            ext = "jpg"
        bid = "image%d" % (len(self.bins) + 1)
        zpath = "BinData/%s.%s" % (bid, ext)
        self.bins.append((bid, zpath, data, MEDIA_TYPES.get(ext, "image/" + ext)))
        info = (bid, data)
        self.bin_by_part[part] = info
        return info

    def rel_target(self, rid):
        rels = self.rels_stack[-1]
        t = rels.get(rid)
        return t[1] if t else None

    def pic_xml(self, rid, cx, cy, crop=None, anchor=None, descr=""):
        """cx, cy: HWPUNIT"""
        part = self.rel_target(rid)
        if not part:
            self.warn("연결된(외부) 그림은 넣지 못했습니다.")
            return None
        info = self.add_bin(part)
        if info is None:
            self.warn("그림 파일을 찾지 못했습니다: %s" % part)
            return None
        bid, data = info
        cx = max(cx, 100)
        cy = max(cy, 100)
        px = image_pixel_size(data)
        if px and px[0] > 0 and px[1] > 0:
            dimw, dimh = px[0] * 75, px[1] * 75
        else:
            dimw, dimh = cx, cy
        cl, ct, cr, cb = 0, 0, dimw, dimh
        if crop:
            l, t, r, b = crop   # 1/1000 % 단위
            cl = int(dimw * l / 100000.0)
            ct = int(dimh * t / 100000.0)
            cr = int(dimw * (1 - r / 100000.0))
            cb = int(dimh * (1 - b / 100000.0))
        oid = self.next_id()
        self.zorder += 1
        if anchor:
            pos = anchor
        else:
            pos = ('<hp:pos treatAsChar="1" affectLSpacing="0" flowWithText="1" allowOverlap="0" '
                   'holdAnchorAndSO="0" vertRelTo="PARA" horzRelTo="PARA" vertAlign="TOP" '
                   'horzAlign="LEFT" vertOffset="0" horzOffset="0"/>')
        wrap = "TOP_AND_BOTTOM"
        m = re.search(r'data-wrap="(\w+)"', pos)
        if m:
            wrap = m.group(1)
            pos = pos.replace(m.group(0), "")
        x = ('<hp:pic id="%d" zOrder="%d" numberingType="PICTURE" textWrap="%s" textFlow="BOTH_SIDES" '
             'lock="0" dropcapstyle="None" href="" groupLevel="0" instid="%d" reverse="0">'
             % (oid, self.zorder, wrap, oid % 100000000))
        x += '<hp:offset x="0" y="0"/>'
        x += '<hp:orgSz width="%d" height="%d"/>' % (cx, cy)
        x += '<hp:curSz width="%d" height="%d"/>' % (cx, cy)
        x += '<hp:flip horizontal="0" vertical="0"/>'
        x += ('<hp:rotationInfo angle="0" centerX="%d" centerY="%d" rotateimage="1"/>'
              % (cx // 2, cy // 2))
        x += ('<hp:renderingInfo><hc:transMatrix e1="1" e2="0" e3="0" e4="0" e5="1" e6="0"/>'
              '<hc:scaMatrix e1="1" e2="0" e3="0" e4="0" e5="1" e6="0"/>'
              '<hc:rotMatrix e1="1" e2="0" e3="0" e4="0" e5="1" e6="0"/></hp:renderingInfo>')
        x += ('<hp:imgRect><hc:pt0 x="0" y="0"/><hc:pt1 x="%d" y="0"/><hc:pt2 x="%d" y="%d"/>'
              '<hc:pt3 x="0" y="%d"/></hp:imgRect>' % (cx, cx, cy, cy))
        x += '<hp:imgClip left="%d" right="%d" top="%d" bottom="%d"/>' % (cl, cr, ct, cb)
        x += '<hp:inMargin left="0" right="0" top="0" bottom="0"/>'
        x += '<hp:imgDim dimwidth="%d" dimheight="%d"/>' % (dimw, dimh)
        x += ('<hc:img binaryItemIDRef="%s" bright="0" contrast="0" effect="REAL_PIC" alpha="0"/>'
              % bid)
        x += ('<hp:sz width="%d" widthRelTo="ABSOLUTE" height="%d" heightRelTo="ABSOLUTE" '
              'protect="0"/>' % (cx, cy))
        x += pos
        x += '<hp:outMargin left="0" right="0" top="0" bottom="0"/>'
        x += '<hp:shapeComment>%s</hp:shapeComment>' % esc(descr or "그림입니다.")
        x += "</hp:pic>"
        return x

    def drawing(self, el, para, cp, font_h):
        """<w:drawing> → 그림 / 글상자. 글상자 안 문단은 별도로 돌려준다."""
        extra = []
        for container in el:
            kind = local(container)
            if kind not in ("inline", "anchor"):
                continue
            ext = container.find("{%s}extent" % WP)
            cx = emu(ext.get("cx")) if ext is not None else 10000
            cy = emu(ext.get("cy")) if ext is not None else 10000
            docpr = container.find("{%s}docPr" % WP)
            descr = (docpr.get("descr") or docpr.get("title") or "") if docpr is not None else ""
            anchor_pos = None
            if kind == "anchor":
                anchor_pos = self._anchor_pos(container)
            blip = container.find(".//{%s}blip" % A)
            graphic_data = container.find(".//{%s}graphicData" % A)
            uri = graphic_data.get("uri", "") if graphic_data is not None else ""
            if blip is not None and uri.endswith("/picture"):
                rid = blip.get("{%s}embed" % R)
                src = container.find(".//{%s}srcRect" % A)
                crop = None
                if src is not None:
                    crop = tuple(to_int(src.get(k, 0)) for k in ("l", "t", "r", "b"))
                    if not any(crop):
                        crop = None
                x = self.pic_xml(rid, cx, cy, crop, anchor_pos, descr) if rid else None
                if x:
                    para.add("obj", cp, x, height=cy if anchor_pos is None else 0,
                             width=cx if anchor_pos is None else 0)
                continue
            # 도형/글상자: 안쪽 글은 문단으로 꺼낸다
            txbx = container.find(".//{%s}txbxContent" % W)
            if txbx is not None:
                extra.append(txbx)
                continue
            if blip is not None:
                rid = blip.get("{%s}embed" % R)
                x = self.pic_xml(rid, cx, cy, None, anchor_pos, descr) if rid else None
                if x:
                    para.add("obj", cp, x, height=cy, width=cx)
                continue
            if "chart" in uri:
                self.warn("차트는 그림으로 옮길 수 없어 '[차트]' 표시로 바꿨습니다.")
                para.add_text(cp, "[차트]", font_h)
            elif "diagram" in uri:
                self.warn("SmartArt 는 '[SmartArt]' 표시로 바꿨습니다.")
                para.add_text(cp, "[SmartArt]", font_h)
        return extra

    def _anchor_pos(self, anc):
        def off(tag):
            p = anc.find("{%s}%s" % (WP, tag))
            if p is None:
                return "PARA", "TOP", 0
            rel = p.get("relativeFrom", "column")
            o = p.find("{%s}posOffset" % WP)
            al = p.find("{%s}align" % WP)
            return rel, (al.text if al is not None else None), emu(o.text) if o is not None else 0

        hrel, halign, hoff = off("positionH")
        vrel, valign, voff = off("positionV")
        hmap = {"page": "PAPER", "margin": "PAGE", "column": "COLUMN", "character": "PARA",
                "leftMargin": "PAPER", "rightMargin": "PAPER", "insideMargin": "PAGE",
                "outsideMargin": "PAGE"}
        vmap = {"page": "PAPER", "margin": "PAGE", "paragraph": "PARA", "line": "PARA",
                "topMargin": "PAPER", "bottomMargin": "PAPER", "insideMargin": "PAGE",
                "outsideMargin": "PAGE"}
        ha = {"left": "LEFT", "center": "CENTER", "right": "RIGHT", "inside": "LEFT",
              "outside": "RIGHT"}.get(halign or "", "LEFT")
        va = {"top": "TOP", "center": "CENTER", "bottom": "BOTTOM", "inside": "TOP",
              "outside": "BOTTOM"}.get(valign or "", "TOP")
        if anc.find("{%s}wrapNone" % WP) is not None:
            wrap = "BEHIND_TEXT" if anc.get("behindDoc") == "1" else "IN_FRONT_OF_TEXT"
        elif anc.find("{%s}wrapSquare" % WP) is not None or \
                anc.find("{%s}wrapTight" % WP) is not None or \
                anc.find("{%s}wrapThrough" % WP) is not None:
            wrap = "SQUARE"
        else:
            wrap = "TOP_AND_BOTTOM"
        return ('<hp:pos treatAsChar="0" affectLSpacing="0" flowWithText="1" allowOverlap="1" '
                'holdAnchorAndSO="0" vertRelTo="%s" horzRelTo="%s" vertAlign="%s" horzAlign="%s" '
                'vertOffset="%d" horzOffset="%d" data-wrap="%s"/>'
                % (vmap.get(vrel, "PARA"), hmap.get(hrel, "COLUMN"), va, ha,
                   voff if valign is None else 0, hoff if halign is None else 0, wrap))

    def vml_pict(self, el, para, cp):
        """옛 형식 <w:pict>/<w:object> 의 그림·글상자"""
        extra = []
        for img in el.iter("{%s}imagedata" % V):
            rid = img.get("{%s}id" % R)
            shape = None
            for sh in el.iter("{%s}shape" % V):
                shape = sh
                break
            wpt, hpt = 72.0, 72.0
            if shape is not None:
                style = shape.get("style", "")
                mw = re.search(r"width:([\d.]+)(pt|in|px|mm|cm)?", style)
                mh = re.search(r"height:([\d.]+)(pt|in|px|mm|cm)?", style)

                def conv(m):
                    v = float(m.group(1))
                    u = m.group(2) or "pt"
                    return v * {"pt": 1, "in": 72, "px": 0.75, "mm": 2.835, "cm": 28.35}[u]
                if mw:
                    wpt = conv(mw)
                if mh:
                    hpt = conv(mh)
            if rid:
                if local(el) == "object":
                    prog = ""
                    ole = el.find("{%s}OLEObject" % O)
                    if ole is not None:
                        prog = ole.get("ProgID", "")
                    if "Equation" in prog:
                        self.warn("옛 수식 편집기(Equation 3.0/MathType) 수식은 그림으로 넣었습니다. "
                                  "Word에서 '수식 변환'(Office 수식)으로 바꾸면 편집 가능한 한글 수식이 됩니다.")
                x = self.pic_xml(rid, pt(wpt), pt(hpt))
                if x:
                    para.add("obj", cp, x, height=pt(hpt), width=pt(wpt))
            break
        for tb in el.iter("{%s}txbxContent" % W):
            extra.append(tb)
        return extra

    # -- 수식 ----------------------------------------------------------------
    def equation_xml(self, omath, font_h, color):
        conv = OmmlConverter()
        script, wem, aem, dem = conv.convert(omath)
        if not script:
            return None, 0, 0
        if self.eq_rules:
            script = eqlearn.apply(script, self.eq_rules)
        self.eq_count += 1
        self.eq_scripts.append(script)
        base = max(500, font_h)
        width = int(wem * base + base * 0.1)
        height = int((aem + dem) * base + base * 0.1)
        height = max(height, int(base * 1.1))
        baseline = int(round(100.0 * (aem + 0.05) / (aem + dem + 0.1)))
        baseline = max(50, min(95, baseline))
        oid = self.next_id()
        self.zorder += 1
        x = ('<hp:equation id="%d" zOrder="%d" numberingType="EQUATION" textWrap="TOP_AND_BOTTOM" '
             'textFlow="BOTH_SIDES" lock="0" dropcapstyle="None" version="Equation Version 60" '
             'baseLine="%d" textColor="%s" baseUnit="%d" lineMode="CHAR" font="HYhwpEQ">'
             % (oid, self.zorder, baseline, color, base))
        x += ('<hp:sz width="%d" widthRelTo="ABSOLUTE" height="%d" heightRelTo="ABSOLUTE" protect="0"/>'
              % (width, height))
        x += ('<hp:pos treatAsChar="1" affectLSpacing="0" flowWithText="1" allowOverlap="0" '
              'holdAnchorAndSO="0" vertRelTo="PARA" horzRelTo="PARA" vertAlign="TOP" horzAlign="LEFT" '
              'vertOffset="0" horzOffset="0"/>')
        x += '<hp:outMargin left="56" right="56" top="0" bottom="0"/>'
        x += '<hp:shapeComment>수식입니다.</hp:shapeComment>'
        x += '<hp:script>%s</hp:script>' % esc(script)
        x += "</hp:equation>"
        return x, width, height

    # -- 번호 매기기 ----------------------------------------------------------
    def list_marker(self, ppr, rpr_base):
        nid = ppr.get("numId")
        if not nid or nid == "0":
            return None
        ilvl = ppr.get("ilvl", 0)
        aid, res = self.dx.num_level(nid, ilvl)
        if res is None or res[0] is None:
            return None
        lvl, start_override = res
        key = aid
        cnt = self.counters.setdefault(key, [None] * 10)
        # 같은 abstract 이지만 다른 numId 가 startOverride 를 가지면 새로 시작
        if start_override is not None and self.last_num_seen.get((key, ilvl)) != nid:
            cnt[ilvl] = None
            self.last_num_seen[(key, ilvl)] = nid
        else:
            self.last_num_seen.setdefault((key, ilvl), nid)
        start = start_override if start_override is not None else lvl["start"]
        cnt[ilvl] = start if cnt[ilvl] is None else cnt[ilvl] + 1
        for deeper in range(ilvl + 1, 10):
            cnt[deeper] = None
        if lvl["fmt"] == "bullet":
            txt = fix_pua(lvl["text"], lvl.get("font") or "symbol") or "•"
            if txt == "o" and (lvl.get("font") or "").lower().startswith("courier"):
                txt = "○"
        else:
            def rep(m):
                k = int(m.group(1)) - 1
                if k == ilvl:
                    n = cnt[ilvl]
                    fmt = lvl["fmt"]
                else:
                    _a, r2 = self.dx.num_level(nid, k)
                    lv2 = r2[0] if r2 else None
                    n = cnt[k] if cnt[k] is not None else (lv2["start"] if lv2 else 1)
                    fmt = lv2["fmt"] if lv2 else "decimal"
                if lvl.get("isLgl"):
                    fmt = "decimal"
                return format_number(n, fmt)
            txt = re.sub(r"%(\d)", rep, lvl["text"])
        rpr = dict(rpr_base)
        rpr.update(lvl["rpr"])
        rpr.pop("u", None)
        return txt, lvl["suff"], self.charpr_id(rpr), lvl["ppr"]

    # -- 문단 ----------------------------------------------------------------
    def paragraph(self, p_el, ctx, prev_style=None, next_style=None):
        """<w:p> → [Para, …] (쪽 나눔·글상자 때문에 여러 개일 수 있음)"""
        ppr_el = p_el.find(w("pPr"))
        direct = parse_ppr(ppr_el)
        pstyle = direct.get("pStyle") or self.dx.default_pstyle
        ppr = self.base_ppr(pstyle, ctx)
        rpr_base = self.base_rpr(pstyle, ctx)
        # 번호: 스타일 < 번호 정의 < 직접
        num_ppr_probe = dict(ppr)
        merge_ppr(num_ppr_probe, direct)
        marker = None
        if num_ppr_probe.get("numId") and not ctx.in_header:
            mk = self.list_marker(num_ppr_probe, rpr_base)
            if mk:
                marker = mk
                merge_ppr(ppr, mk[3])
        merge_ppr(ppr, direct)
        # 문단 기호의 글자 모양
        mark_rpr = self.run_rpr(rpr_base, ppr_el.find(w("rPr")) if ppr_el is not None else None)
        # contextualSpacing: 같은 스타일끼리 간격 무시
        if ppr.get("contextualSpacing"):
            if prev_style == pstyle:
                ppr["before"] = 0
            if next_style == pstyle:
                ppr["after"] = 0
        style_id = self.style_id_for(pstyle) if not ctx.in_note else self.STYLE_FOOTNOTE
        pp_id = self.parapr_id(ppr, ctx, list_auto_tab=bool(marker))
        self._cur_ppr = ppr
        self._cur_ctx = ctx
        self._cur_style = style_id
        paras = [Para(pp_id, style_id, page_break=bool(ppr.get("pageBreakBefore")))]
        paras[0].mark_cp = self.charpr_id(mark_rpr)
        self._extra_blocks = []
        if marker:
            txt, suff, mcp, _ = marker
            fh = self.reg.charprs[mcp]["height"]
            paras[0].add_text(mcp, txt, fh)
            if suff == "tab":
                paras[0].add("tab", mcp, 1000, width=1000)
            elif suff == "space":
                paras[0].add_text(mcp, " ", fh)
        self._fields = []
        kids = [c for c in p_el if local(c) not in ("pPr", "bookmarkStart", "bookmarkEnd",
                                                     "proofErr")]
        self._only_math = bool(kids) and all(nsof(c) == M for c in kids)
        self.inline_children(p_el, paras, rpr_base, ctx)
        # 수식 문단(oMathPara)만 있는 문단은 가운데 정렬
        result = paras
        tail = []
        for tb in self._extra_blocks:
            tail.extend(self.blocks(list(tb), ctx))
        return result + tail

    def new_para_after(self, paras, page=False, col=False):
        p = Para(paras[-1].pp, paras[-1].style, page_break=page, col_break=col)
        p.mark_cp = paras[-1].mark_cp
        paras.append(p)
        return p

    def inline_children(self, parent, paras, rpr_base, ctx):
        for ch in list(parent):
            ns, tag = nsof(ch), local(ch)
            if ns == W:
                if tag == "r":
                    self.run(ch, paras, rpr_base, ctx)
                elif tag in ("hyperlink", "smartTag", "customXml", "ins", "moveTo",
                             "dir", "bdo"):
                    self.inline_children(ch, paras, rpr_base, ctx)
                elif tag == "fldSimple":
                    instr = (ch.get(w("instr")) or "").strip().split(" ")[0].upper()
                    if instr in ("PAGE", "NUMPAGES", "SECTIONPAGES"):
                        self._autonum(paras[-1], instr, rpr_base, ch.find(w("r")))
                    else:
                        self.inline_children(ch, paras, rpr_base, ctx)
                elif tag == "sdt":
                    c = ch.find(w("sdtContent"))
                    if c is not None:
                        self.inline_children(c, paras, rpr_base, ctx)
                elif tag in ("del", "moveFrom", "pPr", "bookmarkStart", "bookmarkEnd",
                             "proofErr", "permStart", "permEnd", "commentRangeStart",
                             "commentRangeEnd"):
                    continue
            elif ns == M:
                if tag == "oMathPara":
                    jc = None
                    pr = ch.find("{%s}oMathParaPr" % M)
                    if pr is not None:
                        j = pr.find("{%s}jc" % M)
                        if j is not None:
                            jc = j.get("{%s}val" % M)
                    self._math_para(ch, paras, rpr_base, jc)
                elif tag == "oMath":
                    self._math_inline(ch, paras[-1], rpr_base)
            elif tag == "AlternateContent":
                choice = None
                for c in ch:
                    if local(c) == "Choice" and choice is None:
                        choice = c
                fb = [c for c in ch if local(c) == "Fallback"]
                use = choice if choice is not None else (fb[0] if fb else None)
                if use is not None:
                    self.inline_children(use, paras, rpr_base, ctx)

    def _math_inline(self, omath, para, rpr_base):
        r = rpr_base
        cp = self.charpr_id(r)
        fh = self.reg.charprs[cp]["height"]
        color = self.reg.charprs[cp]["color"]
        x, wd, ht = self.equation_xml(omath, fh, color)
        if x:
            para.add("obj", cp, x, height=ht, width=wd, font_h=fh)

    def _math_para(self, omp, paras, rpr_base, jc):
        para = paras[-1]
        # 문단 전체가 수식이면 정렬을 바꾼 새 문단 모양을 쓴다
        if self._only_math and para.is_empty():
            ppr = dict(self._cur_ppr)
            ppr["jc"] = {"left": "left", "right": "right", "center": "center",
                         "centerGroup": "center"}.get(jc or "centerGroup", "center")
            ppr.pop("ind_firstLine", None)
            ppr.pop("ind_hanging", None)
            para.pp = self.parapr_id(ppr, self._cur_ctx)
        first = True
        for om in omp:
            if local(om) != "oMath":
                continue
            if not first:
                para.add("br", para.items[-1][1] if para.items else para.mark_cp, None)
            self._math_inline(om, para, rpr_base)
            first = False

    def _autonum(self, para, instr, rpr_base, r_el=None):
        rpr = self.run_rpr(rpr_base, r_el.find(w("rPr")) if r_el is not None else None)
        cp = self.charpr_id(rpr)
        typ = {"PAGE": "PAGE", "NUMPAGES": "TOTAL_PAGE", "SECTIONPAGES": "TOTAL_PAGE"}[instr]
        x = ('<hp:ctrl><hp:autoNum num="1" numType="%s"><hp:autoNumFormat type="DIGIT" userChar="" '
             'prefixChar="" suffixChar="" supscript="0"/></hp:autoNum></hp:ctrl>' % typ)
        para.add("obj", cp, x, width=600)

    def run(self, r_el, paras, rpr_base, ctx):
        rpr_el = r_el.find(w("rPr"))
        rpr = self.run_rpr(rpr_base, rpr_el)
        if rpr.get("vanish"):
            return
        cp = None
        fh = int(rpr.get("sz", 20) * 50)

        def cpid():
            return self.charpr_id(rpr)

        for ch in list(r_el):
            ns, tag = nsof(ch), local(ch)
            # 복합 필드 상태
            if ns == W and tag == "fldChar":
                t = ch.get(w("fldCharType"))
                if t == "begin":
                    self._fields.append({"instr": "", "state": "instr"})
                elif t == "separate" and self._fields:
                    f = self._fields[-1]
                    f["state"] = "result"
                    code = f["instr"].strip().split(" ")[0].upper() if f["instr"].strip() else ""
                    if code in ("PAGE", "NUMPAGES", "SECTIONPAGES"):
                        self._autonum(paras[-1], code, rpr_base, r_el)
                        f["suppress"] = True
                elif t == "end" and self._fields:
                    f = self._fields.pop()
                    if f["state"] == "instr":
                        code = f["instr"].strip().split(" ")[0].upper() if f["instr"].strip() else ""
                        if code in ("PAGE", "NUMPAGES", "SECTIONPAGES"):
                            self._autonum(paras[-1], code, rpr_base, r_el)
                continue
            if ns == W and tag == "instrText":
                if self._fields:
                    self._fields[-1]["instr"] += ch.text or ""
                continue
            if self._fields and (self._fields[-1]["state"] == "instr" or
                                 self._fields[-1].get("suppress")):
                continue
            para = paras[-1]
            if ns == W:
                if tag == "t":
                    text = ch.text or ""
                    font = rpr.get("ascii")
                    text = fix_pua(text, font)
                    if rpr.get("caps") or rpr.get("smallCaps"):
                        text = text.upper()
                    para.add_text(cpid(), text, fh)
                    self.plain_text.append(text)
                elif tag == "tab" or tag == "ptab":
                    para.add("tab", cpid(), 4000, width=2000)
                elif tag == "br":
                    bt = ch.get(w("type"))
                    if bt == "page":
                        if ctx.in_header or ctx.in_note:
                            para.add("br", cpid(), None)
                        else:
                            self.new_para_after(paras, page=True)
                    elif bt == "column":
                        self.new_para_after(paras, col=True)
                    else:
                        para.add("br", cpid(), None)
                elif tag == "cr":
                    para.add("br", cpid(), None)
                elif tag == "noBreakHyphen":
                    para.add_text(cpid(), "-", fh)
                elif tag == "sym":
                    code = int(ch.get(w("char"), "20"), 16)
                    para.add_text(cpid(), map_symbol_char(code, ch.get(w("font"))), fh)
                elif tag == "drawing":
                    self._extra_blocks.extend(self.drawing(ch, para, cpid(), fh))
                elif tag in ("pict", "object"):
                    self._extra_blocks.extend(self.vml_pict(ch, para, cpid()))
                elif tag == "footnoteReference":
                    self.note(ch.get(w("id")), para, cpid(), ctx, foot=True)
                elif tag == "endnoteReference":
                    self.note(ch.get(w("id")), para, cpid(), ctx, foot=False)
                elif tag in ("footnoteRef", "endnoteRef"):
                    typ = "FOOTNOTE" if tag == "footnoteRef" else "ENDNOTE"
                    x = ('<hp:ctrl><hp:autoNum num="%d" numType="%s"><hp:autoNumFormat type="DIGIT" '
                         'userChar="" prefixChar="" suffixChar=")" supscript="0"/></hp:autoNum></hp:ctrl>'
                         % (self._note_no, typ))
                    para.add("obj", cpid(), x, width=600)
                elif tag == "separator" or tag == "continuationSeparator":
                    pass
            elif tag == "AlternateContent":
                choice = None
                for c in ch:
                    if local(c) == "Choice" and choice is None:
                        choice = c
                fb = [c for c in ch if local(c) == "Fallback"]
                use = choice if choice is not None else (fb[0] if fb else None)
                if use is not None:
                    for sub in use:
                        if local(sub) == "drawing":
                            self._extra_blocks.extend(self.drawing(sub, para, cpid(), fh))
                        elif local(sub) in ("pict", "object"):
                            self._extra_blocks.extend(self.vml_pict(sub, para, cpid()))
            elif ns == M and tag in ("oMath", "oMathPara"):
                self._math_inline(ch, para, rpr)

    _note_no = 1

    def note(self, nid, para, cp, ctx, foot=True):
        store = self.dx.footnotes if foot else self.dx.endnotes
        el = store.get(nid)
        if el is None:
            return
        if foot:
            self.footnote_no += 1
            no = self.footnote_no
        else:
            self.endnote_no += 1
            no = self.endnote_no
        saved = (self._note_no, getattr(self, "_fields", []), getattr(self, "_cur_ppr", {}),
                 getattr(self, "_cur_ctx", ctx), getattr(self, "_cur_style", 0),
                 getattr(self, "_extra_blocks", []))
        self._note_no = no
        nctx = ctx.sub(in_note=True)
        body = self.blocks(list(el), nctx)
        (self._note_no, self._fields, self._cur_ppr, self._cur_ctx, self._cur_style,
         self._extra_blocks) = saved
        tag = "footNote" if foot else "endNote"
        x = ('<hp:ctrl><hp:%s number="%d" suffixChar=")" instId="%d">'
             '<hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="TOP" '
             'linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0" '
             'hasNumRef="0">%s</hp:subList></hp:%s></hp:ctrl>'
             % (tag, no, self.next_id() % 100000000, "".join(p.xml() for p in body), tag))
        para.add("obj", cp, x, width=400)

    # -- 블록 (문단·표) -------------------------------------------------------
    def blocks(self, children, ctx):
        """본문/셀/머리말 등의 자식 목록 → [Para]"""
        out = []
        flat = []
        for ch in children:
            self._flatten(ch, flat)
        styles = []
        for el in flat:
            if local(el) == "p":
                pp = el.find(w("pPr"))
                styles.append(wval(pp, "pStyle") if pp is not None else None)
            else:
                styles.append("__tbl__")
        for i, el in enumerate(flat):
            tag = local(el)
            if tag == "p":
                dflt = self.dx.default_pstyle
                prev_s = next_s = None
                if i > 0 and styles[i - 1] != "__tbl__":
                    prev_s = styles[i - 1] or dflt
                if i + 1 < len(styles) and styles[i + 1] != "__tbl__":
                    next_s = styles[i + 1] or dflt
                out.extend(self.paragraph(el, ctx, prev_s, next_s))
            elif tag == "tbl":
                out.append(self.table(el, ctx))
        return out

    def _flatten(self, el, out):
        ns, tag = nsof(el), local(el)
        if ns == W and tag in ("p", "tbl"):
            out.append(el)
        elif ns == W and tag == "sdt":
            c = el.find(w("sdtContent"))
            if c is not None:
                for ch in c:
                    self._flatten(ch, out)
        elif ns == W and tag in ("customXml", "ins", "moveTo"):
            for ch in el:
                self._flatten(ch, out)
        elif ns == W and tag == "altChunk":
            self.warn("문서 안에 끼워 넣은 다른 문서(altChunk)는 옮기지 못했습니다.")

    # -- 표 ------------------------------------------------------------------
    def _tbl_style_info(self, tbl_pr):
        sid = wval(tbl_pr, "tblStyle") if tbl_pr is not None else None
        sid = sid or self.dx.default_tstyle
        borders = {}
        cell_mar = {}
        rpr = {}
        ppr = {}
        for st in self.dx.style_chain(sid) if sid else []:
            tp = st.get("tblPr")
            if tp is not None:
                borders.update(self._parse_borders(tp.find(w("tblBorders"))))
                cell_mar.update(self._parse_mar(tp.find(w("tblCellMar"))))
            rpr.update(st["rpr"])
            merge_ppr(ppr, st["ppr"])
        return borders, cell_mar, rpr, ppr

    @staticmethod
    def _parse_borders(el):
        d = {}
        if el is None:
            return d
        for side in ("top", "left", "start", "bottom", "right", "end", "insideH", "insideV"):
            e = el.find(w(side))
            if e is None:
                continue
            key = {"start": "left", "end": "right"}.get(side, side)
            val = e.get(w("val"), "single")
            if val in ("nil", "none"):
                d[key] = None
            else:
                typ = LINE_TYPES.get(val, "SOLID")
                d[key] = (typ, border_width(e.get(w("sz"), "4")), norm_color(e.get(w("color"))))
        return d

    @staticmethod
    def _parse_mar(el):
        d = {}
        if el is None:
            return d
        for side in ("top", "left", "start", "bottom", "right", "end"):
            e = el.find(w(side))
            if e is not None:
                key = {"start": "left", "end": "right"}.get(side, side)
                d[key] = tw(e.get(w("w"), "0"))
        return d

    def table(self, tbl, ctx):
        tbl_pr = tbl.find(w("tblPr"))
        st_borders, st_mar, st_rpr, st_ppr = self._tbl_style_info(tbl_pr)
        borders = dict(st_borders)
        borders.update(self._parse_borders(tbl_pr.find(w("tblBorders")) if tbl_pr is not None else None))
        mar = {"left": 540, "right": 540, "top": 0, "bottom": 0}
        mar.update(st_mar)
        mar.update(self._parse_mar(tbl_pr.find(w("tblCellMar")) if tbl_pr is not None else None))
        jc = wval(tbl_pr, "jc", "left") if tbl_pr is not None else "left"
        ind = 0
        if tbl_pr is not None and tbl_pr.find(w("tblInd")) is not None:
            ind = tw(tbl_pr.find(w("tblInd")).get(w("w"), "0"))

        # 열 너비
        grid = [tw(g.get(w("w"), "0")) for g in tbl.findall("%s/%s" % (w("tblGrid"), w("gridCol")))]
        rows = tbl.findall(w("tr"))
        ncols = len(grid)
        # 셀 배치 계산
        layout = []        # [ [cell dict, …] per row ]
        for r_i, tr in enumerate(rows):
            trpr = tr.find(w("trPr"))
            col = to_int(wval(trpr, "gridBefore", "0")) if trpr is not None else 0
            cells = []
            for tc in self._row_cells(tr):
                tcpr = tc.find(w("tcPr"))
                span = to_int(wval(tcpr, "gridSpan", "1"), 1) if tcpr is not None else 1
                vm = tcpr.find(w("vMerge")) if tcpr is not None else None
                vmerge = None
                if vm is not None:
                    vmerge = vm.get(w("val"), "continue")
                tcw = None
                if tcpr is not None and tcpr.find(w("tcW")) is not None:
                    e = tcpr.find(w("tcW"))
                    if e.get(w("type"), "dxa") == "dxa":
                        tcw = tw(e.get(w("w"), "0"))
                cells.append({"tc": tc, "col": col, "span": max(1, span), "vmerge": vmerge,
                              "row": r_i, "rowspan": 1, "w": tcw})
                col += max(1, span)
            ncols = max(ncols, col)
            layout.append(cells)
        if len(grid) < ncols:
            grid += [0] * (ncols - len(grid))
        # 너비가 없는 열 보완
        if not any(grid):
            total = ctx.width
            tblw = tbl_pr.find(w("tblW")) if tbl_pr is not None else None
            if tblw is not None and tblw.get(w("type")) == "dxa" and to_int(tblw.get(w("w"))):
                total = tw(tblw.get(w("w")))
            grid = [int(total / max(1, ncols))] * ncols
        else:
            missing = [i for i, g in enumerate(grid) if g <= 0]
            for i in missing:
                grid[i] = 500
        # 표 너비 지정(%)이면 맞춤
        tblw = tbl_pr.find(w("tblW")) if tbl_pr is not None else None
        if tblw is not None and tblw.get(w("type")) == "pct":
            v = tblw.get(w("w"), "0")
            pct = float(v.rstrip("%")) if v.endswith("%") else to_int(v) / 50.0
            if pct > 0:
                target = ctx.width * pct / 100.0
                k = target / max(1, sum(grid))
                grid = [int(g * k) for g in grid]
        # 페이지 폭보다 크면 줄인다
        if sum(grid) > ctx.width * 1.15 and ctx.width > 0:
            k = ctx.width / float(sum(grid))
            grid = [int(g * k) for g in grid]

        # 세로 병합 처리
        anchor_at = {}     # col → 시작 셀
        for r_i, cells in enumerate(layout):
            for c in cells:
                if c["vmerge"] == "continue" and c["col"] in anchor_at:
                    a = anchor_at[c["col"]]
                    a["rowspan"] += 1
                    c["merged"] = True
                else:
                    if c["vmerge"] == "restart":
                        anchor_at[c["col"]] = c
                    else:
                        anchor_at.pop(c["col"], None)
            # 이 행에 없는 열의 병합 고리는 끊는다
            present = {c["col"] for c in cells}
            for col in list(anchor_at):
                if col not in present:
                    anchor_at.pop(col)

        nrows = len(rows)
        # 한글 표는 모든 칸이 정확히 한 셀에 속해야 한다 → gridBefore/After 등 빈 곳을 채운다
        covered = [[False] * len(grid) for _ in range(nrows)]
        for cells in layout:
            for c in cells:
                if c.get("merged"):
                    continue
                for rr in range(c["row"], min(nrows, c["row"] + c["rowspan"])):
                    for cc in range(c["col"], min(len(grid), c["col"] + c["span"])):
                        covered[rr][cc] = True
        for r_i, cells in enumerate(layout):
            col = 0
            while col < len(grid):
                if covered[r_i][col]:
                    col += 1
                    continue
                start = col
                while col < len(grid) and not covered[r_i][col]:
                    covered[r_i][col] = True
                    col += 1
                cells.append({"tc": None, "col": start, "span": col - start, "vmerge": None,
                              "row": r_i, "rowspan": 1, "w": None, "filler": True})
            cells.sort(key=lambda c: c["col"])
            # 표 범위를 넘는 셀은 잘라낸다
            for c in cells:
                if c["col"] + c["span"] > len(grid):
                    c["span"] = max(1, len(grid) - c["col"])
        # 행 높이(지정값) + 셀 내용
        row_h = [0] * nrows
        row_rule = [None] * nrows
        header_rows = set()
        for r_i, tr in enumerate(rows):
            trpr = tr.find(w("trPr"))
            if trpr is not None:
                th = trpr.find(w("trHeight"))
                if th is not None:
                    row_h[r_i] = tw(th.get(w("val"), "0"))
                    row_rule[r_i] = th.get(w("hRule"), "atLeast")
                if wbool(trpr, "tblHeader"):
                    header_rows.add(r_i)

        tctx_rpr = st_rpr
        tctx_ppr = st_ppr
        cell_xml = {}
        content_h = {}
        for cells in layout:
            for c in cells:
                if c.get("merged"):
                    continue
                tc = c["tc"]
                if tc is None:
                    p = Para(self.parapr_id({"line": 240}, ctx), 0)
                    p.mark_cp = self.charpr_id(self.base_rpr(None, ctx))
                    c.update({"cmar": dict(mar), "has_mar": False, "paras": [p],
                              "cw": sum(grid[c["col"]:c["col"] + c["span"]])})
                    content_h[id(c)] = 600
                    continue
                tcpr = tc.find(w("tcPr"))
                cmar = dict(mar)
                has_mar = False
                if tcpr is not None and tcpr.find(w("tcMar")) is not None:
                    cmar.update(self._parse_mar(tcpr.find(w("tcMar"))))
                    has_mar = True
                cw = sum(grid[c["col"]:c["col"] + c["span"]])
                inner = max(1000, cw - cmar["left"] - cmar["right"])
                cctx = Ctx(inner, tctx_rpr, tctx_ppr, ctx.in_note, ctx.in_header)
                paras = self.blocks([e for e in tc if local(e) in ("p", "tbl", "sdt", "customXml")], cctx)
                if not paras:
                    p = Para(0, 0)
                    paras = [p]
                h = sum(self.est_height(p, inner) for p in paras) + cmar["top"] + cmar["bottom"]
                content_h[id(c)] = h
                c["cmar"] = cmar
                c["has_mar"] = has_mar
                c["paras"] = paras
                c["cw"] = cw
        # 높이 배분
        for cells in layout:
            for c in cells:
                if c.get("merged"):
                    continue
                need = content_h[id(c)]
                r0, rs = c["row"], c["rowspan"]
                have = sum(row_h[r0:r0 + rs])
                if need > have:
                    if row_rule[r0 + rs - 1] == "exact":
                        continue
                    add = (need - have)
                    for k in range(r0, r0 + rs):
                        if row_rule[k] == "exact":
                            continue
                        row_h[k] += int(add / rs) + 1
        row_h = [max(h, 600) for h in row_h]

        # XML
        oid = self.next_id()
        self.zorder += 1
        total_w = sum(grid)
        total_h = sum(row_h)
        out = []
        out.append('<hp:tbl id="%d" zOrder="%d" numberingType="TABLE" textWrap="TOP_AND_BOTTOM" '
                   'textFlow="BOTH_SIDES" lock="0" dropcapstyle="None" pageBreak="CELL" '
                   'repeatHeader="%d" rowCnt="%d" colCnt="%d" cellSpacing="0" borderFillIDRef="%d" '
                   'noAdjust="0">' % (oid, self.zorder, 1 if header_rows else 0, nrows, len(grid),
                                      self.reg.borderfill(None, None, None, None, "none")))
        out.append('<hp:sz width="%d" widthRelTo="ABSOLUTE" height="%d" heightRelTo="ABSOLUTE" protect="0"/>'
                   % (total_w, total_h))
        out.append('<hp:pos treatAsChar="1" affectLSpacing="0" flowWithText="1" allowOverlap="0" '
                   'holdAnchorAndSO="0" vertRelTo="PARA" horzRelTo="COLUMN" vertAlign="TOP" '
                   'horzAlign="LEFT" vertOffset="0" horzOffset="0"/>')
        out.append('<hp:outMargin left="0" right="0" top="0" bottom="0"/>')
        out.append('<hp:inMargin left="%d" right="%d" top="%d" bottom="%d"/>'
                   % (mar["left"], mar["right"], mar["top"], mar["bottom"]))
        for r_i, cells in enumerate(layout):
            out.append("<hp:tr>")
            for c in cells:
                if c.get("merged"):
                    continue
                tc = c["tc"]
                tcpr = tc.find(w("tcPr")) if tc is not None else None
                cb = dict(borders) if tc is not None else {}
                last_col = c["col"] + c["span"] >= len(grid)
                last_row = c["row"] + c["rowspan"] >= nrows
                edge = {
                    "left": cb.get("left") if c["col"] == 0 else cb.get("insideV"),
                    "right": cb.get("right") if last_col else cb.get("insideV"),
                    "top": cb.get("top") if c["row"] == 0 else cb.get("insideH"),
                    "bottom": cb.get("bottom") if last_row else cb.get("insideH"),
                }
                fill = "none"
                valign = "TOP"
                if tcpr is not None:
                    own = self._parse_borders(tcpr.find(w("tcBorders")))
                    for k in ("left", "right", "top", "bottom"):
                        if k in own:
                            edge[k] = own[k]
                    shd = tcpr.find(w("shd"))
                    if shd is not None:
                        f = shd.get(w("fill"))
                        if f and f != "auto":
                            fill = norm_color(f, "none")
                    va = wval(tcpr, "vAlign")
                    valign = {"center": "CENTER", "bottom": "BOTTOM"}.get(va, "TOP")
                bf = self.reg.borderfill(edge["left"], edge["right"], edge["top"], edge["bottom"], fill)
                cmar = c["cmar"]
                ch_ = sum(row_h[c["row"]:c["row"] + c["rowspan"]])
                out.append('<hp:tc name="" header="%d" hasMargin="%d" protect="0" editable="0" dirty="0" '
                           'borderFillIDRef="%d">' % (1 if c["row"] in header_rows else 0,
                                                       1 if c["has_mar"] else 0, bf))
                out.append('<hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="%s" '
                           'linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" '
                           'hasTextRef="0" hasNumRef="0">' % valign)
                out.append("".join(p.xml() for p in c["paras"]))
                out.append("</hp:subList>")
                out.append('<hp:cellAddr colAddr="%d" rowAddr="%d"/>' % (c["col"], c["row"]))
                out.append('<hp:cellSpan colSpan="%d" rowSpan="%d"/>' % (c["span"], c["rowspan"]))
                out.append('<hp:cellSz width="%d" height="%d"/>' % (c["cw"], ch_))
                out.append('<hp:cellMargin left="%d" right="%d" top="%d" bottom="%d"/>'
                           % (cmar["left"], cmar["right"], cmar["top"], cmar["bottom"]))
                out.append("</hp:tc>")
            out.append("</hp:tr>")
        out.append("</hp:tbl>")
        # 표를 담는 문단 (줄 간격 최소, 정렬은 표 정렬)
        holder = {"jc": {"center": "center", "right": "right", "end": "right"}.get(jc, "left"),
                  "line": 240, "lineRule": "auto", "before": 0, "after": 0,
                  "ind_left": int(ind / 5) if jc not in ("center", "right", "end") else 0}
        pp = self.parapr_id(holder, ctx)
        para = Para(pp, 0)
        para.mark_cp = self.charpr_id(self.base_rpr(None, ctx))
        para.add("obj", para.mark_cp, "".join(out), height=total_h + 200, width=total_w)
        para.is_table = True
        return para

    @staticmethod
    def _row_cells(tr):
        out = []
        for ch in tr:
            tag = local(ch)
            if tag == "tc":
                out.append(ch)
            elif tag in ("sdt", "customXml"):
                c = ch.find(w("sdtContent")) if tag == "sdt" else ch
                if c is not None:
                    out.extend(Converter._row_cells(c))
        return out

    def est_height(self, para, width):
        """문단 높이 대략값 (표의 행 높이를 잡는 데만 사용)"""
        pp = self.reg.paraprs[para.pp] if para.pp < len(self.reg.paraprs) else {}
        fh = para.max_font
        if pp.get("ls_type") == "PERCENT":
            line_h = fh * pp.get("ls_val", 160) / 100.0
        else:
            line_h = max(fh, pp.get("ls_val", fh))
        avail = max(1000, width - pp.get("left", 0) - pp.get("right", 0))
        lines = max(1, int(math.ceil(para.width_units / float(avail)))) if para.width_units else 1
        lines += sum(1 for k, _c, _p in para.items if k == "br")
        h = lines * line_h + pp.get("prev", 0) + pp.get("next", 0)
        if para.obj_heights:
            h = max(h, max(para.obj_heights) + pp.get("prev", 0) + pp.get("next", 0))
        return int(h)

    # -- 구역 / 머리말 -------------------------------------------------------
    def section_props(self, sect):
        d = {"w": 59528, "h": 84188, "landscape": False, "top": 1440 * 5, "bottom": 1440 * 5,
             "left": 1800 * 5, "right": 1800 * 5, "header": 851 * 5, "footer": 992 * 5,
             "gutter": 0, "cols": 1, "col_space": 425 * 5, "col_sep": False, "type": "nextPage",
             "titlePg": False, "headers": {}, "footers": {}, "pg_start": None, "pg_fmt": None}
        if sect is None:
            return d
        sz = sect.find(w("pgSz"))
        if sz is not None:
            d["w"] = tw(sz.get(w("w"), "11906"))
            d["h"] = tw(sz.get(w("h"), "16838"))
            d["landscape"] = sz.get(w("orient")) == "landscape" or d["w"] > d["h"]
        mg = sect.find(w("pgMar"))
        if mg is not None:
            for k in ("top", "bottom", "left", "right", "header", "footer", "gutter"):
                v = mg.get(w(k))
                if v is not None:
                    d[k] = abs(tw(v))
        cols = sect.find(w("cols"))
        if cols is not None:
            d["cols"] = max(1, to_int(cols.get(w("num"), "1"), 1))
            d["col_space"] = tw(cols.get(w("space"), "425"))
            d["col_sep"] = cols.get(w("sep")) in ("1", "true", "on")
        d["type"] = wval(sect, "type", "nextPage")
        d["titlePg"] = bool(wbool(sect, "titlePg"))
        for ref in sect.findall(w("headerReference")):
            d["headers"][ref.get(w("type"), "default")] = ref.get("{%s}id" % R)
        for ref in sect.findall(w("footerReference")):
            d["footers"][ref.get(w("type"), "default")] = ref.get("{%s}id" % R)
        pg = sect.find(w("pgNumType"))
        if pg is not None:
            if pg.get(w("start")) is not None:
                d["pg_start"] = to_int(pg.get(w("start")), 1)
            d["pg_fmt"] = pg.get(w("fmt"))
        return d

    def secpr_xml(self, sp, header_ctrls):
        if sp["landscape"]:
            pw, ph, land = min(sp["w"], sp["h"]), max(sp["w"], sp["h"]), "NARROWLY"
        else:
            pw, ph, land = sp["w"], sp["h"], "WIDELY"
        # Word: top = 본문 시작 위치, header = 종이 끝~머리말 / 한글: top = 종이 끝~머리말, header = 머리말 영역
        h_top = min(sp["header"], sp["top"])
        h_header = max(0, sp["top"] - sp["header"])
        h_bottom = min(sp["footer"], sp["bottom"])
        h_footer = max(0, sp["bottom"] - sp["footer"])
        hide_first_header = 1 if sp.get("hide_first_header") else 0
        hide_first_footer = 1 if sp.get("hide_first_footer") else 0
        tab = self.dx.default_tab * 5
        x = ('<hp:secPr id="" textDirection="HORIZONTAL" spaceColumns="%d" tabStop="%d" '
             'tabStopVal="%d" tabStopUnit="HWPUNIT" outlineShapeIDRef="1" memoShapeIDRef="0" '
             'textVerticalWidthHead="0" masterPageCnt="0">' % (sp["col_space"], tab, tab // 2))
        x += '<hp:grid lineGrid="0" charGrid="0" wonggojiFormat="0"/>'
        x += ('<hp:startNum pageStartsOn="BOTH" page="%d" pic="0" tbl="0" equation="0"/>'
              % (sp["pg_start"] or 0))
        x += ('<hp:visibility hideFirstHeader="%d" hideFirstFooter="%d" hideFirstMasterPage="0" '
              'border="SHOW_ALL" fill="SHOW_ALL" hideFirstPageNum="0" hideFirstEmptyLine="0" '
              'showLineNumber="0"/>' % (hide_first_header, hide_first_footer))
        x += '<hp:lineNumberShape restartType="0" countBy="0" distance="0" startNumber="0"/>'
        x += ('<hp:pagePr landscape="%s" width="%d" height="%d" gutterType="LEFT_ONLY">'
              '<hp:margin header="%d" footer="%d" gutter="%d" left="%d" right="%d" top="%d" bottom="%d"/>'
              '</hp:pagePr>' % (land, pw, ph, h_header, h_footer, sp["gutter"], sp["left"],
                                sp["right"], h_top, h_bottom))
        x += ('<hp:footNotePr><hp:autoNumFormat type="DIGIT" userChar="" prefixChar="" suffixChar=")" '
              'supscript="0"/><hp:noteLine length="-1" type="SOLID" width="0.12 mm" color="#000000"/>'
              '<hp:noteSpacing betweenNotes="283" belowLine="567" aboveLine="850"/>'
              '<hp:numbering type="CONTINUOUS" newNum="1"/><hp:placement place="EACH_COLUMN" '
              'beneathText="0"/></hp:footNotePr>')
        x += ('<hp:endNotePr><hp:autoNumFormat type="DIGIT" userChar="" prefixChar="" suffixChar=")" '
              'supscript="0"/><hp:noteLine length="14692344" type="SOLID" width="0.12 mm" '
              'color="#000000"/><hp:noteSpacing betweenNotes="0" belowLine="567" aboveLine="850"/>'
              '<hp:numbering type="CONTINUOUS" newNum="1"/><hp:placement place="END_OF_DOCUMENT" '
              'beneathText="0"/></hp:endNotePr>')
        for t in ("BOTH", "EVEN", "ODD"):
            x += ('<hp:pageBorderFill type="%s" borderFillIDRef="1" textBorder="PAPER" headerInside="0" '
                  'footerInside="0" fillArea="PAPER"><hp:offset left="1417" right="1417" top="1417" '
                  'bottom="1417"/></hp:pageBorderFill>' % t)
        x += "</hp:secPr>"
        x += self.colpr_xml(sp)
        x += header_ctrls
        return x

    @staticmethod
    def colpr_xml(sp):
        if sp["cols"] <= 1:
            return ('<hp:ctrl><hp:colPr id="" type="NEWSPAPER" layout="LEFT" colCount="1" sameSz="1" '
                    'sameGap="0"/></hp:ctrl>')
        line = ('<hp:colLine type="SOLID" width="0.12 mm" color="#000000"/>' if sp["col_sep"] else "")
        return ('<hp:ctrl><hp:colPr id="" type="NEWSPAPER" layout="LEFT" colCount="%d" sameSz="1" '
                'sameGap="%d">%s</hp:colPr></hp:ctrl>' % (sp["cols"], sp["col_space"], line))

    def header_footer_ctrls(self, sp, text_w):
        out = []
        for kind, refs in (("header", sp["headers"]), ("footer", sp["footers"])):
            area = max(0, (sp["top"] - sp["header"]) if kind == "header" else (sp["bottom"] - sp["footer"]))
            plan = []
            if "default" in refs:
                plan.append(("ODD" if (self.dx.even_odd and "even" in refs) else "BOTH", refs["default"]))
            if self.dx.even_odd and "even" in refs:
                plan.append(("EVEN", refs["even"]))
            if sp["titlePg"]:
                first = refs.get("first")
                if first is None or self._part_is_empty(first):
                    sp["hide_first_" + kind] = True
            for apply_type, rid in plan:
                rel = self.dx.rels.get(rid)
                if not rel:
                    continue
                part = rel[1]
                root = self.dx.xml(part)
                if root is None:
                    continue
                self.rels_stack.append(self.dx.load_rels(part))
                saved_fields = getattr(self, "_fields", [])
                paras = self.blocks(list(root), Ctx(text_w, in_header=True))
                self._fields = saved_fields
                self.rels_stack.pop()
                if not paras:
                    continue
                for p in paras:
                    if not getattr(p, "is_table", False):
                        p.style = self.STYLE_HEADER if p.style == 0 else p.style
                out.append('<hp:ctrl><hp:%s id="%d" applyPageType="%s"><hp:subList id="" '
                           'textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="%s" linkListIDRef="0" '
                           'linkListNextIDRef="0" textWidth="%d" textHeight="%d" hasTextRef="0" '
                           'hasNumRef="0">%s</hp:subList></hp:%s></hp:ctrl>'
                           % (kind, self.next_id() % 100000, apply_type,
                              "TOP" if kind == "header" else "BOTTOM", text_w, area,
                              "".join(p.xml() for p in paras), kind))
        return "".join(out)

    def _part_is_empty(self, rid):
        rel = self.dx.rels.get(rid)
        if not rel:
            return True
        root = self.dx.xml(rel[1])
        if root is None:
            return True
        txt = "".join(t.text or "" for t in root.iter(w("t"))).strip()
        has_obj = any(True for _ in root.iter(w("drawing"))) or any(True for _ in root.iter(w("pict")))
        has_fld = any(True for _ in root.iter(w("fldSimple"))) or any(True for _ in root.iter(w("instrText")))
        return not txt and not has_obj and not has_fld

    # -- 전체 ----------------------------------------------------------------
    def convert(self):
        body = self.dx.doc.find(w("body"))
        if body is None:
            raise ValueError("document.xml 에 본문이 없습니다.")
        # Word 구역 단위로 자식 나누기
        sections = []
        cur = []
        for ch in list(body):
            if local(ch) == "sectPr":
                continue
            cur.append(ch)
            if local(ch) == "p":
                ppr = ch.find(w("pPr"))
                sp = ppr.find(w("sectPr")) if ppr is not None else None
                if sp is not None:
                    sections.append((cur, sp))
                    cur = []
        sections.append((cur, body.find(w("sectPr"))))

        hwp_sections = []    # [(sp, [Para])]
        prev_sp = None
        inherit = {"headers": {}, "footers": {}}
        for children, sect in sections:
            sp = self.section_props(sect)
            # Word: 머리말/꼬리말을 지정하지 않은 구역은 앞 구역 것을 이어받는다
            own_hf = bool(sp["headers"] or sp["footers"])
            for k in ("headers", "footers"):
                merged = dict(inherit[k])
                merged.update(sp[k])
                sp[k] = merged
                inherit[k] = merged
            if sp["landscape"]:
                text_w = max(sp["w"], sp["h"]) - sp["left"] - sp["right"]
            else:
                text_w = sp["w"] - sp["left"] - sp["right"]
            col_w = int((text_w - sp["col_space"] * (sp["cols"] - 1)) / max(1, sp["cols"]))
            paras = self.blocks(children, Ctx(col_w))
            if not paras:
                p = Para(0, 0)
                p.mark_cp = self.charpr_id(self.base_rpr(None))
                paras = [p]
            same_page = prev_sp is not None and all(
                sp[k] == prev_sp[k] for k in ("w", "h", "landscape", "top", "bottom", "left", "right"))
            # Word 는 sectPr 의 type 이 '다음 구역의 시작'을 뜻함
            if (sp["type"] == "continuous" and same_page and hwp_sections and not own_hf):
                # 같은 쪽에서 다단만 바뀌면 한글의 '다단 설정 나누기'로
                paras[0].prefix = self.colpr_xml(sp) + paras[0].prefix
                hwp_sections[-1][1].extend(paras)
            else:
                hdr = self.header_footer_ctrls(sp, text_w)
                hwp_sections.append((sp, paras, hdr))
            prev_sp = sp

        # Word 의 sectPr.type 은 '이 구역이 어떻게 시작하는가'라서 첫 구역 외에는
        # 한글 새 구역 = 새 쪽. (연속 구역 + 쪽 설정 변경은 새 쪽으로 처리)
        sec_xml = []
        for idx, (sp, paras, hdr) in enumerate(hwp_sections):
            paras[0].prefix = self.secpr_xml(sp, hdr) + paras[0].prefix
            body_xml = "".join(p.xml() for p in paras)
            sec_xml.append(XML_HEAD + '<hs:sec %s>%s</hs:sec>' % (HP_NS_DECL, body_xml))
        return sec_xml

    def build(self, out_path):
        sections = self.convert()
        header = self.reg.xml(len(sections))
        write_hwpx(out_path, header, sections, self.bins,
                   title=os.path.splitext(os.path.basename(self.dx.path))[0],
                   preview_text="".join(self.plain_text)[:1000])
        return out_path


# ---------------------------------------------------------------------------
# HWPX 패키지 쓰기
# ---------------------------------------------------------------------------
def write_hwpx(path, header_xml, section_xmls, bins, title="", preview_text=""):
    now = datetime.datetime.now(datetime.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    d = os.path.dirname(os.path.abspath(path))
    if not os.path.isdir(d):
        os.makedirs(d)
    version = (XML_HEAD + '<hv:HCFVersion xmlns:hv="http://www.hancom.co.kr/hwpml/2011/version" '
               'tagetApplication="WORDPROCESSOR" major="5" minor="0" micro="5" buildNumber="0" '
               'os="1" xmlVersion="1.4" application="Hancom Office Hangul" '
               'appVersion="9, 1, 1, 5656 WIN32LEWindows_Unknown_Version"/>')
    settings = (XML_HEAD + '<ha:HWPApplicationSetting xmlns:ha="http://www.hancom.co.kr/hwpml/2011/app" '
                'xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0">'
                '<ha:CaretPosition listIDRef="0" paraIDRef="0" pos="0"/></ha:HWPApplicationSetting>')
    container = (XML_HEAD + '<ocf:container xmlns:ocf="urn:oasis:names:tc:opendocument:xmlns:container" '
                 'xmlns:hpf="http://www.hancom.co.kr/schema/2011/hpf"><ocf:rootfiles>'
                 '<ocf:rootfile full-path="Contents/content.hpf" media-type="application/hwpml-package+xml"/>'
                 '<ocf:rootfile full-path="Preview/PrvText.txt" media-type="text/plain"/>'
                 '</ocf:rootfiles></ocf:container>')
    manifest = (XML_HEAD + '<odf:manifest xmlns:odf="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"/>')
    items = ['<opf:item id="header" href="Contents/header.xml" media-type="application/xml"/>']
    for bid, zpath, _data, mt in bins:
        items.append('<opf:item id="%s" href="%s" media-type="%s" isEmbeded="1"/>' % (bid, zpath, mt))
    for i in range(len(section_xmls)):
        items.append('<opf:item id="section%d" href="Contents/section%d.xml" media-type="application/xml"/>'
                     % (i, i))
    items.append('<opf:item id="settings" href="settings.xml" media-type="application/xml"/>')
    spine = ['<opf:itemref idref="header" linear="yes"/>']
    for i in range(len(section_xmls)):
        spine.append('<opf:itemref idref="section%d" linear="yes"/>' % i)
    hpf = (XML_HEAD + '<opf:package %s version="" unique-identifier="" id="">'
           '<opf:metadata><opf:title>%s</opf:title><opf:language>ko</opf:language>'
           '<opf:meta name="creator" content="text">docx2hwp</opf:meta>'
           '<opf:meta name="subject" content="text"/><opf:meta name="description" content="text"/>'
           '<opf:meta name="lastsaveby" content="text">docx2hwp</opf:meta>'
           '<opf:meta name="CreatedDate" content="text">%s</opf:meta>'
           '<opf:meta name="ModifiedDate" content="text">%s</opf:meta>'
           '<opf:meta name="keyword" content="text"/></opf:metadata>'
           '<opf:manifest>%s</opf:manifest><opf:spine>%s</opf:spine></opf:package>'
           % (HP_NS_DECL, esc(title), now, now, "".join(items), "".join(spine)))

    tmp = path + ".tmp"
    with zipfile.ZipFile(tmp, "w") as z:
        def put(name, data, compress=True):
            zi = zipfile.ZipInfo(name, date_time=(1980, 1, 1, 0, 0, 0))
            zi.compress_type = zipfile.ZIP_DEFLATED if compress else zipfile.ZIP_STORED
            if isinstance(data, str):
                data = data.encode("utf-8")
            z.writestr(zi, data)
        put("mimetype", "application/hwp+zip", compress=False)   # 반드시 첫 항목, 무압축
        put("version.xml", version, compress=False)
        put("Contents/header.xml", header_xml)
        for i, s in enumerate(section_xmls):
            put("Contents/section%d.xml" % i, s)
        for _bid, zpath, data, _mt in bins:
            put(zpath, data, compress=False)
        put("Preview/PrvText.txt", preview_text or " ", compress=False)
        put("settings.xml", settings)
        put("META-INF/container.xml", container)
        put("Contents/content.hpf", hpf)
        put("META-INF/manifest.xml", manifest)
    if os.path.exists(path):
        os.remove(path)
    os.rename(tmp, path)


# ---------------------------------------------------------------------------
# 한컴오피스 한글(Windows) 연동: .hwpx → .hwp, 또는 한글로 직접 docx 열기
# ---------------------------------------------------------------------------
def hangul_available():
    if os.name != "nt":
        return False
    try:
        import win32com.client  # noqa: F401
    except ImportError:
        return False
    try:
        import winreg
        winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "HWPFrame.HwpObject")
        return True
    except Exception:
        return False


class HangulApp(object):
    """한글 COM 자동화 (pywin32 필요: pip install pywin32)"""

    def __init__(self, visible=False):
        import win32com.client as win32
        try:
            self.hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
        except Exception:
            self.hwp = win32.Dispatch("HWPFrame.HwpObject")
        try:
            # 보안 승인 모듈이 등록돼 있으면 '파일 접근 허용' 창이 뜨지 않는다
            self.hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
        except Exception:
            pass
        try:
            self.hwp.XHwpWindows.Item(0).Visible = visible
        except Exception:
            pass

    def save_as(self, src, dst, fmt="HWP", src_fmt=""):
        src = os.path.abspath(src)
        dst = os.path.abspath(dst)
        ok = self.hwp.Open(src, src_fmt, "forceopen:true;suspendpassword:true")
        if not ok:
            raise RuntimeError("한글에서 파일을 열지 못했습니다: %s" % src)
        ok = self.hwp.SaveAs(dst, fmt, "")
        try:
            self.hwp.Clear(1)
        except Exception:
            pass
        if not ok:
            raise RuntimeError("한글에서 저장하지 못했습니다: %s" % dst)

    def quit(self):
        try:
            self.hwp.Quit()
        except Exception:
            pass


# ---------------------------------------------------------------------------
# 실행
# ---------------------------------------------------------------------------
def convert_file(src, out=None, fmt="auto", engine="builtin", hangul=None, log=print, eq_rules=None):
    """하나의 docx 변환. fmt: hwpx | hwp | both | auto"""
    base = os.path.splitext(out or src)[0]
    want_hwp = fmt in ("hwp", "both") or (out and out.lower().endswith(".hwp"))
    want_hwpx = fmt in ("hwpx", "both", "auto") or (out and out.lower().endswith(".hwpx"))
    if fmt == "auto" and not out:
        want_hwpx = True
    hwpx_path = out if (out and out.lower().endswith(".hwpx")) else base + ".hwpx"
    hwp_path = out if (out and out.lower().endswith(".hwp")) else base + ".hwp"
    results = []
    warnings = []

    if engine == "hangul":
        # 한글이 직접 docx 를 읽는 방식 (한글 자체 변환기 사용)
        if hangul is None:
            raise RuntimeError("한글(한컴오피스)과 pywin32 가 필요합니다.")
        if want_hwp:
            hangul.save_as(src, hwp_path, "HWP")
            results.append(hwp_path)
        if want_hwpx and fmt != "hwp":
            hangul.save_as(src, hwpx_path, "HWPX")
            results.append(hwpx_path)
        return results, warnings

    conv = Converter(src, eq_rules=eq_rules)
    tmp_hwpx = hwpx_path if want_hwpx else base + ".__tmp__.hwpx"
    conv.build(tmp_hwpx)
    warnings = conv.warnings
    if want_hwpx:
        results.append(tmp_hwpx)
    if want_hwp:
        if hangul is None:
            warnings.append(".hwp 저장에는 Windows + 한컴오피스 한글 + pywin32 가 필요합니다. "
                            ".hwpx 로 저장했습니다 (한글 2014 이상에서 그대로 열림).")
            if not want_hwpx:
                os.rename(tmp_hwpx, hwpx_path)
                results.append(hwpx_path)
        else:
            hangul.save_as(tmp_hwpx, hwp_path, "HWP", "HWPX")
            results.append(hwp_path)
            if not want_hwpx:
                try:
                    os.remove(tmp_hwpx)
                except OSError:
                    pass
    return results, warnings


def collect_inputs(paths):
    files = []
    for p in paths:
        if os.path.isdir(p):
            for f in sorted(glob.glob(os.path.join(p, "**", "*.docx"), recursive=True)):
                if not os.path.basename(f).startswith("~$"):
                    files.append(f)
        else:
            files.extend(sorted(glob.glob(p)) or [p])
    return files


def run_gui():
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk

    root = tk.Tk()
    root.title("docx → 한글 변환기 %s" % __version__)
    root.geometry("560x380")
    files = []
    fmt_var = tk.StringVar(value="both" if hangul_available() else "hwpx")
    status = tk.StringVar(value="변환할 .docx 파일을 고르세요.")

    def pick():
        sel = filedialog.askopenfilenames(filetypes=[("Word 문서", "*.docx")])
        if sel:
            files[:] = list(sel)
            lb.delete(0, tk.END)
            for f in files:
                lb.insert(tk.END, f)
            status.set("%d개 파일 선택됨" % len(files))

    def go():
        if not files:
            messagebox.showinfo("알림", "먼저 파일을 고르세요.")
            return
        fmt = fmt_var.get()
        hangul = None
        if fmt in ("hwp", "both"):
            if hangul_available():
                hangul = HangulApp()
            else:
                messagebox.showwarning("알림", ".hwp 저장은 한글이 설치된 Windows 에서만 됩니다. .hwpx 로 저장합니다.")
                fmt = "hwpx"
        done, errs = 0, []
        for f in files:
            status.set("변환 중: %s" % os.path.basename(f))
            root.update()
            try:
                convert_file(f, fmt=fmt, hangul=hangul)
                done += 1
            except Exception as e:  # noqa
                errs.append("%s: %s" % (os.path.basename(f), e))
        if hangul:
            hangul.quit()
        status.set("완료: %d개 / 실패 %d개" % (done, len(errs)))
        if errs:
            messagebox.showerror("일부 실패", "\n".join(errs))
        else:
            messagebox.showinfo("완료", "%d개 파일을 변환했습니다.\n(원본과 같은 폴더에 저장)" % done)

    frm = ttk.Frame(root, padding=12)
    frm.pack(fill="both", expand=True)
    ttk.Button(frm, text="파일 고르기…", command=pick).pack(anchor="w")
    lb = tk.Listbox(frm, height=10)
    lb.pack(fill="both", expand=True, pady=8)
    opt = ttk.Frame(frm)
    opt.pack(anchor="w")
    ttk.Label(opt, text="저장 형식: ").pack(side="left")
    for v, t in (("hwpx", ".hwpx"), ("hwp", ".hwp (한글 필요)"), ("both", "둘 다")):
        ttk.Radiobutton(opt, text=t, value=v, variable=fmt_var).pack(side="left")
    ttk.Button(frm, text="변환", command=go).pack(anchor="e", pady=6)
    ttk.Label(frm, textvariable=status).pack(anchor="w")
    ttk.Label(frm, text=HANCOM_NOTICE, foreground="gray", wraplength=520).pack(anchor="w", pady=(8, 0))
    root.mainloop()


def run_dropped(paths):
    """창 앱(.exe/.app)에 파일을 끌어다 놓았을 때: 바로 변환하고 결과를 창으로 알림"""
    import tkinter as tk
    from tkinter import messagebox
    root = tk.Tk()
    root.withdraw()
    done, errs, notes = [], [], []
    for f in collect_inputs(paths):
        try:
            res, warns = convert_file(f, fmt="hwpx")
            done.extend(res)
            notes.extend(warns)
        except Exception as e:  # noqa
            errs.append("%s: %s" % (os.path.basename(f), e))
    msg = "%d개 파일을 변환했습니다.\n\n%s" % (len(done), "\n".join(os.path.basename(d) for d in done))
    if notes:
        msg += "\n\n참고:\n" + "\n".join(sorted(set(notes)))
    if errs:
        messagebox.showerror("docx2hwp", msg + "\n\n실패:\n" + "\n".join(errs))
    else:
        messagebox.showinfo("docx2hwp", msg)
    root.destroy()


def main(argv=None):
    ap = argparse.ArgumentParser(
        prog="docx2hwp",
        description="Word(.docx) → 한글(.hwpx/.hwp) 변환기. 수식은 편집 가능한 한글 수식으로 바뀝니다.",
        epilog=HANCOM_NOTICE)
    ap.add_argument("inputs", nargs="*", help=".docx 파일 또는 폴더 (여러 개 가능)")
    ap.add_argument("-o", "--output", help="출력 파일 경로 (.hwpx 또는 .hwp). 입력이 하나일 때만")
    ap.add_argument("-d", "--outdir", help="출력 폴더")
    ap.add_argument("-f", "--format", choices=["hwpx", "hwp", "both", "auto"], default="auto",
                    help="저장 형식 (기본: hwpx, 한글 설치 시 -f hwp 또는 both 가능)")
    ap.add_argument("--engine", choices=["builtin", "hangul"], default="builtin",
                    help="builtin: 이 프로그램의 변환기(기본) / hangul: 한글 프로그램의 docx 불러오기 사용")
    ap.add_argument("--gui", action="store_true", help="파일 선택 창으로 실행")
    ap.add_argument("--rules", help="수식 교정 규칙 파일(.json, 웹 앱에서 내보낸 것)")
    ap.add_argument("-q", "--quiet", action="store_true")
    ap.add_argument("--version", action="version", version=__version__)
    # 콘솔 없는 창 앱(PyInstaller --windowed)에서 파일을 끌어다 놓은 경우
    raw = sys.argv[1:] if argv is None else argv
    if getattr(sys, "frozen", False) and sys.stdout is None:
        paths = [a for a in raw if not a.startswith("-")]
        if paths:
            run_dropped(paths)
        else:
            run_gui()
        return 0
    args = ap.parse_args(argv)

    if args.gui or not args.inputs:
        if not args.inputs:
            try:
                run_gui()
                return 0
            except Exception:
                ap.print_help()
                return 1
        run_gui()
        return 0

    files = collect_inputs(args.inputs)
    if not files:
        print("변환할 .docx 파일이 없습니다.", file=sys.stderr)
        return 1
    if args.output and len(files) > 1:
        print("-o 는 입력이 하나일 때만 쓸 수 있습니다. 여러 개는 -d 를 쓰세요.", file=sys.stderr)
        return 1

    need_hangul = args.engine == "hangul" or args.format in ("hwp", "both") or \
        (args.output or "").lower().endswith(".hwp")
    hangul = None
    if need_hangul:
        if hangul_available():
            hangul = HangulApp()
        elif args.engine == "hangul":
            print("--engine hangul 은 Windows + 한컴오피스 한글 + pywin32 가 필요합니다.", file=sys.stderr)
            return 1

    rules = eqlearn.load(args.rules) if args.rules else None
    ok = 0
    try:
        for f in files:
            out = args.output
            if args.outdir:
                os.makedirs(args.outdir, exist_ok=True)
                stem = os.path.splitext(os.path.basename(f))[0]
                ext = ".hwp" if args.format == "hwp" else ".hwpx"
                out = os.path.join(args.outdir, stem + ext)
                fmt = args.format
            else:
                fmt = args.format
            try:
                res, warns = convert_file(f, out, fmt, args.engine, hangul, eq_rules=rules)
                ok += 1
                if not args.quiet:
                    for r in res:
                        print("[완료] %s -> %s" % (f, r))
                    for wmsg in warns:
                        print("  ! %s" % wmsg)
            except Exception as e:
                print("[실패] %s: %s" % (f, e), file=sys.stderr)
                if os.environ.get("DOCX2HWP_DEBUG"):
                    raise
    finally:
        if hangul is not None:
            hangul.quit()
    return 0 if ok == len(files) else 2


if __name__ == "__main__":
    sys.exit(main())
