# -*- coding: utf-8 -*-
"""
OMML(Word 수식) → 한글(HWP/HWPX) 수식 스크립트 변환기

Word 의 <m:oMath> 트리를 한글 수식 편집기 문법(예: {a} over {b}, sqrt {x},
sum _{i=1} ^{n}, matrix {a & b # c & d})으로 바꾸고, 한글이 수식 개체의 크기를
잡는 데 쓸 대략적인 가로/세로 크기(em 단위)도 함께 계산한다.

표준 라이브러리만 사용한다.
"""
import re

M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _m(tag):
    return "{%s}%s" % (M_NS, tag)


def _local(el):
    t = el.tag
    return t.split("}", 1)[1] if "}" in t else t


def _ns(el):
    t = el.tag
    return t[1:].split("}", 1)[0] if t.startswith("{") else ""


# ---------------------------------------------------------------------------
# 기호 표
# ---------------------------------------------------------------------------
GREEK = {
    "α": "alpha", "β": "beta", "γ": "gamma", "δ": "delta", "ε": "epsilon",
    "ϵ": "epsilon", "ζ": "zeta", "η": "eta", "θ": "theta", "ϑ": "vartheta",
    "ι": "iota", "κ": "kappa", "λ": "lambda", "μ": "mu", "ν": "nu", "ξ": "xi",
    "ο": "omicron", "π": "pi", "ϖ": "varpi", "ρ": "rho", "ϱ": "rho",
    "σ": "sigma", "ς": "varsigma", "τ": "tau", "υ": "upsilon", "φ": "phi",
    "ϕ": "varphi", "χ": "chi", "ψ": "psi", "ω": "omega",
    "Α": "ALPHA", "Β": "BETA", "Γ": "GAMMA", "Δ": "DELTA", "Ε": "EPSILON",
    "Ζ": "ZETA", "Η": "ETA", "Θ": "THETA", "Ι": "IOTA", "Κ": "KAPPA",
    "Λ": "LAMBDA", "Μ": "MU", "Ν": "NU", "Ξ": "XI", "Ο": "OMICRON", "Π": "PI",
    "Ρ": "RHO", "Σ": "SIGMA", "Τ": "TAU", "Υ": "UPSILON", "Φ": "PHI",
    "Χ": "CHI", "Ψ": "PSI", "Ω": "OMEGA",
}

# 한글 수식 공식 문서(한글문서파일형식 수식 rev1.3)에 있는 명령어 위주
SYMBOLS = {
    "±": "+-", "∓": "-+", "×": "times", "÷": "DIV",
    "·": "cdot", "⋅": "cdot", "∙": "BULLET", "∘": "CIRC", "°": "DEG",
    "∗": "AST", "⋆": "STAR", "∩": "CAP", "∪": "CUP", "⊕": "OPLUS",
    "⊖": "OMINUS", "⊗": "OTIMES", "⊙": "ODOT", "⊘": "OSLASH", "∨": "VEE",
    "∧": "WEDGE", "¬": "LNOT", "⊂": "SUBSET", "⊃": "SUPSET",
    "⊆": "SUBSETEQ", "⊇": "SUPSETEQ", "⊏": "SQSUBSET", "⊐": "SQSUPSET",
    "∈": "IN", "∋": "OWNS", "∉": "notin", "≤": "LEQ", "≦": "LEQ",
    "≥": "GEQ", "≧": "GEQ", "≪": "<<", "≫": ">>", "≺": "PREC", "≻": "SUCC",
    "≠": "!=", "≐": "DOTEQ", "∼": "SIM", "≈": "APPROX", "≃": "SIMEQ",
    "≅": "CONG", "≡": "EQUIV", "≍": "ASYMP", "∝": "PROPTO", "⋄": "DIAMOND",
    "∅": "EMPTYSET", "∴": "THEREFORE", "∵": "BECAUSE", "∃": "EXIST",
    "∀": "FORALL", "′": "prime", "″": "prime prime", "‴": "prime prime prime",
    "∂": "PARTIAL", "∞": "inf", "†": "DAGGER", "‡": "DDAGGER",
    "△": "TRIANGLE", "∠": "ANGLE", "⊢": "VDASH", "⊥": "BOT", "⊤": "TOP",
    "⊨": "MODELS", "ℵ": "ALEPH", "ℏ": "HBAR", "ı": "IMATH", "ȷ": "JMATH",
    "ℓ": "ELL", "℘": "WP", "ℑ": "IMAG", "Å": "ANGSTROM", "℃": "CENTIGRADE",
    "℉": "FAHRENHEIT",
    "←": "<-", "→": "->", "↑": "uparrow", "↓": "downarrow",
    "↕": "udarrow", "↔": "<->", "↖": "nwarrow", "↘": "searrow",
    "↗": "nearrow", "↙": "swarrow", "↩": "hookleft", "↪": "hookright",
    "↦": "mapsto", "⇐": "LARROW", "⇒": "RARROW", "⇑": "UPARROW",
    "⇓": "DOWNARROW", "⇕": "UDARROW", "⇔": "LRARROW", "⟵": "larrow",
    "⟶": "rarrow", "⟹": "RARROW", "⟸": "LARROW", "⟺": "LRARROW",
    "⋯": "cdots", "…": "LDOTS", "⋮": "VDOTS", "⋱": "DDOTS",
    "‖": "VERT", "−": "-", "–": "-", "∕": "/", "∣": "|", "⁄": "/",
    "＝": "=", "＋": "+", "\u2146": "d", "\u2147": "e", "\u2148": "i",
    "\u2061": "", "\u2062": "", "\u2063": "", "\u2064": "", "\u200b": "",
}

DOUBLE_STRUCK = {"R": "ℝ", "N": "ℕ", "Z": "ℤ", "Q": "ℚ", "C": "ℂ", "P": "ℙ", "H": "ℍ"}

# 큰 연산자 (m:nary 의 chr)
NARY = {
    "∑": "sum", "∏": "prod", "∐": "COPROD", "∫": "int", "∬": "dint",
    "∭": "tint", "∮": "oint", "∯": "ODINT", "∰": "OTINT", "⋃": "UNION",
    "⋂": "INTER", "⋁": "VEE", "⋀": "WEDGE", "⨁": "OPLUS", "⨂": "OTIMES",
}

# 한글이 로만체로 자동 처리하는 함수/예약어
FUNCS = {
    "sin", "cos", "tan", "cot", "sec", "csc", "cosec", "coth", "log", "ln",
    "lg", "exp", "Exp", "max", "min", "arcsin", "arccos", "arctan", "arc",
    "sinh", "cosh", "tanh", "det", "gcd", "lcm", "mod", "lim", "Lim",
    "asin", "acos", "atan", "hom", "ker", "deg", "arg", "dim", "Pr",
}

# 악센트 (m:acc chr)
ACCENTS = {
    "̂": "hat", "^": "hat", "̌": "check", "́": "acute",
    "̀": "grave", "̇": "dot", "̈": "ddot", "̄": "bar",
    "̅": "bar", "¯": "bar", "⃗": "vec", "→": "vec",
    "⃖": "vec", "⃡": "dyad", "̃": "tilde", "~": "tilde",
    "̆": "arch", "⌢": "arch", "̑": "arch",
}

ACCENTS.update({"\u203e": "bar", "\u0332": "under", "\u02c6": "hat",
                "\u02dc": "tilde", "\u02d9": "dot", "\u00a8": "ddot"})

# 괄호 문자 → 한글 수식에서 left/right 뒤에 쓸 표현
BRACKETS = {
    "(": "(", ")": ")", "[": "[", "]": "]", "{": "lbrace", "}": "rbrace",
    "|": "|", "‖": "VERT", "⟨": "<", "⟩": ">", "〈": "<", "〉": ">",
    "\u2329": "<", "\u232a": ">", "\u27e8": "<", "\u27e9": ">",
    "⌊": "⌊", "⌋": "⌋", "⌈": "⌈", "⌉": "⌉",
}

# 크기 추정용 (em)
_REL_WIDTH = 0.85      # 관계·이항 연산자(앞뒤 여백 포함)
_FRAC_SCALE = 0.6      # 분수 안쪽 글자 축소 비율 (한글 수식 실측 기준)
_SCRIPT_SCALE = 0.62   # 첨자 축소 비율


class Box(object):
    """수식 조각: 한글 스크립트 + 대략적 크기(em)."""
    __slots__ = ("s", "w", "a", "d")

    def __init__(self, s="", w=0.0, a=0.0, d=0.0):
        self.s = s
        self.w = w
        self.a = a
        self.d = d

    @property
    def h(self):
        return self.a + self.d

    def group(self):
        """다른 명령의 인자로 쓸 때 {…}로 감싼 문자열."""
        s = self.s.strip()
        if not s:
            return "{}"
        if s.startswith("{") and _balanced_outer(s):
            return s
        return "{" + s + "}"


def _balanced_outer(s):
    depth = 0
    for i, ch in enumerate(s):
        if ch == "{":
            depth += 1
        elif ch == "}":
            depth -= 1
            if depth == 0 and i != len(s) - 1:
                return False
    return depth == 0


def _hjoin(boxes):
    boxes = [b for b in boxes if b is not None and b.s.strip()]
    if not boxes:
        return Box("", 0, 0.7, 0.2)
    s = " ".join(b.s.strip() for b in boxes)
    return Box(s, sum(b.w for b in boxes), max(b.a for b in boxes),
               max(b.d for b in boxes))


def _scaled(b, k):
    return Box(b.s, b.w * k, b.a * k, b.d * k)


# ---------------------------------------------------------------------------
# 텍스트 토큰화
# ---------------------------------------------------------------------------
_HANGUL_RE = re.compile(r"[ᄀ-ᇿ㄰-㆏가-힣一-鿿]")


def _char_width(ch):
    if _HANGUL_RE.match(ch):
        return 1.0
    if ch in "il.,;:!|'`":
        return 0.3
    if ch in "mwMW":
        return 0.8
    return 0.55


def _is_latin_letter(ch):
    if ("a" <= ch <= "z") or ("A" <= ch <= "Z"):
        return True
    return ch.isalpha() and "\u00c0" <= ch <= "\u024f"


def text_to_tokens(text, roman=False, bold=False):
    """Word 수식 run 의 글자를 한글 수식 토큰(Box) 목록으로."""
    out = []
    i = 0
    n = len(text)
    while i < n:
        ch = text[i]
        # 공백: 한글 수식은 공백을 무시하므로 버린다 (일반 텍스트 모드는 ~)
        if ch in "      ":
            if roman and out:
                out.append(Box("~", 0.3, 0.7, 0.2))
            i += 1
            continue
        if ch.isdigit() or (ch == "." and i + 1 < n and text[i + 1].isdigit()):
            j = i
            while j < n and (text[j].isdigit() or text[j] == "."):
                j += 1
            num = text[i:j]
            out.append(Box(num, 0.45 * len(num), 0.7, 0.05))
            i = j
            continue
        if _is_latin_letter(ch):
            j = i
            while j < n and _is_latin_letter(text[j]):
                j += 1
            word = text[i:j]
            w = sum(_char_width(c) for c in word)
            if word in FUNCS:
                out.append(Box(word, w + 0.2, 0.72, 0.2))
            elif roman:
                out.append(Box('"%s"' % word, w, 0.72, 0.2))
            else:
                # 이탤릭 낱말은 그대로(4ac). 예약어(over, in, pi …)와 겹칠 때만 띄운다
                if word.lower() in RESERVED:
                    out.append(Box(" ".join(word), w, 0.72, 0.2))
                else:
                    out.append(Box(word, w, 0.72, 0.2))
            i = j
            continue
        if _HANGUL_RE.match(ch):
            j = i
            while j < n and (_HANGUL_RE.match(text[j]) or text[j] == " "):
                j += 1
            word = text[i:j].strip()
            parts = [p for p in word.split(" ") if p]
            s = "~".join('"%s"' % p for p in parts)
            out.append(Box(s, sum(_char_width(c) for c in word), 0.8, 0.2))
            i = j
            continue
        if ch in GREEK:
            out.append(Box(GREEK[ch], 0.55, 0.72, 0.2))
        elif ch in NARY:
            out.append(Box(NARY[ch], 1.0, 1.0, 0.4))
        elif ch in SYMBOLS:
            kw = SYMBOLS[ch]
            out.append(Box(kw, _REL_WIDTH, 0.7, 0.15))
        elif ch == "{":
            out.append(Box("lbrace", 0.4, 0.75, 0.25))
        elif ch == "}":
            out.append(Box("rbrace", 0.4, 0.75, 0.25))
        elif ch == "~":
            out.append(Box("SIM", _REL_WIDTH, 0.6, 0.1))
        elif ch == "&":
            out.append(Box("&", 0, 0, 0))
        elif ch in '#^_`"\\':
            out.append(Box('"%s"' % ch if ch != '"' else "″", 0.5, 0.7, 0.1))
        elif ch in "+-=<>":
            out.append(Box(ch, _REL_WIDTH, 0.6, 0.1))
        elif ch in "()[]|":
            out.append(Box(ch, 0.35, 0.75, 0.25))
        elif ch in ",;:!./'*":
            out.append(Box(ch, 0.3, 0.7, 0.15))
        else:
            # 그 밖의 유니코드 문자는 그대로 둔다 (한글이 수식 글꼴로 그림)
            out.append(Box(ch, 0.6, 0.72, 0.2))
        i += 1
    if bold and out:
        joined = _hjoin(out)
        return [Box("bold{" + joined.s + "}", joined.w * 1.05, joined.a, joined.d)]
    return out


# ---------------------------------------------------------------------------
# OMML 변환기
# ---------------------------------------------------------------------------
def _child(el, name):
    if el is None:
        return None
    return el.find(_m(name))


def _val(el, name, default=None):
    """<m:xxxPr><m:name m:val="…"/></…> 형태의 값."""
    if el is None:
        return default
    c = el.find(_m(name))
    if c is None:
        return default
    v = c.get(_m("val"))
    if v is None:
        v = c.get("{%s}val" % W_NS)
    return default if v is None else v


def _on(el, name, default=False):
    v = _val(el, name, None)
    if v is None:
        if el is not None and el.find(_m(name)) is not None:
            return True
        return default
    return v not in ("0", "off", "false")


class OmmlConverter(object):
    def __init__(self):
        self.warnings = []

    # -- 진입점 ------------------------------------------------------------
    def convert(self, omath):
        """<m:oMath> 하나 → (script, width_em, ascent_em, descent_em)"""
        box = self.seq(omath)
        script = tidy(box.s)
        return script, max(box.w, 0.5), max(box.a, 0.7), max(box.d, 0.2)

    # -- 공통 --------------------------------------------------------------
    def seq(self, el):
        """여러 자식을 차례대로 이어 붙인다."""
        if el is None:
            return Box("", 0, 0.7, 0.2)
        parts = []
        for c in list(el):
            b = self.node(c)
            if b is not None:
                parts.append(b)
        return _hjoin(parts)

    def arg(self, el, name):
        return self.seq(_child(el, name))

    def node(self, el):
        ns = _ns(el)
        tag = _local(el)
        if ns == M_NS:
            fn = getattr(self, "m_" + tag, None)
            if fn is not None:
                return fn(el)
            if tag.endswith("Pr"):
                return None
            return self.seq(el)
        if ns == W_NS:
            if tag == "r":   # 수식 안의 일반 Word run
                txt = "".join(t.text or "" for t in el.iter("{%s}t" % W_NS))
                return _hjoin(text_to_tokens(txt, roman=True))
            if tag in ("rPr", "bookmarkStart", "bookmarkEnd", "proofErr",
                       "del", "moveFrom"):
                return None
            return self.seq(el)
        if tag == "AlternateContent":
            for c in el:
                if _local(c) == "Fallback":
                    return self.seq(c)
            if len(el):
                return self.seq(el[0])
        return None

    # -- 원소별 ------------------------------------------------------------
    def m_oMath(self, el):
        return self.seq(el)

    def m_oMathPara(self, el):
        rows = [self.seq(c) for c in el if _local(c) == "oMath"]
        return self._pile(rows, "pile")

    def m_r(self, el):
        rpr = _child(el, "rPr")
        sty = _val(rpr, "sty", "i")
        nor = _on(rpr, "nor")
        roman = nor or sty in ("p", "b")
        bold = sty in ("b", "bi")
        text = "".join(t.text or "" for t in el.iter(_m("t")))
        if not text:
            return None
        scr = _val(rpr, "scr")
        if scr == "double-struck":
            text = "".join(DOUBLE_STRUCK.get(c, c) for c in text)
        # 일반 텍스트 모드의 문장은 통째로 따옴표
        if nor and re.search(r"[A-Za-zㄱ-힣]", text):
            toks = []
            for chunk in re.split(r"(\s+)", text):
                if not chunk:
                    continue
                if chunk.isspace():
                    toks.append(Box("~", 0.3, 0.7, 0.2))
                else:
                    toks.append(Box('"%s"' % chunk.replace('"', "″"),
                                    sum(_char_width(c) for c in chunk), 0.75, 0.2))
            return _hjoin(toks)
        return _hjoin(text_to_tokens(text, roman=roman, bold=bold))

    def m_f(self, el):
        pr = _child(el, "fPr")
        typ = _val(pr, "type", "bar")
        num = self.arg(el, "num")
        den = self.arg(el, "den")
        if typ == "lin":
            return Box("%s / %s" % (num.group(), den.group()),
                       num.w + den.w + 0.5, max(num.a, den.a), max(num.d, den.d))
        if typ == "skw":
            return Box("%s / %s" % (num.group(), den.group()),
                       (num.w + den.w) * 0.8 + 0.5, max(num.a, den.a) + 0.2,
                       max(num.d, den.d) + 0.2)
        op = "atop" if typ == "noBar" else "over"
        k = _FRAC_SCALE
        w = max(num.w, den.w) * k + 0.25
        axis = 0.28
        a = axis + 0.12 + num.h * k
        d = den.h * k + 0.12 - axis
        return Box("%s %s %s" % (num.group(), op, den.group()), w, a, max(d, 0.25))

    def m_rad(self, el):
        pr = _child(el, "radPr")
        e = self.arg(el, "e")
        deg = self.arg(el, "deg")
        hide = _on(pr, "degHide")
        if hide or not deg.s.strip():
            return Box("sqrt %s" % e.group(), e.w + 0.6, e.a + 0.2, e.d + 0.05)
        return Box("root %s of %s" % (deg.group(), e.group()),
                   e.w + 0.75 + deg.w * _SCRIPT_SCALE, e.a + 0.25, e.d + 0.05)

    def _script(self, base, sub=None, sup=None):
        bs = base.s.strip()
        if not bs:
            s = "{}"
        elif re.match(r"^[A-Za-z0-9.]+$", bs) or (bs.startswith("{") and _balanced_outer(bs)):
            s = bs            # x ^{2}, sin ^{2}: 한 토큰이면 괄호 불필요
        else:
            s = base.group()
        w = base.w
        a = base.a
        d = base.d
        k = _SCRIPT_SCALE
        extra = 0.0
        if sub is not None and sub.s.strip():
            s += " _" + sub.group()
            extra = max(extra, sub.w * k)
            d = max(d, 0.25 + sub.h * k * 0.75)
        if sup is not None and sup.s.strip():
            s += " ^" + sup.group()
            extra = max(extra, sup.w * k)
            a = max(a, base.a * 0.55 + sup.h * k)
        return Box(s, w + extra + 0.05, a, d)

    def m_sSup(self, el):
        return self._script(self.arg(el, "e"), sup=self.arg(el, "sup"))

    def m_sSub(self, el):
        return self._script(self.arg(el, "e"), sub=self.arg(el, "sub"))

    def m_sSubSup(self, el):
        return self._script(self.arg(el, "e"), sub=self.arg(el, "sub"),
                            sup=self.arg(el, "sup"))

    def m_sPre(self, el):
        pre = self._script(Box("", 0, 0.7, 0.2), sub=self.arg(el, "sub"),
                           sup=self.arg(el, "sup"))
        e = self.arg(el, "e")
        return _hjoin([pre, e])

    def m_nary(self, el):
        pr = _child(el, "naryPr")
        ch = _val(pr, "chr", "∫")
        op = NARY.get(ch)
        if op is None:
            op = SYMBOLS.get(ch, ch)
        sub = None if _on(pr, "subHide") else self.arg(el, "sub")
        sup = None if _on(pr, "supHide") else self.arg(el, "sup")
        e = self.arg(el, "e")
        under_over = _val(pr, "limLoc", "undOvr" if op not in
                          ("int", "dint", "tint", "oint", "ODINT", "OTINT")
                          else "subSup") == "undOvr"
        s = op
        k = _SCRIPT_SCALE
        opw = 1.05
        a, d = 1.3, 0.6
        lim_w = 0.0
        if sub is not None and sub.s.strip():
            s += " _" + sub.group()
            lim_w = max(lim_w, sub.w * k)
            d += sub.h * k + 0.1 if under_over else 0.25
        if sup is not None and sup.s.strip():
            s += " ^" + sup.group()
            lim_w = max(lim_w, sup.w * k)
            a += sup.h * k + 0.1 if under_over else 0.25
        w = max(opw, lim_w) if under_over else opw + lim_w
        body = e.s.strip()   # 한글은 sum 다음 한 항을 그대로 받음 (sum _{k=1} ^{n} k ^{2})
        if body and (" over " in body or " atop " in body):
            body = e.group()
        return Box((s + " " + body).strip(), w + e.w + 0.15,
                   max(a, e.a), max(d, e.d))

    def m_d(self, el):
        pr = _child(el, "dPr")
        beg = _val(pr, "begChr", "(")
        end = _val(pr, "endChr", ")")
        sep = _val(pr, "sepChr", "|")
        es = [self.seq(c) for c in el if _local(c) == "e"]

        # 행렬을 감싼 괄호 → pmatrix / bmatrix / dmatrix
        if len(es) == 1:
            inner = [c for c in el if _local(c) == "e"][0]
            kids = [c for c in inner if _local(c) not in ("ctrlPr", "argPr")]
            if len(kids) == 1 and _local(kids[0]) == "m":
                kind = {("(", ")"): "pmatrix", ("[", "]"): "bmatrix",
                        ("|", "|"): "dmatrix"}.get((beg, end))
                if kind:
                    return self.m_m(kids[0], kind=kind)
            # { 와 eqArr / 행렬 → cases
            if len(kids) == 1 and beg == "{" and not end:
                if _local(kids[0]) == "eqArr":
                    rows = [self.seq(r) for r in kids[0] if _local(r) == "e"]
                    return self._pile(rows, "cases")
                if _local(kids[0]) == "m":
                    rows = []
                    for mr in kids[0]:
                        if _local(mr) != "mr":
                            continue
                        cells = [self.seq(c) for c in mr if _local(c) == "e"]
                        rows.append(Box(" & ".join(c.s.strip() or "~" for c in cells),
                                        sum(c.w for c in cells) + 1.0 * (len(cells) - 1),
                                        max([c.a for c in cells] + [0.7]),
                                        max([c.d for c in cells] + [0.2])))
                    return self._pile(rows, "cases")

        sep_tok = BRACKETS.get(sep, sep) if sep else ""
        parts = []
        for i, b in enumerate(es):
            if i and sep_tok:
                parts.append(Box(sep_tok, 0.4, b.a, b.d))
            parts.append(b)
        inner = _hjoin(parts) if parts else Box("", 0, 0.7, 0.2)
        btok = BRACKETS.get(beg, beg) if beg else ""
        etok = BRACKETS.get(end, end) if end else ""
        tall = inner.h > 1.25
        body = inner.s.strip() or "~"
        if btok and etok:
            # LEFT/RIGHT: 안쪽 높이에 맞춰 괄호 크기가 자동 조절된다
            s = "LEFT %s %s RIGHT %s" % (btok, body, etok)
        elif btok:
            s = "%s %s" % (btok, body)
        elif etok:
            s = "%s %s" % (body, etok)
        else:
            s = body
        pad = (0.35 if btok else 0) + (0.35 if etok else 0)
        return Box(s, inner.w + pad, inner.a + (0.1 if tall else 0),
                   inner.d + (0.1 if tall else 0))

    def _pile(self, rows, kind):
        if not rows:
            return Box("", 0, 0.7, 0.2)
        rows = [r if r.s.strip() else Box("~", 0.3, 0.7, 0.2) for r in rows]
        has_amp = any("&" in r.s for r in rows)
        if kind == "pile" and has_amp:
            kind = "eqalign"
        if kind == "cases":
            body = " # ".join(r.s.strip() for r in rows)
        else:
            body = " # ".join(r.s.strip() for r in rows)
        h = sum(r.h for r in rows) + 0.25 * (len(rows) - 1)
        w = max(r.w for r in rows) + (0.6 if kind == "cases" else 0)
        axis = 0.28
        return Box("%s {%s}" % (kind, body), w, h / 2 + axis, h / 2 - axis + 0.1)

    def m_eqArr(self, el):
        rows = [self.seq(r) for r in el if _local(r) == "e"]
        return self._pile(rows, "pile")

    def m_m(self, el, kind="matrix"):
        rows = []
        for mr in el:
            if _local(mr) != "mr":
                continue
            rows.append([self.seq(c) for c in mr if _local(c) == "e"])
        if not rows:
            return Box("", 0, 0.7, 0.2)
        ncol = max(len(r) for r in rows)
        colw = [0.0] * ncol
        for r in rows:
            for j, c in enumerate(r):
                colw[j] = max(colw[j], c.w)
        body = " # ".join(" & ".join((c.s.strip() or "~") for c in r) for r in rows)
        h = sum(max([c.h for c in r] + [0.9]) for r in rows) + 0.3 * (len(rows) - 1)
        w = sum(colw) + 0.9 * (ncol - 1) + (0.8 if kind != "matrix" else 0)
        axis = 0.28
        return Box("%s {%s}" % (kind, body), w, h / 2 + axis, h / 2 - axis + 0.1)

    def m_func(self, el):
        name = self.arg(el, "fName")
        e = self.arg(el, "e")
        return Box("%s %s" % (name.s.strip(), e.group() if e.s.strip() else ""),
                   name.w + e.w + 0.2, max(name.a, e.a), max(name.d, e.d))

    def m_limLow(self, el):
        e = self.arg(el, "e")
        lim = self.arg(el, "lim")
        return Box("%s _%s" % (e.group() if not _is_word(e.s) else e.s.strip(),
                               lim.group()),
                   max(e.w, lim.w * _SCRIPT_SCALE), e.a, e.d + lim.h * _SCRIPT_SCALE)

    def m_limUpp(self, el):
        e = self.arg(el, "e")
        lim = self.arg(el, "lim")
        return Box("%s ^%s" % (e.group(), lim.group()),
                   max(e.w, lim.w * _SCRIPT_SCALE), e.a + lim.h * _SCRIPT_SCALE, e.d)

    def m_acc(self, el):
        pr = _child(el, "accPr")
        ch = _val(pr, "chr", "̂")
        kw = ACCENTS.get(ch, "hat")
        e = self.arg(el, "e")
        return Box("%s %s" % (kw, e.group()), e.w, e.a + 0.25, e.d)

    def m_bar(self, el):
        pr = _child(el, "barPr")
        pos = _val(pr, "pos", "bot")
        e = self.arg(el, "e")
        if pos == "top":
            return Box("bar %s" % e.group(), e.w, e.a + 0.2, e.d)
        return Box("under %s" % e.group(), e.w, e.a, e.d + 0.2)

    def m_groupChr(self, el):
        pr = _child(el, "groupChrPr")
        pos = _val(pr, "pos", "bot")
        ch = _val(pr, "chr", "⏟")
        e = self.arg(el, "e")
        if ch in ("←", "→", "↔", "⇒", "⇐", "⇔"):
            # 화살표 위/아래 글자는 상위 limUpp/limLow 가 처리 → 기호만
            return Box(SYMBOLS.get(ch, ch) + " " + e.s, e.w + 1, e.a, e.d)
        if pos == "top":
            return Box("bar %s" % e.group(), e.w, e.a + 0.3, e.d)
        return Box("under %s" % e.group(), e.w, e.a, e.d + 0.3)

    def m_box(self, el):
        return self.arg(el, "e")

    def m_borderBox(self, el):
        return self.arg(el, "e")

    def m_phant(self, el):
        pr = _child(el, "phantPr")
        if pr is not None and not _on(pr, "show", True):
            e = self.arg(el, "e")
            return Box("~", e.w, e.a, e.d)
        return self.arg(el, "e")

    def m_e(self, el):
        return self.seq(el)

    def m_ctrlPr(self, el):
        return None

    def m_argPr(self, el):
        return None


# 낱말로 쓰면 명령어로 읽히는 것들 (이탤릭 변수 이름과 겹칠 때 글자 사이를 띄움)
RESERVED = {k.lower() for k in list(GREEK.values()) + list(FUNCS) +
            [w for v in list(SYMBOLS.values()) + list(NARY.values()) for w in v.split()] + [
    "over", "atop", "sqrt", "root", "of", "sum", "int", "prod", "lim", "from", "to",
    "left", "right", "matrix", "pmatrix", "bmatrix", "dmatrix", "cases", "pile",
    "lpile", "rpile", "eqalign", "rm", "it", "bold", "hat", "bar", "vec", "dot",
    "ddot", "tilde", "under", "acute", "grave", "check", "arch", "dyad", "times",
    "cdot", "inf", "in", "and", "or", "if", "for", "not", "le", "ge", "ne",
    "leq", "geq", "neq", "sub", "sup", "lsub", "lsup", "union", "inter", "cap",
    "cup", "div", "deg", "prime", "color", "scale", "binom", "choose", "big",
    "bigg", "partial", "exist", "forall", "angle", "top", "bot", "circ", "ell",
    "image", "base", "well", "att", "hund", "thou", "vert", "owns", "notin",
    "subset", "supset", "sim", "approx", "equiv", "dint", "tint", "oint",
    "larrow", "rarrow", "uparrow", "downarrow", "col", "lcol", "rcol"]}

_OPS = set("+-=<>!,;:/()[]|^_'*.")
_PAIR = set("+-=<>!")
_ALNUM = re.compile(r"^[A-Za-z0-9.]+$")


_REL = {"=", "<", ">", "<=", ">=", "!=", "==", "->", "<-", "<->", "<<", ">>"}


def _is_wordtok(t):
    return bool(t) and (t[0].isalnum() or t[0] == '"' or t[0] == "_" and len(t) > 1
                        or ord(t[0]) > 0x7f)


def tidy(script):
    """한글 수식 편집기가 저장할 때와 같은 모양으로 공백을 정리한다.
    예) sum _{k=1} ^{n} k ^{2} = {n LEFT ( n+1 RIGHT ) LEFT ( 2n+1 RIGHT )} over {6}
      - 관계 기호(=, <, <= …)와 명령어 앞뒤는 띄움
      - 첨자 기호 _ ^ 앞은 띄우고 뒤는 붙임
      - LEFT/RIGHT 와 그 괄호는 띄움
      - +, - 같은 연산자와 변수는 붙임 (n+1, 2n+1)"""
    toks = re.findall(r'"[^"]*"|[{}]|[^\s{}"]+', script)
    if not toks:
        return ""
    # 4 a c → 4ac, d x → dx (명령어와 겹치지 않을 때만)
    merged = []
    for tok in toks:
        if (merged and _ALNUM.match(tok) and _ALNUM.match(merged[-1])
                and merged[-1].lower() not in RESERVED and tok.lower() not in RESERVED
                and (merged[-1] + tok).lower() not in RESERVED
                and not (merged[-1][-1].isdigit() and tok[0].isdigit())):
            merged[-1] += tok
        else:
            merged.append(tok)
    toks = merged
    out = toks[0]
    depth = 1 if toks[0] == "{" else 0
    for k in range(1, len(toks)):
        a, b = toks[k - 1], toks[k]
        prev2 = toks[k - 2] if k >= 2 else ""
        if b in ("_", "^"):
            sp = True
        elif a in ("_", "^"):
            sp = False
        elif a == "{" or b == "}":
            sp = False
        elif a in ("LEFT", "RIGHT") or b in ("LEFT", "RIGHT"):
            sp = True
        elif prev2 in ("LEFT", "RIGHT"):     # LEFT ( n , RIGHT ) LEFT
            sp = True
        elif a in ("#", "&") or b in ("#", "&"):
            sp = True
        elif a in _REL or b in _REL:
            sp = depth == 0                  # 맨 바깥의 = 만 띄움 (첨자 안 k=1 은 붙임)
        elif a[-1] in _PAIR and b[0] in _PAIR:
            sp = True                     # <- , +- 처럼 붙어서 다른 기호가 되는 것 방지
        elif _is_wordtok(a) and _is_wordtok(b):
            sp = True
        elif (_is_wordtok(a) and b == "{") or (a == "}" and (_is_wordtok(b) or b == "{")):
            sp = True
        else:
            sp = False
        if b == "{":
            depth += 1
        elif b == "}":
            depth -= 1
        out += (" " if sp else "") + b
    return out


def _is_word(s):
    s = s.strip()
    return bool(re.match(r"^[A-Za-z]+$", s))


def omml_to_hwp(omath):
    """편의 함수: <m:oMath> → (script, w_em, a_em, d_em)"""
    return OmmlConverter().convert(omath)


# ===========================================================================
# 수식 교정 학습 (사용자가 고친 수식에서 규칙을 배워 다음 변환에 적용)
# ===========================================================================
"""
수식 교정 학습

사용자가 "변환기가 만든 수식"과 "한글에서 고친 수식"을 한 쌍 주면, 둘의 차이를
규칙으로 바꿔 두었다가 다음 변환부터 자동으로 적용한다.

규칙 종류
  exact : 수식 전체가 똑같으면 고친 수식을 그대로 쓴다 (띄어쓰기까지 그대로)
  token : 바뀐 부분만 바꾼다
          - 기호·명령어만 바뀐 경우 (예: -> → rarrow): 모든 수식에 적용
          - 변수·숫자가 섞인 경우: 앞뒤 한 토큰이 같을 때만 적용 (엉뚱한 곳 방지)
  space : 토큰은 같고 띄어쓰기만 다르면, 그 띄어쓰기 습관을 기억한다
          (예: "a+b" 를 "a + b" 로 고치면 + 앞뒤를 띄움)

규칙은 JSON 목록으로 주고받는다. 표준 라이브러리만 사용.
"""
import difflib  # noqa: E402
import json  # noqa: E402

START, END = "<시작>", "<끝>"
_TOK = re.compile(r'"[^"]*"|[{}]|[A-Za-z]+|[0-9.]+|[+\-=<>!]+|\S')


def tokens(script):
    return _TOK.findall(script or "")


def _norm(script):
    return " ".join(tokens(script))


# 한컴 수식 문서의 명령어·기호 이름 (사용자가 고칠 때 쓸 만한 것들)
KEYWORDS = {k.lower() for k in """
PLUSMINUS MINUSPLUS times DIV DIVIDE CIRC BULLET DEG AST STAR BIGCIRC EMPTYSET THEREFORE
BECAUSE IDENTICAL EXIST neq DOTEQ image REIMAGE SIM APPROX SIMEQ CONG EQUIV ASYMP ISO
DIAMOND DSUM FORALL prime PARTIAL inf LNOT PROPTO XOR TRIANGLED DAGGER DDAGGER larrow
rarrow uparrow downarrow LARROW RARROW UPARROW DOWNARROW udarrow lrarrow UDARROW LRARROW
nwarrow searrow nearrow swarrow hookleft hookright mapsto vert VERT cdots LDOTS VDOTS DDOTS
TRIANGLE ANGLE MSANGLE SANGLE RTANGLE VDASH HLEFT BOT TOP MODELS LAPLACE CENTIGRADE
FAHRENHEIT LSLANT RSLANT att hund thou well base benzene INTER CAP SQCAP SQCUP OPLUS
OMINUS OTIMES ODOT OSLASH VEE WEDGE SUBSET SUPSET SUBSETEQ SUPSETEQ IN OWNS notin LEQ GEQ
SQSUBSET SQSUPSET SQSUBSETEQ SQSUPSETEQ LLL PREC SUCC UPLUS PROD COPROD UNION SMALLUNION
SMALLINTER SUM INT OINT DINT TINT ODINT OTINT ALEPH HBAR IMATH JMATH OHM ELL LITER WP IMAG
ANGSTROM vartheta varpi varsigma varupsilon varphi varepsilon OVER ATOP SQRT ROOT OF SUP SUB
LSUB LSUP BIGG CASES PILE LPILE RPILE EQALIGN CHOOSE BINOM MATRIX PMATRIX BMATRIX DMATRIX
HAT CHECK TILDE ACUTE GRAVE DOT DDOT BAR VEC DYAD UNDER ARCH LEFT RIGHT FROM TO REL BUILDREL
NOT SCALE COLOR rm it bold cdot le ge ne leq geq
""".split()} | RESERVED


def _is_var(tok):
    """변수 이름·숫자 (규칙을 전체에 적용하면 위험한 토큰)"""
    if tok[0].isdigit():
        return True
    return tok.isalpha() and tok.lower() not in KEYWORDS


def _gaps(script):
    """토큰 사이 공백 여부 목록 (토큰 i 와 i+1 사이)"""
    toks = list(_TOK.finditer(script))
    return [toks[i + 1].start() > toks[i].end() for i in range(len(toks) - 1)]


def learn(before, after):
    """(변환 결과, 한글에서 고친 결과) → 규칙 목록"""
    before = (before or "").strip()
    after = (after or "").strip()
    if not before or not after or before == after:
        return []
    rules = [{"type": "exact", "from": _norm(before), "to": after, "show": before}]
    ta, tb = tokens(before), tokens(after)
    if ta == tb:
        # 띄어쓰기만 다름 → 어떤 토큰 앞뒤를 띄우는지/붙이는지 배운다
        ga, gb = _gaps(before), _gaps(after)
        seen = set()
        for i, (x, y) in enumerate(zip(ga, gb)):
            if x != y:
                key = (_cls(ta[i]), _cls(ta[i + 1]))
                # 연산자 쪽만 기억 (a + b 를 배우면 x + 1 에도 적용)
                if _is_op(key[1]) and not _is_op(key[0]) and key[0] != "{":
                    key = ("*", key[1])
                elif _is_op(key[0]) and not _is_op(key[1]) and key[1] != "}":
                    key = (key[0], "*")
                if key in seen:
                    continue
                seen.add(key)
                rules.append({"type": "space", "left": key[0], "right": key[1], "space": y})
        return rules
    sm = difflib.SequenceMatcher(a=ta, b=tb, autojunk=False)
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == "equal":
            continue
        frm, to = ta[i1:i2], tb[j1:j2]
        general = tag == "replace" and frm and not any(_is_var(t) for t in frm) \
            and not any(_is_var(t) for t in to)
        if general:
            rules.append({"type": "token", "from": frm, "to": to})
        else:
            left = ta[i1 - 1] if i1 > 0 else START
            right = ta[i2] if i2 < len(ta) else END
            rules.append({"type": "token", "from": frm, "to": to,
                          "left": left, "right": right})
    return rules


def _is_op(c):
    return bool(re.match(r"^[+\-=<>!,/|*]+$", c or ""))


def _cls(tok):
    """띄어쓰기 규칙에 쓰는 토큰 분류: 기호는 그대로, 나머지는 종류로"""
    if tok in ("{", "}"):
        return tok
    if re.match(r"^[+\-=<>!]+$", tok) or len(tok) == 1 and not tok.isalnum():
        return tok
    if tok.startswith('"'):
        return "TEXT"
    if tok[0].isdigit():
        return "NUM"
    if len(tok) > 1 and tok.lower() in KEYWORDS:
        return tok          # 명령어는 이름 그대로 (over, LEFT …)
    return "VAR"


def apply(script, rules):
    """규칙을 수식 하나에 적용"""
    if not rules:
        return script
    key = _norm(script)
    for r in rules:
        if r.get("type") == "exact" and r.get("from") == key:
            return r["to"]
    toks = tokens(script)
    changed = False
    for r in rules:
        if r.get("type") != "token":
            continue
        frm, to = r.get("from") or [], r.get("to") or []
        left, right = r.get("left"), r.get("right")
        out = []
        n = len(toks)
        m = len(frm)
        if m:
            i = 0
            while i < n:
                prev = out[-1] if out else START
                nxt = toks[i + m] if i + m < n else END
                if toks[i:i + m] == frm and (left is None or prev == left) \
                        and (right is None or nxt == right):
                    out.extend(to)
                    i += m
                    changed = True
                    continue
                out.append(toks[i])
                i += 1
        else:
            for i in range(n + 1):
                prev = toks[i - 1] if i > 0 else START
                nxt = toks[i] if i < n else END
                if prev == left and nxt == right:
                    out.extend(to)
                    changed = True
                if i < n:
                    out.append(toks[i])
        toks = out
    result = tidy(" ".join(toks)) if changed else script
    spaces = [r for r in rules if r.get("type") == "space"]
    if spaces:
        result = _apply_spaces(result, spaces)
    return result


def _apply_spaces(script, rules):
    want = {(r["left"], r["right"]): r["space"] for r in rules}
    toks = tokens(script)
    if not toks:
        return script
    gaps = _gaps(script)
    out = toks[0]
    for i in range(1, len(toks)):
        sp = gaps[i - 1]
        a, b = _cls(toks[i - 1]), _cls(toks[i])
        for k in ((a, b), ("*", b), (a, "*")):
            if k in want and not (k[0] == "*" and a == "{") and not (k[1] == "*" and b == "}"):
                sp = want[k]
                break
        out += (" " if sp else "") + toks[i]
    return out


def describe(rule):
    """사람이 읽을 수 있는 규칙 설명"""
    t = rule.get("type")
    if t == "exact":
        return "수식 전체: %s  →  %s" % (rule.get("show", rule["from"]), rule["to"])
    if t == "space":
        return "띄어쓰기: '%s' 와 '%s' 사이를 %s" % (
            rule["left"], rule["right"], "띄움" if rule["space"] else "붙임")
    frm = " ".join(rule.get("from") or []) or "(없음)"
    to = " ".join(rule.get("to") or []) or "(지움)"
    ctx = ""
    if rule.get("left") is not None:
        ctx = "  (앞: %s, 뒤: %s 일 때만)" % (rule.get("left"), rule.get("right"))
    return "바꾸기: %s  →  %s%s" % (frm, to, ctx)


def merge(rules, new):
    """중복 없이 합치기 (새 규칙이 같은 대상을 덮어씀)"""
    def key(r):
        if r.get("type") == "exact":
            return ("exact", r["from"])
        if r.get("type") == "space":
            return ("space", r["left"], r["right"])
        return ("token", tuple(r.get("from") or []), r.get("left"), r.get("right"))
    d = {key(r): r for r in rules}
    for r in new:
        d[key(r)] = r
    return list(d.values())


def load(path):
    with open(path, encoding="utf-8") as f:
        data = json.load(f)
    return data.get("rules", data) if isinstance(data, dict) else data
