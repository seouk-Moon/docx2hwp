# -*- coding: utf-8 -*-
"""HWPX 구조 검사: 패키지 규칙, XML 형식, 모든 ID 참조가 실제로 있는지."""
import re
import sys
import zipfile
import xml.etree.ElementTree as ET

HH = "{http://www.hancom.co.kr/hwpml/2011/head}"
HP = "{http://www.hancom.co.kr/hwpml/2011/paragraph}"
HC = "{http://www.hancom.co.kr/hwpml/2011/core}"


def check(path):
    errs = []
    z = zipfile.ZipFile(path)
    infos = z.infolist()
    if infos[0].filename != "mimetype" or infos[0].compress_type != zipfile.ZIP_STORED:
        errs.append("mimetype 이 첫 항목(무압축)이 아님")
    if z.read("mimetype") != b"application/hwp+zip":
        errs.append("mimetype 내용 이상")
    trees = {}
    for i in infos:
        if i.filename.endswith((".xml", ".hpf")):
            try:
                trees[i.filename] = ET.fromstring(z.read(i.filename))
            except ET.ParseError as e:
                errs.append("XML 오류 %s: %s" % (i.filename, e))
    head = trees["Contents/header.xml"]
    ids = {}
    for tag in ("charPr", "paraPr", "borderFill", "style", "tabPr", "numbering"):
        els = list(head.iter(HH + tag))
        ids[tag] = {e.get("id") for e in els}
        parent = {"charPr": "charProperties", "paraPr": "paraProperties",
                  "borderFill": "borderFills", "style": "styles",
                  "tabPr": "tabProperties", "numbering": "numberings"}[tag]
        p = head.find(".//" + HH + parent)
        if p is not None and int(p.get("itemCnt")) != len(els):
            errs.append("%s itemCnt=%s 실제=%d" % (parent, p.get("itemCnt"), len(els)))
    nfont = {ff.get("lang"): len(ff.findall(HH + "font")) for ff in head.iter(HH + "fontface")}
    for ff in head.iter(HH + "fontface"):
        if int(ff.get("fontCnt")) != len(ff.findall(HH + "font")):
            errs.append("fontCnt 불일치 %s" % ff.get("lang"))
    for cp in head.iter(HH + "charPr"):
        fr = cp.find(HH + "fontRef")
        for k, v in fr.attrib.items():
            if int(v) >= nfont[k.upper()]:
                errs.append("charPr %s fontRef %s=%s 범위 밖" % (cp.get("id"), k, v))
        if cp.get("borderFillIDRef") not in ids["borderFill"]:
            errs.append("charPr borderFill 참조 오류")
    for pp in head.iter(HH + "paraPr"):
        if pp.get("tabPrIDRef") not in ids["tabPr"]:
            errs.append("paraPr %s tabPr 참조 오류" % pp.get("id"))
        b = pp.find(HH + "border")
        if b is not None and b.get("borderFillIDRef") not in ids["borderFill"]:
            errs.append("paraPr %s borderFill 참조 오류" % pp.get("id"))
    for st in head.iter(HH + "style"):
        if st.get("paraPrIDRef") not in ids["paraPr"] or st.get("charPrIDRef") not in ids["charPr"]:
            errs.append("style %s 참조 오류" % st.get("name"))
    hpf = trees["Contents/content.hpf"]
    items = {i.get("id"): i.get("href") for i in hpf.iter("{http://www.idpf.org/2007/opf/}item")}
    names = set(z.namelist())
    for iid, href in items.items():
        if href not in names:
            errs.append("manifest 항목 없음: %s" % href)
    stats = {"p": 0, "tbl": 0, "pic": 0, "equation": 0, "footNote": 0, "header": 0, "footer": 0,
             "autoNum": 0}
    for name, t in trees.items():
        if not re.match(r"Contents/section\d+\.xml", name):
            continue
        for el in t.iter():
            tag = el.tag.split("}")[1]
            if tag in stats:
                stats[tag] += 1
            for a, refset in (("charPrIDRef", "charPr"), ("paraPrIDRef", "paraPr"),
                              ("styleIDRef", "style"), ("borderFillIDRef", "borderFill")):
                v = el.get(a)
                if v is not None and v not in ids[refset]:
                    errs.append("%s: <%s %s=%s> 없는 ID" % (name, tag, a, v))
            if tag == "img":
                if el.get("binaryItemIDRef") not in items:
                    errs.append("그림 binaryItemIDRef 없음")
        # 각 구역 첫 문단에 secPr
        first = t.find(HP + "p")
        if first is None or first.find(".//" + HP + "secPr") is None:
            errs.append("%s 첫 문단에 secPr 없음" % name)
        # 표 셀 주소/병합 검사
        for tbl in t.iter(HP + "tbl"):
            rows, cols = int(tbl.get("rowCnt")), int(tbl.get("colCnt"))
            grid = [[0] * cols for _ in range(rows)]
            for tc in tbl.iter(HP + "tc"):
                if tc.find("..") is not None:
                    pass
            for tr in tbl.findall(HP + "tr"):
                for tc in tr.findall(HP + "tc"):
                    a = tc.find(HP + "cellAddr")
                    s = tc.find(HP + "cellSpan")
                    c0, r0 = int(a.get("colAddr")), int(a.get("rowAddr"))
                    for r in range(r0, r0 + int(s.get("rowSpan"))):
                        for c in range(c0, c0 + int(s.get("colSpan"))):
                            if r >= rows or c >= cols:
                                errs.append("표 셀이 범위 밖 (%d,%d)" % (r, c))
                            else:
                                grid[r][c] += 1
            holes = sum(1 for r in grid for v in r if v != 1)
            if holes:
                errs.append("표 칸 겹침/빈칸 %d개 (rowCnt=%d colCnt=%d)" % (holes, rows, cols))
    return errs, stats


if __name__ == "__main__":
    bad = 0
    for p in sys.argv[1:]:
        errs, stats = check(p)
        print("%s  %s" % ("OK " if not errs else "ERR", p))
        print("     ", " ".join("%s=%d" % kv for kv in stats.items()))
        for e in errs[:30]:
            print("   -", e)
        bad += bool(errs)
    sys.exit(1 if bad else 0)
