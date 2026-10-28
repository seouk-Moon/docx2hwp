# -*- coding: utf-8 -*-
"""브라우저(Pyodide)에서 부르는 진입점. 결과는 JSON 문자열로 돌려준다."""
import json
import os
import zipfile

import docx2hwp
import omml2hwpeq as eqlearn


def convert(src_path, out_path, rules_json="[]"):
    rules = json.loads(rules_json or "[]")
    conv = docx2hwp.Converter(src_path, eq_rules=rules)
    conv.build(out_path)
    info = {
        "ok": True,
        "equations": conv.eq_scripts,
        "images": len(conv.bins),
        "footnotes": conv.footnote_no,
        "warnings": conv.warnings,
        "size": os.path.getsize(out_path),
    }
    return json.dumps(info, ensure_ascii=False)


def learn(before, after, rules_json="[]"):
    """고친 수식 한 쌍을 배워 전체 규칙 목록과, 새로 생긴 규칙 설명을 돌려준다"""
    old = json.loads(rules_json or "[]")
    new = eqlearn.learn(before, after)
    merged = eqlearn.merge(old, new)
    return json.dumps({"rules": merged, "added": [eqlearn.describe(r) for r in new]},
                      ensure_ascii=False)


def describe(rules_json):
    return json.dumps([eqlearn.describe(r) for r in json.loads(rules_json or "[]")],
                      ensure_ascii=False)


def preview(script, rules_json):
    return eqlearn.apply(script, json.loads(rules_json or "[]"))


def make_zip(pairs_json, out_path):
    """[[파일경로, 압축 안 이름], …] → zip"""
    pairs = json.loads(pairs_json)
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as z:
        for path, name in pairs:
            z.write(path, name)
    return out_path
