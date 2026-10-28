# -*- coding: utf-8 -*-
"""브라우저(Pyodide)에서 부르는 진입점. 결과는 JSON 문자열로 돌려준다."""
import json
import os
import zipfile

import docx2hwp


def convert(src_path, out_path):
    conv = docx2hwp.Converter(src_path)
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


def make_zip(pairs_json, out_path):
    """[[파일경로, 압축 안 이름], …] → zip"""
    pairs = json.loads(pairs_json)
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as z:
        for path, name in pairs:
            z.write(path, name)
    return out_path
