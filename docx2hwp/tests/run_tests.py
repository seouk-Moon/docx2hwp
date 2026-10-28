# -*- coding: utf-8 -*-
"""간단한 회귀 테스트: 시험 문서를 변환하고 HWPX 구조를 검사한다."""
import os
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(HERE))
sys.path.insert(0, HERE)

import docx2hwp  # noqa: E402
import check_hwpx  # noqa: E402

EXPECT = {"test.docx": 12, "wordmath.docx": 9}   # 들어 있어야 할 수식 개수

failed = 0
out_dir = tempfile.mkdtemp()
for name, n_eq in EXPECT.items():
    src = os.path.join(HERE, name)
    out = os.path.join(out_dir, name.replace(".docx", ".hwpx"))
    conv = docx2hwp.Converter(src)
    conv.build(out)
    errs, stats = check_hwpx.check(out)
    ok = not errs and stats["equation"] == n_eq
    print("%s %s  %s" % ("OK  " if ok else "FAIL", name, stats))
    for e in errs:
        print("   -", e)
    failed += not ok
sys.exit(1 if failed else 0)
