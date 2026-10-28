# docx2hwp — Word(.docx) → 한글(.hwpx) 변환기

**웹에서 바로 쓰기: https://seouk-moon.github.io/docx2hwp/**

Word 문서를 한글 문서로 바꿉니다. Word 수식은 그림이 아니라 **한글에서 바로 고칠 수 있는 한글 수식**으로 바뀝니다.

- 설치할 것 없음: 브라우저에서 `.docx`를 끌어다 놓으면 `.hwpx`가 나옵니다.
- 파일이 서버로 올라가지 않습니다. 변환은 브라우저 안에서 파이썬(Pyodide)으로 이뤄집니다.
- 결과 `.hwpx`는 한글 2014 이상, 한글 뷰어, 한컴독스에서 열립니다. `.hwp`가 필요하면 한글에서 열어 다른 이름으로 저장하세요.

## 옮겨지는 것

- 수식: 분수, 근호, 첨자, 합·적분·극한, 행렬, cases, 연립 정렬, 악센트, 그리스 문자, 각종 기호 (한컴 수식 명령어 표기)
- 글자·문단 모양, 번호·글머리표 목록, 표(셀 병합·테두리·배경), 그림, 머리말·꼬리말·쪽 번호, 각주·미주, 용지 방향·여백·다단

## 내 컴퓨터에서 실행 (파이썬)

```
python docx2hwp.py 문서.docx                 # → 문서.hwpx
python docx2hwp.py 폴더/ -d 출력폴더           # 여러 개 한꺼번에
python docx2hwp.py                           # 파일 선택 창
python docx2hwp.py 문서.docx -f hwp          # .hwp (Windows + 한글 + pywin32)
```

Windows에서는 `.docx`를 `convert.bat` 위로 끌어다 놓아도 됩니다.

## 폴더 구성

| 경로 | 내용 |
|---|---|
| `docx2hwp.py`, `omml2hwpeq.py` | 변환기 본체 (표준 라이브러리만 사용) |
| `web/` | GitHub Pages 웹 앱 (`index.html`, `style.css`, `app.js`, `py/web_entry.py`) |
| `.github/workflows/pages.yml` | main에 올리면 테스트 후 Pages에 자동 배포 |
| `tests/` | 시험 문서와 검사 스크립트 (`python tests/run_tests.py`) |

---

본 제품은 한글과컴퓨터의 한글 문서 파일(.hwp) 공개 문서를 참고하여 개발하였습니다.
