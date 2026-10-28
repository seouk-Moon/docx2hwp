/* docx2hwp 웹 앱: 브라우저 안의 파이썬(Pyodide)으로 docx → hwpx 변환 */
"use strict";

const $ = (id) => document.getElementById(id);
const els = {
  engine: $("engine"), engineText: $("engineText"), drop: $("drop"), picker: $("picker"),
  files: $("files"), empty: $("emptyRow"), count: $("count"), zipBtn: $("zipBtn"),
  clearBtn: $("clearBtn"), eqs: $("eqs"), eqFile: $("eqFile"), toast: $("toast"),
  lBefore: $("lBefore"), lAfter: $("lAfter"), learnBtn: $("learnBtn"), learnMsg: $("learnMsg"),
  rules: $("rules"), ruleCount: $("ruleCount"), reconvBtn: $("reconvBtn"),
  exportBtn: $("exportBtn"), importInput: $("importInput"),
};
const RULES_KEY = "docx2hwp.rules.v1";
let rules = loadRules();

const PY_FILES = ["docx2hwp.py", "omml2hwpeq.py", "web_entry.py"];
const WORK = "/home/pyodide/work";
let py = null;          // pyodide
let entry = null;       // web_entry 모듈
let queue = Promise.resolve();
let seq = 0;
const results = [];     // {id, name, outName, outPath, blob, info, li}
let selectedId = null;

/* ---------- 엔진 준비 ---------- */
async function boot() {
  try {
    py = await loadPyodide();
    py.FS.mkdirTree("/home/pyodide/app");
    py.FS.mkdirTree(WORK);
    for (const f of PY_FILES) {
      const res = await fetch("py/" + f, { cache: "no-cache" });
      if (!res.ok) throw new Error(f + " 을(를) 불러오지 못했습니다 (" + res.status + ")");
      py.FS.writeFile("/home/pyodide/app/" + f, await res.text());
    }
    py.runPython("import sys; sys.path.insert(0, '/home/pyodide/app')");
    entry = py.pyimport("web_entry");
    setEngine("ready", "준비됨 · 파일은 이 컴퓨터 밖으로 나가지 않습니다");
    renderRules();
  } catch (e) {
    console.error(e);
    setEngine("error", "엔진을 불러오지 못했습니다 — 새로고침해 보세요");
  }
}

function setEngine(state, text) {
  els.engine.dataset.state = state;
  els.engineText.textContent = text;
}

/* ---------- 파일 받기 ---------- */
function addFiles(fileList) {
  const files = [...fileList];
  const docx = files.filter((f) => /\.docx$/i.test(f.name));
  const skipped = files.length - docx.length;
  if (skipped) toast(`.docx 가 아닌 파일 ${skipped}개는 건너뛰었습니다`);
  for (const f of docx) {
    const item = createRow(f);
    queue = queue.then(() => convertOne(f, item));
  }
}

function createRow(file) {
  const id = ++seq;
  const outName = file.name.replace(/\.docx$/i, ".hwpx");
  const li = document.createElement("li");
  li.className = "file";
  li.innerHTML = `
    <div class="name"></div>
    <div class="meta"><span class="pill run">대기 중</span><span class="size"></span></div>
    <div class="act"><button class="btn small primary" disabled>받기</button></div>`;
  li.querySelector(".name").textContent = file.name;
  li.querySelector(".size").textContent = fmtSize(file.size);
  els.empty.hidden = true;
  els.files.prepend(li);
  const item = { id, name: file.name, outName, li, info: null, blob: null, outPath: null, file };
  results.push(item);
  li.addEventListener("click", (ev) => {
    if (ev.target.closest("button")) return;
    select(item);
  });
  li.querySelector("button").addEventListener("click", () => download(item.blob, item.outName));
  refreshCounts();
  return item;
}

async function waitEngine() {
  while (!entry) {
    if (els.engine.dataset.state === "error") throw new Error("변환 엔진을 쓸 수 없습니다");
    await new Promise((r) => setTimeout(r, 200));
  }
}

async function convertOne(file, item) {
  const pill = item.li.querySelector(".pill");
  try {
    setPill(pill, "run", "변환 중…");
    await waitEngine();
    const buf = new Uint8Array(await file.arrayBuffer());
    const src = `${WORK}/in_${item.id}.docx`;
    const out = `${WORK}/out_${item.id}.hwpx`;
    py.FS.writeFile(src, buf);
    await new Promise((r) => setTimeout(r, 20));      // 화면 갱신할 틈
    const info = JSON.parse(entry.convert(src, out, JSON.stringify(rules)));
    py.FS.unlink(src);
    const bytes = py.FS.readFile(out);
    item.outPath = out;
    item.info = info;
    item.blob = new Blob([bytes], { type: "application/hwp+zip" });

    const meta = item.li.querySelector(".meta");
    setPill(pill, info.warnings.length ? "warn" : "ok", info.warnings.length ? "완료 · 참고 있음" : "완료");
    meta.querySelector(".size").textContent = `${fmtSize(file.size)} → ${fmtSize(bytes.length)}`;
    addStat(meta, `수식 ${info.equations.length}`);
    if (info.images) addStat(meta, `그림 ${info.images}`);
    if (info.footnotes) addStat(meta, `각주 ${info.footnotes}`);
    if (info.warnings.length) {
      const ul = document.createElement("ul");
      ul.className = "warns";
      for (const w of info.warnings) {
        const wli = document.createElement("li");
        wli.textContent = w;
        ul.appendChild(wli);
      }
      item.li.appendChild(ul);
    }
    item.li.querySelector("button").disabled = false;
    if (!selectedId || results.filter((r) => r.info).length === 1) select(item);
  } catch (e) {
    console.error(e);
    setPill(pill, "err", "실패");
    const msg = String(e && e.message ? e.message : e).split("\n").filter(Boolean).pop();
    addStat(item.li.querySelector(".meta"), msg.slice(0, 120));
  }
  refreshCounts();
}

function addStat(meta, text) {
  const s = document.createElement("span");
  s.textContent = "· " + text;
  meta.appendChild(s);
}
function setPill(p, kind, text) { p.className = "pill " + kind; p.textContent = text; }

/* ---------- 수식 미리보기 ---------- */
function select(item) {
  selectedId = item.id;
  for (const r of results) r.li.classList.toggle("selected", r.id === item.id);
  els.eqs.innerHTML = "";
  els.eqFile.textContent = item.name;
  if (!item.info) {
    els.eqs.innerHTML = `<li class="none">아직 변환되지 않았습니다.</li>`;
    return;
  }
  const eqs = item.info.equations;
  if (!eqs.length) {
    els.eqs.innerHTML = `<li class="none">이 문서에는 수식이 없습니다.</li>`;
    return;
  }
  eqs.forEach((s, i) => {
    const li = document.createElement("li");
    li.innerHTML = `<span class="no">${i + 1}</span><code></code>
      <span class="btns"><button class="btn small copy">복사</button><button class="btn small copy fix">고치기</button></span>`;
    li.querySelector("code").textContent = s;
    li.querySelector(".copy").addEventListener("click", async () => {
      try { await navigator.clipboard.writeText(s); toast("수식을 복사했습니다"); }
      catch { toast("복사하지 못했습니다"); }
    });
    li.querySelector(".fix").addEventListener("click", () => {
      els.lBefore.value = s;
      els.lAfter.value = s;
      els.learnMsg.textContent = "아래 칸을 한글에서 고친 모양으로 바꾼 뒤 '배우기'를 누르세요";
      els.lAfter.focus();
      els.lAfter.scrollIntoView({ behavior: "smooth", block: "center" });
    });
    els.eqs.appendChild(li);
  });
}

/* ---------- 받기 ---------- */
function download(blob, name) {
  if (!blob) return;
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = name;
  document.body.appendChild(a);
  a.click();
  setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
}

function zipAll() {
  const done = results.filter((r) => r.blob);
  if (!done.length) return;
  if (done.length === 1) return download(done[0].blob, done[0].outName);
  const used = new Set();
  const pairs = done.map((r) => {
    let n = r.outName, k = 2;
    while (used.has(n)) n = r.outName.replace(/\.hwpx$/i, ` (${k++}).hwpx`);
    used.add(n);
    return [r.outPath, n];
  });
  const out = `${WORK}/all.zip`;
  entry.make_zip(JSON.stringify(pairs), out);
  download(new Blob([py.FS.readFile(out)], { type: "application/zip" }), "hwpx-변환결과.zip");
}

function clearAll() {
  for (const r of results) {
    r.li.remove();
    if (r.outPath) { try { py.FS.unlink(r.outPath); } catch { /* 없음 */ } }
  }
  results.length = 0;
  selectedId = null;
  els.eqs.innerHTML = "";
  els.eqFile.textContent = "파일을 고르면 한글 수식 스크립트가 여기 나옵니다";
  els.empty.hidden = false;
  refreshCounts();
}

/* ---------- 수식 교정 학습 ---------- */
function loadRules() {
  try { return JSON.parse(localStorage.getItem(RULES_KEY) || "[]"); } catch { return []; }
}
function saveRules() {
  try { localStorage.setItem(RULES_KEY, JSON.stringify(rules)); } catch { /* 저장 불가 환경 */ }
}
function renderRules() {
  els.rules.innerHTML = "";
  els.ruleCount.textContent = rules.length ? `${rules.length}개` : "";
  if (!rules.length) {
    els.rules.innerHTML = `<li class="none">아직 없습니다. 수식 미리보기에서 '고치기'를 눌러 시작하세요.</li>`;
    return;
  }
  const texts = entry ? JSON.parse(entry.describe(JSON.stringify(rules))) : rules.map((r) => r.type);
  texts.forEach((t, i) => {
    const li = document.createElement("li");
    li.innerHTML = `<span></span><button class="btn small del" title="이 규칙 지우기">✕</button>`;
    li.querySelector("span").textContent = t;
    li.querySelector("button").addEventListener("click", () => {
      rules.splice(i, 1);
      saveRules();
      renderRules();
      toast("규칙을 지웠습니다");
    });
    els.rules.appendChild(li);
  });
}
function learnOne() {
  const before = els.lBefore.value.trim();
  const after = els.lAfter.value.trim();
  if (!entry) return toast("변환 엔진이 준비되면 다시 눌러 주세요");
  if (!before || !after) return toast("두 칸을 모두 채워 주세요");
  if (before === after) return toast("두 수식이 같습니다. 아래 칸을 고친 모양으로 바꿔 주세요");
  const res = JSON.parse(entry.learn(before, after, JSON.stringify(rules)));
  rules = res.rules;
  saveRules();
  renderRules();
  els.learnMsg.textContent = `새 규칙 ${res.added.length}개를 배웠습니다. '다시 변환'을 누르면 바로 적용됩니다.`;
  els.lBefore.value = "";
  els.lAfter.value = "";
  refreshCounts();
}
function reconvertAll() {
  const files = results.map((r) => r.file).filter(Boolean).reverse();
  if (!files.length) return;
  clearAll();
  addFiles(files);
  toast("배운 규칙을 적용해 다시 변환합니다");
}
function exportRules() {
  const blob = new Blob([JSON.stringify({ app: "docx2hwp", rules }, null, 2)], { type: "application/json" });
  download(blob, "docx2hwp-수식규칙.json");
}
async function importRules(file) {
  try {
    const data = JSON.parse(await file.text());
    const incoming = Array.isArray(data) ? data : data.rules;
    if (!Array.isArray(incoming)) throw new Error("형식이 맞지 않습니다");
    rules = dedupe(rules.concat(incoming));
    saveRules();
    renderRules();
    toast(`규칙 ${incoming.length}개를 불러왔습니다`);
  } catch (e) {
    toast("불러오지 못했습니다: " + e.message);
  }
}
function dedupe(list) {
  const seen = new Map();
  for (const r of list) seen.set(JSON.stringify([r.type, r.from, r.left, r.right]), r);
  return [...seen.values()];
}

function refreshCounts() {
  const done = results.filter((r) => r.blob).length;
  els.count.textContent = results.length ? `${done} / ${results.length}` : "";
  els.zipBtn.disabled = done === 0;
  els.clearBtn.disabled = results.length === 0;
  els.reconvBtn.disabled = results.length === 0 || !rules.length;
}

/* ---------- 유틸 ---------- */
function fmtSize(n) {
  if (n < 1024) return n + " B";
  if (n < 1024 * 1024) return (n / 1024).toFixed(0) + " KB";
  return (n / 1024 / 1024).toFixed(1) + " MB";
}
let toastTimer = null;
function toast(msg) {
  els.toast.textContent = msg;
  els.toast.classList.add("show");
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => els.toast.classList.remove("show"), 1800);
}

/* ---------- 이벤트 ---------- */
els.picker.addEventListener("change", () => { addFiles(els.picker.files); els.picker.value = ""; });
els.drop.addEventListener("keydown", (e) => {
  if (e.key === "Enter" || e.key === " ") { e.preventDefault(); els.picker.click(); }
});
["dragenter", "dragover"].forEach((t) => els.drop.addEventListener(t, (e) => {
  e.preventDefault(); els.drop.classList.add("over");
}));
["dragleave", "drop"].forEach((t) => els.drop.addEventListener(t, (e) => {
  e.preventDefault(); els.drop.classList.remove("over");
}));
els.drop.addEventListener("drop", (e) => addFiles(e.dataTransfer.files));
// 창 아무 데나 떨어뜨려도 브라우저가 파일을 열어 버리지 않게
window.addEventListener("dragover", (e) => e.preventDefault());
window.addEventListener("drop", (e) => {
  e.preventDefault();
  if (!els.drop.contains(e.target)) addFiles(e.dataTransfer.files);
});
els.zipBtn.addEventListener("click", zipAll);
els.clearBtn.addEventListener("click", clearAll);
els.learnBtn.addEventListener("click", learnOne);
els.reconvBtn.addEventListener("click", reconvertAll);
els.exportBtn.addEventListener("click", exportRules);
els.importInput.addEventListener("change", () => {
  if (els.importInput.files[0]) importRules(els.importInput.files[0]);
  els.importInput.value = "";
});
renderRules();

boot();
