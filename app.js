/* global JSZip */

const state = {
  papers: [],
  metricMode: "published",
  latestYear: new Date().getFullYear(),
  selected: new Set(),
  datasets: { impact: null, cas: null, jcr: null },
  workflow: { parsed: false, pubmed: false, exported: false },
  sort: null,
  journalKeys: null,
  journalNameMap: null,
  journalKeysLoose: null,
  journalNameMapLoose: null,
  ui: {
    authorEditOpen: new Set(),
    authorDraft: new Map(),
    pubmedManualOpen: new Set(),
    pubmedManualValue: new Map(),
    sampleDocx: null, // { name: string, buffer: ArrayBuffer }
  }
};

const elements = {
  docxFile: document.getElementById("docxFile"),
  pickFileBtn: document.getElementById("pickFileBtn"),
  fileNameText: document.getElementById("fileNameText"),
  sampleBtn: document.getElementById("sampleBtn"),
  parseBtn: document.getElementById("parseBtn"),
  pubmedBtn: document.getElementById("pubmedBtn"),
  metricMode: document.getElementById("metricMode"),
  sortYearBtn: document.getElementById("sortYearBtn"),
  sortImpactBtn: document.getElementById("sortImpactBtn"),
  toggleAllBtn: document.getElementById("toggleAllBtn"),
  exportBtn: document.getElementById("exportBtn"),
  statusText: document.getElementById("statusText"),
  paperList: document.getElementById("paperList"),
  heroCountText: document.getElementById("heroCountText"),
  aiHint: document.getElementById("aiHint"),
  aiHintText: document.getElementById("aiHintText"),
  stepParse: document.getElementById("stepParse"),
  stepPubmed: document.getElementById("stepPubmed"),
  stepExport: document.getElementById("stepExport"),
  progressMeta: document.getElementById("progressMeta"),
};

function computeAiStats() {
  const total = state.papers.length || 0;
  const matched = state.papers.filter((p) => p.pubmed?.pmid).length;
  const rate = total ? Math.round((matched / total) * 100) : 0;
  return { total, matched, rate };
}

function updateWorkflowUi() {
  const { total, matched, rate } = computeAiStats();
  if (elements.heroCountText) elements.heroCountText.textContent = `已成功解析 ${total} 篇文献`;
  if (elements.aiHint && elements.aiHintText) {
    if (total) {
      elements.aiHint.style.display = "";
      elements.aiHintText.textContent = `AI 已自动识别并规范化 ${total} 篇文献（PubMed 匹配率 ${rate}%）。`;
    } else {
      elements.aiHint.style.display = "none";
    }
  }

  const setStep = (el, stateName) => {
    if (!el) return;
    el.classList.remove("active", "done");
    if (stateName === "active") el.classList.add("active");
    if (stateName === "done") el.classList.add("done");
  };
  setStep(elements.stepParse, state.workflow.parsed ? "done" : "active");
  setStep(elements.stepPubmed, state.workflow.pubmed ? "done" : state.workflow.parsed ? "active" : "");
  setStep(elements.stepExport, state.workflow.exported ? "done" : state.workflow.pubmed ? "active" : "");

  if (elements.progressMeta) {
    if (!state.workflow.parsed) elements.progressMeta.textContent = "请先上传文档并完成解析。";
    else if (!state.workflow.pubmed) elements.progressMeta.textContent = "已完成解析，可进行 PubMed 校验与补全。";
    else if (!state.workflow.exported) elements.progressMeta.textContent = `PubMed 校验完成（已匹配 ${matched}/${total}），可导出投稿格式。`;
    else elements.progressMeta.textContent = "已导出结果。";
  }
}

function updateStatus(message) {
  if (elements.statusText) {
    elements.statusText.style.display = "";
    elements.statusText.textContent = message;
  }
}

function setChosenFileName(name) {
  if (!elements.fileNameText) return;
  elements.fileNameText.textContent = name || "未选择任何文件";
}

function escapeHtml(value) {
  return String(value || "").replace(/[&<>"']/g, (ch) => {
    const map = { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" };
    return map[ch] || ch;
  });
}

function parsePmidFromInput(input) {
  const s = String(input || "").trim();
  if (!s) return "";
  const m1 = s.match(/pubmed\.ncbi\.nlm\.nih\.gov\/(\d{6,10})/i);
  if (m1) return m1[1];
  const m2 = s.match(/\b(\d{6,10})\b/);
  return m2 ? m2[1] : "";
}

function normalizeTitle(title) {
  return title.toLowerCase().replace(/\s+/g, "").replace(/[.,:;()'"“”’‘\[\]{}]/g, "");
}

function normalizeJournal(journal) {
  return journal.toLowerCase().replace(/\s+/g, "").replace(/&/g, "and").replace(/[.,:;()'"“”’‘\[\]{}\-]/g, "");
}

function normalizeJournalLoose(journal) {
  return normalizeJournal(journal).replace(/and/g, "");
}

function decodeXml(text) {
  return text.replace(/&amp;/g, "&").replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&apos;/g, "'");
}

async function decodeTextFromBuffer(buffer) {
  const tryDecode = (encoding) => {
    try { return new TextDecoder(encoding).decode(buffer); } catch { return null; }
  };
  return (tryDecode("gbk") || tryDecode("gb18030") || tryDecode("utf-8") || "").replace(/\r\n/g, "\n").replace(/\r/g, "\n");
}

const SAMPLE_DOCX_URL = "测试1.docx";

async function loadTextFromUrl(url) {
  const response = await fetch(url);
  if (!response.ok) throw new Error(`无法加载 ${url}`);
  const buffer = await response.arrayBuffer();
  return decodeTextFromBuffer(buffer);
}

function parseImpactTxt(text) {
  const lines = text.split("\n").map((line) => line.trim()).filter(Boolean);
  if (!lines.length) return null;
  const header = lines[0].split("\t").map((item) => item.trim());
  const yearColumns = header.filter((item) => /^\d{4}$/.test(item));
  const latestYear = yearColumns.length ? yearColumns.map((year) => Number(year)).sort((a, b) => b - a)[0].toString() : "";
  const yearIndexMap = {};
  header.forEach((item, idx) => { if (/^\d{4}$/.test(item)) yearIndexMap[item] = idx; });
  const map = new Map();
  lines.slice(1).forEach((line) => {
    const parts = line.split("\t");
    const name = parts[0]?.trim();
    if (!name) return;
    const record = { _name: name };
    yearColumns.forEach((year) => {
      const raw = parts[yearIndexMap[year]]?.trim();
      record[year] = raw && raw !== "#N/A" && raw !== "Not Available" ? raw : "";
    });
    map.set(normalizeJournal(name), record);
  });
  return { years: yearColumns, latestYear, map };
}

function parseZoneTxt(text) {
  const lines = text.split("\n").map((line) => line.trim()).filter(Boolean);
  if (!lines.length) return null;
  const header = lines[0].split("\t").map((item) => item.trim());
  const yearColumns = header.filter((item) => /^\d{4}/.test(item));
  const latestYear = yearColumns.length ? yearColumns.map((year) => Number(year.slice(0, 4))).sort((a, b) => b - a)[0].toString() : "";
  const yearIndexMap = {};
  header.forEach((item, idx) => { if (/^\d{4}/.test(item)) yearIndexMap[item.slice(0, 4)] = idx; });
  const topIndex = header.findIndex((item) => /top/i.test(item));
  const valueIndex = header.length >= 2 ? 1 : -1;
  const map = new Map();
  lines.slice(1).forEach((line) => {
    const parts = line.split("\t");
    const name = parts[0]?.trim();
    if (!name) return;
    const record = { _name: name };
    if (yearColumns.length) {
      yearColumns.forEach((year) => {
        const raw = parts[yearIndexMap[year]]?.trim();
        record[year] = raw ? raw : "";
      });
    } else if (valueIndex !== -1) {
      const raw = parts[valueIndex]?.trim();
      record.latest = raw ? raw : "";
    }
    if (topIndex !== -1) record.top = parts[topIndex]?.trim() || "";
    map.set(normalizeJournal(name), record);
  });
  return { years: yearColumns, latestYear, map };
}

async function extractDocxText(file) {
  const data = await file.arrayBuffer();
  return extractDocxTextFromArrayBuffer(data);
}

async function extractDocxTextFromArrayBuffer(data) {
  const zip = await JSZip.loadAsync(data);
  const xml = await zip.file("word/document.xml").async("text");
  const pMatches = xml.match(/<w:p[^>]*>.*?<\/w:p>/g) || [];
  const paragraphs = pMatches.map((pXml) => {
    const tokenRe = /<w:t[^>]*>(.*?)<\/w:t>|<w:br[^>]*\/>|<w:cr[^>]*\/>/g;
    let m; const parts = [];
    while ((m = tokenRe.exec(pXml))) {
      if (m[1] !== undefined) parts.push(decodeXml(m[1])); else parts.push("\n");
    }
    return parts.join("");
  });
  return paragraphs.join("\n")
    .replace(/\u00a0/g, " ")
    .replace(/\r\n/g, "\n")
    .replace(/\r/g, "\n")
    .replace(/[ \t]+\n/g, "\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

async function loadSampleDocx(options = {}) {
  const { silent = false } = options;
  try {
    if (!silent) updateStatus(`正在加载示例文档：${SAMPLE_DOCX_URL} ...`);

    const resp = await fetch(SAMPLE_DOCX_URL);
    if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
    const buf = await resp.arrayBuffer();

    // 仅“选择”示例，不解析；解析由“解析文档”按钮触发
    state.ui.sampleDocx = { name: SAMPLE_DOCX_URL, buffer: buf };
    if (elements.docxFile) elements.docxFile.value = "";
    setChosenFileName(SAMPLE_DOCX_URL);

    if (!silent) updateStatus(`已选择示例文档：${SAMPLE_DOCX_URL}（点击“解析文档”开始解析）`);
    return true;
  } catch (e) {
    console.warn("loadSampleDocx failed", e);
    if (!silent) {
      updateStatus("示例文档加载失败（可能是浏览器限制本地 file:// 读取）。建议用本地服务器打开本页，或手动选择文件。");
    }
    return false;
  }
}

function parseEntryMeta(metaText) {
  const typeMatch = metaText.match(/\[([^\]]+)\]/);
  const type = typeMatch ? typeMatch[1].trim() : "";
  let authorRole = "";
  if (metaText.includes("共同通讯作者")) authorRole = "共同通讯作者";
  else if (metaText.includes("通讯作者")) authorRole = "通讯作者";
  else if (metaText.includes("第一作者")) authorRole = "第一作者";
  const ifMatch = metaText.match(/IF\s*=?\s*([0-9.]+)/i);
  const casMatch = metaText.match(/中科院\s*([0-9]+)\s*区/);
  const jcrMatch = metaText.match(/JCR\s*([0-9]+)\s*区/i);
  const topMatch = metaText.match(/Top\s*期刊/i);
  return { type, authorRole, impactFactor: ifMatch ? ifMatch[1] : "", casZone: casMatch ? casMatch[1] : "", jcrZone: topMatch ? "Top" : jcrMatch ? jcrMatch[1] : "" };
}

function splitAuthors(text) {
  const TITLE_KEYWORDS = /(comment on|insights|protective|deciphering|integrated|multiomics|multivariate|meta-analysis|analysis|verification|characterization|promotes|pandemic|randomization|identifies|implications|prediction|therapy|treatment|prognosis|microenvironment)/i;
  const looksLikeAuthorList = (segment) => {
    const s = (segment || "").trim(); if (!s) return false;
    const commaCount = (s.match(/,/g) || []).length + (s.match(/，/g) || []).length;
    const hasEtAl = /\bet\s+al\b/i.test(s) || /等/.test(s);
    const hasSeparator = /;|；|、/.test(s);
    if (!(commaCount >= 2 || hasEtAl || hasSeparator)) return false;
    if (TITLE_KEYWORDS.test(s)) return false;
    return true;
  };
  const match = text.match(/^(.*?)\.\s+(.*)$/);
  if (match && looksLikeAuthorList(match[1])) return { authors: match[1].trim(), rest: match[2].trim() };
  return { authors: "", rest: text };
}

function splitTitleJournalNoYear(text) {
  const s = (text || "").trim(); if (!s) return { title: "", journal: "" };
  const looksLikeJournalTail = (tail) => {
    const t = (tail || "").trim(); if (!t || t.length > 80 || /\d/.test(t)) return false;
    const words = t.split(/\s+/).filter(Boolean);
    if (words.length === 1) {
      if (t.length < 4 || t.length > 20 || !/^[A-Za-z][A-Za-z0-9\-]*$/.test(t)) return false;
      return !/(study|analysis|trial|pandemic|disease)/i.test(t);
    }
    if (words.length > 8) return false;
    const okWord = (w) => /^(of|and|in|on|for|the|a|an|to|with|from)$/i.test(w) || /^[A-Z][A-Za-z0-9\-]*$/.test(w);
    return words.filter(okWord).length / words.length >= 0.7;
  };
  const parts = s.split(/\.\s+/).map((p) => p.trim()).filter(Boolean);
  if (parts.length >= 2) {
    const tail = parts[parts.length - 1]; const head = parts.slice(0, -1).join(". ").trim();
    if (looksLikeJournalTail(tail, head)) return { title: head, journal: tail };
  }
  return { title: s, journal: "" };
}

function stripTrailingJournal(title, journal) {
  const t = (title || "").trim(); const j = (journal || "").trim(); if (!t || !j) return t;
  const escapeRegExp = (str) => String(str).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const re = new RegExp(`[\\s·\\-–—,:;]*${escapeRegExp(j)}\\s*$`, "i");
  return t.replace(re, "").replace(/[.\s]+$/g, "").trim() || t;
}

function stripLeadingIndexAndSymbols(text) {
  return (text || "").trim().replace(/^\(?\s*\d+\s*[\.\)]\s*/g, "").replace(/^[\-–—•·]+\s*/g, "").trim();
}

function splitTitleJournal(rest) {
  const yearMatch = rest.match(/\b(19|20)\d{2}\b/);
  if (!yearMatch) return { title: rest, journal: "", year: "", volumeIssuePages: "" };
  const yearIndex = yearMatch.index; const beforeYear = rest.slice(0, yearIndex); const journalInfo = rest.slice(yearIndex).trim();
  let titlePart = ""; let journal = "";
  const lastPeriod = beforeYear.lastIndexOf(". ");
  if (lastPeriod !== -1) { titlePart = beforeYear.slice(0, lastPeriod).trim(); journal = beforeYear.slice(lastPeriod + 2).trim().replace(/,$/, ""); }
  else {
    const firstComma = beforeYear.indexOf(",");
    if (firstComma !== -1) { titlePart = beforeYear.slice(0, firstComma).trim(); journal = beforeYear.slice(firstComma + 1).trim().replace(/,$/, ""); }
    else { titlePart = beforeYear.trim().replace(/,$/, ""); }
  }
  let year = ""; let volumeIssuePages = "";
  const journalMatch = journalInfo.match(/^(\d{4})\s*,?\s*([^.]*)/);
  if (journalMatch) { year = journalMatch[1].trim(); volumeIssuePages = journalMatch[2].trim().replace(/\.$/, ""); } else { year = yearMatch[0]; }
  return { title: titlePart, journal, year, volumeIssuePages };
}

function isKnownJournal(name) {
  const key = normalizeJournal(name);
  return state.datasets.impact?.map.has(key) || state.datasets.cas?.map.has(key) || state.datasets.jcr?.map.has(key);
}

function findJournalByDataset(mainText) {
  const textLower = normalizeJournal(mainText);
  let bestMatch = ""; let maxLength = 0;
  if (state.journalKeys) {
    for (const key of state.journalKeys) { if (textLower.includes(key) && key.length > maxLength) { maxLength = key.length; bestMatch = state.journalNameMap.get(key); } }
  }
  if (!bestMatch && state.journalKeysLoose) {
    const textLoose = normalizeJournalLoose(mainText);
    for (const key of state.journalKeysLoose) { if (textLoose.includes(key) && key.length > maxLength) { maxLength = key.length; bestMatch = state.journalNameMapLoose.get(key); } }
  }
  return bestMatch;
}

function parseEntryText(entryText) {
  const normalized = entryText.replace(/（/g, "(").replace(/）/g, ")");
  const metaMatch = normalized.match(/\(([^()]*)\)\s*$/);
  const metaText = metaMatch ? metaMatch[1].trim() : "";
  const meta = parseEntryMeta(metaText);
  const isActuallyMeta = meta.impactFactor || meta.casZone || meta.jcrZone || meta.type || meta.authorRole || /if\s*[:=]\s*[0-9.]+/i.test(metaText);
  const mainText = metaMatch && isActuallyMeta ? normalized.slice(0, metaMatch.index).trim() : normalized.trim();
  const { authors, rest } = splitAuthors(mainText);
  let { title, journal: parsedJournal, year, volumeIssuePages } = splitTitleJournal(rest);
  if (!year) { const tj = splitTitleJournalNoYear(rest); title = tj.title; parsedJournal = tj.journal || parsedJournal; volumeIssuePages = ""; }
  const datasetJournal = findJournalByDataset(mainText);
  const journalCandidate = parsedJournal || datasetJournal || "";
  let finalTitle = stripLeadingIndexAndSymbols(title || rest || "(标题待补充)");
  if (journalCandidate) finalTitle = stripTrailingJournal(finalTitle, journalCandidate);
  if (datasetJournal && datasetJournal !== journalCandidate) finalTitle = stripTrailingJournal(finalTitle, datasetJournal);
  const finalJournal = !isKnownJournal(parsedJournal) && datasetJournal ? datasetJournal : parsedJournal || datasetJournal;
  return { id: crypto.randomUUID(), authors, title: finalTitle, journal: finalJournal || "(期刊待补充)", year: year || "", volumeIssuePages, paperType: meta.type || "", myAuthorRole: meta.authorRole || "", metrics: { published: { impactFactor: "", casZone: "", jcrZone: "" }, latest: { impactFactor: "", casZone: "", jcrZone: "" } }, pubmed: { pmid: "", coFirst: [], coCorresponding: [] }, rawText: mainText };
}

function getImpactFactor(journal, year) {
  const dataset = state.datasets.impact; if (!dataset) return "";
  const record = dataset.map.get(normalizeJournal(journal)); if (!record) return "";
  if (year && record[year]) return record[year];
  if (dataset.latestYear && record[dataset.latestYear]) return record[dataset.latestYear];
  return dataset.years[0] ? record[dataset.years[0]] || "" : "";
}

function getZoneValue(dataset, journal, year) {
  if (!dataset) return "";
  const record = dataset.map.get(normalizeJournal(journal)); if (!record) return "";
  if (year && record[year]) return record[year];
  if (record.latest) return record.latest;
  if (dataset.latestYear && record[dataset.latestYear]) return record[dataset.latestYear];
  return dataset.years[0] ? record[dataset.years[0]] || "" : "";
}

function getTopFlag(dataset, journal) {
  if (!dataset) return false;
  const record = dataset.map.get(normalizeJournal(journal)); if (!record?.top) return false;
  return /是/i.test(record.top) || /top/i.test(record.top);
}

function resolveMetrics(paper, mode) {
  const pubYear = paper.year ? String(paper.year) : ""; const latestYear = String(state.latestYear); const journal = paper.journal;
  const targetYear = mode === "latest" ? latestYear : pubYear;
  const impact = getImpactFactor(journal, targetYear);
  const cas = getZoneValue(state.datasets.cas, journal, targetYear);
  const jcr = getZoneValue(state.datasets.jcr, journal, targetYear);
  return { impactFactor: impact || "", casZone: cas || "", jcrZone: jcr || "", casTop: getTopFlag(state.datasets.cas, journal) };
}

function getSortedPapers() {
  if (!state.sort?.field) return [...state.papers];
  const direction = state.sort.direction === "asc" ? 1 : -1;
  return [...state.papers].sort((a, b) => {
    if (state.sort.field === "year") return ((Number(a.year) || 0) - (Number(b.year) || 0)) * direction;
    if (state.sort.field === "impact") {
      const av = Number(resolveMetrics(a, state.metricMode).impactFactor) || 0;
      const bv = Number(resolveMetrics(b, state.metricMode).impactFactor) || 0;
      return (av - bv) * direction;
    }
    return 0;
  });
}

function parseDocxEntries(text) {
  const lines = text.split("\n").map((l) => l.trim()).filter(Boolean);
  const rawEntries = [];
  lines.forEach((line) => {
    const maybeMulti = line.split(/\s+(?=\d+\s*[\.\)]\s+)/g).filter(Boolean);
    maybeMulti.forEach((chunk) => {
      const parts = chunk.split(/\)\s*(?=[A-Z\u4e00-\u9fa5])/).filter(Boolean);
      parts.forEach((p, idx) => {
        let entry = p.trim(); if (idx < parts.length - 1) entry += ")";
        rawEntries.push(entry);
      });
    });
  });
  return rawEntries.map(parseEntryText);
}

function setPapers(papers) {
  state.papers = papers;
  state.selected = new Set(papers.map((p) => p.id));
  state.workflow.parsed = papers.length > 0;
  state.workflow.pubmed = false;
  state.workflow.exported = false;
  updateStatus(`已解析 ${papers.length} 条论文。`);
  updateWorkflowUi();
}

function normalizeForMatch(name) {
  return (name || "").toLowerCase().replace(/&/g, "and").replace(/[^a-z0-9\u4e00-\u9fa5]+/g, "");
}

function formatAuthorsWithMarkers(authors, coFirst, coCorresponding) {
  const coFirstSet = new Set(coFirst.map(normalizeForMatch));
  const coCorrSet = new Set(coCorresponding.map(normalizeForMatch));
  return authors.map((name) => {
    const trimmedName = name.trim();
    const key = normalizeForMatch(trimmedName);
    const firstMark = coFirstSet.has(key) ? '<span class="mark">#</span>' : "";
    const corrMark = coCorrSet.has(key) ? '<span class="mark">*</span>' : "";
    // 使用 white-space: nowrap 确保名字与标记永远连在一起
    return `<span style="white-space: nowrap;">${trimmedName}${firstMark}${corrMark}</span>`;
  }).join(", ");
}

function render() {
  if (!state.papers.length) {
    elements.paperList.innerHTML = `<div class="empty"><i class="lucide-file-x-2"></i><p>请上传论文 docx 或从浏览器本地加载数据。</p></div>`;
    updateWorkflowUi(); return;
  }
  const sortedPapers = getSortedPapers();
  elements.paperList.innerHTML = sortedPapers.map((paper, index) => {
    const isEnriched = Boolean(paper.pubmed?.pmid || paper.pubmed?.doi);
    const isSelected = state.selected.has(paper.id);

    // 状态判断
    const matchStatus = paper.pubmed?.pmid
      ? (paper.pubmed?.fullAuthors?.length ? { cls: "success", text: "已匹配", icon: "lucide-check" } : { cls: "warn", text: "部分匹配", icon: "lucide-alert-triangle" })
      : (paper.pubmed?.error ? { cls: "fail", text: "未找到", icon: "lucide-x" } : { cls: "warn", text: "待校验", icon: "lucide-clock" });

    const pmidText = paper.pubmed?.pmid ? String(paper.pubmed.pmid) : "-";
    const metrics = resolveMetrics(paper, state.metricMode);

    // 逻辑控制：是否展示详细信息
    const showDetails = state.workflow.pubmed || isEnriched || paper.pubmed?.error;

    const canMarkAuthors = Boolean(paper.pubmed?.fullAuthors?.length);
    const isEditingAuthors = canMarkAuthors && state.ui.authorEditOpen.has(paper.id);

    let editorHtml = "";
    if (isEditingAuthors) {
      const draft = state.ui.authorDraft.get(paper.id);
      const rows = draft.baseAuthors.map((name, idx) => {
        const key = normalizeForMatch(name);
        const cf = draft.coFirstSet.has(key) ? "checked" : "";
        const cc = draft.coCorrSet.has(key) ? "checked" : "";
        return `<div class="author-editor-row"><div class="author-editor-name">${escapeHtml(name)}</div><label class="author-editor-flag"><input type="checkbox" data-action="authorFlag" data-id="${paper.id}" data-role="cofirst" data-idx="${idx}" ${cf} />共一</label><label class="author-editor-flag"><input type="checkbox" data-action="authorFlag" data-id="${paper.id}" data-role="cocorr" data-idx="${idx}" ${cc} />通讯</label></div>`;
      }).join("");
      editorHtml = `<div class="author-editor"><div class="author-editor-head">勾选需要标注的作者（蓝色 # / * 会显示在名单中）</div><div class="author-editor-grid">${rows}</div><div class="author-editor-actions"><button class="btn-mini" data-action="saveAuthorEdit" data-id="${paper.id}">保存</button><button class="btn-mini" data-action="cancelAuthorEdit" data-id="${paper.id}">取消</button></div></div>`;
    }

    const hasPubmedError = !isEnriched && paper.pubmed?.error;
    const manualOpen = !isEnriched && state.ui.pubmedManualOpen.has(paper.id);
    const manualValue = !isEnriched ? state.ui.pubmedManualValue.get(paper.id) || "" : "";
    const manualBlock = manualOpen ? `
      <div class="manual-pubmed">
        <div class="manual-pubmed-row">
          <input type="text" value="${escapeHtml(manualValue)}" data-action="manualPmidInput" data-id="${paper.id}" placeholder="粘贴 PubMed 链接或 PMID" />
          <button class="btn-mini" data-action="applyManualPmid" data-id="${paper.id}">应用</button>
          <button class="btn-mini" data-action="clipboardApplyManualPmid" data-id="${paper.id}">从剪贴板读取并补全</button>
          <button class="btn-mini" data-action="toggleManualPmid" data-id="${paper.id}">取消</button>
        </div>
        <div class="manual-pubmed-help">推荐流程：打开 PubMed 搜索 → 在 PubMed 点进正确条目并复制地址栏链接 → 回到本页点“从剪贴板读取并补全”。</div>
      </div>` : "";

    // 操作按钮逻辑：仅补全失败时显示三个重试/手动标签；成功后显示标注标签
    let extraActions = "";
    if (paper.pubmed?.error && !isEnriched) {
      extraActions = `
        <div class="meta-line" style="margin-top:8px">
          <button class="btn-mini" data-action="retryTitleOnly" data-id="${paper.id}"><i class="lucide-rotate-cw"></i> 强力检索标题</button>
          <button class="btn-mini" data-action="openPubmedSearch" data-id="${paper.id}"><i class="lucide-search"></i> 打开 PubMed 搜索</button>
          <button class="btn-mini" data-action="toggleManualPmid" data-id="${paper.id}"><i class="lucide-link"></i> 手动指定PubMed</button>
        </div>`;
    } else if (isEnriched) {
      extraActions = `<div class="meta-line" style="margin-top:8px"><button class="btn-mini" data-action="markAuthors" data-id="${paper.id}"><i class="lucide-user-cog"></i> ${isEditingAuthors ? "收起标注" : "标注共一/共通讯"}</button></div>`;
    }

    return `
      <article class="paper selectable ${isSelected ? "selected" : ""}" data-action="toggleSelect" data-id="${paper.id}">
        <div class="paper-header">
          <div class="check" aria-hidden="true"><i class="lucide-check" style="font-size:14px"></i></div>
          <div style="flex:1">
            <div class="paper-title">
              ${index + 1}. ${paper.title}
              <span class="status-badge ${matchStatus.cls}"><i class="${matchStatus.icon}" style="font-size:12px"></i> ${matchStatus.text}</span>
            </div>

            ${showDetails ? `
              <div class="paper-meta">
                <div class="paper-authors">${paper.authors || "(作者列表待补全)"}</div>
                <div class="meta-line">
                  <span><i class="lucide-calendar" style="font-size:12px"></i> 年份：${paper.year || "-"}</span>
                  <span><i class="lucide-book-open" style="font-size:12px"></i> 期刊：${paper.journal || "-"}</span>
                  <span><i class="lucide-hash" style="font-size:12px"></i> PMID：${pmidText}</span>
                </div>
              </div>
              ${hasPubmedError ? `<div class="paper-meta" style="color:var(--fail); font-size:13px; margin-top:4px">${escapeHtml(paper.pubmed.error)}</div>` : ""}
              ${extraActions}
              ${manualBlock}
              ${editorHtml}
            ` : ""}
          </div>

          ${showDetails ? `
            <div class="metrics-panel">
              <div class="metric-tag"><div class="label">年份</div><div class="val">${state.metricMode === "latest" ? state.latestYear : (paper.year || "-")}</div></div>
              <div class="metric-tag"><div class="label">IF</div><div class="val">${metrics.impactFactor || "-"}</div></div>
              <div class="metric-tag"><div class="label">JCR</div><div class="val">${metrics.jcrZone ? metrics.jcrZone + "区" : "-"}</div></div>
              <div class="metric-tag"><div class="label">中科院</div><div class="val">${metrics.casZone ? metrics.casZone + "区" : "-"}${metrics.casTop ? " (Top)" : ""}</div></div>
            </div>
          ` : ""}
        </div>
      </article>`;
  }).join("");
  updateWorkflowUi();
}

// Keep existing logic for PubMed queries, enrichment, file loading, etc.
// (Simplified for brevity but including all core functions)
const STOP_WORDS = new Set(["a", "an", "and", "are", "as", "at", "be", "between", "by", "do", "for", "from", "how", "in", "into", "is", "it", "its", "not", "of", "on", "or", "the", "their", "this", "to", "via", "was", "were", "who", "with", "among", "study", "studies", "analysis", "meta", "comment", "comments", "trend", "insight", "insights", "evidence", "implications", "reveals", "reveal", "based", "using", "association", "associations"]);
function titleToPubMedKeywordQuery(title, options = {}) {
  const { maxTerms = 6, field = "" } = options;
  const words = (title || "").toLowerCase().replace(/&/g, "and").replace(/[^a-z0-9\u4e00-\u9fa5]+/g, " ").split(/\s+/g).map(w => w.trim()).filter(w => w && !STOP_WORDS.has(w) && !/^\d+$/.test(w) && (w.length >= 3 || /\d/.test(w))).slice(0, maxTerms);
  if (!words.length) return "";
  return field ? words.map(w => `${w}[${field}]`).join(" AND ") : words.join(" AND ");
}
function buildPubMedQueries(paper) {
  const title = (paper.title || "").trim(); const journal = (paper.journal && paper.journal.length >= 6 && /\s/.test(paper.journal)) ? `"${paper.journal}"[Journal]` : "";
  const queries = [];
  if (title) { queries.push(journal ? `"${title}" AND ${journal}` : `"${title}"`); queries.push(journal ? `${title} AND ${journal}` : title); }
  const kwAll = titleToPubMedKeywordQuery(title, { maxTerms: 6 }); if (kwAll) queries.push(journal ? `${kwAll} AND ${journal}` : kwAll);
  const kwTitle = titleToPubMedKeywordQuery(title, { maxTerms: 4, field: "Title" }); if (kwTitle) queries.push(journal ? `${kwTitle} AND ${journal}` : kwTitle);
  return [...new Set(queries.filter(Boolean))];
}
function buildPubMedQueriesTitleOnly(paper) {
  const title = (paper.title || "").trim(); const queries = [];
  if (title) { queries.push(`"${title}"`); queries.push(title); queries.push(`"${title}"[Title]`); }
  const kwAll = titleToPubMedKeywordQuery(title, { maxTerms: 10 }); if (kwAll) queries.push(kwAll);
  const kwTitle = titleToPubMedKeywordQuery(title, { maxTerms: 6, field: "Title" }); if (kwTitle) queries.push(kwTitle);
  return [...new Set(queries.filter(Boolean))];
}
function getPubMedWebsiteSearchUrl(paper) {
  const title = (paper?.title || "").trim(); const journal = (paper?.journal || "").trim();
  const kwAll = titleToPubMedKeywordQuery(title, { maxTerms: 8 });
  const term = (journal && !/待补充/.test(journal)) ? `${kwAll || title} AND "${journal}"[Journal]` : (kwAll || title);
  return `https://pubmed.ncbi.nlm.nih.gov/?term=${encodeURIComponent(term || title || "")}`;
}

const PUBMED_PROXIES = [(url) => `https://api.allorigins.win/raw?url=${encodeURIComponent(url)}`, (url) => `https://corsproxy.io/?${encodeURIComponent(url)}` ];
async function fetchWithFallback(url) {
  try { const r = await fetch(url); if (r.ok) return r; } catch (e) {}
  for (const proxy of PUBMED_PROXIES) { try { const r = await fetch(proxy(url)); if (r.ok) return r; } catch (e) {} }
  throw new Error("PubMed 请求失败");
}
async function fetchPubMedJson(url) { const r = await fetchWithFallback(url); return r.json(); }
async function fetchPubMedXml(url) { const r = await fetchWithFallback(url); return r.text(); }
async function fetchPubMedMedline(url) { const r = await fetchWithFallback(url); return r.text(); }
async function fetchHtmlText(url) { const r = await fetchWithFallback(url); return r.text(); }

function parsePubMedAuthors(xmlText) {
  const doc = new DOMParser().parseFromString(xmlText, "text/xml"); const authors = [...doc.querySelectorAll("Author")];
  const fullAuthors = []; const coFirst = []; const coCorresponding = [];
  const emailRegex = /[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/gi;
  const isCorresponding = (author) => {
    const affs = [...author.querySelectorAll("Affiliation")].map(a => a.textContent || "");
    return affs.some(t => emailRegex.test(t) || /(correspondence|corresponding|通讯)/i.test(t));
  };
  authors.forEach(author => {
    const last = author.querySelector("LastName")?.textContent || ""; const fore = author.querySelector("ForeName")?.textContent || ""; const collective = author.querySelector("CollectiveName")?.textContent || "";
    const name = collective || `${last} ${fore}`.trim(); if (!name) return;
    fullAuthors.push(name);
    if (author.querySelector("EqualContrib")?.textContent?.trim() === "Y") coFirst.push(name);
    if (isCorresponding(author)) coCorresponding.push(name);
  });
  return { fullAuthors, coFirst, coCorresponding };
}

function parseMedlineAuthors(text) {
  return text.split("\n").filter(l => l.startsWith("FAU  - ")).map(l => {
    let n = l.replace("FAU  - ", "").trim();
    if (n.includes(",")) { const [la, fo] = n.split(",").map(s => s.trim()); n = `${la} ${fo}`.trim(); }
    return n;
  }).filter(Boolean);
}
function parseMedlineDoi(text) { const m = text.match(/AID  - (.*?) \[doi\]/); return m ? m[1].trim() : ""; }
function parseMedlineEmails(text) { const emailRegex = /[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/gi; const m = text.match(emailRegex); return m ? [...new Set(m)] : []; }
function getDoiFromSummaryRecord(r) { return (r?.articleids?.find(i => i.idtype === "doi")?.value || "").trim(); }

async function tryEnrichFromDoiPublisher(paper) {
  const doi = paper.pubmed?.doi; if (!doi || !doi.startsWith("10.3389/")) return false;
  try {
    const html = await fetchHtmlText(`https://doi.org/${doi}`); const doc = new DOMParser().parseFromString(html, "text/html");
    const metaAuthors = [...doc.querySelectorAll('meta[name="citation_author"]')].map(m => {
      let n = m.getAttribute("content")?.trim(); if (n && n.includes(",")) { const [la, fo] = n.split(",").map(s => s.trim()); n = `${la} ${fo}`.trim(); }
      return n;
    }).filter(Boolean);
    if (!metaAuthors.length) return false;
    const headText = (doc.body?.textContent || "").replace(/\s+/g, " ").slice(0, 9000);
    const coFirst = []; const coCorr = [];
    metaAuthors.forEach(n => {
      const idx = headText.indexOf(n); if (idx === -1) return;
      const tail = headText.slice(idx + n.length, idx + n.length + 12);
      if (tail.includes("†") || tail.includes("\u2020")) coFirst.push(n);
      if (tail.includes("*")) coCorr.push(n);
    });
    paper.pubmed.fullAuthors = metaAuthors; paper.pubmed.coFirst = coFirst; paper.pubmed.coCorresponding = coCorr;
    paper.authors = formatAuthorsWithMarkers(metaAuthors, coFirst, coCorr);
    return true;
  } catch (e) { return false; }
}

async function searchPubMedRecords(paper, options = {}) {
  const { titleOnly = false, retmax = 5 } = options;
  const terms = titleOnly ? buildPubMedQueriesTitleOnly(paper) : buildPubMedQueries(paper);
  for (const term of terms) {
    const search = await fetchPubMedJson(`https://eutils.ncbi.nlm.nih.gov/entrez/eutils/esearch.fcgi?db=pubmed&retmode=json&retmax=${retmax}&term=${encodeURIComponent(term)}`);
    const ids = search.esearchresult?.idlist || []; if (!ids.length) continue;
    const summary = await fetchPubMedJson(`https://eutils.ncbi.nlm.nih.gov/entrez/eutils/esummary.fcgi?db=pubmed&retmode=json&id=${ids.join(",")}`);
    return ids.map(id => ({ id, record: summary.result?.[id] })).filter(x => x.record);
  }
  return [];
}

function scorePubMedRecord(paper, record) {
  const pT = normalizeTitle(paper.title); const rT = normalizeTitle(record.title);
  let s = 0; if (pT === rT) s += 5; else if (rT.includes(pT) || pT.includes(rT)) s += 3;
  const pJ = normalizeJournal(paper.journal); const rJ = normalizeJournal(record.fulljournalname || record.source || "");
  if (pJ === rJ) s += 3; else if (rJ.includes(pJ) || pJ.includes(rJ)) s += 1;
  if (paper.year && record.pubdate?.startsWith(paper.year)) s += 1;
  return s;
}

async function enrichPaperFromPubMed(paper, options = {}) {
  paper.pubmed = paper.pubmed || {}; paper.pubmed.error = ""; const { titleOnly = false } = options;
  let results = await searchPubMedRecords(paper, { titleOnly, retmax: titleOnly ? 25 : 5 });
  if (!titleOnly && !results.length) results = await searchPubMedRecords(paper, { titleOnly: true, retmax: 25 });
  if (!results.length) { paper.pubmed.error = "未找到匹配的 PubMed 结果。"; return false; }
  let best = results[0]; let bs = scorePubMedRecord(paper, best.record);
  results.slice(1).forEach(c => { const s = scorePubMedRecord(paper, c.record); if (s > bs) { best = c; bs = s; } });
  const r = best.record; paper.title = r.title || paper.title; paper.journal = r.fulljournalname || r.source || paper.journal;
  paper.year = r.pubdate ? r.pubdate.slice(0, 4) : paper.year;
  const vol = [r.volume, r.issue].filter(Boolean).join(r.issue ? `(${r.issue})` : "").trim();
  paper.volumeIssuePages = r.pages ? (vol ? `${vol}:${r.pages}` : r.pages) : vol;
  const sumAuthors = (r.authors || []).map(a => {
    let n = a.name; if (n.includes(",")) { const [la, fo] = n.split(",").map(s => s.trim()); n = `${la} ${fo}`; }
    return n;
  }).filter(Boolean);
  if (sumAuthors.length) paper.authors = sumAuthors.join(", ");
  paper.pubmed.pmid = best.id; paper.pubmed.doi = getDoiFromSummaryRecord(r);
  const xml = await fetchPubMedXml(`https://eutils.ncbi.nlm.nih.gov/entrez/eutils/efetch.fcgi?db=pubmed&retmode=xml&id=${best.id}`);
  const parsed = parsePubMedAuthors(xml);
  paper.pubmed.coFirst = parsed.coFirst; paper.pubmed.coCorresponding = parsed.coCorresponding; paper.pubmed.fullAuthors = parsed.fullAuthors;
  if (parsed.fullAuthors.length) paper.authors = formatAuthorsWithMarkers(parsed.fullAuthors, parsed.coFirst, parsed.coCorresponding);
  else {
    const med = await fetchPubMedMedline(`https://eutils.ncbi.nlm.nih.gov/entrez/eutils/efetch.fcgi?db=pubmed&rettype=medline&retmode=text&id=${best.id}`);
    const mAuthors = parseMedlineAuthors(med);
    if (mAuthors.length) { paper.pubmed.fullAuthors = mAuthors; paper.authors = formatAuthorsWithMarkers(mAuthors, parsed.coFirst, parsed.coCorresponding); }
    const doi = parseMedlineDoi(med); if (doi) paper.pubmed.doi = doi;
    const emails = parseMedlineEmails(med);
    if (!parsed.coCorresponding.length && emails.length && mAuthors.length) {
      paper.pubmed.coCorresponding = [mAuthors[mAuthors.length - 1]];
      paper.authors = formatAuthorsWithMarkers(mAuthors, parsed.coFirst, paper.pubmed.coCorresponding);
    }
  }
  if (paper.pubmed.doi) await tryEnrichFromDoiPublisher(paper);
  return true;
}

async function enrichPaperFromPmid(paper, pmid) {
  paper.pubmed = paper.pubmed || {}; paper.pubmed.error = ""; const id = String(pmid || "").trim(); if (!id) return false;
  const summary = await fetchPubMedJson(`https://eutils.ncbi.nlm.nih.gov/entrez/eutils/esummary.fcgi?db=pubmed&retmode=json&id=${encodeURIComponent(id)}`);
  const r = summary?.result?.[id]; if (!r) { paper.pubmed.error = "PMID 无效。"; return false; }
  paper.title = r.title || paper.title; paper.journal = r.fulljournalname || r.source || paper.journal;
  paper.year = r.pubdate ? r.pubdate.slice(0, 4) : paper.year;
  const vol = [r.volume, r.issue].filter(Boolean).join(r.issue ? `(${r.issue})` : "").trim();
  paper.volumeIssuePages = r.pages ? (vol ? `${vol}:${r.pages}` : r.pages) : vol;
  paper.pubmed.pmid = id; paper.pubmed.doi = getDoiFromSummaryRecord(r);
  const xml = await fetchPubMedXml(`https://eutils.ncbi.nlm.nih.gov/entrez/eutils/efetch.fcgi?db=pubmed&retmode=xml&id=${id}`);
  const parsed = parsePubMedAuthors(xml);
  paper.pubmed.coFirst = parsed.coFirst; paper.pubmed.coCorresponding = parsed.coCorresponding; paper.pubmed.fullAuthors = parsed.fullAuthors;
  if (parsed.fullAuthors.length) paper.authors = formatAuthorsWithMarkers(parsed.fullAuthors, parsed.coFirst, parsed.coCorresponding);
  if (paper.pubmed.doi) await tryEnrichFromDoiPublisher(paper);
  return true;
}

async function enrichSelectedFromPubMed() {
  const targets = state.papers.filter(p => state.selected.has(p.id));
  if (!targets.length) { updateStatus("请先选择论文。"); return; }
  updateStatus("正在通过 PubMed 校验与补全...");
  let success = 0; let done = 0;
  for (const p of targets) {
    try { if (await enrichPaperFromPubMed(p)) success++; } catch(e) {}
    done++; updateStatus(`校验进度：${done}/${targets.length}`);
    await new Promise(r => setTimeout(r, 350));
  }
  state.workflow.pubmed = true; render();
  updateStatus(`校验完成，成功匹配 ${success} 条。对于未成功的条目，您可以尝试“强力检索标题”或手动指定。`);
}

async function exportSelected() {
  const sel = state.papers.filter(p => state.selected.has(p.id));
  if (!sel.length) { updateStatus("请选择要导出的论文。"); return; }

  const ensureTitleEndingPunctuation = (title) => {
    const t = String(title || "").trim();
    if (!t) return t;
    // 如果标题本身已以句末标点结束，就不要再补一个 "."
    if (/[.!?。！？…]$/.test(t)) return t;
    return `${t}.`;
  };

  const formatEntryExport = (paper) => {
    const m = resolveMetrics(paper, state.metricMode);
    const authors = (paper.authors || "").replace(/<[^>]+>/g, "");
    const meta = [];
    if (paper.paperType) meta.push(`[${paper.paperType}]`);
    if (paper.myAuthorRole) meta.push(paper.myAuthorRole);
    if (m.impactFactor) meta.push(`IF=${m.impactFactor}`);
    if (m.casZone) meta.push(`中科院${m.casZone}区`);
    if (m.jcrZone) meta.push(`JCR ${m.jcrZone}区`);
    if (m.casTop) meta.push("Top期刊");
    const title = ensureTitleEndingPunctuation(paper.title);
    return `${authors ? authors + "." : ""} ${title} ${paper.journal || ""}, ${paper.year || ""}${paper.volumeIssuePages ? ", " + paper.volumeIssuePages : ""}. ${meta.length ? "(" + meta.join(", ") + ")" : ""}`;
  };

  const escapeXml = (s) => String(s ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");

  const makeParagraphXml = (text) => {
    // Word 会折叠空格；统一使用 xml:space="preserve"
    return `<w:p><w:r><w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r></w:p>`;
  };

  try {
    updateStatus("正在导出为 DOCX...");
    const lines = sel.map(formatEntryExport);

    const docBody = [
      ...lines.flatMap((line, idx) => {
        const parts = [makeParagraphXml(line)];
        // 条目间空一行
        if (idx < lines.length - 1) parts.push("<w:p/>");
        return parts;
      }),
      // 页面设置（A4）
      `<w:sectPr>
         <w:pgSz w:w="11906" w:h="16838"/>
         <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/>
       </w:sectPr>`
    ].join("");

    const documentXml =
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
      `<w:body>${docBody}</w:body>` +
      `</w:document>`;

    const contentTypesXml =
      `<?xml version="1.0" encoding="UTF-8"?>` +
      `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
      `<Default Extension="xml" ContentType="application/xml"/>` +
      `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
      `</Types>`;

    const relsXml =
      `<?xml version="1.0" encoding="UTF-8"?>` +
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>` +
      `</Relationships>`;

    const zip = new JSZip();
    zip.file("[Content_Types].xml", contentTypesXml);
    zip.folder("_rels").file(".rels", relsXml);
    zip.folder("word").file("document.xml", documentXml);

    const blob = await zip.generateAsync({
      type: "blob",
      mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    });

    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = "整理后论文格式.docx";
    link.click();

    state.workflow.exported = true; updateWorkflowUi();
    updateStatus("已导出：整理后论文格式.docx");
  } catch (e) {
    console.error(e);
    updateStatus("导出失败，请重试。");
  }
}

// Event Listeners
if (elements.pickFileBtn && elements.docxFile) {
  elements.pickFileBtn.addEventListener("click", () => elements.docxFile.click());
}

if (elements.docxFile) {
  elements.docxFile.addEventListener("change", () => {
    const f = elements.docxFile.files?.[0];
    state.ui.sampleDocx = null;
    setChosenFileName(f ? f.name : "");
  });
}

if (elements.sampleBtn) {
  elements.sampleBtn.addEventListener("click", () => loadSampleDocx({ silent: false }));
}

elements.parseBtn.addEventListener("click", async () => {
  try {
    updateStatus("正在解析文档...");
    let text = "";
    const file = elements.docxFile?.files?.[0] || null;
    if (file) {
      const ln = file.name.toLowerCase();
      if (ln.endsWith(".docx")) text = await extractDocxText(file);
      else { const buf = await file.arrayBuffer(); text = await decodeTextFromBuffer(buf); }
    } else if (state.ui.sampleDocx?.buffer) {
      text = await extractDocxTextFromArrayBuffer(state.ui.sampleDocx.buffer);
    } else {
      updateStatus("请先选择文件，或点击“加载示例文件”。");
      return;
    }
    setPapers(parseDocxEntries(text));
    elements.pubmedBtn.disabled = state.papers.length === 0; render();
  } catch (e) { console.error(e); updateStatus("解析失败，请检查文件格式。"); }
});

elements.pubmedBtn.addEventListener("click", enrichSelectedFromPubMed);
elements.metricMode.addEventListener("change", e => { state.metricMode = e.target.value; render(); });
elements.sortYearBtn.addEventListener("click", () => {
  if (state.sort?.field === "year") state.sort.direction = state.sort.direction === "asc" ? "desc" : "asc";
  else state.sort = { field: "year", direction: "desc" }; render();
});
elements.sortImpactBtn.addEventListener("click", () => {
  if (state.sort?.field === "impact") state.sort.direction = state.sort.direction === "asc" ? "desc" : "asc";
  else state.sort = { field: "impact", direction: "desc" }; render();
});
elements.toggleAllBtn.addEventListener("click", () => {
  if (state.selected.size === state.papers.length) state.selected.clear();
  else state.selected = new Set(state.papers.map(p => p.id)); render();
});
elements.exportBtn.addEventListener("click", exportSelected);

elements.paperList.addEventListener("click", e => {
  const actionEl = e.target.closest("[data-action]"); const action = actionEl?.dataset?.action; const id = actionEl?.dataset?.id;
  if (action === "toggleSelect" || e.target.closest(".paper")) {
    const card = e.target.closest(".paper"); const cid = card.dataset.id;
    if (e.target.closest("button") || e.target.closest("input") || e.target.closest("select") || e.target.closest("a")) return;
    if (state.selected.has(cid)) state.selected.delete(cid); else state.selected.add(cid);
    render(); return;
  }
  if (!action || !id) return;
  if (action === "markAuthors") {
    const p = state.papers.find(x => x.id === id); if (!p) return;
    if (state.ui.authorEditOpen.has(id)) { state.ui.authorEditOpen.delete(id); state.ui.authorDraft.delete(id); }
    else {
      state.ui.authorEditOpen.add(id);
      const authors = (p.pubmed?.fullAuthors?.length ? [...p.pubmed.fullAuthors] : (p.authors || "").replace(/<[^>]+>/g, "").split(",").map(s => s.trim().replace(/[#*]+$/g, "")).filter(Boolean));
      state.ui.authorDraft.set(id, { baseAuthors: authors, coFirstSet: new Set((p.pubmed?.coFirst || []).map(normalizeForMatch)), coCorrSet: new Set((p.pubmed?.coCorresponding || []).map(normalizeForMatch)) });
    }
    render(); return;
  }
  if (action === "authorFlag") {
    const draft = state.ui.authorDraft.get(id); const idx = Number(actionEl.dataset.idx); const role = actionEl.dataset.role;
    const key = normalizeForMatch(draft.baseAuthors[idx]);
    if (role === "cofirst") { if (actionEl.checked) draft.coFirstSet.add(key); else draft.coFirstSet.delete(key); }
    else { if (actionEl.checked) draft.coCorrSet.add(key); else draft.coCorrSet.delete(key); }
    const p = state.papers.find(x => x.id === id);
    const cf = draft.baseAuthors.filter(n => draft.coFirstSet.has(normalizeForMatch(n)));
    const cc = draft.baseAuthors.filter(n => draft.coCorrSet.has(normalizeForMatch(n)));
    p.authors = formatAuthorsWithMarkers(draft.baseAuthors, cf, cc);
    render(); return;
  }
  if (action === "saveAuthorEdit" || action === "cancelAuthorEdit") {
    state.ui.authorEditOpen.delete(id); state.ui.authorDraft.delete(id); render(); return;
  }
  if (action === "toggleManualPmid") {
    if (state.ui.pubmedManualOpen.has(id)) state.ui.pubmedManualOpen.delete(id);
    else { state.ui.pubmedManualOpen.add(id); state.ui.pubmedManualValue.set(id, ""); }
    render(); return;
  }
  if (action === "applyManualPmid" || action === "clipboardApplyManualPmid") {
    const p = state.papers.find(x => x.id === id); if (!p) return;
    (async () => {
      const val = action === "clipboardApplyManualPmid" ? await navigator.clipboard.readText() : state.ui.pubmedManualValue.get(id);
      const pmid = parsePmidFromInput(val);
      if (!pmid) { updateStatus("未识别到 PMID"); return; }
      updateStatus(`正在按 PMID ${pmid} 补全...`);
      if (await enrichPaperFromPmid(p, pmid)) { state.ui.pubmedManualOpen.delete(id); updateStatus("补全成功"); }
      else updateStatus("补全失败");
      render();
    })();
    return;
  }
  if (action === "openPubmedSearch") {
    const p = state.papers.find(x => x.id === id); window.open(getPubMedWebsiteSearchUrl(p), "_blank"); return;
  }
  if (action === "retryTitleOnly") {
    const p = state.papers.find(x => x.id === id);
    (async () => {
      updateStatus("正在进行强力检索...");
      if (await enrichPaperFromPubMed(p, { titleOnly: true })) updateStatus("强力检索成功");
      else updateStatus("强力检索未找到结果");
      render();
    })();
    return;
  }
});

elements.paperList.addEventListener("input", (e) => {
  if (e.target.dataset.action === "manualPmidInput") state.ui.pubmedManualValue.set(e.target.dataset.id, e.target.value);
});

async function loadBuiltinMetrics() {
  try {
    updateStatus("正在加载内置指标表...");
    const [iT, cT, jT] = await Promise.all([loadTextFromUrl("影响因子.txt"), loadTextFromUrl("中科院分区.txt"), loadTextFromUrl("JCR分区.txt")]);
    state.datasets.impact = parseImpactTxt(iT); state.datasets.cas = parseZoneTxt(cT); state.datasets.jcr = parseZoneTxt(jT);
    const jNM = new Map(); const jNML = new Map();
    const add = (d) => { if (!d?.map) return; d.map.forEach((r, k) => { if (!jNM.has(k)) jNM.set(k, r._name || ""); const lK = normalizeJournalLoose(r._name || ""); if (lK && !jNML.has(lK)) jNML.set(lK, r._name || ""); }); };
    add(state.datasets.impact); add(state.datasets.cas); add(state.datasets.jcr);
    state.journalNameMap = jNM; state.journalNameMapLoose = jNML;
    state.journalKeys = [...jNM.keys()].sort((a,b)=>b.length-a.length);
    state.journalKeysLoose = [...jNML.keys()].sort((a,b)=>b.length-a.length);
    const years = new Set([...(state.datasets.impact?.years||[]), ...(state.datasets.cas?.years||[]), ...(state.datasets.jcr?.years||[])]);
    if (years.size) state.latestYear = [...years].map(Number).sort((a,b)=>b-a)[0];
    updateStatus("系统就绪。"); render();
    return true;
  } catch (e) { updateStatus("指标表加载失败。"); }
  return false;
}

(async () => {
  setChosenFileName("");
  await loadBuiltinMetrics();
})();
