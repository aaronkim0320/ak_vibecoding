(() => {
  "use strict";

  const DATA_PATH = "data/";
  const SALES_SOURCES = [
    { key: "01", filename: "01_영업실적.xlsx", label: "영업 실적", channel: "영업" },
    { key: "02", filename: "02_온라인실적.xlsx", label: "온라인 실적", channel: "온라인" },
    { key: "03", filename: "03_팝업실적.xlsx", label: "팝업 실적", channel: "팝업" },
  ];
  const REQUIRED = ["record_id", "기준주", "점포", "브랜드", "장르", "채널", "매출_만원", "목표_만원", "비용_만원", "상태", "출처", "기준일"];
  const ALIASES = {
    record_id: ["record_id", "레코드ID", "기록번호"], 기준주: ["기준주", "주차", "주간기준일"],
    점포: ["점포", "점포명", "영업점"], 브랜드: ["브랜드", "브랜드명", "행사브랜드"],
    장르: ["장르", "카테고리", "상품군"], 채널: ["채널", "유통채널", "판매채널"],
    매출_만원: ["매출_만원", "실적(만원)", "매출액_만원"], 목표_만원: ["목표_만원", "목표매출(만원)", "목표액_만원"],
    비용_만원: ["비용_만원", "집행비용(만원)", "사용비용_만원"], 상태: ["상태", "진행상태", "진행현황"],
    출처: ["출처", "근거자료", "자료출처"], 기준일: ["기준일", "확인일", "기준일자"],
  };
  const TEMPLATE_HEADERS = ["record_id", "채널", "기준주", "점포", "브랜드", "장르", "매출_만원", "목표_만원", "목표차이_만원", "달성률", "비용_만원", "비용차감액_만원", "상태", "메모", "출처", "기준일"];
  const MEMO_KEY = "moment-room-dashboard-memos-v1";
  const LOCATION_KEY = "moment-room-dashboard-weather-location-v1";
  const sources = Object.fromEntries(SALES_SOURCES.map((s) => [s.key, { ...s, rows: [], status: "불러오는 중" }]));
  let checklist = [];
  let reportHeaders = [...TEMPLATE_HEADERS];
  let reportTemplateReady = false;
  let memos = readStorage(MEMO_KEY, {});
  let filteredSales = [];
  let page = 1;
  let weatherLocations = [];
  let weatherMode = "sample";
  let toastTimer;
  let memoTimers = new Map();

  const $ = (selector) => document.querySelector(selector);
  const $$ = (selector) => Array.from(document.querySelectorAll(selector));
  const esc = (value) => String(value ?? "").replace(/[&<>"']/g, (ch) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[ch]);
  const number = (value) => {
    if (value === null || value === undefined || value === "") return 0;
    if (typeof value === "number") return Number.isFinite(value) ? value : 0;
    const parsed = Number(String(value).replace(/,/g, "").trim());
    return Number.isFinite(parsed) ? parsed : 0;
  };
  const sum = (rows, field) => rows.reduce((total, row) => total + number(row[field]), 0);
  const fmt = (value, digits = 1) => Number(value || 0).toLocaleString("ko-KR", { maximumFractionDigits: digits });
  const dateLabel = (value, options = { month: "2-digit", day: "2-digit" }) => {
    if (!value) return "날짜 미정";
    const parts = String(value).slice(0, 10).split("-").map(Number);
    if (parts.length !== 3 || parts.some(Number.isNaN)) return String(value);
    return new Intl.DateTimeFormat("ko-KR", options).format(new Date(parts[0], parts[1] - 1, parts[2]));
  };
  function readStorage(key, fallback) {
    try { return JSON.parse(localStorage.getItem(key) || "null") ?? fallback; } catch { return fallback; }
  }
  function toast(message, duration = 3300) {
    const el = $("#toast");
    el.textContent = message;
    el.classList.add("visible");
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => el.classList.remove("visible"), duration);
  }
  function localSave(key, value) {
    try { localStorage.setItem(key, JSON.stringify(value)); return true; }
    catch { toast("브라우저 저장 공간에 기록하지 못했습니다. 브라우저 설정을 확인해 주세요.", 5000); return false; }
  }
  function setTodayLabel() {
    $("#todayLabel").textContent = new Intl.DateTimeFormat("ko-KR", { timeZone: "Asia/Seoul", year: "numeric", month: "long", day: "numeric", weekday: "short" }).format(new Date());
  }

  function normalizeHeader(raw) {
    const normalized = {};
    for (const field of REQUIRED) {
      const aliases = ALIASES[field];
      const found = aliases.find((name) => Object.prototype.hasOwnProperty.call(raw, name));
      if (found) normalized[field] = raw[found];
    }
    return normalized;
  }
  function workbookRows(workbook, diagnosticOnly = false) {
    if (!workbook.SheetNames.includes("원본") && !diagnosticOnly) throw new Error("필수 시트 누락: 원본");
    const sheet = workbook.Sheets.원본 || workbook.Sheets[workbook.SheetNames[0]];
    if (!sheet) throw new Error("시트가 비어 있습니다.");
    return XLSX.utils.sheet_to_json(sheet, { defval: "", raw: false, blankrows: false });
  }
  async function parseSalesFile(file, expectedSource = null) {
    if (!window.XLSX) throw new Error("엑셀 읽기 라이브러리를 불러오지 못했습니다.");
    const buffer = file instanceof ArrayBuffer ? file : await file.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: "array", cellDates: false });
    // 미인식 추가 파일은 첫 시트에서 누락 열을 진단하지만 적재하지 않습니다.
    const rawRows = workbookRows(workbook, !expectedSource);
    if (!rawRows.length) throw new Error("데이터 행이 없습니다.");
    const headers = Object.keys(rawRows[0]);
    const missing = REQUIRED.filter((field) => !ALIASES[field].some((alias) => headers.includes(alias)));
    if (missing.length) throw new Error(`필수 열 누락: ${missing.join(", ")}`);
    if (!workbook.SheetNames.includes("원본")) throw new Error("필수 시트 누락: 원본");
    const filename = file.name || expectedSource?.filename || "입력 파일";
    const prepared = rawRows.map((raw, index) => ({ normalized: normalizeHeader(raw), line: (raw.__rowNum__ ?? index + 1) + 1 }));
    const badChannels = expectedSource ? prepared.filter(({ normalized }) => String(normalized.채널).trim() !== expectedSource.channel) : [];
    if (badChannels.length) {
      const rows = badChannels.slice(0, 6).map((item) => item.line).join(", ");
      throw new Error(`${filename}: 채널 값 불일치 (${expectedSource.channel} 파일 기준, ${badChannels.length}행; 원본 시트 행 ${rows})`);
    }
    const emptyId = prepared.filter(({ normalized }) => !String(normalized.record_id ?? "").trim());
    if (emptyId.length) throw new Error(`${filename}: 빈 record_id, 원본 시트 행 ${emptyId.map((item) => item.line).join(", ")}`);
    const idLines = new Map();
    for (const item of prepared) {
      const id = String(item.normalized.record_id).trim();
      if (idLines.has(id)) throw new Error(`${filename}: record_id 중복 ${id}, 원본 시트 행 ${idLines.get(id)}, ${item.line}`);
      idLines.set(id, item.line);
    }
    const output = prepared.map(({ normalized, line }) => ({
      ...normalized,
      record_id: String(normalized.record_id).trim(),
      기준주: String(normalized.기준주 ?? "").slice(0, 10),
      기준일: String(normalized.기준일 ?? "").slice(0, 10),
      점포: String(normalized.점포 ?? ""), 브랜드: String(normalized.브랜드 ?? ""), 장르: String(normalized.장르 ?? ""),
      채널: String(normalized.채널 ?? "").trim(), 상태: String(normalized.상태 ?? ""), 출처: String(normalized.출처 || file.name),
      매출_만원: number(normalized.매출_만원), 목표_만원: number(normalized.목표_만원), 비용_만원: number(normalized.비용_만원),
      _sourceFile: filename, _sourceSheet: "원본", _sourceRow: line,
    }));
    return output;
  }
  function sourceByName(filename) {
    const base = filename.split(/[\\/]/).pop() || "";
    return SALES_SOURCES.find((source) => base.startsWith(source.key + "_") || base === source.filename) || null;
  }
  function allRows() { return SALES_SOURCES.flatMap((source) => sources[source.key].rows); }
  function ingestRows(source, rows) {
    if (!source || !sources[source.key]) throw new Error("세 실적 원본 파일명(01_, 02_, 03_)을 확인해 주세요.");
    const existing = new Map(SALES_SOURCES.filter((item) => item.key !== source.key).flatMap((item) => sources[item.key].rows).map((row) => [row.record_id, row]));
    for (const row of rows) {
      const other = existing.get(row.record_id);
      if (other) throw new Error(`전체 자료 record_id 중복 ${row.record_id}: ${other._sourceFile} 원본 행 ${other._sourceRow} / ${row._sourceFile} 원본 행 ${row._sourceRow}`);
    }
    sources[source.key].rows = rows;
    sources[source.key].status = "정상";
    return rows.length;
  }
  async function loadBundledSales() {
    for (const item of SALES_SOURCES) {
      const state = sources[item.key];
      state.status = "자료 읽는 중";
      updateSourceStatus();
      try {
        const response = await fetch(`${DATA_PATH}${encodeURIComponent(item.filename)}`);
        if (!response.ok) throw new Error(`파일을 읽지 못했습니다 (${response.status}).`);
        const rows = await parseSalesFile(await response.arrayBuffer(), item);
        ingestRows(item, rows);
      } catch (error) {
        state.status = `오류: ${error.message}`;
      }
      updateSourceStatus();
    }
    updateSalesFilters();
    updateSourceStatus();
    renderSales();
  }
  async function loadReportTemplate() {
    try {
      const response = await fetch(`${DATA_PATH}04_%EB%B3%B4%EA%B3%A0%EC%96%91%EC%8B%9D.xlsx`);
      if (!response.ok) throw new Error(`보고 양식 조회 실패 (${response.status})`);
      const workbook = XLSX.read(await response.arrayBuffer(), { type: "array" });
      // 제공된 보고 양식의 실제 원본 시트 이름은 '보고'입니다.
      const sheetName = workbook.SheetNames.includes("보고") ? "보고" : "원본";
      if (!workbook.Sheets[sheetName]) throw new Error("보고 양식 시트 누락 (보고 또는 원본)");
      const table = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, defval: "", raw: false, blankrows: false });
      const headers = (table.find((row) => row.some((cell) => String(cell).trim() !== "")) || []).map((cell) => String(cell).trim());
      if (headers.length !== 16 || new Set(headers).size !== 16 || !headers.every((field) => TEMPLATE_HEADERS.includes(field))) throw new Error("보고 양식의 필수 16개 헤더를 확인해 주세요.");
      reportHeaders = headers;
      reportTemplateReady = true;
    } catch (error) { toast(`보고 파일 저장 준비 실패: ${error.message}`, 6500); }
  }
  function updateSalesFilters() {
    const rows = allRows();
    const storeSelect = $("#storeFilter");
    const weekSelect = $("#weekFilter");
    const oldStore = storeSelect.value;
    const oldWeek = weekSelect.value;
    const storesList = [...new Set(rows.map((row) => row.점포).filter(Boolean))].sort((a, b) => a.localeCompare(b, "ko"));
    const weeksList = [...new Set(rows.map((row) => row.기준주).filter(Boolean))].sort();
    storeSelect.innerHTML = `<option value="all">전체 점포</option>${storesList.map((v) => `<option value="${esc(v)}">${esc(v)}</option>`).join("")}`;
    weekSelect.innerHTML = `<option value="all">전체 기간</option>${weeksList.map((v) => `<option value="${esc(v)}">${esc(dateLabel(v))} 주</option>`).join("")}`;
    if (storesList.includes(oldStore)) storeSelect.value = oldStore;
    if (weeksList.includes(oldWeek)) weekSelect.value = oldWeek;
    for (const [selector, field, label] of [["#brandFilter", "브랜드", "전체 브랜드"], ["#genreFilter", "장르", "전체 장르"]]) {
      const select = $(selector);
      const previous = select.value;
      const values = [...new Set(rows.map((row) => row[field]).filter(Boolean))].sort((a, b) => a.localeCompare(b, "ko"));
      select.innerHTML = `<option value="all">${label}</option>${values.map((value) => `<option value="${esc(value)}">${esc(value)}</option>`).join("")}`;
      if (values.includes(previous)) select.value = previous;
    }
  }
  function updateSourceStatus() {
    const target = $("#sourceStatus");
    target.innerHTML = SALES_SOURCES.map((item) => {
      const state = sources[item.key];
      const success = state.status === "정상";
      return `<div class="source-chip"><strong>${esc(item.label)}</strong><span class="${success ? "source-ok" : "source-error"}">${success ? `${state.rows.length}행 · 연결됨` : esc(state.status)}</span></div>`;
    }).join("");
    const loaded = allRows().length;
    $("#recordCountLabel").textContent = `${fmt(loaded, 0)} ROWS`;
    const connected = SALES_SOURCES.every((item) => sources[item.key].status === "정상");
    const label = $("#connectionLabel");
    if (label) label.textContent = connected ? "실습 데이터 연결됨" : SALES_SOURCES.some((item) => sources[item.key].status.startsWith("오류")) ? "일부 자료 확인 필요" : "자료 연결 확인 중";
  }

  function getSalesFilter() {
    const store = $("#storeFilter").value;
    const channel = $("#channelFilter").value;
    const week = $("#weekFilter").value;
    return { store, channel, week, brand: $("#brandFilter").value, genre: $("#genreFilter").value };
  }
  function filterSales(rows, filters = {}) {
    const fields = { store: "점포", channel: "채널", week: "기준주", brand: "브랜드", genre: "장르" };
    return rows.filter((row) => Object.entries(fields).every(([key, field]) => !filters[key] || filters[key] === "all" || row[field] === filters[key]));
  }
  function summarizeSales(rows) {
    const revenue = sum(rows, "매출_만원"), goal = sum(rows, "목표_만원"), cost = sum(rows, "비용_만원");
    return { count: rows.length, revenue, goal, cost, gap: revenue - goal, net: revenue - cost, attainment: goal === 0 ? null : revenue / goal };
  }
  function renderSales() {
    const filters = getSalesFilter();
    filteredSales = filterSales(allRows(), filters);
    const storeName = filters.store === "all" ? "전체 점포" : filters.store;
    const channelName = filters.channel === "all" ? "전체 채널" : filters.channel;
    const periodName = filters.week === "all" ? "전체 기간" : `${dateLabel(filters.week)} 주`;
    const brandName = filters.brand === "all" ? "전체 브랜드" : filters.brand;
    const genreName = filters.genre === "all" ? "전체 장르" : filters.genre;
    $("#filterSummary").textContent = `${storeName} · ${channelName} · ${periodName} · ${brandName} · ${genreName}`;
    const revenue = sum(filteredSales, "매출_만원");
    const goal = sum(filteredSales, "목표_만원");
    const cost = sum(filteredSales, "비용_만원");
    const attainment = goal === 0 ? null : (revenue / goal) * 100;
    $("#salesKpi").textContent = fmt(revenue);
    $("#targetKpi").textContent = fmt(goal);
    $("#rateKpi").textContent = attainment === null ? "—" : `${fmt(attainment)}%`;
    $("#netKpi").textContent = fmt(revenue - cost);
    $("#rowsKpi").textContent = `${fmt(filteredSales.length, 0)}개 기록`;
    const gap = revenue - goal;
    $("#targetGapKpi").textContent = `목표 차이 ${gap >= 0 ? "+" : "−"}${fmt(Math.abs(gap))}`;
    $("#rateCaption").textContent = attainment === null ? "목표 합계가 0입니다" : attainment >= 100 ? "목표 달성" : "목표 미달성";
    $("#costKpi").textContent = `비용 ${fmt(cost)}`;
    $("#chartRows").textContent = `${fmt(filteredSales.length, 0)}개 기록 포함`;
    renderChannelChart(filteredSales);
    page = 1;
    renderRecordTable();
  }
  function renderChannelChart(rows) {
    const stores = [...new Set(rows.map((row) => row.점포))].sort((a, b) => a.localeCompare(b, "ko"));
    const totals = stores.map((store) => {
      const subset = rows.filter((row) => row.점포 === store);
      return { channel: store, sales: sum(subset, "매출_만원"), target: sum(subset, "목표_만원") };
    });
    const max = Math.max(0, ...totals.map((item) => Math.max(item.sales, item.target)));
    const target = $("#channelChart");
    if (!rows.length) { target.innerHTML = `<div class="chart-empty">필터 조건에 해당하는 실적이 없습니다.</div>`; return; }
    target.innerHTML = totals.map((item) => {
      const salesWidth = max ? item.sales / max * 100 : 0;
      const targetPosition = max ? item.target / max * 100 : 0;
      return `<div class="chart-row"><span class="chart-label">${item.channel}</span><div class="bar-track"><span class="bar-fill" style="width:${salesWidth}%"></span><span class="target-mark" style="left:${targetPosition}%"></span></div><span class="chart-value">${fmt(item.sales, 0)} <small>/ ${fmt(item.target, 0)}</small></span></div>`;
    }).join("");
  }
  function recordRowsForPage() {
    const query = $("#recordSearch").value.trim().toLocaleLowerCase("ko");
    const rows = filteredSales.filter((row) => !query || [row.record_id, row.브랜드, row.점포, row.채널].some((item) => String(item).toLocaleLowerCase("ko").includes(query)));
    const size = Number($("#pageSize").value) || 15;
    const pages = Math.max(1, Math.ceil(rows.length / size));
    page = Math.min(page, pages);
    return { rows, pages, size, shown: rows.slice((page - 1) * size, page * size) };
  }
  function renderRecordTable() {
    const { rows, pages, size, shown } = recordRowsForPage();
    const body = $("#recordTableBody");
    if (!shown.length) {
      body.innerHTML = `<tr><td colspan="11" class="chart-empty">표시할 실적이 없습니다.</td></tr>`;
    } else {
      body.innerHTML = shown.map((row) => {
        const rate = number(row.목표_만원) ? `${fmt(number(row.매출_만원) / number(row.목표_만원) * 100)}%` : "—";
        const net = number(row.매출_만원) - number(row.비용_만원);
        const channelClass = row.채널 === "온라인" ? "online" : row.채널 === "팝업" ? "popup" : "";
        return `<tr><td>${esc(dateLabel(row.기준주))}</td><td>${esc(row.점포)}</td><td>${esc(row.브랜드)}</td><td><span class="channel-pill ${channelClass}">${esc(row.채널)}</span></td><td class="numeric">${fmt(row.매출_만원, 0)}</td><td class="numeric">${fmt(row.목표_만원, 0)}</td><td class="numeric">${rate}</td><td class="numeric">${fmt(net, 0)}</td><td><input type="text" maxlength="300" value="${esc(memos[row.record_id] || "")}" placeholder="메모 추가" data-record-id="${esc(row.record_id)}" aria-label="${esc(row.record_id)} 메모"></td><td>${esc(row.출처)}</td><td title="${esc(row.record_id)}">${esc(row._sourceFile)} · ${esc(row._sourceSheet)} ${row._sourceRow}행</td></tr>`;
      }).join("");
    }
    $("#recordPageLabel").textContent = `${rows.length ? (page - 1) * size + 1 : 0}–${Math.min(page * size, rows.length)} / ${fmt(rows.length, 0)}개`;
    $("#pageNumber").textContent = `${page} / ${pages}`;
    $("#prevPage").disabled = page <= 1;
    $("#nextPage").disabled = page >= pages;
  }

  function parseCSV(text) {
    const clean = text.replace(/^\uFEFF/, "");
    const output = [];
    let row = [], field = "", quoted = false;
    for (let i = 0; i < clean.length; i++) {
      const char = clean[i];
      if (quoted) {
        if (char === '"' && clean[i + 1] === '"') { field += '"'; i++; }
        else if (char === '"') quoted = false;
        else field += char;
      } else if (char === '"') quoted = true;
      else if (char === ",") { row.push(field); field = ""; }
      else if (char === "\n") { row.push(field.replace(/\r$/, "")); if (row.some((cell) => cell !== "")) output.push(row); row = []; field = ""; }
      else field += char;
    }
    row.push(field.replace(/\r$/, ""));
    if (row.some((cell) => cell !== "")) output.push(row);
    if (!output.length) return [];
    const headers = output.shift().map((header) => header.trim());
    return output.map((cells) => Object.fromEntries(headers.map((header, index) => [header, cells[index] ?? ""])));
  }
  async function loadChecklist() {
    try {
      const response = await fetch(`${DATA_PATH}08_%ED%96%89%EC%82%AC%EC%A4%80%EB%B9%84.csv`);
      if (!response.ok) throw new Error("CSV를 읽지 못했습니다.");
      checklist = parseCSV(await response.text());
      const csvCriterion = checklist.find((task) => task.점검기준일)?.점검기준일;
      const criterionSelect = $("#criterionFilter");
      if (csvCriterion && Array.from(criterionSelect.options).some((option) => option.value === csvCriterion)) criterionSelect.value = csvCriterion;
      const areas = [...new Set(checklist.map((task) => task.업무영역).filter(Boolean))].sort((a, b) => a.localeCompare(b, "ko"));
      $("#areaFilter").innerHTML = `<option value="all">전체 업무</option>${areas.map((area) => `<option value="${esc(area)}">${esc(area)}</option>`).join("")}`;
      renderTasks();
    } catch (error) {
      $("#taskTableBody").innerHTML = `<tr><td colspan="7">행사 준비 자료를 불러오지 못했습니다: ${esc(error.message)}</td></tr>`;
      $("#taskCountLabel").textContent = "0개 항목";
    }
  }
  function dateOffset(dateString, days) {
    const [year, month, day] = dateString.split("-").map(Number);
    const date = new Date(year, month - 1, day + days);
    return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`;
  }
  function deadlineCategory(task, criterion) {
    const due = String(task.마감일 || "").trim();
    if (!due) return { name: "마감일 확인 필요", className: "missing" };
    if (due < criterion) return { name: "지연", className: "late" };
    if (due === criterion) return { name: "오늘 마감", className: "today" };
    if (due <= dateOffset(criterion, 2)) return { name: "마감 임박", className: "soon" };
    return { name: "이후 마감", className: "later" };
  }
  const isOpenTask = (task) => !["완료", "취소"].includes(task.상태);
  function renderTasks() {
    const criterion = $("#criterionFilter").value;
    const open = checklist.filter(isOpenTask);
    const categories = ["지연", "오늘 마감", "마감 임박", "이후 마감", "마감일 확인 필요"];
    const counts = Object.fromEntries(categories.map((name) => [name, 0]));
    open.forEach((task) => { counts[deadlineCategory(task, criterion).name]++; });
    const statusCounts = { "미완료": open.length, "지연": counts.지연, "오늘 마감": counts["오늘 마감"], "마감 임박": counts["마감 임박"], "이후 마감": counts["이후 마감"], "날짜 확인": counts["마감일 확인 필요"] };
    $("#taskStats").innerHTML = Object.entries(statusCounts).map(([name, count]) => `<div class="task-stat ${name === "지연" ? "overdue" : name === "오늘 마감" ? "today" : name === "마감 임박" ? "soon" : name === "미완료" ? "pending" : ""}"><span>${esc(name)}</span><strong>${count}</strong></div>`).join("");
    const statusFilter = $("#taskStatusFilter").value;
    const areaFilter = $("#areaFilter").value;
    let shown = checklist.filter((task) => (statusFilter === "all" || (statusFilter === "open" ? isOpenTask(task) : task.상태 === statusFilter)) && (areaFilter === "all" || task.업무영역 === areaFilter));
    const priorityWeight = { "높음": 0, "보통": 1, "낮음": 2 };
    shown = shown.slice().sort((a, b) => {
      const dueA = a.마감일 || "9999-99-99", dueB = b.마감일 || "9999-99-99";
      return dueA.localeCompare(dueB) || (priorityWeight[a.우선순위] ?? 3) - (priorityWeight[b.우선순위] ?? 3) || a.task_id.localeCompare(b.task_id);
    });
    const statusClass = { "미착수": "pending-status", "진행 중": "progress-status", "검토 대기": "review-status", "완료": "done-status", "취소": "cancelled-status" };
    $("#taskTableBody").innerHTML = shown.length ? shown.map((task) => {
      const deadline = deadlineCategory(task, criterion);
      const priorityClass = task.우선순위 === "높음" ? "high" : task.우선순위 === "낮음" ? "low" : "normal";
      return `<tr><td title="${esc(task.확인사항)}">${esc(task.준비항목)}</td><td>${esc(task.업무영역)}</td><td>${esc(task.담당역할)}</td><td>${esc(task.마감일 ? dateLabel(task.마감일, { year: "numeric", month: "2-digit", day: "2-digit" }) : "날짜 미정")}</td><td><span class="status-pill ${statusClass[task.상태] || ""}">${esc(task.상태)}</span></td><td><span class="deadline-pill ${deadline.className}">${deadline.name}</span></td><td><span class="priority-pill ${priorityClass}">${esc(task.우선순위)}</span></td></tr>`;
    }).join("") : `<tr><td colspan="7" class="chart-empty">조건에 맞는 준비 항목이 없습니다.</td></tr>`;
    $("#taskCountLabel").textContent = `${shown.length}개 항목 · 전체 ${checklist.length}건`;
  }

  function weatherFor(code) {
    const id = Number(code);
    if (id >= 200 && id < 300) return ["뇌우", "⛈"];
    if (id >= 300 && id < 400) return ["이슬비", "🌦"];
    if (id >= 500 && id < 600) return ["비", "🌧"];
    if (id >= 600 && id < 700) return ["눈", "🌨"];
    if (id >= 700 && id < 800) return ["안개", "🌫"];
    if (id === 800) return ["맑음", "☀"];
    if (id === 801) return ["구름 조금", "🌤"];
    if (id === 802) return ["구름 많음", "⛅"];
    if (id >= 803 && id <= 804) return ["흐림", "☁"];
    const mappings = [
      [[0], "맑음", "☀"], [[1], "대체로 맑음", "🌤"], [[2], "구름 조금", "⛅"], [[3], "흐림", "☁"],
      [[45, 48], "안개", "🌫"], [[51, 53, 55, 56, 57], "이슬비", "🌦"], [[61, 63, 65, 66, 67], "비", "🌧"],
      [[71, 73, 75, 77], "눈", "🌨"], [[80, 81, 82], "소나기", "🌦"], [[85, 86], "눈 소나기", "🌨"], [[95, 96, 99], "뇌우", "⛈"],
    ];
    return mappings.find(([codes]) => codes.includes(Number(code)))?.slice(1) || ["날씨 정보", "☁"];
  }
  function percentChance(value) { return value === undefined || value === null ? "" : `강수 ${Math.round(Number(value) * (Number(value) <= 1 ? 100 : 1))}%`; }
  function renderForecast(items) {
    $("#forecastList").innerHTML = items.map((item) => {
      const [description, symbol] = item.description ? [item.description, weatherFor(item.conditionId)[1]] : weatherFor(item.conditionId);
      const chance = item.pop === undefined || item.pop === null ? "" : percentChance(item.pop);
      return `<div class="forecast-day"><time>${esc(dateLabel(item.date, { weekday: "short", month: "2-digit", day: "2-digit" }))}</time><div class="forecast-symbol" aria-label="${esc(description)}">${symbol}</div><strong>${fmt(item.max, 0)}° <span style="color:#a1aaa2;font-weight:400">${fmt(item.min, 0)}°</span></strong><small>${esc(chance || description)}</small></div>`;
    }).join("");
  }
  function formatWeatherTime(value, timezone) {
    if (!value) return "조회 시각 미제공";
    const date = new Date(value);
    if (Number.isNaN(date.getTime())) return String(value);
    return `조회 ${new Intl.DateTimeFormat("ko-KR", { timeZone: timezone || "Asia/Seoul", month: "2-digit", day: "2-digit", hour: "2-digit", minute: "2-digit", hour12: false }).format(date)}`;
  }
  async function loadSampleWeather() {
    try {
      const response = await fetch(`${DATA_PATH}11_%EB%82%A0%EC%94%A8%EC%9D%91%EB%8B%B5_%EC%98%88%EC%A0%9C.json`);
      if (!response.ok) throw new Error("예제 응답을 읽지 못했습니다.");
      const data = await response.json();
      weatherMode = "sample";
      try { localStorage.removeItem(LOCATION_KEY); } catch { /* 브라우저 저장소가 차단되어도 예제 날씨는 표시합니다. */ }
      const current = data.current;
      $("#weatherNotice").innerHTML = `<strong>교육용 예제</strong><span>${esc(data.source.name)} · 실제 점포 위치가 아닙니다.</span>`;
      $("#weatherPlace").textContent = data.location.name;
      $("#weatherTemp").textContent = `${Math.round(current.temp)}°`;
      $("#weatherDesc").textContent = `${current.description} · 체감 ${Math.round(current.feelsLike)}°`;
      $("#weatherIcon").textContent = weatherFor(current.conditionId)[1];
      $("#weatherHumidity").textContent = `습도 ${current.humidity}%`;
      $("#weatherWind").textContent = `바람 ${current.wind} m/s`;
      $("#weatherFetched").textContent = formatWeatherTime(data.fetchedAt, data.location.timezone);
      renderForecast(data.forecast.slice(0, 5));
      $("#weatherModeButton").textContent = "실제 날씨 검색";
      $("#weatherSearchForm").hidden = true;
    } catch (error) { toast(`예제 날씨를 불러오지 못했습니다. ${error.message}`); }
  }
  async function searchWeatherLocations(query) {
    const endpoint = new URL("https://geocoding-api.open-meteo.com/v1/search");
    endpoint.search = new URLSearchParams({ name: query, count: "6", language: "ko", format: "json", countryCode: "KR" }).toString();
    const response = await fetch(endpoint);
    if (!response.ok) throw new Error("지역을 검색하지 못했습니다.");
    const data = await response.json();
    weatherLocations = data.results || [];
    const select = $("#weatherLocationResults");
    if (!weatherLocations.length) {
      select.hidden = true;
      throw new Error("검색 결과가 없습니다. 도시나 지역 이름을 확인해 주세요.");
    }
    select.innerHTML = weatherLocations.map((place, index) => {
      const label = [place.name, place.admin1, place.country].filter(Boolean).join(" · ");
      return `<option value="${index}">${esc(label)}</option>`;
    }).join("");
    select.hidden = false;
    $("#weatherSearchStatus").textContent = `${weatherLocations.length}개 위치를 찾았습니다. 목록에서 선택한 뒤 날씨를 불러오세요.`;
    toast(`${weatherLocations.length}개 지역을 찾았습니다. 목록에서 선택해 주세요.`);
  }
  async function loadRealWeather(place = null) {
    if (!place) {
      const index = Number($("#weatherLocationResults").value);
      place = weatherLocations[index];
    }
    if (!place) { toast("먼저 점포가 있는 지역을 검색하고 선택해 주세요."); return; }
    const endpoint = new URL("https://api.open-meteo.com/v1/forecast");
    endpoint.search = new URLSearchParams({
      latitude: String(place.latitude), longitude: String(place.longitude), current: "temperature_2m,relative_humidity_2m,apparent_temperature,precipitation,weather_code,wind_speed_10m",
      daily: "weather_code,temperature_2m_max,temperature_2m_min,precipitation_probability_max", forecast_days: "5", timezone: place.timezone || "auto",
    }).toString();
    const button = $("#loadRealWeather");
    button.disabled = true;
    button.textContent = "불러오는 중…";
    try {
      const response = await fetch(endpoint);
      const data = await response.json();
      if (!response.ok || data.error) throw new Error(data.reason || "날씨 응답을 가져오지 못했습니다.");
      const current = data.current;
      const placeName = [place.name, place.admin1, place.country].filter(Boolean).join(" · ");
      const daily = data.daily;
      const forecast = daily.time.map((date, index) => ({ date, min: daily.temperature_2m_min[index], max: daily.temperature_2m_max[index], pop: daily.precipitation_probability_max?.[index], conditionId: daily.weather_code[index] }));
      weatherMode = "real";
      $("#weatherNotice").innerHTML = `<strong>Open-Meteo 예보</strong><span>검색한 지역 기준 · 예보 기반 현재값 · <a href="https://open-meteo.com/" target="_blank" rel="noopener noreferrer">출처</a></span>`;
      $("#weatherPlace").textContent = placeName;
      $("#weatherTemp").textContent = `${Math.round(current.temperature_2m)}°`;
      $("#weatherDesc").textContent = `${weatherFor(current.weather_code)[0]} · 체감 ${Math.round(current.apparent_temperature)}°`;
      $("#weatherIcon").textContent = weatherFor(current.weather_code)[1];
      $("#weatherHumidity").textContent = `습도 ${Math.round(current.relative_humidity_2m)}%`;
      $("#weatherWind").textContent = `바람 ${fmt(current.wind_speed_10m)} km/h`;
      $("#weatherFetched").textContent = formatWeatherTime(data.current.time, data.timezone);
      renderForecast(forecast);
      $("#weatherModeButton").textContent = "위치 변경";
      $("#weatherSearchStatus").textContent = `현재 표시 위치: ${placeName}. 실제 점포 위치와의 일치 여부는 사용자가 확인해 주세요.`;
      $("#weatherSearchForm").hidden = true;
      localSave(LOCATION_KEY, { ...place, displayName: placeName });
      toast(`${placeName} 날씨와 5일 예보를 불러왔습니다.`);
    } catch (error) {
      toast(`날씨를 불러오지 못했습니다. ${error.message}`);
    } finally {
      button.disabled = false;
      button.textContent = "선택 위치 날씨 불러오기";
    }
  }

  function makeReportRow(row) {
    const revenue = number(row.매출_만원), goal = number(row.목표_만원), cost = number(row.비용_만원);
    const values = {
      record_id: row.record_id, 채널: row.채널, 기준주: row.기준주, 점포: row.점포, 브랜드: row.브랜드, 장르: row.장르,
      매출_만원: revenue, 목표_만원: goal, 목표차이_만원: revenue - goal,
      달성률: goal === 0 ? "" : revenue / goal,
      비용_만원: cost, 비용차감액_만원: revenue - cost, 상태: row.상태, 메모: memos[row.record_id] || "", 출처: row.출처, 기준일: row.기준일,
    };
    return reportHeaders.map((header) => values[header] ?? "");
  }
  function buildReportWorkbook(rows = filteredSales) {
    if (!reportTemplateReady) throw new Error("04_보고양식.xlsx의 원본 헤더를 불러온 뒤 저장할 수 있습니다.");
    const aoa = [reportHeaders, ...rows.map(makeReportRow)];
    const worksheet = XLSX.utils.aoa_to_sheet(aoa);
    const rateColumn = reportHeaders.indexOf("달성률");
    if (rateColumn >= 0) {
      for (let i = 1; i < aoa.length; i++) {
        const value = aoa[i][rateColumn];
        const cell = XLSX.utils.encode_cell({ r: i, c: rateColumn });
        if (value === "") delete worksheet[cell];
        else worksheet[cell].z = "0.0%";
      }
    }
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "운영회의보고");
    return workbook;
  }
  function exportReport() {
    if (!filteredSales.length) { toast("저장할 실적이 없습니다."); return; }
    if (!window.XLSX) { toast("엑셀 저장 모듈을 사용할 수 없습니다."); return; }
    let workbook;
    try { workbook = buildReportWorkbook(); }
    catch (error) { toast(error.message, 6500); return; }
    const filename = `운영회의_실적보고_${new Intl.DateTimeFormat("en-CA", { timeZone: "Asia/Seoul" }).format(new Date())}.xlsx`;
    XLSX.writeFile(workbook, filename);
    toast(`현재 필터의 ${filteredSales.length}개 행을 보고서로 저장했습니다. 메모는 이 기기에만 저장됩니다.`);
  }
  async function handleFileUpload(fileList) {
    const files = Array.from(fileList || []);
    if (!files.length) return;
    let accepted = 0;
    const errors = [];
    for (const file of files) {
      const source = sourceByName(file.name);
      try {
        const rows = await parseSalesFile(file, source);
        if (!source) {
          errors.push(`${file.name}: 이 파일은 세 실적 원본 중 하나로 인식되지 않았습니다. 누락 열이 있으면 함께 확인하세요.`);
          continue;
        }
        ingestRows(source, rows);
        accepted++;
      } catch (error) {
        if (source) sources[source.key].status = `오류: ${error.message}`;
        errors.push(`${file.name}: ${error.message}`);
      }
    }
    updateSalesFilters();
    updateSourceStatus();
    renderSales();
    if (errors.length) toast(`${accepted ? `${accepted}개 파일 반영 · ` : ""}${errors.join(" / ")} 기존 정상 자료는 유지했습니다.`, 6500);
    else if (accepted) toast(`${accepted}개 파일을 확인해 반영했습니다. 총 ${allRows().length}개 행입니다.`);
    $("#fileInput").value = "";
  }
  function openSourceDialog() {
    const details = [
      ...SALES_SOURCES.map((source) => `<div class="source-detail"><strong>${esc(source.filename)} · ${sources[source.key].rows.length}행</strong><p>${esc(source.channel)} 채널. 기준주·점포·브랜드 등 필수 12개 열을 이름으로 맞춥니다.</p></div>`),
      `<div class="source-detail"><strong>04_보고양식.xlsx</strong><p>보고서 열 순서 기준으로만 사용합니다. 결과는 새 엑셀 파일로 다운로드합니다.</p></div>`,
      `<div class="source-detail"><strong>08_행사준비.csv · ${checklist.length}건</strong><p>점검 기준일은 2026-10-20이며 2026-10-22 재점검을 선택할 수 있습니다.</p></div>`,
      `<div class="source-detail"><strong>11_날씨응답_예제.json</strong><p>교육용 예제이며 실제 관측·예보가 아닙니다. 예제A점의 실제 위치 정보도 아닙니다.</p></div>`,
    ];
    $("#sourceDetailList").innerHTML = details.join("");
    $("#dataDialog").showModal();
  }
  function bindEvents() {
    $$('[data-view]').forEach((button) => button.addEventListener("click", (event) => {
      event.preventDefault();
      const view = button.dataset.view;
      $$('[data-view-panel]').forEach((panel) => { panel.hidden = panel.dataset.viewPanel !== view; });
      $$('.main-nav [data-view], .sidebar-nav [data-view], nav [data-view]').forEach((link) => {
        const active = link.dataset.view === view;
        link.classList.toggle("active", active);
        if (active) link.setAttribute("aria-current", "page"); else link.removeAttribute("aria-current");
      });
    }));
    ["#storeFilter", "#channelFilter", "#weekFilter", "#brandFilter", "#genreFilter"].forEach((selector) => $(selector).addEventListener("change", renderSales));
    $("#resetFilters").addEventListener("click", () => {
      ["#storeFilter", "#channelFilter", "#weekFilter", "#brandFilter", "#genreFilter"].forEach((selector) => { $(selector).value = "all"; });
      $("#recordSearch").value = ""; renderSales();
    });
    $("#recordSearch").addEventListener("input", () => { page = 1; renderRecordTable(); });
    $("#pageSize").addEventListener("change", () => { page = 1; renderRecordTable(); });
    $("#prevPage").addEventListener("click", () => { page--; renderRecordTable(); });
    $("#nextPage").addEventListener("click", () => { page++; renderRecordTable(); });
    $("#recordTableBody").addEventListener("input", (event) => {
      const input = event.target.closest("input[data-record-id]");
      if (!input) return;
      const id = input.dataset.recordId;
      memos[id] = input.value;
      clearTimeout(memoTimers.get(id));
      memoTimers.set(id, setTimeout(() => { localSave(MEMO_KEY, memos); toast("메모를 이 브라우저에 저장했습니다.", 1300); }, 400));
    });
    $("#recordTableBody").addEventListener("change", (event) => {
      const input = event.target.closest("input[data-record-id]");
      if (!input) return;
      clearTimeout(memoTimers.get(input.dataset.recordId));
      localSave(MEMO_KEY, memos);
    });
    $("#exportButton").addEventListener("click", exportReport);
    $("#fileInput").addEventListener("change", (event) => handleFileUpload(event.target.files));
    $("#criterionFilter").addEventListener("change", renderTasks);
    $("#areaFilter").addEventListener("change", renderTasks);
    $("#taskStatusFilter").addEventListener("change", renderTasks);
    $("#weatherModeButton").addEventListener("click", () => { $("#weatherSearchForm").hidden = !$("#weatherSearchForm").hidden; });
    $("#weatherSearchForm").addEventListener("submit", async (event) => {
      event.preventDefault();
      const query = $("#weatherSearch").value.trim();
      if (query.length < 2) { toast("지역 이름을 두 글자 이상 입력해 주세요."); return; }
      $("#weatherSearchStatus").textContent = "위치를 검색하고 있습니다…";
      try { await searchWeatherLocations(query); } catch (error) { $("#weatherSearchStatus").textContent = error.message; toast(error.message); }
    });
    $("#loadRealWeather").addEventListener("click", () => loadRealWeather());
    $("#loadSampleWeather").addEventListener("click", loadSampleWeather);
    $("#showDataSources").addEventListener("click", openSourceDialog);
    $("#manageFiles").addEventListener("click", openSourceDialog);
    $("#closeDataDialog").addEventListener("click", () => $("#dataDialog").close());
    $("#dialogCloseBottom").addEventListener("click", () => $("#dataDialog").close());
    $("#dialogUpload").addEventListener("click", () => { $("#dataDialog").close(); $("#fileInput").click(); });
  }

  async function init() {
    setTodayLabel();
    bindEvents();
    await Promise.all([loadBundledSales(), loadReportTemplate(), loadChecklist(), loadSampleWeather()]);
    renderSales();
  }
  window.SalesTestAPI = {
    parseSalesFile, ingestRows, allRows, sourceByName, normalizeHeader, filterSales, summarizeSales,
    makeReportRow, buildReportWorkbook, exportReport, handleFileUpload, renderSales,
    loadReportTemplate, loadBundledSales, init,
    getSources: () => sources,
    getFilteredRows: () => filteredSales,
    getReportHeaders: () => [...reportHeaders],
    setMemo: (id, value) => { memos[id] = value; localSave(MEMO_KEY, memos); },
    setFilters: (filters = {}) => {
      $("#storeFilter").value = filters.store || "all";
      $("#channelFilter").value = filters.channel || "all";
      $("#weekFilter").value = filters.week || "all";
      $("#brandFilter").value = filters.brand || "all";
      $("#genreFilter").value = filters.genre || "all";
      renderSales();
    },
  };
  document.addEventListener("DOMContentLoaded", init);
})();
