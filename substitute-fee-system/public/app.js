// Phase 5 管理介面：純 vanilla JS，沒有建置流程，呼叫 /api/* 取得/送出資料。
// 目標是「行政人員容易操作」，不是視覺華麗。

const WEEKDAY_LABEL = { MON: "星期一", TUE: "星期二", WED: "星期三", THU: "星期四", FRI: "星期五" };
const CLASSIFICATION_LABEL = { GENERAL: "一般公費", OVERTIME: "超鐘點", PROJECT: "專案" };
const STAFF_TYPE_LABEL = { BD: "編制內(BD)", NON_BD: "編制外(非BD)", UNKNOWN: "未指定" };
const FUNDING_SOURCE_LABEL = { GENERAL: "一般公費", OVERTIME: "超鐘點", PROJECT: "專案", UNDETERMINED: "待確認" };
const CLASSIFICATION_METHOD_LABEL = {
  WEEKLY_RULE: "週規則",
  DATE_EXCEPTION: "單日例外",
  GENERAL_DEFAULT: "無規則(一般公費)",
  MANUAL_OVERRIDE: "人工覆寫",
  CONFLICT: "規則衝突",
  TEACHER_UNMATCHED: "原教師未配對",
};

const WEEKDAY_ORDER = ["MON", "TUE", "WED", "THU", "FRI", "SAT", "SUN"];
const WEEKDAY_LABEL_SHORT = { MON: "一", TUE: "二", WED: "三", THU: "四", FRI: "五", SAT: "六", SUN: "日" };

const PENDING_ISSUE_TYPE_LABEL = {
  TEACHER_UNMATCHED: "原教師未配對",
  CONFLICT: "規則衝突",
  AMOUNT_MISSING: "金額無法計算",
  IMPORT_ERROR: "匯入錯誤",
};

const RECONCILIATION_STATUS_LABEL = {
  MATCH: "✅ 一致",
  SYSTEM_LESS: "⬇️ 系統較少",
  SYSTEM_MORE: "⬆️ 系統較多",
  ONLY_SYSTEM: "僅系統有",
  ONLY_ORIGINAL: "僅原始有",
  UNCERTAIN: "待確認",
};

const PAGE_TITLES = {
  dashboard: "🏠 月結首頁",
  importExcel: "📥 公費代課匯入",
  classification: "👥 分類預覽",
  pendingIssues: "⚠️ 待處理",
  feeCalculation: "💰 費用計算",
  selfFunded: "💵 自費代課",
  reconciliation: "📊 對帳",
  chuna: "🧾 給出納",
  feeRules: "⚙️ 基本設定／費用規則",
  projects: "⚙️ 基本設定／專案",
  weekly: "⚙️ 基本設定／每週固定規則",
  dateRules: "⚙️ 基本設定／單日例外",
  calendar: "⚙️ 基本設定／學校上課日曆",
};

// 統一的狀態徽章，取代裸露的英文 enum（例如 TEACHER_UNMATCHED、CONFLICT），
// 讓行政人員一眼看懂目前狀態，不用先學系統內部的術語。
function badge(type, text) {
  return `<span class="badge badge-${type}">${text}</span>`;
}

// Implementation Batch：月結首頁／待處理／自費代課／給出納／對帳共用同一組「學期＋年月」，
// 這五個分頁本質上都是在看／處理同一個月結期間的狀態，所以用同一個 state 欄位，
// 在任一分頁切換年月時，其餘分頁的選單也一併同步。
const CLOSE_PERIOD_SELECT_IDS = ["closePeriodSelect", "pendingPeriodSelect", "selfFundedPeriodSelect", "chunaPeriodSelect", "reconciliationPeriodSelect"];

const state = {
  semesterId: null,
  semesters: [],
  changedBy: localStorage.getItem("changedBy") || "",
  ruleType: "OVERTIME",
  periodSlots: [],
  projects: [],
  calendarYear: null,
  calendarMonth: null,
  importPeriod: null, // { year, month }
  currentImportId: null,
  classificationBatchId: null,
  feeCalcPeriod: null, // { year, month }
  closePeriod: null, // { year, month }，月結首頁／待處理／自費代課／給出納／對帳共用
  chunaBatchIds: [],
  pendingFilter: "ALL",
  pendingIssuesCache: [],
};

const FEE_TYPE_LABEL = {
  SUBSTITUTE_PERIOD: "一般公費／專案代課鐘點",
  OVERTIME_PERIOD: "超鐘點",
};

function getCurrentSemester() {
  return state.semesters.find((s) => s.id === state.semesterId) || null;
}

async function api(path, options) {
  const res = await fetch(path, {
    headers: { "Content-Type": "application/json" },
    ...options,
  });
  const body = await res.json().catch(() => ({}));
  if (!res.ok) {
    throw new Error(body.error || `請求失敗（${res.status}）`);
  }
  return body;
}

function qs(sel) { return document.querySelector(sel); }
function qsa(sel) { return Array.from(document.querySelectorAll(sel)); }

function dateOnly(iso) { return iso ? iso.slice(0, 10) : ""; }

// ---------- 初始化 ----------

async function init() {
  qs("#changedByInput").value = state.changedBy;
  qs("#changedByInput").addEventListener("change", (e) => {
    state.changedBy = e.target.value.trim();
    localStorage.setItem("changedBy", state.changedBy);
  });

  qsa(".tab-btn").forEach((btn) =>
    btn.addEventListener("click", () => {
      switchTab(btn.dataset.tab);
      const loader = TAB_LOADERS[btn.dataset.tab];
      if (loader) loader();
    })
  );

  qs("#settingsToggle").addEventListener("click", () => {
    qs("#settingsToggle").classList.toggle("collapsed");
    qs("#settingsSubgroup").classList.toggle("collapsed");
  });
  qs("#menuToggle").addEventListener("click", () => {
    qs("#sidebar").classList.toggle("open");
  });
  qsa(".subtab-btn").forEach((btn) =>
    btn.addEventListener("click", () => {
      state.ruleType = btn.dataset.ruletype;
      qsa(".subtab-btn").forEach((b) => b.classList.toggle("active", b === btn));
      loadWeeklyRules();
    })
  );

  qs("#semesterSelect").addEventListener("change", async (e) => {
    state.semesterId = e.target.value;
    updateSemesterBadge();
    await loadForSemester();
  });

  qsa("#pendingFilterTabs .subtab-btn").forEach((btn) =>
    btn.addEventListener("click", () => {
      state.pendingFilter = btn.dataset.filter;
      qsa("#pendingFilterTabs .subtab-btn").forEach((b) => b.classList.toggle("active", b === btn));
      renderPendingIssues();
    })
  );

  qs("#filterTeacher").addEventListener("input", debounce(loadWeeklyRules, 300));
  qs("#filterWeekday").addEventListener("change", loadWeeklyRules);
  qs("#filterProject").addEventListener("change", loadWeeklyRules);
  qs("#btnNewWeeklyRule").addEventListener("click", () => openWeeklyRuleForm());

  qs("#btnNewProject").addEventListener("click", () => openProjectForm());

  qs("#filterIncludeCancelled").addEventListener("change", loadDateRules);
  qs("#btnNewDateRule").addEventListener("click", () => openDateRuleForm());

  qs("#calendarMonthSelect").addEventListener("change", (e) => {
    const [year, month] = e.target.value.split("-").map(Number);
    state.calendarYear = year;
    state.calendarMonth = month;
    loadCalendar();
  });
  qs("#btnGenerateCalendar").addEventListener("click", async () => {
    const result = await api("/api/calendar/generate", {
      method: "POST",
      body: JSON.stringify({ semesterId: state.semesterId, changedBy: state.changedBy || undefined }),
    });
    alert(`已產生 ${result.createdCount} 天（略過已存在 ${result.skippedCount} 天）`);
    await setupCalendarMonthOptions();
    await loadCalendar();
  });

  qs("#btnInspectFile").addEventListener("click", inspectImportFile);
  qs("#btnUploadImport").addEventListener("click", uploadImportFile);
  qs("#btnAutoApplyMatches").addEventListener("click", async () => {
    if (!state.currentImportId) return;
    const result = await api(`/api/monthly-imports/${state.currentImportId}/auto-apply-matches`, {
      method: "POST",
      body: JSON.stringify({ changedBy: state.changedBy || undefined }),
    });
    alert(`已套用 ${result.appliedCount} 筆，剩餘待處理 ${result.remainingCount} 筆`);
    await loadUnmatchedForImport(state.currentImportId);
  });

  qs("#classificationBatchSelect").addEventListener("change", (e) => {
    state.classificationBatchId = e.target.value;
    loadClassificationPreview();
  });
  qs("#btnRunClassification").addEventListener("click", async () => {
    if (!state.classificationBatchId) {
      alert("請先選擇要分類的匯入批次");
      return;
    }
    const summary = await api(`/api/monthly-imports/${state.classificationBatchId}/classify`, {
      method: "POST",
      body: JSON.stringify({ changedBy: state.changedBy || undefined }),
    });
    const summaryEl = qs("#classificationSummary");
    summaryEl.hidden = false;
    summaryEl.textContent =
      `共 ${summary.total} 筆｜一般公費 ${summary.general}｜超鐘點 ${summary.overtime}｜專案 ${summary.project}｜` +
      `規則衝突 ${summary.conflict}｜原教師未配對 ${summary.teacherUnmatched}｜維持人工覆寫 ${summary.manualPreserved}`;
    await loadClassificationPreview();
  });
  ["filterFundingSource", "filterClassificationMethod", "filterManualOverride", "filterStaffType"].forEach((id) =>
    qs(`#${id}`).addEventListener("change", loadClassificationPreview)
  );

  qs("#feeRuleTypeSelect").addEventListener("change", loadFeeRules);
  qs("#btnNewFeeRule").addEventListener("click", () => openFeeRuleForm());

  qs("#feeCalcPeriodSelect").addEventListener("change", (e) => {
    const [year, month] = e.target.value.split("-").map(Number);
    state.feeCalcPeriod = { year, month };
  });
  qs("#btnRunFeeCalculation").addEventListener("click", runFeeCalculation);

  qs("#btnRefreshDashboard").addEventListener("click", loadDashboardAndHistory);
  qs("#btnRefreshPending").addEventListener("click", loadPendingIssues);
  qs("#btnNewSelfFunded").addEventListener("click", () => openSelfFundedForm());
  qs("#btnLoadChuna").addEventListener("click", loadChuna);
  qs("#btnExportChuna").addEventListener("click", exportChuna);
  qs("#btnRunReconciliation").addEventListener("click", runReconciliation);

  const semesters = await api("/api/semesters");
  state.semesters = semesters;
  const select = qs("#semesterSelect");
  select.innerHTML = semesters
    .map((s) => `<option value="${s.id}">${s.schoolYear}學年度第${s.term}學期${s.isCurrent ? "（使用中）" : ""}</option>`)
    .join("");
  const current = semesters.find((s) => s.isCurrent) || semesters[0];
  if (current) {
    select.value = current.id;
    state.semesterId = current.id;
  }
  updateSemesterBadge();

  state.periodSlots = await api("/api/period-slots");

  await loadForSemester();
}

// 側邊欄上方的學期／年月標示：一進系統就先讓使用者確認「我現在在看哪個學期、現在是幾月」。
function updateSemesterBadge() {
  const s = getCurrentSemester();
  const el = qs("#semesterBadge");
  const now = new Date();
  const todayLabel = `${now.getFullYear()}年${now.getMonth() + 1}月`;
  el.textContent = s ? `${s.schoolYear}學年度第${s.term}學期｜${todayLabel}` : `尚未選擇學期｜${todayLabel}`;
}

function debounce(fn, ms) {
  let timer;
  return (...args) => {
    clearTimeout(timer);
    timer = setTimeout(() => fn(...args), ms);
  };
}

function switchTab(tab) {
  qsa(".tab-btn").forEach((b) => b.classList.toggle("active", b.dataset.tab === tab));
  qsa(".tab-panel").forEach((p) => p.classList.toggle("active", p.id === `tab-${tab}`));
  const title = PAGE_TITLES[tab];
  if (title) qs("#pageTitle").textContent = title;
  qs("#sidebar").classList.remove("open"); // 手機版：切換分頁後自動收起側邊欄選單
}

function loadDashboardAndHistory() {
  loadDashboard();
  loadHistoryTable();
}

const TAB_LOADERS = {
  dashboard: loadDashboardAndHistory,
  pendingIssues: loadPendingIssues,
  selfFunded: loadSelfFunded,
};

async function loadForSemester() {
  if (!state.semesterId) return;
  state.projects = await api(`/api/projects?semesterId=${state.semesterId}`);
  const projectFilter = qs("#filterProject");
  projectFilter.innerHTML =
    `<option value="">全部專案</option>` + state.projects.map((p) => `<option value="${p.id}">${p.name}</option>`).join("");
  await setupCalendarMonthOptions();
  setupImportPeriodOptions();
  await Promise.all([loadWeeklyRules(), loadProjects(), loadDateRules(), loadCalendar(), loadImportBatches()]);
  await setupClassificationBatchOptions();
  await loadFeeRules();
  setupFeeCalcPeriodOptions();
  setupClosePeriodOptions();
  // 月結首頁是進站後第一眼看到的畫面，資料要在初始載入時就準備好，不能等使用者手動點分頁。
  await loadDashboard();
  await loadHistoryTable();
}

// ---------- 每週固定規則 ----------

async function loadWeeklyRules() {
  if (!state.semesterId) return;
  const params = new URLSearchParams({ semesterId: state.semesterId, ruleType: state.ruleType });
  const teacher = qs("#filterTeacher").value.trim();
  const weekday = qs("#filterWeekday").value;
  const projectId = qs("#filterProject").value;
  if (weekday) params.set("weekday", weekday);
  if (projectId) params.set("projectId", projectId);

  const [rules, conflicts] = await Promise.all([
    api(`/api/weekly-rules?${params.toString()}`),
    api(`/api/weekly-rules/conflicts?semesterId=${state.semesterId}`),
  ]);

  const conflictRuleIds = new Set(conflicts.flatMap((c) => c.ruleIds));
  const filtered = teacher ? rules.filter((r) => r.person.name.includes(teacher)) : rules;

  const banner = qs("#conflictBanner");
  if (conflicts.length > 0) {
    banner.hidden = false;
    banner.textContent = `⚠️ 偵測到 ${conflicts.length} 組規則衝突（同一教師、同一星期、同一節次，期間重疊），請確認下方標紅列。`;
  } else {
    banner.hidden = true;
  }

  const today = new Date().toISOString().slice(0, 10);
  const tbody = qs("#weeklyRuleTable tbody");
  tbody.innerHTML = filtered
    .map((r) => {
      const isConflict = conflictRuleIds.has(r.id);
      const isInactive = r.endDate && dateOnly(r.endDate) < today;
      const status = isInactive ? "已停用/已過期" : "生效中";
      return `
        <tr class="${isConflict ? "conflict-row" : ""} ${isInactive ? "inactive-row" : ""}" data-id="${r.id}">
          <td>${r.person.name}</td>
          <td>${WEEKDAY_LABEL[r.weekday]}</td>
          <td>${r.periodCode}</td>
          <td>${r.subject ?? ""}</td>
          <td>${r.ruleType === "OVERTIME" ? "超鐘點" : "專案"}</td>
          <td>${r.project ? r.project.name : ""}</td>
          <td>${r.weeklyPeriods}</td>
          <td>${dateOnly(r.effectiveDate)}</td>
          <td>${dateOnly(r.endDate) || "－"}</td>
          <td>${isConflict ? "⚠️ 衝突 / " : ""}${status}</td>
          <td>
            <button data-action="edit">編輯</button>
            <button data-action="deactivate" class="danger">停用</button>
          </td>
        </tr>`;
    })
    .join("");

  tbody.querySelectorAll("button[data-action='edit']").forEach((btn) =>
    btn.addEventListener("click", (e) => {
      const id = e.target.closest("tr").dataset.id;
      const rule = rules.find((r) => r.id === id);
      openWeeklyRuleForm(rule);
    })
  );
  tbody.querySelectorAll("button[data-action='deactivate']").forEach((btn) =>
    btn.addEventListener("click", (e) => {
      const id = e.target.closest("tr").dataset.id;
      openDeactivateWeeklyRuleForm(id);
    })
  );
}

function periodOptions(selected) {
  return state.periodSlots
    .map((p) => `<option value="${p.code}" ${p.code === selected ? "selected" : ""}>${p.displayName}</option>`)
    .join("");
}

function projectOptions(selected) {
  return (
    `<option value="">（請選擇專案）</option>` +
    state.projects.map((p) => `<option value="${p.id}" ${p.id === selected ? "selected" : ""}>${p.name}</option>`).join("")
  );
}

function openWeeklyRuleForm(existing) {
  const ruleType = existing ? existing.ruleType : state.ruleType;
  const isProject = ruleType === "PROJECT";
  const body = `
    <h2>${existing ? "編輯" : "新增"}${isProject ? "專案減課" : "超鐘點"}規則</h2>
    <label>教師姓名（新增規則時請填完整姓名）
      <input id="f-personName" type="text" value="${existing ? existing.person.name : ""}" ${existing ? "disabled" : ""} />
    </label>
    ${isProject ? `<label>專案<select id="f-projectId">${projectOptions(existing?.projectId)}</select></label>` : ""}
    <label>星期<select id="f-weekday">
      ${Object.entries(WEEKDAY_LABEL).map(([v, l]) => `<option value="${v}" ${existing?.weekday === v ? "selected" : ""}>${l}</option>`).join("")}
    </select></label>
    <label>節次<select id="f-periodCode">${periodOptions(existing?.periodCode)}</select></label>
    <label>科目（選填，是否需要比對科目由系統設定決定，這裡先保存即可）
      <input id="f-subject" type="text" value="${existing?.subject ?? ""}" />
    </label>
    <label>每週節數<input id="f-weeklyPeriods" type="number" min="0" step="1" value="${existing?.weeklyPeriods ?? 1}" /></label>
    <label>生效日期<input id="f-effectiveDate" type="date" value="${dateOnly(existing?.effectiveDate) || ""}" /></label>
    <label>結束日期（選填）<input id="f-endDate" type="date" value="${dateOnly(existing?.endDate) || ""}" /></label>
    <label>備註<textarea id="f-note">${existing?.note ?? ""}</textarea></label>
    <div class="error-text" id="f-error" hidden></div>
    <div class="modal-actions">
      <button type="button" class="secondary" id="f-cancel">取消</button>
      <button type="button" id="f-submit">儲存</button>
    </div>`;
  showModal(body);

  qs("#f-cancel").addEventListener("click", closeModal);
  qs("#f-submit").addEventListener("click", async () => {
    try {
      let personId = existing?.personId;
      if (!existing) {
        const name = qs("#f-personName").value.trim();
        if (!name) throw new Error("請輸入教師姓名");
        const matches = await api(`/api/persons?search=${encodeURIComponent(name)}`);
        const exact = matches.find((p) => p.name === name);
        if (!exact) {
          throw new Error(`系統找不到教師「${name}」，請先在人員管理建立此人員（Phase 5 不在此處自動建立人員）`);
        }
        personId = exact.id;
      }

      const payload = {
        semesterId: state.semesterId,
        personId,
        ruleType,
        projectId: isProject ? qs("#f-projectId").value : undefined,
        weekday: qs("#f-weekday").value,
        periodCode: qs("#f-periodCode").value,
        subject: qs("#f-subject").value.trim() || undefined,
        weeklyPeriods: Number(qs("#f-weeklyPeriods").value),
        effectiveDate: qs("#f-effectiveDate").value,
        endDate: qs("#f-endDate").value || undefined,
        note: qs("#f-note").value.trim() || undefined,
        changedBy: state.changedBy || undefined,
      };
      if (!payload.effectiveDate) throw new Error("請填寫生效日期");

      const result = existing
        ? await api(`/api/weekly-rules/${existing.id}`, { method: "PATCH", body: JSON.stringify({ ...payload, reason: "透過管理介面修改" }) })
        : await api("/api/weekly-rules", { method: "POST", body: JSON.stringify(payload) });

      closeModal();
      await loadWeeklyRules();
      if (result.conflicts && result.conflicts.length > 0) {
        alert(`⚠️ 這筆規則與其他規則衝突，請在列表中確認標紅的資料列並處理。`);
      }
    } catch (err) {
      showFormError(err.message);
    }
  });
}

function openDeactivateWeeklyRuleForm(ruleId) {
  const body = `
    <h2>停用規則</h2>
    <p>停用不會刪除資料，歷史仍可查詢。</p>
    <label>停用生效日（此日之後不再視為生效）<input id="f-endDate" type="date" value="${new Date().toISOString().slice(0, 10)}" /></label>
    <label>原因<input id="f-reason" type="text" placeholder="例如：學期結束、規則異動" /></label>
    <div class="error-text" id="f-error" hidden></div>
    <div class="modal-actions">
      <button type="button" class="secondary" id="f-cancel">取消</button>
      <button type="button" id="f-submit">確認停用</button>
    </div>`;
  showModal(body);
  qs("#f-cancel").addEventListener("click", closeModal);
  qs("#f-submit").addEventListener("click", async () => {
    try {
      await api(`/api/weekly-rules/${ruleId}/deactivate`, {
        method: "POST",
        body: JSON.stringify({ endDate: qs("#f-endDate").value, reason: qs("#f-reason").value.trim(), changedBy: state.changedBy || undefined }),
      });
      closeModal();
      await loadWeeklyRules();
    } catch (err) {
      showFormError(err.message);
    }
  });
}

// ---------- 專案 ----------

async function loadProjects() {
  if (!state.semesterId) return;
  const projects = await api(`/api/projects?semesterId=${state.semesterId}`);
  const tbody = qs("#projectTable tbody");
  tbody.innerHTML = projects
    .map(
      (p) => `
      <tr class="${p.isActive ? "" : "inactive-row"}" data-id="${p.id}">
        <td>${p.name}</td>
        <td>${p.isActive ? "啟用中" : "已停用"}</td>
        <td>${p.note ?? ""}</td>
        <td>
          <button data-action="edit">編輯</button>
          <button data-action="toggle" class="${p.isActive ? "danger" : ""}">${p.isActive ? "停用" : "啟用"}</button>
        </td>
      </tr>`
    )
    .join("");

  tbody.querySelectorAll("button[data-action='edit']").forEach((btn) =>
    btn.addEventListener("click", (e) => {
      const id = e.target.closest("tr").dataset.id;
      openProjectForm(projects.find((p) => p.id === id));
    })
  );
  tbody.querySelectorAll("button[data-action='toggle']").forEach((btn) =>
    btn.addEventListener("click", async (e) => {
      const id = e.target.closest("tr").dataset.id;
      const proj = projects.find((p) => p.id === id);
      await api(`/api/projects/${id}/active`, {
        method: "PATCH",
        body: JSON.stringify({ isActive: !proj.isActive, changedBy: state.changedBy || undefined }),
      });
      await loadForSemester();
    })
  );
}

function openProjectForm(existing) {
  const body = `
    <h2>${existing ? "編輯" : "新增"}專案</h2>
    <label>專案名稱<input id="f-name" type="text" value="${existing?.name ?? ""}" /></label>
    <label>備註<textarea id="f-note">${existing?.note ?? ""}</textarea></label>
    <div class="error-text" id="f-error" hidden></div>
    <div class="modal-actions">
      <button type="button" class="secondary" id="f-cancel">取消</button>
      <button type="button" id="f-submit">儲存</button>
    </div>`;
  showModal(body);
  qs("#f-cancel").addEventListener("click", closeModal);
  qs("#f-submit").addEventListener("click", async () => {
    try {
      const name = qs("#f-name").value.trim();
      if (!name) throw new Error("請輸入專案名稱");
      const note = qs("#f-note").value.trim() || undefined;
      if (existing) {
        await api(`/api/projects/${existing.id}`, {
          method: "PATCH",
          body: JSON.stringify({ name, note, changedBy: state.changedBy || undefined, reason: "透過管理介面修改" }),
        });
      } else {
        await api("/api/projects", { method: "POST", body: JSON.stringify({ semesterId: state.semesterId, name, note, changedBy: state.changedBy || undefined }) });
      }
      closeModal();
      await loadForSemester();
    } catch (err) {
      showFormError(err.message);
    }
  });
}

// ---------- 單日例外 ----------

async function loadDateRules() {
  if (!state.semesterId) return;
  const includeCancelled = qs("#filterIncludeCancelled").checked;
  const rules = await api(`/api/date-rules?semesterId=${state.semesterId}&includeCancelled=${includeCancelled}`);
  const tbody = qs("#dateRuleTable tbody");
  tbody.innerHTML = rules
    .map(
      (r) => `
      <tr class="${r.isCancelled ? "inactive-row" : ""}" data-id="${r.id}">
        <td>${dateOnly(r.date)}</td>
        <td>${r.person.name}</td>
        <td>${r.periodCode}</td>
        <td>${CLASSIFICATION_LABEL[r.overrideClassification]}</td>
        <td>${r.project ? r.project.name : ""}</td>
        <td>${r.note ?? ""}</td>
        <td>${r.isCancelled ? "已取消" : "生效中"}</td>
        <td>${r.isCancelled ? "" : `<button data-action="cancel" class="danger">取消</button>`}</td>
      </tr>`
    )
    .join("");

  tbody.querySelectorAll("button[data-action='cancel']").forEach((btn) =>
    btn.addEventListener("click", async (e) => {
      const id = e.target.closest("tr").dataset.id;
      const reason = prompt("取消原因：") || undefined;
      await api(`/api/date-rules/${id}/cancel`, { method: "POST", body: JSON.stringify({ reason, changedBy: state.changedBy || undefined }) });
      await loadDateRules();
    })
  );
}

function openDateRuleForm() {
  const body = `
    <h2>新增單日例外</h2>
    <label>日期<input id="f-date" type="date" /></label>
    <label>教師姓名<input id="f-personName" type="text" /></label>
    <label>節次<select id="f-periodCode">${periodOptions()}</select></label>
    <label>改判為<select id="f-classification">
      <option value="GENERAL">一般公費</option>
      <option value="OVERTIME">超鐘點</option>
      <option value="PROJECT">專案</option>
    </select></label>
    <label id="f-projectWrap" hidden>專案<select id="f-projectId">${projectOptions()}</select></label>
    <label>備註<textarea id="f-note"></textarea></label>
    <div class="error-text" id="f-error" hidden></div>
    <div class="modal-actions">
      <button type="button" class="secondary" id="f-cancel">取消</button>
      <button type="button" id="f-submit">儲存</button>
    </div>`;
  showModal(body);
  qs("#f-classification").addEventListener("change", (e) => {
    qs("#f-projectWrap").hidden = e.target.value !== "PROJECT";
  });
  qs("#f-cancel").addEventListener("click", closeModal);
  qs("#f-submit").addEventListener("click", async () => {
    try {
      const name = qs("#f-personName").value.trim();
      if (!name) throw new Error("請輸入教師姓名");
      const matches = await api(`/api/persons?search=${encodeURIComponent(name)}`);
      const exact = matches.find((p) => p.name === name);
      if (!exact) throw new Error(`系統找不到教師「${name}」`);

      const classification = qs("#f-classification").value;
      await api("/api/date-rules", {
        method: "POST",
        body: JSON.stringify({
          semesterId: state.semesterId,
          date: qs("#f-date").value,
          personId: exact.id,
          periodCode: qs("#f-periodCode").value,
          overrideClassification: classification,
          projectId: classification === "PROJECT" ? qs("#f-projectId").value : undefined,
          note: qs("#f-note").value.trim() || undefined,
          changedBy: state.changedBy || undefined,
        }),
      });
      closeModal();
      await loadDateRules();
    } catch (err) {
      showFormError(err.message);
    }
  });
}

// ---------- 共用 modal ----------

function showModal(html) {
  qs("#modalBox").innerHTML = html;
  qs("#modalBackdrop").hidden = false;
}
function closeModal() {
  qs("#modalBackdrop").hidden = true;
}
function showFormError(message) {
  const el = qs("#f-error");
  if (el) {
    el.textContent = message;
    el.hidden = false;
  } else {
    alert(message);
  }
}

// ---------- 學校上課日曆 ----------

function monthRange(startIso, endIso) {
  const start = new Date(startIso);
  const end = new Date(endIso);
  const months = [];
  let y = start.getUTCFullYear();
  let m = start.getUTCMonth(); // 0-based
  const endY = end.getUTCFullYear();
  const endM = end.getUTCMonth();
  while (y < endY || (y === endY && m <= endM)) {
    months.push({ year: y, month: m + 1 });
    m += 1;
    if (m > 11) {
      m = 0;
      y += 1;
    }
  }
  return months;
}

async function setupCalendarMonthOptions() {
  const semester = getCurrentSemester();
  if (!semester) return;
  const months = monthRange(semester.startDate, semester.endDate);
  const select = qs("#calendarMonthSelect");
  select.innerHTML = months.map((m) => `<option value="${m.year}-${m.month}">${m.year}年${m.month}月</option>`).join("");

  const hasCurrent =
    state.calendarYear && months.some((m) => m.year === state.calendarYear && m.month === state.calendarMonth);
  const chosen = hasCurrent ? { year: state.calendarYear, month: state.calendarMonth } : months[0];
  if (chosen) {
    select.value = `${chosen.year}-${chosen.month}`;
    state.calendarYear = chosen.year;
    state.calendarMonth = chosen.month;
  }
}

function buildCalendarGrid(days) {
  if (days.length === 0) return [];
  const firstIndex = WEEKDAY_ORDER.indexOf(days[0].weekday);
  const cells = new Array(firstIndex).fill(null);
  for (const d of days) cells.push(d);
  while (cells.length % 7 !== 0) cells.push(null);
  return cells;
}

async function loadCalendar() {
  if (!state.semesterId || !state.calendarYear || !state.calendarMonth) return;
  const params = `semesterId=${state.semesterId}&year=${state.calendarYear}&month=${state.calendarMonth}`;
  const [days, summary] = await Promise.all([
    api(`/api/calendar?${params}`),
    api(`/api/calendar/summary?${params}`),
  ]);

  const cells = buildCalendarGrid(days);
  const tbody = qs("#calendarTable tbody");
  const rows = [];
  for (let i = 0; i < cells.length; i += 7) {
    const rowCells = cells
      .slice(i, i + 7)
      .map((day) => {
        if (!day) return `<td class="empty"></td>`;
        const dayNum = Number(day.date.slice(8, 10));
        return `
          <td data-id="${day.id}">
            <div class="day-num">${dayNum}</div>
            <div class="day-status">${day.isTeachingDay ? "🟢" : "⚪"}</div>
            ${day.note ? `<div class="day-note" title="${day.note}">${day.note}</div>` : ""}
          </td>`;
      })
      .join("");
    rows.push(`<tr>${rowCells}</tr>`);
  }
  tbody.innerHTML = rows.join("");

  tbody.querySelectorAll("td[data-id]").forEach((td) =>
    td.addEventListener("click", () => {
      const day = days.find((d) => d.id === td.dataset.id);
      openCalendarDayForm(day);
    })
  );

  const summaryEl = qs("#calendarSummary");
  const rowsHtml = summary.byWeekday
    .filter((w) => w.weekday !== "SAT" && w.weekday !== "SUN")
    .map((w) => `<tr><td>星期${WEEKDAY_LABEL_SHORT[w.weekday]}</td><td>${w.teachingDayCount}</td></tr>`)
    .join("");
  const weekendRows = summary.byWeekday
    .filter((w) => w.weekday === "SAT" || w.weekday === "SUN")
    .map((w) => `<tr><td>星期${WEEKDAY_LABEL_SHORT[w.weekday]}</td><td>${w.teachingDayCount}</td></tr>`)
    .join("");
  summaryEl.innerHTML = `
    <h3>${state.calendarYear}年${state.calendarMonth}月上課日統計</h3>
    <table>${rowsHtml}${weekendRows}<tr class="total-row"><td>總上課日</td><td>${summary.totalTeachingDays}</td></tr></table>`;
}

function openCalendarDayForm(day) {
  const body = `
    <h2>${day.date.slice(0, 10)}（星期${WEEKDAY_LABEL_SHORT[day.weekday]}）</h2>
    <label><input id="f-isTeachingDay" type="checkbox" ${day.isTeachingDay ? "checked" : ""} /> 是上課日</label>
    <label>備註<textarea id="f-note">${day.note ?? ""}</textarea></label>
    <div class="error-text" id="f-error" hidden></div>
    <div class="modal-actions">
      <button type="button" class="secondary" id="f-cancel">取消</button>
      <button type="button" id="f-submit">儲存</button>
    </div>`;
  showModal(body);
  qs("#f-cancel").addEventListener("click", closeModal);
  qs("#f-submit").addEventListener("click", async () => {
    try {
      await api(`/api/calendar/${day.id}`, {
        method: "PATCH",
        body: JSON.stringify({
          isTeachingDay: qs("#f-isTeachingDay").checked,
          note: qs("#f-note").value.trim(),
          changedBy: state.changedBy || undefined,
          reason: "透過管理介面修改上課日狀態",
        }),
      });
      closeModal();
      await loadCalendar();
    } catch (err) {
      showFormError(err.message);
    }
  });
}

// ---------- 公費代課 Excel 匯入 ----------

function setupImportPeriodOptions() {
  const semester = getCurrentSemester();
  if (!semester) return;
  const months = monthRange(semester.startDate, semester.endDate);
  const select = qs("#importPeriodSelect");
  select.innerHTML = months.map((m) => `<option value="${m.year}-${m.month}">${m.year}年${m.month}月</option>`).join("");
  const hasCurrent = state.importPeriod && months.some((m) => m.year === state.importPeriod.year && m.month === state.importPeriod.month);
  const chosen = hasCurrent ? state.importPeriod : months[0];
  if (chosen) {
    select.value = `${chosen.year}-${chosen.month}`;
    state.importPeriod = chosen;
  }
  select.onchange = () => {
    const [year, month] = select.value.split("-").map(Number);
    state.importPeriod = { year, month };
  };
}

// 步驟一：上傳前先讓管理者看到這份 Excel 有哪些工作表，選擇要匯入哪一個。
// 真實的公費代課 Excel 通常同時包含教師代碼對照表、編制內/外明細、給出納彙總等多個
// 工作表，不能假設固定是第一個，也不能把「給出納」這種已彙總結果的工作表當原始資料匯入。
async function inspectImportFile() {
  const fileInput = qs("#importFile");
  if (!fileInput.files || fileInput.files.length === 0) {
    alert("請先選擇 Excel 檔案");
    return;
  }
  if (!state.importPeriod) {
    alert("請先選擇匯入年月");
    return;
  }

  const formData = new FormData();
  formData.append("file", fileInput.files[0]);

  let sheets;
  try {
    const res = await fetch("/api/monthly-imports/inspect", { method: "POST", body: formData });
    sheets = await res.json();
    if (!res.ok) throw new Error(sheets.error || "讀取 Excel 失敗");
  } catch (err) {
    alert(err.message);
    return;
  }

  const sheetSelect = qs("#importSheetSelect");
  sheetSelect.innerHTML = sheets
    .map((s) => `<option value="${s.name}" data-suggest="${s.suggestedStaffType}">${s.name}（${s.rowCount} 列）</option>`)
    .join("");
  qs("#sheetPickerForm").hidden = false;

  const applySuggestion = () => {
    const selected = sheetSelect.selectedOptions[0];
    if (selected) qs("#importStaffTypeSelect").value = selected.dataset.suggest || "UNKNOWN";
  };
  sheetSelect.onchange = applySuggestion;
  applySuggestion();
}

// 步驟二：實際上傳並匯入指定的工作表
async function uploadImportFile() {
  const fileInput = qs("#importFile");
  const sheetName = qs("#importSheetSelect").value;
  if (!fileInput.files || fileInput.files.length === 0 || !sheetName) {
    alert("請先完成上一步的工作表選擇");
    return;
  }

  const formData = new FormData();
  formData.append("file", fileInput.files[0]);
  formData.append("semesterId", state.semesterId);
  formData.append("year", String(state.importPeriod.year));
  formData.append("month", String(state.importPeriod.month));
  formData.append("sheetName", sheetName);
  formData.append("sourceStaffType", qs("#importStaffTypeSelect").value);
  if (state.changedBy) formData.append("changedBy", state.changedBy);

  let result;
  try {
    const res = await fetch("/api/monthly-imports", { method: "POST", body: formData });
    result = await res.json();
    if (!res.ok) throw new Error(result.error || "匯入失敗");
  } catch (err) {
    alert(err.message);
    return;
  }

  state.currentImportId = result.monthlyImport.id;
  renderImportResult(result);
  await loadUnmatchedForImport(result.monthlyImport.id);
  await loadImportBatches();
}

function renderImportResult(result) {
  qs("#importResult").hidden = false;
  const supersededNote =
    result.supersededImportIds && result.supersededImportIds.length > 0
      ? `（已將 ${result.supersededImportIds.length} 個舊批次標記為已取代，資料仍保留）`
      : "";
  qs("#importSummary").textContent =
    `檔案：${result.monthlyImport.fileName}｜總筆數 ${result.totalCount}｜成功 ${result.successCount}｜錯誤 ${result.errorCount}${supersededNote}`;
  qs("#importDetectedHeaders").textContent = `偵測到的欄位：${(result.detectedHeaders || []).join("、")}`;

  const errorTable = qs("#importErrorTable");
  const errorRows = result.errors || [];
  if (errorRows.length > 0) {
    errorTable.hidden = false;
    errorTable.querySelector("tbody").innerHTML = errorRows
      .map((e) => `<tr><td>${e.rowNumber ?? ""}</td><td>${e.fieldName ?? ""}</td><td>${e.message}</td></tr>`)
      .join("");
  } else {
    errorTable.hidden = true;
  }
}

async function loadUnmatchedForImport(importId) {
  const unmatched = await api(`/api/monthly-imports/${importId}/unmatched`);
  const section = qs("#importUnmatchedSection");
  if (unmatched.length === 0) {
    section.hidden = true;
    return;
  }
  section.hidden = false;
  const tbody = qs("#importUnmatchedTable tbody");
  tbody.innerHTML = unmatched
    .map((u, idx) => {
      const options =
        `<option value="">請選擇</option>` +
        u.candidates.map((c) => `<option value="${c.id}">${c.name}</option>`).join("");
      return `
        <tr data-record-id="${u.recordId}" data-field="${u.field}">
          <td>${u.rawName}</td>
          <td>${u.field === "original" ? "原教師" : "代課教師"}</td>
          <td><select class="unmatched-select" data-idx="${idx}">${options}</select></td>
          <td><button data-action="resolve">配對</button></td>
        </tr>`;
    })
    .join("");

  tbody.querySelectorAll("button[data-action='resolve']").forEach((btn) =>
    btn.addEventListener("click", async (e) => {
      const tr = e.target.closest("tr");
      const personId = tr.querySelector(".unmatched-select").value;
      if (!personId) {
        alert("請先選擇要配對的人員（若清單是空的，請先到人員管理新增這位教師）");
        return;
      }
      await api(`/api/substitute-records/${tr.dataset.recordId}/resolve-teacher`, {
        method: "POST",
        body: JSON.stringify({ field: tr.dataset.field, personId, changedBy: state.changedBy || undefined }),
      });
      await loadUnmatchedForImport(importId);
    })
  );
}

async function loadImportBatches() {
  if (!state.semesterId) return;
  const batches = await api(`/api/monthly-imports?semesterId=${state.semesterId}`);
  const tbody = qs("#importBatchTable tbody");
  tbody.innerHTML = batches
    .map(
      (b) => `
      <tr class="${b.status === "ACTIVE" ? "status-active" : "status-superseded"}">
        <td>${b.year}年${b.month}月</td>
        <td>${STAFF_TYPE_LABEL[b.sourceStaffType] || b.sourceStaffType}</td>
        <td>v${b.versionNo}</td>
        <td>${b.fileName}${b.sourceSheetName ? `<br><small>${b.sourceSheetName}</small>` : ""}</td>
        <td>${new Date(b.importedAt).toLocaleString("zh-TW")}</td>
        <td>${b.totalCount}</td>
        <td>${b.successCount}</td>
        <td>${b.errorCount}</td>
        <td>${b.status === "ACTIVE" ? badge("ok", "🟢 生效中") : badge("idle", "⚪ 已取代")}</td>
        <td><button data-id="${b.id}" data-action="view">查看</button></td>
      </tr>`
    )
    .join("");

  tbody.querySelectorAll("button[data-action='view']").forEach((btn) =>
    btn.addEventListener("click", async () => {
      const detail = await api(`/api/monthly-imports/${btn.dataset.id}`);
      state.currentImportId = detail.id;
      renderImportResult({
        monthlyImport: detail,
        totalCount: detail.totalCount,
        successCount: detail.successCount,
        errorCount: detail.errorCount,
        errors: detail.errors,
        detectedHeaders: [],
        supersededImportIds: [],
      });
      qs("#importDetectedHeaders").textContent = "（歷史批次不重新顯示欄位偵測結果，僅顯示錯誤明細）";
      await loadUnmatchedForImport(detail.id);
    })
  );
}

// ---------- Phase 8：分類預覽 ----------

async function setupClassificationBatchOptions() {
  if (!state.semesterId) return;
  const batches = await api(`/api/monthly-imports?semesterId=${state.semesterId}`);
  const select = qs("#classificationBatchSelect");
  select.innerHTML = batches
    .map(
      (b) =>
        `<option value="${b.id}">${b.year}年${b.month}月 ${STAFF_TYPE_LABEL[b.sourceStaffType] || b.sourceStaffType} v${b.versionNo}${b.status === "SUPERSEDED" ? "（已取代）" : ""}</option>`
    )
    .join("");
  const hasCurrent = state.classificationBatchId && batches.some((b) => b.id === state.classificationBatchId);
  if (!hasCurrent) {
    state.classificationBatchId = batches[0]?.id ?? null;
  }
  if (state.classificationBatchId) select.value = state.classificationBatchId;
  await loadClassificationPreview();
}

async function loadClassificationPreview() {
  if (!state.classificationBatchId) {
    qs("#classificationTable tbody").innerHTML = "";
    return;
  }
  const params = new URLSearchParams();
  const fundingSource = qs("#filterFundingSource").value;
  const classificationMethod = qs("#filterClassificationMethod").value;
  const isManuallyModified = qs("#filterManualOverride").value;
  const staffType = qs("#filterStaffType").value;
  if (fundingSource) params.set("fundingSource", fundingSource);
  if (classificationMethod) params.set("classificationMethod", classificationMethod);
  if (isManuallyModified) params.set("isManuallyModified", isManuallyModified);
  if (staffType) params.set("staffType", staffType);

  const records = await api(`/api/monthly-imports/${state.classificationBatchId}/classification-preview?${params.toString()}`);
  const tbody = qs("#classificationTable tbody");
  tbody.innerHTML = records
    .map((r) => {
      let status;
      if (r.isManuallyModified) status = badge("ok", "✅ 已人工覆寫");
      else if (r.classificationMethod === "CONFLICT") status = badge("blocked", "🔴 分類規則衝突");
      else if (r.classificationMethod === "TEACHER_UNMATCHED") status = badge("blocked", "🔴 教師尚未配對");
      else status = badge("idle", "系統判斷");

      return `
        <tr class="${r.classificationMethod === "CONFLICT" || r.classificationMethod === "TEACHER_UNMATCHED" ? "conflict-row" : ""}" data-id="${r.id}">
          <td>${dateOnly(r.date)}</td>
          <td>${r.originalTeacher?.name ?? "（未配對）"}</td>
          <td>${r.periodCode ?? ""}</td>
          <td>${r.className ?? ""}</td>
          <td>${r.subject ?? ""}</td>
          <td>${r.substituteTeacher?.name ?? "（未配對）"}</td>
          <td>${STAFF_TYPE_LABEL[r.staffType] || r.staffType}</td>
          <td>${FUNDING_SOURCE_LABEL[r.fundingSource] || r.fundingSource}</td>
          <td>${CLASSIFICATION_METHOD_LABEL[r.classificationMethod] || r.classificationMethod}</td>
          <td>${r.classificationBasisText ?? ""}</td>
          <td>${r.project?.name ?? ""}</td>
          <td>${status}</td>
          <td>
            <button data-action="override">覆寫</button>
            ${r.isManuallyModified ? `<button data-action="revert" class="danger">復原</button>` : ""}
          </td>
        </tr>`;
    })
    .join("");

  tbody.querySelectorAll("button[data-action='override']").forEach((btn) =>
    btn.addEventListener("click", (e) => {
      const id = e.target.closest("tr").dataset.id;
      const record = records.find((r) => r.id === id);
      openOverrideClassificationForm(record);
    })
  );
  tbody.querySelectorAll("button[data-action='revert']").forEach((btn) =>
    btn.addEventListener("click", async (e) => {
      const id = e.target.closest("tr").dataset.id;
      const reason = prompt("復原原因：");
      if (!reason) return;
      try {
        await api(`/api/substitute-records/${id}/revert-classification`, {
          method: "POST",
          body: JSON.stringify({ changedBy: state.changedBy || undefined, reason }),
        });
        await loadClassificationPreview();
      } catch (err) {
        alert(err.message);
      }
    })
  );
}

function openOverrideClassificationForm(record) {
  const body = `
    <h2>人工覆寫分類</h2>
    <p>${dateOnly(record.date)}｜${record.originalTeacher?.name ?? "（未配對）"}｜${record.periodCode ?? ""}</p>
    <label>分類結果<select id="f-fundingSource">
      <option value="GENERAL" ${record.fundingSource === "GENERAL" ? "selected" : ""}>一般公費</option>
      <option value="OVERTIME" ${record.fundingSource === "OVERTIME" ? "selected" : ""}>超鐘點</option>
      <option value="PROJECT" ${record.fundingSource === "PROJECT" ? "selected" : ""}>專案</option>
      <option value="UNDETERMINED" ${record.fundingSource === "UNDETERMINED" ? "selected" : ""}>待確認</option>
    </select></label>
    <label id="f-projectWrap" ${record.fundingSource === "PROJECT" ? "" : "hidden"}>專案<select id="f-projectId">${projectOptions(record.projectId)}</select></label>
    <label>覆寫原因（必填）<input id="f-reason" type="text" /></label>
    <div class="error-text" id="f-error" hidden></div>
    <div class="modal-actions">
      <button type="button" class="secondary" id="f-cancel">取消</button>
      <button type="button" id="f-submit">儲存</button>
    </div>`;
  showModal(body);
  qs("#f-fundingSource").addEventListener("change", (e) => {
    qs("#f-projectWrap").hidden = e.target.value !== "PROJECT";
  });
  qs("#f-cancel").addEventListener("click", closeModal);
  qs("#f-submit").addEventListener("click", async () => {
    try {
      const fundingSource = qs("#f-fundingSource").value;
      const reason = qs("#f-reason").value.trim();
      if (!reason) throw new Error("請填寫覆寫原因");
      await api(`/api/substitute-records/${record.id}/override-classification`, {
        method: "POST",
        body: JSON.stringify({
          fundingSource,
          projectId: fundingSource === "PROJECT" ? qs("#f-projectId").value : undefined,
          changedBy: state.changedBy || undefined,
          reason,
        }),
      });
      closeModal();
      await loadClassificationPreview();
    } catch (err) {
      showFormError(err.message);
    }
  });
}

// ---------- Phase 9 第一階段：費用規則 ----------

async function loadFeeRules() {
  if (!state.semesterId) return;
  const feeType = qs("#feeRuleTypeSelect").value;
  const rules = await api(`/api/fee-rules?semesterId=${state.semesterId}&feeType=${feeType}`);
  const today = new Date().toISOString().slice(0, 10);
  const tbody = qs("#feeRuleTable tbody");
  tbody.innerHTML = rules
    .slice()
    .reverse()
    .map((r) => {
      const isInactive = r.endDate && dateOnly(r.endDate) < today;
      return `
        <tr class="${isInactive ? "inactive-row" : ""}" data-id="${r.id}">
          <td>${r.amount}</td>
          <td>${dateOnly(r.effectiveDate)}</td>
          <td>${r.endDate ? dateOnly(r.endDate) : ""}</td>
          <td>${r.note ?? ""}</td>
          <td>${isInactive ? "已停用/已過期" : "生效中"}</td>
          <td>${isInactive ? "" : `<button data-action="deactivate">停用</button>`}</td>
        </tr>`;
    })
    .join("");

  tbody.querySelectorAll("button[data-action='deactivate']").forEach((btn) =>
    btn.addEventListener("click", async () => {
      const id = btn.closest("tr").dataset.id;
      const endDate = prompt("請輸入停用日期（YYYY-MM-DD），此日期之後這筆費率就不再生效：");
      if (!endDate) return;
      try {
        await api(`/api/fee-rules/${id}/deactivate`, {
          method: "POST",
          body: JSON.stringify({ endDate, changedBy: state.changedBy || undefined, reason: "透過管理介面停用" }),
        });
        await loadFeeRules();
      } catch (err) {
        alert(err.message);
      }
    })
  );
}

function openFeeRuleForm() {
  const feeType = qs("#feeRuleTypeSelect").value;
  const body = `
    <h2>新增費率版本：${FEE_TYPE_LABEL[feeType] || feeType}</h2>
    <label>金額（元／節）<input id="f-amount" type="number" step="0.01" /></label>
    <label>生效日期<input id="f-effectiveDate" type="date" /></label>
    <label>結束日期（選填）<input id="f-endDate" type="date" /></label>
    <label>備註<textarea id="f-note"></textarea></label>
    <div class="error-text" id="f-error" hidden></div>
    <div class="modal-actions">
      <button type="button" class="secondary" id="f-cancel">取消</button>
      <button type="button" id="f-submit">儲存</button>
    </div>`;
  showModal(body);
  qs("#f-cancel").addEventListener("click", closeModal);
  qs("#f-submit").addEventListener("click", async () => {
    try {
      const amount = qs("#f-amount").value;
      const effectiveDate = qs("#f-effectiveDate").value;
      if (!amount) throw new Error("請輸入金額");
      if (!effectiveDate) throw new Error("請輸入生效日期");
      await api("/api/fee-rules", {
        method: "POST",
        body: JSON.stringify({
          semesterId: state.semesterId,
          feeType,
          amount,
          effectiveDate,
          endDate: qs("#f-endDate").value || undefined,
          note: qs("#f-note").value.trim() || undefined,
          changedBy: state.changedBy || undefined,
        }),
      });
      closeModal();
      await loadFeeRules();
    } catch (err) {
      showFormError(err.message);
    }
  });
}

// ---------- Phase 9 第一階段：節次型費用計算 ----------

function setupFeeCalcPeriodOptions() {
  const semester = getCurrentSemester();
  if (!semester) return;
  const months = monthRange(semester.startDate, semester.endDate);
  const select = qs("#feeCalcPeriodSelect");
  select.innerHTML = months.map((m) => `<option value="${m.year}-${m.month}">${m.year}年${m.month}月</option>`).join("");
  const hasCurrent = state.feeCalcPeriod && months.some((m) => m.year === state.feeCalcPeriod.year && m.month === state.feeCalcPeriod.month);
  const chosen = hasCurrent ? state.feeCalcPeriod : months[0];
  if (chosen) {
    select.value = `${chosen.year}-${chosen.month}`;
    state.feeCalcPeriod = chosen;
  }
}

// 對「這個學期＋這個年月」下所有匯入批次（BD／非BD 都算進去）先算金額，再彙總成
// 「代課教師 × fundingSource」的月結表。只計算 GENERAL/OVERTIME/PROJECT，
// 規則衝突／原教師未配對／尚未分類的資料維持不計算。
async function runFeeCalculation() {
  if (!state.semesterId || !state.feeCalcPeriod) return;
  const msgEl = qs("#feeCalcSummaryMsg");
  msgEl.hidden = false;
  msgEl.textContent = "計算中…";

  const batches = await api(`/api/monthly-imports?semesterId=${state.semesterId}`);
  const targetBatches = batches.filter(
    (b) => b.year === state.feeCalcPeriod.year && b.month === state.feeCalcPeriod.month && b.status === "ACTIVE"
  );
  if (targetBatches.length === 0) {
    msgEl.textContent = "這個年月沒有有效的匯入批次，請先到「公費代課匯入」上傳並完成分類。";
    qs("#feeSummaryTable tbody").innerHTML = "";
    return;
  }

  for (const b of targetBatches) {
    await api(`/api/monthly-imports/${b.id}/calculate-fees`, {
      method: "POST",
      body: JSON.stringify({ changedBy: state.changedBy || undefined }),
    });
  }

  const ids = targetBatches.map((b) => b.id).join(",");
  const summary = await api(`/api/monthly-imports/fee-summary?ids=${ids}`);
  msgEl.textContent = `已計算 ${targetBatches.length} 個批次，共 ${summary.length} 位代課教師有金額紀錄。`;

  const totalAmount = summary.reduce((s, r) => s + Number(r.totalAmount), 0);
  const totalCount = summary.reduce((s, r) => s + r.totalCount, 0);
  const generalAmount = summary.reduce((s, r) => s + Number(r.generalAmount), 0);
  const overtimeAmount = summary.reduce((s, r) => s + Number(r.overtimeAmount), 0);
  const projectAmount = summary.reduce((s, r) => s + Number(r.projectAmount), 0);
  let selfFundedTotal = 0;
  try {
    const selfFundedRecords = await api(`/api/self-funded?semesterId=${state.semesterId}&year=${state.feeCalcPeriod.year}&month=${state.feeCalcPeriod.month}`);
    selfFundedTotal = selfFundedRecords.reduce((s, r) => s + Number(r.amount ?? 0), 0);
  } catch (err) {
    // 自費代課資料讀取失敗不影響公費費用計算本身，安靜略過即可
  }

  const stripEl = qs("#feeCalcSummaryStrip");
  stripEl.hidden = false;
  stripEl.innerHTML = `
    <div class="summary-tile st-accent"><div class="st-label">本月公費總額</div><div class="st-value">$${totalAmount}</div></div>
    <div class="summary-tile"><div class="st-label">代課教師人數</div><div class="st-value">${summary.length}</div></div>
    <div class="summary-tile"><div class="st-label">總節數</div><div class="st-value">${totalCount}</div></div>
    <div class="summary-tile"><div class="st-label">一般公費</div><div class="st-value">$${generalAmount}</div></div>
    <div class="summary-tile"><div class="st-label">超鐘點</div><div class="st-value">$${overtimeAmount}</div></div>
    <div class="summary-tile"><div class="st-label">專案</div><div class="st-value">$${projectAmount}</div></div>
    <div class="summary-tile st-ok"><div class="st-label">自費代課（不列入公費總額）</div><div class="st-value">$${selfFundedTotal}</div></div>`;

  const tbody = qs("#feeSummaryTable tbody");
  tbody.innerHTML = summary
    .map(
      (r) => `
      <tr>
        <td>${r.substituteTeacherName}</td>
        <td>${r.generalCount}</td>
        <td>${r.generalAmount}</td>
        <td>${r.overtimeCount}</td>
        <td>${r.overtimeAmount}</td>
        <td>${r.projectCount}</td>
        <td>${r.projectAmount}</td>
        <td>${r.totalCount}</td>
        <td>${r.totalAmount}</td>
      </tr>`
    )
    .join("");
}

// ---------- Implementation Batch：共用「學期＋年月」選單 ----------

function setupClosePeriodOptions() {
  const semester = getCurrentSemester();
  if (!semester) return;
  const months = monthRange(semester.startDate, semester.endDate);
  const optionsHtml = months.map((m) => `<option value="${m.year}-${m.month}">${m.year}年${m.month}月</option>`).join("");
  const hasCurrent = state.closePeriod && months.some((m) => m.year === state.closePeriod.year && m.month === state.closePeriod.month);
  state.closePeriod = hasCurrent ? state.closePeriod : months[0];

  CLOSE_PERIOD_SELECT_IDS.forEach((id) => {
    const el = qs(`#${id}`);
    if (!el) return;
    el.innerHTML = optionsHtml;
    if (state.closePeriod) el.value = `${state.closePeriod.year}-${state.closePeriod.month}`;
    el.onchange = () => {
      const [year, month] = el.value.split("-").map(Number);
      state.closePeriod = { year, month };
      CLOSE_PERIOD_SELECT_IDS.forEach((otherId) => {
        if (otherId !== id) {
          const otherEl = qs(`#${otherId}`);
          if (otherEl) otherEl.value = el.value;
        }
      });
    };
  });
}

async function getActiveBatchIdsForClosePeriod() {
  if (!state.semesterId || !state.closePeriod) return [];
  const batches = await api(`/api/monthly-imports?semesterId=${state.semesterId}`);
  return batches
    .filter((b) => b.year === state.closePeriod.year && b.month === state.closePeriod.month && b.status === "ACTIVE")
    .map((b) => b.id);
}

// ---------- Implementation Batch：月結首頁 ----------

async function loadDashboard() {
  if (!state.semesterId || !state.closePeriod) return;
  const { year, month } = state.closePeriod;
  const dashboard = await api(`/api/monthly-dashboard?semesterId=${state.semesterId}&year=${year}&month=${month}`);
  renderDashboard(dashboard);
}

// 根據目前狀態，直接告訴使用者「現在應該做什麼」——不要只丟一堆數字讓人自己判斷。
function computeNextStep(d) {
  if (!d.import.hasActiveBatch && !d.selfFunded.exists) {
    return { level: "todo", icon: "📥", title: "尚未匯入公費代課資料", sub: "請先匯入本月 Excel（編制內／編制外各匯入一次）", actionTab: "importExcel", actionLabel: "前往匯入" };
  }
  if (d.lock.isLocked) {
    return { level: "done", icon: "🔒", title: "本月已完成並鎖定", sub: `鎖定人：${d.lock.lockedBy}｜時間：${new Date(d.lock.lockedAt).toLocaleString("zh-TW")}`, actionTab: null };
  }
  const classifiedCount = d.classification.general + d.classification.overtime + d.classification.project + d.classification.conflict + d.classification.teacherUnmatched;
  if (d.import.hasActiveBatch && classifiedCount < d.import.successCount) {
    return { level: "todo", icon: "👥", title: "公費資料已匯入", sub: "下一步：處理教師配對／執行自動分類", actionTab: "classification", actionLabel: "前往分類預覽" };
  }
  if (d.issues.blocking.total > 0) {
    return { level: "blocked", icon: "⚠️", title: `尚有 ${d.issues.blocking.total} 筆問題待處理`, sub: "需要先處理或確認接受，才能鎖定本月", actionTab: "pendingIssues", actionLabel: "前往待處理" };
  }
  if (d.fee.notCalculatedCount > 0) {
    return { level: "todo", icon: "💰", title: "分類與問題都已處理完成", sub: "下一步：執行費用計算", actionTab: "feeCalculation", actionLabel: "前往費用計算" };
  }
  return { level: "ok", icon: "🔓", title: "費用計算完成，本月已經準備好", sub: "建議先進行對帳確認金額，再鎖定本月", actionTab: "reconciliation", actionLabel: "前往對帳" };
}

function renderNextStepBanner(d) {
  const step = computeNextStep(d);
  const el = qs("#nextStepBanner");
  el.innerHTML = `
    <div class="next-step-banner ns-${step.level}">
      <div class="ns-icon">${step.icon}</div>
      <div class="ns-body">
        <div class="ns-title">${step.title}</div>
        <div class="ns-sub">${step.sub}</div>
      </div>
      ${step.actionTab ? `<button id="btnNextStepAction" data-tab="${step.actionTab}">${step.actionLabel} →</button>` : ""}
    </div>`;
  if (step.actionTab) {
    qs("#btnNextStepAction").addEventListener("click", () => {
      switchTab(step.actionTab);
      const loader = TAB_LOADERS[step.actionTab];
      if (loader) loader();
    });
  }
}

function renderDashboard(d) {
  renderNextStepBanner(d);

  const batchesHtml =
    d.import.batches
      .map((b) => `${STAFF_TYPE_LABEL[b.sourceStaffType] || b.sourceStaffType} v${b.versionNo}：總筆數${b.totalCount}｜成功${b.successCount}｜錯誤${b.errorCount}`)
      .join("<br>") || "尚無有效匯入批次";

  const classificationDone = d.import.hasActiveBatch && d.classification.general + d.classification.overtime + d.classification.project + d.classification.conflict + d.classification.teacherUnmatched >= d.import.successCount && d.import.successCount > 0;

  const el = qs("#dashboardContent");
  el.innerHTML = `
    <div class="dashboard-grid">
      <div class="dashboard-card">
        <div class="dc-head"><span class="dc-icon">📥</span><h3>公費代課</h3></div>
        <p>${batchesHtml}</p>
        <p>合計成功 ${d.import.successCount}｜合計錯誤 ${d.import.errorCount}</p>
      </div>
      <div class="dashboard-card">
        <div class="dc-head"><span class="dc-icon">👥</span><h3>分類</h3></div>
        ${classificationDone ? badge("ok", "🟢 已完成") : d.import.hasActiveBatch ? badge("warn", "🟡 尚未完成") : badge("idle", "⚪ 尚未開始")}
        <p style="margin-top: 8px;">一般公費 ${d.classification.general}｜超鐘點 ${d.classification.overtime}｜專案 ${d.classification.project}</p>
        <p>規則衝突 ${d.classification.conflict}｜原教師未配對 ${d.classification.teacherUnmatched}</p>
      </div>
      <div class="dashboard-card">
        <div class="dc-head"><span class="dc-icon">💰</span><h3>費用</h3></div>
        <div class="dc-figure">$${d.fee.totalAmount}<small>　已算 ${d.fee.calculatedCount} 筆／未算 ${d.fee.notCalculatedCount} 筆</small></div>
        <p>自費代課：${d.selfFunded.exists ? `${d.selfFunded.count} 筆（不列入公費金額）` : "無"}</p>
      </div>
      <div class="dashboard-card">
        <div class="dc-head"><span class="dc-icon">⚠️</span><h3>待處理</h3></div>
        ${d.issues.blocking.total > 0 ? badge("blocked", `🔴 ${d.issues.blocking.total} 件需要處理`) : badge("ok", "🟢 沒有待處理事項")}
        <p style="margin-top: 8px;">原教師未配對 ${d.issues.blocking.teacherUnmatched}｜規則衝突 ${d.issues.blocking.conflict}｜金額無法計算 ${d.issues.blocking.amountMissing}｜匯入錯誤 ${d.issues.blocking.importErrors}</p>
        <p>已確認接受 ${d.issues.acknowledgedCount} 筆</p>
      </div>
      <div class="dashboard-card">
        <div class="dc-head"><span class="dc-icon">🧾</span><h3>給出納</h3></div>
        ${d.fee.notCalculatedCount === 0 && d.fee.calculatedCount > 0 ? badge("ok", "🟢 可以產生") : badge("idle", "⚪ 尚未準備好")}
        <p style="margin-top: 8px;">費用計算完成後即可到「給出納」頁面產生彙總並下載 Excel。</p>
      </div>
      <div class="dashboard-card">
        <div class="dc-head"><span class="dc-icon">🔒</span><h3>月結</h3></div>
        ${d.lock.isLocked ? badge("locked", "🔒 本月已鎖定") : badge("idle", "🟢 本月尚未鎖定")}
        <div class="lock-panel" style="margin-top: 10px;">
          <ul class="lock-checklist">
            <li>${d.issues.blocking.teacherUnmatched === 0 ? "🟢" : "🔴"} 教師配對（${d.issues.blocking.teacherUnmatched} 筆未配對）</li>
            <li>${d.issues.blocking.conflict === 0 ? "🟢" : "🔴"} 分類（${d.issues.blocking.conflict} 筆規則衝突）</li>
            <li>${d.fee.notCalculatedCount === 0 ? "🟢" : "🟡"} 費用（${d.fee.notCalculatedCount} 筆尚未計算）</li>
            <li>${d.issues.blocking.total === 0 ? "🟢" : "🔴"} 待處理（合計 ${d.issues.blocking.total} 筆阻擋）</li>
            <li>🔵 對帳（建議鎖定前先核對金額）</li>
          </ul>
          ${d.lock.isLocked ? `<button id="btnUnlockMonth" class="danger">解除鎖定</button>` : `<button id="btnLockMonth">鎖定本月</button>`}
        </div>
      </div>
    </div>`;

  if (d.lock.isLocked) {
    qs("#btnUnlockMonth").addEventListener("click", openUnlockForm);
  } else {
    qs("#btnLockMonth").addEventListener("click", openLockForm);
  }
}

// 歷史月結：讓使用者一眼看到這學期每個月的金額、問題數與鎖定狀態，點「查看」直接切換過去。
async function loadHistoryTable() {
  const semester = getCurrentSemester();
  if (!semester || !state.semesterId) return;
  const months = monthRange(semester.startDate, semester.endDate);
  const results = await Promise.all(
    months.map((m) =>
      api(`/api/monthly-dashboard?semesterId=${state.semesterId}&year=${m.year}&month=${m.month}`)
        .then((dashboard) => ({ ...m, dashboard }))
        .catch(() => ({ ...m, dashboard: null }))
    )
  );

  const tbody = qs("#historyTable tbody");
  tbody.innerHTML = results
    .map(({ year, month, dashboard }) => {
      if (!dashboard || (!dashboard.import.hasActiveBatch && !dashboard.selfFunded.exists)) {
        return `
          <tr class="inactive-row">
            <td>${year}年${month}月</td><td>－</td><td>－</td><td>${badge("idle", "⚪ 尚未開始")}</td>
            <td><button data-action="goto" data-year="${year}" data-month="${month}">查看</button></td>
          </tr>`;
      }
      const statusBadge = dashboard.lock.isLocked
        ? badge("locked", "🔒 已鎖定")
        : dashboard.issues.blocking.total > 0
        ? badge("blocked", `🔴 ${dashboard.issues.blocking.total} 筆問題`)
        : badge("ok", "🟢 準備完成");
      return `
        <tr>
          <td>${year}年${month}月</td>
          <td>$${dashboard.fee.totalAmount}</td>
          <td>${dashboard.issues.blocking.total}</td>
          <td>${statusBadge}</td>
          <td><button data-action="goto" data-year="${year}" data-month="${month}">查看</button></td>
        </tr>`;
    })
    .join("");

  tbody.querySelectorAll("button[data-action='goto']").forEach((btn) =>
    btn.addEventListener("click", () => {
      const year = Number(btn.dataset.year);
      const month = Number(btn.dataset.month);
      state.closePeriod = { year, month };
      CLOSE_PERIOD_SELECT_IDS.forEach((id) => {
        const el = qs(`#${id}`);
        if (el) el.value = `${year}-${month}`;
      });
      loadDashboard();
    })
  );
}

function openLockForm() {
  const body = `
    <h2>鎖定 ${state.closePeriod.year}年${state.closePeriod.month}月</h2>
    <p class="hint">鎖定前系統會自動檢查阻擋性問題（原教師未配對／規則衝突／金額無法計算／匯入錯誤），若有未確認接受的問題會拒絕鎖定，請先到「待處理」處理或確認接受。</p>
    <label>操作人<input id="f-lockedBy" type="text" value="${state.changedBy}" /></label>
    <div class="error-text" id="f-error" hidden></div>
    <div class="modal-actions">
      <button type="button" class="secondary" id="f-cancel">取消</button>
      <button type="button" id="f-submit">確認鎖定</button>
    </div>`;
  showModal(body);
  qs("#f-cancel").addEventListener("click", closeModal);
  qs("#f-submit").addEventListener("click", async () => {
    try {
      const lockedBy = qs("#f-lockedBy").value.trim();
      if (!lockedBy) throw new Error("請輸入操作人");
      await api("/api/monthly-lock/lock", {
        method: "POST",
        body: JSON.stringify({ semesterId: state.semesterId, year: state.closePeriod.year, month: state.closePeriod.month, lockedBy }),
      });
      closeModal();
      await loadDashboard();
      await loadHistoryTable();
    } catch (err) {
      showFormError(err.message);
    }
  });
}

function openUnlockForm() {
  const body = `
    <h2>解除鎖定 ${state.closePeriod.year}年${state.closePeriod.month}月</h2>
    <label>操作人<input id="f-unlockedBy" type="text" value="${state.changedBy}" /></label>
    <label>解鎖理由（必填）<input id="f-reason" type="text" /></label>
    <div class="error-text" id="f-error" hidden></div>
    <div class="modal-actions">
      <button type="button" class="secondary" id="f-cancel">取消</button>
      <button type="button" id="f-submit">確認解鎖</button>
    </div>`;
  showModal(body);
  qs("#f-cancel").addEventListener("click", closeModal);
  qs("#f-submit").addEventListener("click", async () => {
    try {
      const unlockedBy = qs("#f-unlockedBy").value.trim();
      const reason = qs("#f-reason").value.trim();
      if (!unlockedBy) throw new Error("請輸入操作人");
      if (!reason) throw new Error("請輸入解鎖理由");
      await api("/api/monthly-lock/unlock", {
        method: "POST",
        body: JSON.stringify({ semesterId: state.semesterId, year: state.closePeriod.year, month: state.closePeriod.month, unlockedBy, reason }),
      });
      closeModal();
      await loadDashboard();
      await loadHistoryTable();
    } catch (err) {
      showFormError(err.message);
    }
  });
}

// ---------- Implementation Batch：待處理 ----------

async function loadPendingIssues() {
  if (!state.semesterId || !state.closePeriod) return;
  const { year, month } = state.closePeriod;
  state.pendingIssuesCache = await api(`/api/pending-issues?semesterId=${state.semesterId}&year=${year}&month=${month}`);
  renderPendingIssues();
}

function renderPendingIssues() {
  const issues = state.pendingIssuesCache;
  const pendingCount = issues.filter((i) => i.status === "PENDING").length;
  qs("#pendingSummaryLine").textContent =
    issues.length === 0 ? "本月目前沒有任何待處理事項" : `本月共有 ${issues.length} 件事項｜⚠️ 待處理 ${pendingCount} 件｜✅ 已確認 ${issues.length - pendingCount} 件`;

  const visible = issues.filter((r) => state.pendingFilter === "ALL" || r.status === state.pendingFilter);

  const tbody = qs("#pendingIssuesTable tbody");
  tbody.innerHTML = visible
    .map((r) => {
      const idx = issues.indexOf(r);
      const statusText =
        r.status === "ACKNOWLEDGED"
          ? `${badge("ok", "✅ 已確認接受")}<br><small>${r.acknowledgement.reason}（${r.acknowledgement.acknowledgedBy}）</small>`
          : badge("blocked", "⚠️ 待處理");
      return `
        <tr data-idx="${idx}" class="${r.status === "ACKNOWLEDGED" ? "" : "conflict-row"}">
          <td>${r.date ?? ""}</td>
          <td>${r.originalTeacher ?? ""}</td>
          <td>${r.substituteTeacher ?? ""}</td>
          <td>${r.periodCode ?? ""}</td>
          <td>${r.className ?? ""}</td>
          <td>${r.subject ?? ""}</td>
          <td>${PENDING_ISSUE_TYPE_LABEL[r.issueType] || r.issueType}</td>
          <td>${r.description}</td>
          <td>${statusText}</td>
          <td>${r.status === "ACKNOWLEDGED" ? `<button data-action="revoke">撤銷確認</button>` : `<button data-action="acknowledge">確認接受</button>`}</td>
        </tr>`;
    })
    .join("");

  tbody.querySelectorAll("button[data-action='acknowledge']").forEach((btn) =>
    btn.addEventListener("click", (e) => {
      const idx = Number(e.target.closest("tr").dataset.idx);
      openAcknowledgeForm(issues[idx]);
    })
  );
  tbody.querySelectorAll("button[data-action='revoke']").forEach((btn) =>
    btn.addEventListener("click", async (e) => {
      const idx = Number(e.target.closest("tr").dataset.idx);
      const issue = issues[idx];
      try {
        await api("/api/issue-acknowledgements/revoke", {
          method: "POST",
          body: JSON.stringify({ targetTable: issue.targetTable, targetId: issue.targetId, changedBy: state.changedBy || undefined }),
        });
        await loadPendingIssues();
      } catch (err) {
        alert(err.message);
      }
    })
  );
}

function openAcknowledgeForm(issue) {
  const body = `
    <h2>確認接受問題</h2>
    <p>${PENDING_ISSUE_TYPE_LABEL[issue.issueType] || issue.issueType}：${issue.description}</p>
    <p class="hint">確認接受不會刪除或修改原始資料，只是註記「已經看過、同意讓它不再阻擋本月鎖定」，這筆問題仍會保留在清單中可查詢。</p>
    <label>確認人<input id="f-ackBy" type="text" value="${state.changedBy}" /></label>
    <label>確認理由（必填）<input id="f-reason" type="text" /></label>
    <div class="error-text" id="f-error" hidden></div>
    <div class="modal-actions">
      <button type="button" class="secondary" id="f-cancel">取消</button>
      <button type="button" id="f-submit">確認</button>
    </div>`;
  showModal(body);
  qs("#f-cancel").addEventListener("click", closeModal);
  qs("#f-submit").addEventListener("click", async () => {
    try {
      const acknowledgedBy = qs("#f-ackBy").value.trim();
      const reason = qs("#f-reason").value.trim();
      if (!acknowledgedBy) throw new Error("請輸入確認人");
      if (!reason) throw new Error("請輸入確認理由");
      await api("/api/issue-acknowledgements", {
        method: "POST",
        body: JSON.stringify({
          semesterId: state.semesterId,
          year: state.closePeriod.year,
          month: state.closePeriod.month,
          targetTable: issue.targetTable,
          targetId: issue.targetId,
          reason,
          acknowledgedBy,
        }),
      });
      closeModal();
      await loadPendingIssues();
    } catch (err) {
      showFormError(err.message);
    }
  });
}

// ---------- Implementation Batch：自費代課 ----------

async function loadSelfFunded() {
  if (!state.semesterId || !state.closePeriod) return;
  const { year, month } = state.closePeriod;
  const records = await api(`/api/self-funded?semesterId=${state.semesterId}&year=${year}&month=${month}`);
  const tbody = qs("#selfFundedTable tbody");
  tbody.innerHTML = records
    .map(
      (r) => `
      <tr data-id="${r.id}">
        <td>${dateOnly(r.date)}</td>
        <td>${r.originalTeacher?.name ?? ""}</td>
        <td>${r.substituteTeacher?.name ?? ""}</td>
        <td>${r.periodCode ?? ""}</td>
        <td>${r.className ?? ""}</td>
        <td>${r.subject ?? ""}</td>
        <td>${r.amount ?? ""}</td>
        <td>${r.note ?? ""}</td>
        <td>${r.createdBy ?? ""}</td>
        <td>
          <button data-action="edit">編輯</button>
          <button data-action="delete" class="danger">刪除</button>
        </td>
      </tr>`
    )
    .join("");

  tbody.querySelectorAll("button[data-action='edit']").forEach((btn) =>
    btn.addEventListener("click", (e) => {
      const id = e.target.closest("tr").dataset.id;
      openSelfFundedForm(records.find((r) => r.id === id));
    })
  );
  tbody.querySelectorAll("button[data-action='delete']").forEach((btn) =>
    btn.addEventListener("click", async (e) => {
      const id = e.target.closest("tr").dataset.id;
      const reason = prompt("刪除原因（必填）：");
      if (!reason) return;
      try {
        await api(`/api/self-funded/${id}`, {
          method: "DELETE",
          body: JSON.stringify({ deletedBy: state.changedBy || undefined, reason }),
        });
        await loadSelfFunded();
      } catch (err) {
        alert(err.message);
      }
    })
  );
}

async function findPersonByExactName(name) {
  if (!name) return null;
  const matches = await api(`/api/persons?search=${encodeURIComponent(name)}`);
  return matches.find((p) => p.name === name) || null;
}

function openSelfFundedForm(existing) {
  if (!state.closePeriod) {
    alert("請先在畫面上方選擇學期＋年月");
    return;
  }
  const body = `
    <h2>${existing ? "編輯" : "新增"}自費代課</h2>
    <label>日期<input id="f-date" type="date" value="${dateOnly(existing?.date) || ""}" /></label>
    <label>原教師姓名（選填）<input id="f-originalTeacherName" type="text" value="${existing?.originalTeacher?.name ?? ""}" /></label>
    <label>代課教師姓名（必填）<input id="f-substituteTeacherName" type="text" value="${existing?.substituteTeacher?.name ?? ""}" /></label>
    <label>節次<select id="f-periodCode"><option value="">（未填）</option>${periodOptions(existing?.periodCode)}</select></label>
    <label>班級<input id="f-className" type="text" value="${existing?.className ?? ""}" /></label>
    <label>科目<input id="f-subject" type="text" value="${existing?.subject ?? ""}" /></label>
    <label>金額（必填）<input id="f-amount" type="number" step="0.01" value="${existing?.amount ?? ""}" /></label>
    <label>備註<textarea id="f-note">${existing?.note ?? ""}</textarea></label>
    <div class="error-text" id="f-error" hidden></div>
    <div class="modal-actions">
      <button type="button" class="secondary" id="f-cancel">取消</button>
      <button type="button" id="f-submit">儲存</button>
    </div>`;
  showModal(body);
  qs("#f-cancel").addEventListener("click", closeModal);
  qs("#f-submit").addEventListener("click", async () => {
    try {
      const date = qs("#f-date").value;
      const substituteName = qs("#f-substituteTeacherName").value.trim();
      const originalName = qs("#f-originalTeacherName").value.trim();
      const amount = qs("#f-amount").value;
      if (!date) throw new Error("請填寫日期");
      if (!substituteName) throw new Error("請填寫代課教師姓名");
      if (!amount) throw new Error("請填寫金額");

      const substitutePerson = await findPersonByExactName(substituteName);
      if (!substitutePerson) throw new Error(`系統找不到教師「${substituteName}」，請先在人員管理建立此人員`);
      let originalPerson = null;
      if (originalName) {
        originalPerson = await findPersonByExactName(originalName);
        if (!originalPerson) throw new Error(`系統找不到教師「${originalName}」，請先在人員管理建立此人員`);
      }

      const payload = {
        date,
        originalTeacherId: originalPerson?.id,
        substituteTeacherId: substitutePerson.id,
        periodCode: qs("#f-periodCode").value || undefined,
        className: qs("#f-className").value.trim() || undefined,
        subject: qs("#f-subject").value.trim() || undefined,
        amount,
        note: qs("#f-note").value.trim() || undefined,
      };

      if (existing) {
        await api(`/api/self-funded/${existing.id}`, {
          method: "PATCH",
          body: JSON.stringify({ ...payload, updatedBy: state.changedBy || undefined }),
        });
      } else {
        await api("/api/self-funded", {
          method: "POST",
          body: JSON.stringify({
            ...payload,
            semesterId: state.semesterId,
            year: state.closePeriod.year,
            month: state.closePeriod.month,
            createdBy: state.changedBy || undefined,
          }),
        });
      }
      closeModal();
      await loadSelfFunded();
    } catch (err) {
      showFormError(err.message);
    }
  });
}

// ---------- Implementation Batch：給出納 ----------

async function loadChuna() {
  if (!state.semesterId || !state.closePeriod) return;
  const ids = await getActiveBatchIdsForClosePeriod();
  state.chunaBatchIds = ids;
  if (ids.length === 0) {
    qs("#chunaContent").innerHTML = "<p>這個年月沒有有效的匯入批次，請先到「公費代課匯入」上傳。</p>";
    return;
  }
  const summary = await api(`/api/chuna-summary?ids=${ids.join(",")}`);
  renderChuna(summary);
}

function renderChunaSection(title, rows) {
  if (rows.length === 0) return "";
  const totalCount = rows.reduce((s, r) => s + r.totalCount, 0);
  const totalAmount = rows.reduce((s, r) => s + Number(r.totalAmount), 0);
  return `
    <h3>${title}</h3>
    <table>
      <thead><tr><th>薪資代碼</th><th>教師</th><th>一般公費（節數／金額）</th><th>超鐘點（節數／金額）</th><th>專案（節數／金額）</th><th>總節數</th><th>總金額</th></tr></thead>
      <tbody>
        ${rows
          .map(
            (r) => `
          <tr>
            <td>${r.payrollCode ?? ""}</td>
            <td>${r.substituteTeacherName}</td>
            <td>${r.generalCount} ／ ${r.generalAmount}</td>
            <td>${r.overtimeCount} ／ ${r.overtimeAmount}</td>
            <td>${r.projectCount} ／ ${r.projectAmount}</td>
            <td>${r.totalCount}</td>
            <td>${r.totalAmount}</td>
          </tr>`
          )
          .join("")}
        <tr class="total-row"><td colspan="5">總計</td><td>${totalCount}</td><td>${totalAmount}</td></tr>
      </tbody>
    </table>`;
}

function renderChuna(summary) {
  const allRows = [...summary.bd, ...summary.nonBd, ...summary.unknown];
  const stripEl = qs("#chunaSummaryStrip");
  if (allRows.length > 0) {
    const totalAmount = allRows.reduce((s, r) => s + Number(r.totalAmount), 0);
    const totalCount = allRows.reduce((s, r) => s + r.totalCount, 0);
    stripEl.hidden = false;
    stripEl.innerHTML = `
      <div class="summary-tile st-accent"><div class="st-label">給出納總額</div><div class="st-value">$${totalAmount}</div></div>
      <div class="summary-tile"><div class="st-label">代課教師人數</div><div class="st-value">${allRows.length}</div></div>
      <div class="summary-tile"><div class="st-label">總節數</div><div class="st-value">${totalCount}</div></div>`;
  } else {
    stripEl.hidden = true;
  }

  const html =
    renderChunaSection("編制外（非BD）", summary.nonBd) +
    renderChunaSection("編制內（BD）", summary.bd) +
    renderChunaSection("未標示來源", summary.unknown);
  qs("#chunaContent").innerHTML = html || "<p>這個年月目前沒有可以彙總的資料。</p>";
}

async function exportChuna() {
  if (!state.semesterId || !state.closePeriod) return;
  const ids = state.chunaBatchIds.length > 0 ? state.chunaBatchIds : await getActiveBatchIdsForClosePeriod();
  if (ids.length === 0) {
    alert("這個年月沒有有效的匯入批次");
    return;
  }
  const url = `/api/chuna-export?ids=${ids.join(",")}&year=${state.closePeriod.year}&month=${state.closePeriod.month}`;
  window.location.href = url;
}

// ---------- Implementation Batch：對帳 ----------

async function runReconciliation() {
  const fileInput = qs("#reconciliationFile");
  if (!fileInput.files || fileInput.files.length === 0) {
    alert("請先選擇要比對的原始給出納 Excel 檔案");
    return;
  }
  if (!state.semesterId || !state.closePeriod) return;
  const ids = await getActiveBatchIdsForClosePeriod();
  if (ids.length === 0) {
    alert("這個年月沒有有效的匯入批次，無法比對");
    return;
  }

  const formData = new FormData();
  formData.append("file", fileInput.files[0]);
  formData.append("ids", ids.join(","));

  let result;
  try {
    const res = await fetch("/api/reconciliation", { method: "POST", body: formData });
    result = await res.json();
    if (!res.ok) throw new Error(result.error || "比對失敗");
  } catch (err) {
    alert(err.message);
    return;
  }
  renderReconciliation(result);
}

function renderReconciliation(result) {
  const matchCount = result.rows.filter((r) => r.status === "MATCH").length;
  const diffCount = result.rows.length - matchCount;
  const diffTileClass = result.totals.diff === 0 ? "st-ok" : "st-danger";

  const summaryEl = qs("#reconciliationSummary");
  summaryEl.hidden = false;
  summaryEl.innerHTML = `
    <div class="summary-tile"><div class="st-label">教師人數</div><div class="st-value">${result.rows.length}</div></div>
    <div class="summary-tile st-ok"><div class="st-label">🟢 一致</div><div class="st-value">${matchCount}</div></div>
    <div class="summary-tile ${diffCount > 0 ? 'st-danger' : ''}"><div class="st-label">🔴 有差異</div><div class="st-value">${diffCount}</div></div>
    <div class="summary-tile"><div class="st-label">系統總額</div><div class="st-value">$${result.totals.systemAmount}</div></div>
    <div class="summary-tile"><div class="st-label">原始總額</div><div class="st-value">$${result.totals.originalAmount}</div></div>
    <div class="summary-tile ${diffTileClass}"><div class="st-label">差額</div><div class="st-value">$${result.totals.diff}</div></div>`;

  const tbody = qs("#reconciliationTable tbody");
  tbody.innerHTML = result.rows
    .map(
      (r) => `
      <tr class="${r.status === "MATCH" ? "" : "conflict-row"}">
        <td>${r.name}</td>
        <td>${r.systemPeriodCount ?? "－"}</td>
        <td>${r.originalPeriodCount ?? "－"}</td>
        <td>${r.systemAmount ?? "－"}</td>
        <td>${r.originalAmount ?? "－"}</td>
        <td>${r.amountDiff ?? "－"}</td>
        <td>${RECONCILIATION_STATUS_LABEL[r.status] || r.status}</td>
        <td>${r.possibleReason ?? ""}</td>
      </tr>`
    )
    .join("");
}

init();
