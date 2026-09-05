// Phase 5 管理介面：純 vanilla JS，沒有建置流程，呼叫 /api/* 取得/送出資料。
// 目標是「行政人員容易操作」，不是視覺華麗。

const WEEKDAY_LABEL = { MON: "星期一", TUE: "星期二", WED: "星期三", THU: "星期四", FRI: "星期五" };
const CLASSIFICATION_LABEL = { GENERAL: "一般公費", OVERTIME: "超鐘點", PROJECT: "專案" };

const state = {
  semesterId: null,
  changedBy: localStorage.getItem("changedBy") || "",
  ruleType: "OVERTIME",
  periodSlots: [],
  projects: [],
};

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

  qsa(".tab-btn").forEach((btn) => btn.addEventListener("click", () => switchTab(btn.dataset.tab)));
  qsa(".subtab-btn").forEach((btn) =>
    btn.addEventListener("click", () => {
      state.ruleType = btn.dataset.ruletype;
      qsa(".subtab-btn").forEach((b) => b.classList.toggle("active", b === btn));
      loadWeeklyRules();
    })
  );

  qs("#semesterSelect").addEventListener("change", async (e) => {
    state.semesterId = e.target.value;
    await loadForSemester();
  });

  qs("#filterTeacher").addEventListener("input", debounce(loadWeeklyRules, 300));
  qs("#filterWeekday").addEventListener("change", loadWeeklyRules);
  qs("#filterProject").addEventListener("change", loadWeeklyRules);
  qs("#btnNewWeeklyRule").addEventListener("click", () => openWeeklyRuleForm());

  qs("#btnNewProject").addEventListener("click", () => openProjectForm());

  qs("#filterIncludeCancelled").addEventListener("change", loadDateRules);
  qs("#btnNewDateRule").addEventListener("click", () => openDateRuleForm());

  const semesters = await api("/api/semesters");
  const select = qs("#semesterSelect");
  select.innerHTML = semesters
    .map((s) => `<option value="${s.id}">${s.schoolYear}學年度第${s.term}學期${s.isCurrent ? "（使用中）" : ""}</option>`)
    .join("");
  const current = semesters.find((s) => s.isCurrent) || semesters[0];
  if (current) {
    select.value = current.id;
    state.semesterId = current.id;
  }

  state.periodSlots = await api("/api/period-slots");

  await loadForSemester();
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
}

async function loadForSemester() {
  if (!state.semesterId) return;
  state.projects = await api(`/api/projects?semesterId=${state.semesterId}`);
  const projectFilter = qs("#filterProject");
  projectFilter.innerHTML =
    `<option value="">全部專案</option>` + state.projects.map((p) => `<option value="${p.id}">${p.name}</option>`).join("");
  await Promise.all([loadWeeklyRules(), loadProjects(), loadDateRules()]);
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

init();
