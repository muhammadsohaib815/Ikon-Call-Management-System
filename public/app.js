const state = {
  attendance: { id: null, filename: "", columns: [], rows: [], offset: 0, total: 0, hasNext: false },
  fmcgMobileLiveCall: { id: null, filename: "", size: 0 },
};

const $ = (selector, root = document) => root.querySelector(selector);
const $$ = (selector, root = document) => [...root.querySelectorAll(selector)];

const elements = {
  attendanceInput: $("#attendance-file-input"),
  attendanceDropZone: $("#attendance-drop-zone"),
  attendanceStatus: $("#attendance-status"),
  attendanceDataCard: $("#attendance-data-card"),
  attendanceFilename: $("#attendance-filename"),
  attendanceCount: $("#attendance-record-count"),
  attendanceHead: $("#attendance-table thead"),
  attendanceBody: $("#attendance-table tbody"),
  attendancePrev: $("#attendance-prev"),
  attendanceNext: $("#attendance-next"),
  attendancePageIndicator: $("#attendance-page-indicator"),
  callNatureResults: $("#call-nature-results"),
  callNatureSummary: $("#call-nature-summary"),
  callNatureCounts: $("#call-nature-counts"),
  callNaturePreview: $("#call-nature-preview tbody"),
  downloadCallNature: $("#download-call-nature"),
  downloadCallNatureCsv: $("#download-call-nature-csv"),
  fmcgMobileLiveCallInput: $("#fmcg-mobile-live-call-file-input"),
  fmcgMobileLiveCallDropZone: $("#fmcg-mobile-live-call-drop-zone"),
  fmcgMobileLiveCallStatus: $("#fmcg-mobile-live-call-status"),
};

const MAX_UPLOAD_BYTES = 200 * 1024 * 1024;
const FROZEN_REPORT_COLUMN_KEYS = new Set(["desg", "position_code", "sas_id", "ff_name"]);

function escapeHtml(value) {
  return String(value ?? "").replace(/[&<>'"]/g, (character) => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", "'": "&#39;", '"': "&quot;",
  })[character]);
}

function attendanceStatus(message, error = false) {
  elements.attendanceStatus.textContent = message;
  elements.attendanceStatus.classList.toggle("error", error);
}

async function api(url, options = {}) {
  const response = await fetch(url, options);
  if (!response.ok) {
    let message = "The request could not be completed.";
    try {
      const payload = await response.json();
      message = payload.error || payload.message || message;
    } catch {}
    const error = new Error(message);
    error.status = response.status;
    throw error;
  }
  return response;
}

function renderAttendance() {
  const { columns, rows } = state.attendance;
  elements.attendanceHead.innerHTML = `<tr>${columns.map((column) => `<th>${escapeHtml(column)}</th>`).join("")}</tr>`;
  elements.attendanceBody.innerHTML = rows.length
    ? rows.map((row) => `<tr>${columns.map((column) => `<td title="${escapeHtml(row[column])}">${escapeHtml(row[column])}</td>`).join("")}</tr>`).join("")
    : `<tr><td colspan="${Math.max(columns.length, 1)}">No records found.</td></tr>`;
  elements.attendanceCount.textContent = state.attendance.total.toLocaleString();
  const firstRow = rows.length ? state.attendance.offset + 1 : 0;
  const lastRow = state.attendance.offset + rows.length;
  elements.attendancePageIndicator.textContent = rows.length ? `Showing ${firstRow.toLocaleString()}–${lastRow.toLocaleString()}` : "No rows";
  elements.attendancePrev.disabled = state.attendance.offset === 0;
  elements.attendanceNext.disabled = !state.attendance.hasNext;
}

async function uploadAttendanceSummary(file) {
  if (!file) return;
  const extension = file.name.toLowerCase().split(".").pop();
  if (!["csv", "xlsx", "xls"].includes(extension)) {
    attendanceStatus("Unable to process file. Please check the Mobile Live Calls Summary format and required columns.", true);
    return;
  }
  const formData = new FormData();
  formData.append("file", file);
  attendanceStatus("Processing Call Nature and saving rows to live_call_details…");
  try {
    const response = await api("/api/attendance/mobile-live-call-summary", { method: "POST", body: formData });
    const result = await response.json();
    state.attendance = {
      id: result.id,
      filename: result.filename,
      columns: result.columns,
      rows: result.rows,
      offset: result.offset,
      total: result.total,
      hasNext: result.hasNext,
    };
    elements.attendanceFilename.textContent = result.filename;
    elements.attendanceDataCard.classList.remove("hidden");
    if (result.callNature) {
      const { validation } = result.callNature;
      elements.callNatureSummary.textContent = `${validation.totalRows.toLocaleString()} rows processed · ${validation.unrecognizedCallTypeRows.toLocaleString()} blank or unrecognized CallType row(s).`;
      elements.callNatureCounts.innerHTML = Object.entries(validation.counts)
        .map(([label, count]) => `<div class="validation-count"><strong>${escapeHtml(String(count))}</strong><span>${escapeHtml(label)}</span></div>`)
        .join("");
      elements.downloadCallNature.onclick = () => {
        const link = document.createElement("a");
        link.href = `/api/attendance/mobile-live-call-summary/${encodeURIComponent(result.id)}/download`;
        link.download = result.callNature.filename;
        link.click();
      };
      elements.downloadCallNatureCsv.onclick = () => {
        const link = document.createElement("a");
        link.href = `/api/attendance/mobile-live-call-summary/${encodeURIComponent(result.id)}/download-csv`;
        link.download = result.callNature.filename.replace(/\.xlsx$/i, ".csv");
        link.click();
      };
      elements.callNatureResults.classList.remove("hidden");
    } else {
      elements.callNatureResults.classList.add("hidden");
    }
    renderAttendance();
    const firstRow = result.rows.length ? result.offset + 1 : 0;
    const lastRow = result.offset + result.rows.length;
    const databaseMessage = result.databaseUpload?.message
      || "Database upload status was not returned.";
    attendanceStatus(
      `Call Nature applied. ${databaseMessage} Showing rows ${firstRow.toLocaleString()}–${lastRow.toLocaleString()} of ${result.total.toLocaleString()}.`
    );
    if (typeof loadDashboardAnalytics === "function") {
      dashState.filtersPopulated = false;
      loadDashboardAnalytics();
      loadLiveCallDetailsInteractiveTable();
    }
  } catch (error) {
    attendanceStatus(error.message || "Unable to process file. Please check the Mobile Live Calls Summary format and required columns.", true);
  } finally {
    elements.attendanceInput.value = "";
  }
}

async function loadAttendancePage(offset) {
  if (!state.attendance.id || offset < 0 || offset >= state.attendance.total) return;
  attendanceStatus("Loading rows…");
  try {
    const response = await api(`/api/attendance/mobile-live-call-summary/${state.attendance.id}?offset=${offset}`);
    const result = await response.json();
    state.attendance = { ...state.attendance, columns: result.columns, rows: result.rows, offset: result.offset, total: result.total, hasNext: result.hasNext };
    renderAttendance();
    const firstRow = result.rows.length ? result.offset + 1 : 0;
    const lastRow = result.offset + result.rows.length;
    attendanceStatus(
      `Mobile Live Calls Summary: Showing rows ${firstRow.toLocaleString()}–${lastRow.toLocaleString()} of ${result.total.toLocaleString()} total rows. Upload another file to compare call activity and attendance records.`
    );
  } catch (error) {
    attendanceStatus(error.message || "Unable to process file. Please check the Mobile Live Calls Summary format and required columns.", true);
  }
}

elements.attendanceInput.addEventListener("change", () => uploadAttendanceSummary(elements.attendanceInput.files[0]));
["dragenter", "dragover"].forEach((eventName) => elements.attendanceDropZone.addEventListener(eventName, (event) => {
  event.preventDefault();
  elements.attendanceDropZone.classList.add("dragging");
}));
["dragleave", "drop"].forEach((eventName) => elements.attendanceDropZone.addEventListener(eventName, (event) => {
  event.preventDefault();
  elements.attendanceDropZone.classList.remove("dragging");
}));
elements.attendanceDropZone.addEventListener("drop", (event) => uploadAttendanceSummary(event.dataTransfer.files[0]));
elements.attendancePrev.addEventListener("click", () => loadAttendancePage(state.attendance.offset - 50));
elements.attendanceNext.addEventListener("click", () => loadAttendancePage(state.attendance.offset + 50));

function fmcgMobileLiveCallStatus(message, error = false) {
  elements.fmcgMobileLiveCallStatus.textContent = message;
  elements.fmcgMobileLiveCallStatus.classList.toggle("error", error);
}

function formatFileSize(bytes) {
  if (bytes < 1024 * 1024) return `${Math.max(1, Math.round(bytes / 1024))} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

async function uploadFmcgMobileLiveCallSummary(file) {
  if (!file) return;
  const extension = file.name.toLowerCase().split(".").pop();
  if (!["csv", "xlsx", "xls"].includes(extension)) {
    fmcgMobileLiveCallStatus("Choose a valid CSV, XLSX, or XLS file for the Mobile Live Call Attendence.", true);
    return;
  }
  if (file.size > MAX_UPLOAD_BYTES) {
    fmcgMobileLiveCallStatus("The Mobile Live Call Attendence file exceeds the 200 MB upload limit.", true);
    return;
  }
  const formData = new FormData();
  formData.append("file", file);
  fmcgMobileLiveCallStatus("Uploading Mobile Live Call Attendence…");
  try {
    const response = await api("/api/attendance/fmcg-mobile-live-call-summary", { method: "POST", body: formData });
    const result = await response.json();
    state.fmcgMobileLiveCall = { id: result.id, filename: result.filename, size: result.size };
    fmcgMobileLiveCallStatus(result.message || `${result.filename} uploaded to call_attendence.`);
    if (typeof loadDashboardAnalytics === "function") {
      dashState.filtersPopulated = false;
      loadDashboardAnalytics();
      loadCallAttendenceInteractiveTable();
    }
  } catch (error) {
    fmcgMobileLiveCallStatus(error.message || "Unable to upload the Mobile Live Call Attendence file.", true);
  } finally {
    elements.fmcgMobileLiveCallInput.value = "";
  }
}

elements.fmcgMobileLiveCallInput.addEventListener("change", () => uploadFmcgMobileLiveCallSummary(elements.fmcgMobileLiveCallInput.files[0]));
["dragenter", "dragover"].forEach((eventName) => elements.fmcgMobileLiveCallDropZone.addEventListener(eventName, (event) => {
  event.preventDefault();
  elements.fmcgMobileLiveCallDropZone.classList.add("dragging");
}));
["dragleave", "drop"].forEach((eventName) => elements.fmcgMobileLiveCallDropZone.addEventListener(eventName, (event) => {
  event.preventDefault();
  elements.fmcgMobileLiveCallDropZone.classList.remove("dragging");
}));
elements.fmcgMobileLiveCallDropZone.addEventListener("drop", (event) => uploadFmcgMobileLiveCallSummary(event.dataTransfer.files[0]));

function buildBaseTeamSummaryColumns(regionLabel = "FMCG Region") {
  return [
    { key: "desg", label: "Legend", num: false },
    { key: "zone", label: "Zone", num: false },
    { key: "region", label: regionLabel, num: false },
    { key: "dept_territory_code", label: "Dept. Territory Code", num: false },
    { key: "position_code", label: "Position_Code", num: false },
    { key: "sas_id", label: "SAS_ID", num: false },
    { key: "ff_name", label: "SPO Name", num: false },
    { key: "base_town", label: "Base Town", num: false },
    { key: "emp_type", label: "Emp. Type", num: false },
    { key: "emp_code", label: "Emp. Code", num: false },
    { key: "payroll_status", label: "Payroll Status", num: false },
    { key: "doj", label: "D.O.J", num: false },
    { key: "reps", label: "Reps", num: true },
    { key: "vacant", label: "# Of Vacant", num: true },
    { key: "day_att_used", label: "Active SPOs Today Attn", num: true },
    { key: "mtd_att_used", label: "Active SPOs MTD Attn", num: true },
    { key: "day_call_used", label: "Active SPOs Today Calls", num: true },
    { key: "mtd_calls_used", label: "Active SPOs MTD Calls", num: true },
    { key: "ikon_used", label: "IKON Used", num: true, groupEnd: true },
  ];
}

const BASE_FMCG_COLUMNS = buildBaseTeamSummaryColumns("FMCG Region");
const BASE_PHARMA_COLUMNS = buildBaseTeamSummaryColumns("PHARMA Region");

function formatDisplayDate(isoDate) {
  if (!isoDate) return "";
  const parts = isoDate.split("-");
  if (parts.length !== 3) return isoDate;
  const months = ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"];
  const monthIdx = Number(parts[1]) - 1;
  return `${parts[2]}-${months[monthIdx] || parts[1]}-${parts[0]}`;
}

function formatMonthLabel(isoDate) {
  if (!isoDate) return "OCT-2026";
  const parts = isoDate.split("-");
  if (parts.length !== 3) return "OCT-2026";
  const months = ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"];
  const monthIdx = Number(parts[1]) - 1;
  return `${months[monthIdx] || parts[1]}-${parts[0]}`;
}

function getDaysInMonthFromIso(isoDate) {
  if (!isoDate) return 31;
  const parts = isoDate.split("-").map(Number);
  if (parts.length !== 3 || !parts[0] || !parts[1]) return 31;
  return new Date(parts[0], parts[1], 0).getDate();
}

function populateSelectOptions(selectEl, defaultLabel, items, currentValue) {
  selectEl.innerHTML = `<option value="">${escapeHtml(defaultLabel)}</option>${items
    .map((item) => `<option value="${escapeHtml(item)}">${escapeHtml(item)}</option>`)
    .join("")}`;
  selectEl.value = currentValue;
}

const MULTI_FILTER_SEPARATOR = "|";

function splitMultiFilterValue(value) {
  return String(value || "").split(MULTI_FILTER_SEPARATOR).map((item) => item.trim()).filter(Boolean);
}

function createExcelMultiSelect(source, defaultLabel, onApply) {
  const state = { items: [], active: new Set(), draft: new Set(), open: false };
  const host = document.createElement("div");
  host.className = "excel-multiselect";
  host.innerHTML = `
    <button type="button" class="excel-multiselect-trigger" aria-haspopup="listbox" aria-expanded="false">
      <span class="excel-multiselect-label"></span><span class="excel-multiselect-arrow" aria-hidden="true">▾</span>
    </button>
    <div class="excel-multiselect-menu" role="dialog" aria-label="${escapeHtml(defaultLabel)} filter">
      <label class="excel-multiselect-search"><input type="search" autocomplete="off" aria-label="Search ${escapeHtml(defaultLabel)}" placeholder="Search..." /></label>
      <label class="excel-multiselect-all"><input type="checkbox" /> <span>Select all search results</span></label>
      <div class="excel-multiselect-options" role="listbox" aria-multiselectable="true"></div>
      <div class="excel-multiselect-actions">
        <button type="button" class="excel-multiselect-clear">Clear</button>
        <button type="button" class="excel-multiselect-apply">Apply</button>
      </div>
    </div>`;
  source.classList.add("excel-multiselect-source");
  source.setAttribute("tabindex", "-1");
  source.insertAdjacentElement("afterend", host);

  const trigger = $(".excel-multiselect-trigger", host);
  const label = $(".excel-multiselect-label", host);
  const menu = $(".excel-multiselect-menu", host);
  const searchInput = $(".excel-multiselect-search input", host);
  const selectAll = $(".excel-multiselect-all input", host);
  const options = $(".excel-multiselect-options", host);
  const clearButton = $(".excel-multiselect-clear", host);
  const applyButton = $(".excel-multiselect-apply", host);

  const serialize = (values) => state.items.filter((item) => values.has(item)).join(MULTI_FILTER_SEPARATOR);
  const setOpen = (open) => {
    state.open = open;
    host.classList.toggle("is-open", open);
    trigger.setAttribute("aria-expanded", String(open));
    if (open) {
      state.draft = new Set(state.active);
      renderOptions();
    }
  };
  const writeSource = () => {
    const value = serialize(state.active);
    source.value = value;
    if (source.tagName === "SELECT") {
      [...source.options].forEach((option) => { option.selected = state.active.has(option.value); });
    }
  };
  const refreshLabel = () => {
    const selected = state.items.filter((item) => state.active.has(item));
    label.textContent = selected.length === 0
      ? defaultLabel
      : selected.length === 1 ? selected[0] : `${selected.length} selected`;
    trigger.title = selected.length ? selected.join(", ") : defaultLabel;
  };
  const matchingItems = () => {
    const query = searchInput.value.trim().toLocaleLowerCase();
    return query
      ? state.items.filter((item) => item.toLocaleLowerCase().includes(query))
      : state.items;
  };
  const refreshSelectAll = (items = matchingItems()) => {
    const selectedCount = items.filter((item) => state.draft.has(item)).length;
    selectAll.checked = items.length > 0 && selectedCount === items.length;
    selectAll.indeterminate = selectedCount > 0 && selectedCount < items.length;
    selectAll.disabled = items.length === 0;
  };
  const renderOptions = () => {
    const visibleItems = matchingItems();
    options.innerHTML = visibleItems.length ? visibleItems.map((item) => `
      <label class="excel-multiselect-option">
        <input type="checkbox" value="${escapeHtml(item)}" ${state.draft.has(item) ? "checked" : ""} />
        <span>${escapeHtml(item)}</span>
      </label>`).join("") : '<p class="excel-multiselect-empty">No matching results</p>';
    refreshSelectAll(visibleItems);
  };

  trigger.addEventListener("click", () => {
    setOpen(!state.open);
    if (state.open) searchInput.focus();
  });
  searchInput.addEventListener("input", renderOptions);
  options.addEventListener("change", (event) => {
    const checkbox = event.target.closest("input[type=checkbox]");
    if (!checkbox) return;
    if (checkbox.checked) state.draft.add(checkbox.value);
    else state.draft.delete(checkbox.value);
    refreshSelectAll();
  });
  selectAll.addEventListener("change", () => {
    const visibleItems = matchingItems();
    visibleItems.forEach((item) => {
      if (selectAll.checked) state.draft.add(item);
      else state.draft.delete(item);
    });
    renderOptions();
  });
  clearButton.addEventListener("click", () => {
    state.draft.clear();
    renderOptions();
  });
  menu.addEventListener("keydown", (event) => {
    if (event.key !== "Enter" || event.isComposing || event.target.closest("button")) return;
    event.preventDefault();
    applyButton.click();
  });
  applyButton.addEventListener("click", () => {
    state.active = new Set(state.draft);
    writeSource();
    refreshLabel();
    setOpen(false);
    onApply(serialize(state.active));
  });
  document.addEventListener("click", (event) => {
    if (state.open && !host.contains(event.target)) setOpen(false);
  });
  refreshLabel();

  return {
    setOptions(items) {
      state.items = [...new Set(items.map((item) => String(item).trim()).filter(Boolean))];
      state.active = new Set([...state.active].filter((item) => state.items.includes(item)));
      if (source.tagName === "SELECT") {
        source.innerHTML = state.items.map((item) => `<option value="${escapeHtml(item)}">${escapeHtml(item)}</option>`).join("");
      }
      writeSource();
      refreshLabel();
      if (state.open) renderOptions();
    },
    setValue(value) {
      state.active = new Set(splitMultiFilterValue(value).filter((item) => state.items.includes(item)));
      writeSource();
      refreshLabel();
      if (state.open) renderOptions();
    },
    getValue() {
      return serialize(state.active);
    },
  };
}

function formatMetricCell(value) {
  if (value === null || value === undefined || value === "") {
    return '<span class="missing-metric">-</span>';
  }
  return escapeHtml(Number(value).toLocaleString());
}

function createTeamIkonMtdSummary({
  prefix,
  teamLabel,
  regionLabel,
  apiPath,
  sourcePath,
  exportPath,
}) {
  const baseColumns = buildBaseTeamSummaryColumns(regionLabel);
  const teamState = {
    loaded: false,
    loading: false,
    reportDate: "",
    formattedDate: "",
    monthLabel: "",
    daysInMonth: 31,
    search: "",
    designation: "",
    zone: "",
    region: "",
    baseTown: "",
    sortBy: "",
    sortDir: "asc",
    filtersPopulated: false,
    liveCallSource: { status: "unknown", table: "", total: 0, columns: [] },
    filterOptions: {},
    deptTerritoryCode: "",
    positionCode: "",
    sasId: "",
    employeeName: "",
    empType: "",
    empCode: "",
    payrollStatus: "",
    doj: "",
  };

  const dom = {
    headingDate: $(`#${prefix}-heading-date`),
    bannerTitle: $(`#${prefix}-banner-title`),
    bannerSub: $(`#${prefix}-banner-sub`),
    reportSheet: $(`#${prefix}-report-sheet`),
    fullscreenBtn: $(`#${prefix}-fullscreen-btn`),
    groupBaseTitle: $(`#${prefix}-group-base-title`),
    groupAttTitle: $(`#${prefix}-group-att-title`),
    groupCallsTitle: $(`#${prefix}-group-calls-title`),
    table: $(`#${prefix}-summary-table`),
    colHeaderRow: $(`#${prefix}-col-header-row`),
    rowBadge: $(`#${prefix}-row-badge`),
    dateFilter: $(`#${prefix}-date-filter`),
    exportButton: $(`#${prefix}-export-excel`),
    searchInput: $(`#${prefix}-search-input`),
    desgFilter: $(`#${prefix}-desg-filter`),
    zoneFilter: $(`#${prefix}-zone-filter`),
    regionFilter: $(`#${prefix}-region-filter`),
    townFilter: $(`#${prefix}-town-filter`),
    resetButton: $(`#${prefix}-reset-filters`),
    statusBar: $(`#${prefix}-status-bar`),
    statusText: $(`#${prefix}-status-text`),
    tbody: $(`#${prefix}-summary-tbody`),
    tfootRow: $(`#${prefix}-tfoot-row`),
    managementPrintCard: $(`#${prefix}-management-print-card`),
    managementPrintTitle: $(`#${prefix}-management-print-title`),
    managementPrintBannerTitle: $(`#${prefix}-management-banner-title`),
    managementPrintBannerSub: $(`#${prefix}-management-banner-sub`),
    managementPrintRowBadge: $(`#${prefix}-management-row-badge`),
    managementPrintFullscreenBtn: $(`#${prefix}-management-fullscreen-btn`),
    managementPrintBody: $(`#${prefix}-management-print-tbody`),
    managementPrintFooter: $(`#${prefix}-management-print-tfoot`),
    fullscreenFilterToolbar: $(`#${prefix}-fullscreen-filter-toolbar`),
    fullscreenFilterGrid: $(`#${prefix}-fullscreen-filter-grid`),
    fullscreenFilterCount: $(`#${prefix}-fullscreen-filter-count`),
    fullscreenFilterChips: $(`#${prefix}-fullscreen-filter-chips`),
    fullscreenFilterToggle: $(`#${prefix}-fullscreen-filter-toggle`),
    fullscreenFilterApply: $(`#${prefix}-fullscreen-filter-apply`),
    fullscreenFilterReset: $(`#${prefix}-fullscreen-filter-reset`),
  };

  const fullscreenFilterDefinitions = [
    { stateKey: "search", label: "Search Employee / Position / SAS ID / Emp. Code", input: "search" },
    { stateKey: "designation", label: "Designation", input: "designation" },
    { stateKey: "zone", label: "Zone", input: "zone" },
    { stateKey: "region", label: "Region", input: "region" },
    { stateKey: "baseTown", label: "Base Town", input: "town" },
  ];
  const fullscreenFilterInputs = {
    reportDate: $(`#${prefix}-fs-report-date`),
    ...Object.fromEntries(fullscreenFilterDefinitions.map((definition) => [
      definition.stateKey,
      $(`#${prefix}-fs-${definition.input}`),
    ])),
  };
  const multiFilterDefinitions = [
    { stateKey: "designation", optionKey: "designations", source: dom.desgFilter, defaultLabel: "All Designations" },
    { stateKey: "zone", optionKey: "zones", source: dom.zoneFilter, defaultLabel: "All Zones" },
    { stateKey: "region", optionKey: "regions", source: dom.regionFilter, defaultLabel: "All Regions" },
    { stateKey: "baseTown", optionKey: "baseTowns", source: dom.townFilter, defaultLabel: "All Base Towns" },
  ];
  const mainMultiFilters = Object.fromEntries(multiFilterDefinitions.map((definition) => [
    definition.stateKey,
    createExcelMultiSelect(definition.source, definition.defaultLabel, (value) => {
      teamState[definition.stateKey] = value;
      syncFullscreenFilterInputs();
      loadSummary();
    }),
  ]));
  const fullscreenMultiFilters = Object.fromEntries(multiFilterDefinitions.map((definition) => [
    definition.stateKey,
    createExcelMultiSelect(
      fullscreenFilterInputs[definition.stateKey],
      definition.defaultLabel,
      renderFullscreenFilterChips,
    ),
  ]));

  function getTotalColumns(daysInMonth = teamState.daysInMonth) {
    return baseColumns.length + (daysInMonth + 1) + (daysInMonth + 2);
  }

  function setFullscreenFilterOptions(filters = teamState.filterOptions) {
    for (const definition of multiFilterDefinitions) {
      const options = filters[definition.optionKey] || [];
      mainMultiFilters[definition.stateKey].setOptions(options);
      fullscreenMultiFilters[definition.stateKey].setOptions(options);
      mainMultiFilters[definition.stateKey].setValue(teamState[definition.stateKey]);
      fullscreenMultiFilters[definition.stateKey].setValue(teamState[definition.stateKey]);
    }
  }

  function syncFullscreenFilterInputs() {
    if (!fullscreenFilterInputs.reportDate) return;
    fullscreenFilterInputs.reportDate.value = teamState.reportDate || "";
    for (const definition of fullscreenFilterDefinitions) {
      if (definition.stateKey === "search") {
        fullscreenFilterInputs.search.value = teamState.search || "";
      } else {
        fullscreenMultiFilters[definition.stateKey].setValue(teamState[definition.stateKey]);
      }
    }
    renderFullscreenFilterChips();
  }

  function readFullscreenFilterInputs() {
    const values = { reportDate: fullscreenFilterInputs.reportDate?.value || "" };
    for (const definition of fullscreenFilterDefinitions) {
      values[definition.stateKey] = definition.stateKey === "search"
        ? fullscreenFilterInputs.search?.value.trim() || ""
        : fullscreenMultiFilters[definition.stateKey].getValue();
    }
    return values;
  }

  function renderFullscreenFilterChips() {
    if (!dom.fullscreenFilterChips) return;
    const active = fullscreenFilterDefinitions.filter((definition) => {
      const value = definition.stateKey === "search"
        ? fullscreenFilterInputs.search?.value || ""
        : fullscreenMultiFilters[definition.stateKey].getValue();
      return String(value).trim();
    });
    dom.fullscreenFilterCount.textContent = `${active.length} active`;
    dom.fullscreenFilterChips.innerHTML = active.map((definition) => {
      const value = definition.stateKey === "search"
        ? fullscreenFilterInputs.search.value.trim()
        : splitMultiFilterValue(fullscreenMultiFilters[definition.stateKey].getValue()).join(", ");
      return `<span class="fmcg-filter-chip"><span>${escapeHtml(definition.label)}: ${escapeHtml(value)}</span><button type="button" data-clear-filter="${definition.stateKey}" aria-label="Clear ${escapeHtml(definition.label)}">×</button></span>`;
    }).join("");
  }

  function resetTeamFilters(load = true, resetDate = false) {
    teamState.search = "";
    teamState.designation = "";
    teamState.zone = "";
    teamState.region = "";
    teamState.baseTown = "";
    teamState.deptTerritoryCode = "";
    teamState.positionCode = "";
    teamState.sasId = "";
    teamState.employeeName = "";
    teamState.empType = "";
    teamState.empCode = "";
    teamState.payrollStatus = "";
    teamState.doj = "";
    if (resetDate) {
      teamState.reportDate = "";
      dom.dateFilter.value = teamState.reportDate;
    }
    teamState.sortBy = "";
    teamState.sortDir = "asc";
    dom.searchInput.value = "";
    for (const definition of multiFilterDefinitions) {
      mainMultiFilters[definition.stateKey].setValue("");
    }
    syncFullscreenFilterInputs();
    if (load) loadSummary();
  }

  function applyFullscreenFilters() {
    const values = readFullscreenFilterInputs();
    teamState.reportDate = values.reportDate || "";
    for (const definition of fullscreenFilterDefinitions) {
      teamState[definition.stateKey] = values[definition.stateKey];
    }
    dom.dateFilter.value = teamState.reportDate;
    dom.searchInput.value = teamState.search;
    for (const definition of multiFilterDefinitions) {
      mainMultiFilters[definition.stateKey].setValue(teamState[definition.stateKey]);
    }
    loadSummary();
  }

  function bindFullscreenFilters() {
    fullscreenFilterInputs.search?.addEventListener("input", renderFullscreenFilterChips);
    fullscreenFilterInputs.reportDate?.addEventListener("change", renderFullscreenFilterChips);
    dom.fullscreenFilterGrid?.addEventListener("keydown", (event) => {
      if (event.key !== "Enter" || event.isComposing) return;
      if (event.target.closest(".excel-multiselect-trigger, .excel-multiselect-actions")) return;
      event.preventDefault();
      applyFullscreenFilters();
    });
    dom.fullscreenFilterChips?.addEventListener("click", (event) => {
      const button = event.target.closest("button[data-clear-filter]");
      if (!button) return;
      if (button.dataset.clearFilter === "search") {
        fullscreenFilterInputs.search.value = "";
      } else {
        fullscreenMultiFilters[button.dataset.clearFilter]?.setValue("");
      }
      renderFullscreenFilterChips();
    });
    dom.fullscreenFilterToggle?.addEventListener("click", () => {
      const collapsed = dom.fullscreenFilterToolbar.classList.toggle("is-collapsed");
      dom.fullscreenFilterToggle.textContent = collapsed ? "Show Filters" : "Hide Filters";
    });
    dom.fullscreenFilterApply?.addEventListener("click", applyFullscreenFilters);
    dom.fullscreenFilterReset?.addEventListener("click", () => resetTeamFilters(true, true));
  }

  function updateTitles(formattedDate, monthLabel, daysInMonth) {
    const display = formattedDate || formatDisplayDate(teamState.reportDate) || "Latest Available";
    const mLabel = monthLabel || formatMonthLabel(teamState.reportDate);
    const days = daysInMonth || getDaysInMonthFromIso(teamState.reportDate);
    teamState.monthLabel = mLabel;
    teamState.daysInMonth = days;

    dom.headingDate.textContent = display;
    const fullTitle = `${teamLabel} Team – IKON MTD Call Performance Summary Till ${display}`;
    dom.bannerTitle.textContent = fullTitle;
    if (dom.groupBaseTitle) {
      dom.groupBaseTitle.textContent = fullTitle;
    }
    if (dom.groupAttTitle) {
      dom.groupAttTitle.colSpan = days + 1;
      dom.groupAttTitle.textContent = `Attendance (${mLabel})`;
    }
    if (dom.groupCallsTitle) {
      dom.groupCallsTitle.colSpan = days + 2;
      dom.groupCallsTitle.textContent = `Calls + Exceptions Other Calls (${mLabel})`;
    }
  }

  function freezeCellAttribute(key) {
    return FROZEN_REPORT_COLUMN_KEYS.has(key)
      ? ` data-freeze-key="${key}"`
      : "";
  }

  function updateFrozenColumns() {
    if (!dom.table || !dom.colHeaderRow) return;

    let left = 0;
    for (const key of FROZEN_REPORT_COLUMN_KEYS) {
      const header = dom.colHeaderRow.querySelector(`th[data-freeze-key="${key}"]`);
      const width = header ? header.getBoundingClientRect().width : 0;
      dom.table.style.setProperty(`--fmcg-freeze-left-${key}`, `${left}px`);
      left += width;
    }
  }

  function renderColumnHeaders(daysInMonth) {
    const ths = [];
    for (const col of baseColumns) {
      const isSorted = teamState.sortBy === col.key;
      const classes = ["sortable"];
      if (col.num) classes.push("num-col");
      if (col.groupEnd) classes.push("group-end-col");
      if (isSorted) classes.push("sorted");
      const icon = isSorted ? (teamState.sortDir === "asc" ? "↑" : "↓") : "↕";
      ths.push(`<th data-sort="${col.key}" class="${classes.join(" ")}"><span>${escapeHtml(col.label)}</span><i class="sort-icon" aria-hidden="true">${icon}</i></th>`);
    }

    for (let day = 1; day <= daysInMonth; day += 1) {
      const key = `att_day_${day}`;
      const isSorted = teamState.sortBy === key;
      const icon = isSorted ? (teamState.sortDir === "asc" ? "↑" : "↓") : "↕";
      ths.push(`<th data-sort="${key}" class="sortable att-col-th day-col${isSorted ? " sorted" : ""}"><span>${day}</span><i class="sort-icon" aria-hidden="true">${icon}</i></th>`);
    }
    for (const [key, label, isEnd] of [["attendance_total", "Attendance", true]]) {
      const isSorted = teamState.sortBy === key;
      const icon = isSorted ? (teamState.sortDir === "asc" ? "↑" : "↓") : "↕";
      ths.push(`<th data-sort="${key}" class="sortable att-col-th num-col${isEnd ? " group-end-col" : ""}${isSorted ? " sorted" : ""}"><span>${label}</span><i class="sort-icon" aria-hidden="true">${icon}</i></th>`);
    }

    for (let day = 1; day <= daysInMonth; day += 1) {
      const key = `call_day_${day}`;
      const isSorted = teamState.sortBy === key;
      const icon = isSorted ? (teamState.sortDir === "asc" ? "↑" : "↓") : "↕";
      ths.push(`<th data-sort="${key}" class="sortable call-col-th day-col${isSorted ? " sorted" : ""}"><span>${day}</span><i class="sort-icon" aria-hidden="true">${icon}</i></th>`);
    }
    for (const [key, label] of [["calls_total", "Calls"], ["avg_call", "Avg Call"]]) {
      const isSorted = teamState.sortBy === key;
      const icon = isSorted ? (teamState.sortDir === "asc" ? "↑" : "↓") : "↕";
      ths.push(`<th data-sort="${key}" class="sortable call-col-th num-col${isSorted ? " sorted" : ""}"><span>${label}</span><i class="sort-icon" aria-hidden="true">${icon}</i></th>`);
    }

    dom.colHeaderRow.innerHTML = ths.join("");
    for (const key of FROZEN_REPORT_COLUMN_KEYS) {
      const header = dom.colHeaderRow.querySelector(`th[data-sort="${key}"]`);
      if (header) header.dataset.freezeKey = key;
    }
  }

  function buildQueryParams() {
    const params = new URLSearchParams();
    if (teamState.reportDate) params.set("report_date", teamState.reportDate);
    if (teamState.search.trim()) params.set("search", teamState.search.trim());
    if (teamState.designation) params.set("designation", teamState.designation);
    if (teamState.zone) params.set("zone", teamState.zone);
    if (teamState.region) params.set("region", teamState.region);
    if (teamState.baseTown) params.set("base_town", teamState.baseTown);
    if (teamState.deptTerritoryCode) params.set("dept_territory_code", teamState.deptTerritoryCode);
    if (teamState.positionCode) params.set("position_code", teamState.positionCode);
    if (teamState.sasId) params.set("sas_id", teamState.sasId);
    if (teamState.employeeName) params.set("employee_name", teamState.employeeName);
    if (teamState.empType) params.set("emp_type", teamState.empType);
    if (teamState.empCode) params.set("emp_code", teamState.empCode);
    if (teamState.payrollStatus) params.set("payroll_status", teamState.payrollStatus);
    if (teamState.doj) params.set("doj", teamState.doj);
    if (teamState.sortBy) {
      params.set("sort_by", teamState.sortBy);
      params.set("sort_dir", teamState.sortDir);
    }
    return params.toString();
  }

  function renderFooter(totals, daysInMonth) {
    const extraCells = [];
    const attendanceDays = totals.attendance_days || {};
    const callDays = totals.call_days || {};
    for (let day = 1; day <= daysInMonth; day += 1) {
      extraCells.push(`<td class="day-col">${formatMetricCell(attendanceDays[day])}</td>`);
    }
    extraCells.push(`<td class="num-col group-end-col">${formatMetricCell(totals.attendance_total)}</td>`);
    for (let day = 1; day <= daysInMonth; day += 1) {
      extraCells.push(`<td class="day-col">${formatMetricCell(callDays[day])}</td>`);
    }
    extraCells.push(`<td class="num-col">${formatMetricCell(totals.calls_total)}</td>`);
    extraCells.push(`<td class="num-col">${formatMetricCell(totals.avg_call)}</td>`);

    dom.tfootRow.innerHTML = `
      <td colspan="6" class="tfoot-label">Total Summary</td>
      <td id="${prefix}-total-rows-label">${totals.rows.toLocaleString()} Record(s) (${totals.activeFilled.toLocaleString()} Filled · ${totals.vacant.toLocaleString()} Vacant)</td>
      <td colspan="5"></td>
      <td id="${prefix}-total-reps" class="num-col">${totals.reps.toLocaleString()}</td>
      <td id="${prefix}-total-vacant" class="num-col">${totals.vacant.toLocaleString()}</td>
      <td class="num-col">${formatMetricCell(totals.day_att_used)}</td>
      <td class="num-col">${formatMetricCell(totals.mtd_att_used)}</td>
      <td class="num-col">${formatMetricCell(totals.day_call_used)}</td>
      <td class="num-col">${formatMetricCell(totals.mtd_calls_used)}</td>
      <td class="num-col group-end-col">${formatMetricCell(totals.ikon_used)}</td>
      ${extraCells.join("")}
    `;
  }

  function renderManagementPrint(data) {
    if (!dom.managementPrintCard || !dom.managementPrintBody || !dom.managementPrintFooter) return;

    const allRows = data.rows || [];
    const managementDesignations = new Set(["ASM", "ASMTE", "CM", "RSM", "DGM", "HOD", "TE"]);
    const managementRows = allRows.filter((row) => managementDesignations.has(String(row.desg || "").trim().toUpperCase()));
    const rows = managementRows.length ? managementRows : allRows;
    const reportDate = data.reportDate || teamState.reportDate || "";
    const dateParts = reportDate.split("-").map(Number);
    const reportDay = dateParts.length === 3 && Number.isFinite(dateParts[2]) ? dateParts[2] : null;
    const yesterdayDay = reportDay && reportDay > 1 ? reportDay - 1 : null;

    const displayDate = data.formattedDate || formatDisplayDate(reportDate) || "Latest Available";
    if (dom.managementPrintTitle) {
      dom.managementPrintTitle.textContent = `${teamLabel} Team - IKON MTD Call Performance Summary Till ${displayDate}`;
    }
    if (dom.managementPrintBannerTitle) {
      dom.managementPrintBannerTitle.textContent = `${teamLabel} Team – IKON MTD Call Performance Summary · MANAGEMENT PRINT Till ${displayDate}`;
    }
    if (dom.managementPrintBannerSub) {
      dom.managementPrintBannerSub.textContent = "Management-level attendance, calls, and exception performance for the selected report filters.";
    }
    if (dom.managementPrintRowBadge) {
      dom.managementPrintRowBadge.textContent = `${rows.length.toLocaleString()} Row${rows.length === 1 ? "" : "s"}`;
    }

    const textCell = (value) => escapeHtml(value === null || value === undefined || String(value).trim() === "" ? "-" : String(value));
    const dayMetric = (row, prefixName, day) => day ? formatMetricCell(row[`${prefixName}_${day}`]) : formatMetricCell(null);
    const managementReps = (row) => row.management_reps ?? row.reps;
    const managementVacant = (row) => row.management_vacant ?? row.vacant;
    const notUsedToday = (row) => {
      const reps = Number(managementReps(row));
      const activeToday = Number(row.day_att_used);
      if (![reps, activeToday].every(Number.isFinite)) return null;
      return Math.max(reps - activeToday, 0);
    };
    const designationClass = (row) => {
      const designation = String(row.desg || "").trim().toUpperCase();
      if (designation === "ASM" || designation === "ASMTE") return "management-designation-asm";
      if (designation === "RSM") return "management-designation-rsm";
      return {
        TE: "management-designation-te",
        CM: "management-designation-cm",
        DGM: "management-designation-dgm",
        HOD: "management-designation-hod",
        PS: "management-designation-ps",
        SPOWS: "management-designation-spows",
        SPO: "management-designation-spo",
        MIO: "management-designation-mio",
      }[designation] || "";
    };

    if (!rows.length) {
      dom.managementPrintBody.innerHTML = `<tr><td colspan="18" class="fmcg-state-cell">No management records found for the selected filters.</td></tr>`;
      dom.managementPrintFooter.innerHTML = "";
      return;
    }

    dom.managementPrintBody.innerHTML = rows.map((row) => `
      <tr class="${designationClass(row)}${row.is_vacant ? " management-vacant-row" : ""}">
        <td class="designation-cell">${textCell(row.desg)}</td>
        <td>${textCell(row.zone)}</td>
        <td>${textCell(row.region)}</td>
        <td title="${textCell(row.ff_name)}">${textCell(row.ff_name)}</td>
        <td class="num-col">${formatMetricCell(managementReps(row))}</td>
        <td class="num-col">${formatMetricCell(managementVacant(row))}</td>
        <td class="num-col">${formatMetricCell(notUsedToday(row))}</td>
        <td class="num-col">${dayMetric(row, "att_day", yesterdayDay)}</td>
        <td class="num-col">${formatMetricCell(row.day_att_used)}</td>
        <td class="num-col">${formatMetricCell(row.attendance_total)}</td>
        <td class="num-col">${formatMetricCell(row.mtd_att_used)}</td>
        <td class="num-col">${dayMetric(row, "call_day", yesterdayDay)}</td>
        <td class="num-col">${dayMetric(row, "call_day", reportDay)}</td>
        <td class="num-col">${formatMetricCell(row.day_call_used)}</td>
        <td class="num-col">${formatMetricCell(row.calls_total)}</td>
        <td class="num-col">${formatMetricCell(row.mtd_calls_used)}</td>
        <td class="num-col">${formatMetricCell(row.avg_call)}</td>
        <td class="management-comment-cell"></td>
      </tr>
    `).join("");

    const sum = (getter) => {
      const values = rows.map(getter).map(Number).filter(Number.isFinite);
      return values.length ? values.reduce((total, value) => total + value, 0) : null;
    };
    const average = (getter) => {
      const values = rows.map(getter).map(Number).filter(Number.isFinite);
      return values.length ? values.reduce((total, value) => total + value, 0) / values.length : null;
    };
    const summaryValues = [
      sum((row) => managementReps(row)),
      sum((row) => managementVacant(row)),
      sum((row) => notUsedToday(row)),
      sum((row) => yesterdayDay ? row[`att_day_${yesterdayDay}`] : null),
      sum((row) => row.day_att_used),
      sum((row) => row.attendance_total),
      sum((row) => row.mtd_att_used),
      sum((row) => yesterdayDay ? row[`call_day_${yesterdayDay}`] : null),
      sum((row) => reportDay ? row[`call_day_${reportDay}`] : null),
      sum((row) => row.day_call_used),
      sum((row) => row.calls_total),
      sum((row) => row.mtd_calls_used),
      average((row) => row.avg_call),
    ];
    dom.managementPrintFooter.innerHTML = `<tr>
      <td colspan="4" class="management-total-label">TOTAL SUMMARY</td>
      ${summaryValues.slice(0, 3).map((value) => `<td class="num-col">${formatMetricCell(value)}</td>`).join("")}
      ${summaryValues.slice(3).map((value) => `<td class="num-col">${formatMetricCell(value)}</td>`).join("")}
      <td class="management-comment-cell"></td>
    </tr>`;
  }

  function renderSummary(data) {
    if (data.reportDate) {
      teamState.reportDate = data.reportDate;
      dom.dateFilter.value = data.reportDate;
      syncFullscreenFilterInputs();
    }
    const daysInMonth = data.daysInMonth || getDaysInMonthFromIso(teamState.reportDate);
    teamState.formattedDate = data.formattedDate;
    updateTitles(data.formattedDate, data.monthLabel, daysInMonth);
    renderColumnHeaders(daysInMonth);

    if (data.filters) {
      teamState.filterOptions = data.filters;
      setFullscreenFilterOptions(data.filters);
      teamState.filtersPopulated = true;
    }

    const rows = data.rows || [];
    const totals = data.totals || { rows: 0, reps: 0, vacant: 0, activeFilled: 0, unmapped: 0 };
    const totalCols = getTotalColumns(daysInMonth);

    dom.rowBadge.textContent = `${totals.rows.toLocaleString()} Row${totals.rows === 1 ? "" : "s"}`;
    if (dom.bannerSub) {
      dom.bannerSub.textContent = `${totals.rows.toLocaleString()} records · ${totals.activeFilled.toLocaleString()} filled · ${totals.vacant.toLocaleString()} vacant`;
    }
    renderFooter(totals, daysInMonth);

    if (!rows.length) {
      dom.tbody.innerHTML = `<tr><td colspan="${totalCols}" class="fmcg-state-cell">No matching ${escapeHtml(teamLabel)} employee records found for the selected filters.</td></tr>`;
    } else {
      dom.tbody.innerHTML = rows.map((row) => {
        const rowClasses = [];
        if (row.is_vacant) rowClasses.push("vacant-row");
        if (row.is_unmapped) rowClasses.push("unmapped-row");
        const designation = String(row.desg || "").trim().toUpperCase();
        const designationClasses = {
          ASM: ["designation-asm", "designation-asm-row"],
          ASMTE: ["designation-asm", "designation-asm-row"],
          RSM: ["designation-rsm", "designation-rsm-row"],
          TE: ["designation-te-row"],
          CM: ["designation-cm-row"],
          DGM: ["designation-dgm-row"],
          HOD: ["designation-hod-row"],
          PS: ["designation-ps-row"],
          SPOWS: ["designation-spows-row"],
          SPO: ["designation-spo-row"],
          MIO: ["designation-mio-row"],
        }[designation] || [];
        rowClasses.push(...designationClasses);
        const designationClass = designationClasses[0] || "";

        const ffNameTooltip = row.last_employee && row.last_employee !== row.ff_name
          ? `${row.ff_name} (Prev / IKON: ${row.last_employee})`
          : row.ff_name;

        const ffNameDisplay = row.is_vacant
          ? `<span class="vacant-pill" title="Vacant Position">⚠ ${escapeHtml(row.ff_name || "VACANT")}</span>`
          : escapeHtml(row.ff_name || "-");

        const empCodeDisplay = row.is_vacant && String(row.emp_code).toUpperCase() === "VACANT"
          ? `<span class="vacant-pill">${escapeHtml(row.emp_code)}</span>`
          : escapeHtml(row.emp_code || "-");

        const dojDisplay = row.is_vacant && String(row.doj).toUpperCase() === "RESIGNED"
          ? `<span class="vacant-count-highlight">${escapeHtml(row.doj)}</span>`
          : escapeHtml(row.doj || "-");

        const vacantDisplay = row.vacant > 0
          ? `<span class="vacant-count-highlight">${escapeHtml(String(row.vacant))}</span>`
          : escapeHtml(String(row.vacant));

        const attDayCells = [];
        for (let day = 1; day <= daysInMonth; day += 1) {
          attDayCells.push(`<td class="day-col">${formatMetricCell(row[`att_day_${day}`])}</td>`);
        }
        attDayCells.push(`<td class="num-col group-end-col">${formatMetricCell(row.attendance_total)}</td>`);

        const callDayCells = [];
        for (let day = 1; day <= daysInMonth; day += 1) {
          callDayCells.push(`<td class="day-col">${formatMetricCell(row[`call_day_${day}`])}</td>`);
        }
        callDayCells.push(`<td class="num-col">${formatMetricCell(row.calls_total)}</td>`);
        callDayCells.push(`<td class="num-col">${formatMetricCell(row.avg_call)}</td>`);

        return `<tr class="${rowClasses.join(" ")}">
          <td${freezeCellAttribute("desg")} class="designation-cell ${designationClass}" title="${escapeHtml(row.desg)}">${escapeHtml(row.desg || "-")}</td>
          <td title="${escapeHtml(row.zone)}">${escapeHtml(row.zone || "-")}</td>
          <td title="${escapeHtml(row.region)}">${escapeHtml(row.region || "-")}</td>
          <td title="${escapeHtml(row.dept_territory_code)}">${escapeHtml(row.dept_territory_code || "-")}</td>
          <td${freezeCellAttribute("position_code")} title="${escapeHtml(row.position_code)}">${escapeHtml(row.position_code || "-")}</td>
          <td${freezeCellAttribute("sas_id")} title="${escapeHtml(row.sas_id)}">${escapeHtml(row.sas_id || "-")}</td>
          <td${freezeCellAttribute("ff_name")} title="${escapeHtml(ffNameTooltip)}">${ffNameDisplay}</td>
          <td title="${escapeHtml(row.base_town)}">${escapeHtml(row.base_town || "-")}</td>
          <td title="${escapeHtml(row.emp_type)}">${escapeHtml(row.emp_type || "-")}</td>
          <td title="${escapeHtml(row.emp_code)}">${empCodeDisplay}</td>
          <td title="${escapeHtml(row.payroll_status)}">${escapeHtml(row.payroll_status || "-")}</td>
          <td title="${escapeHtml(row.doj)}">${dojDisplay}</td>
          <td class="num-col">${escapeHtml(String(row.reps))}</td>
          <td class="num-col">${vacantDisplay}</td>
          <td class="num-col">${formatMetricCell(row.day_att_used)}</td>
          <td class="num-col">${formatMetricCell(row.mtd_att_used)}</td>
          <td class="num-col">${formatMetricCell(row.day_call_used)}</td>
          <td class="num-col">${formatMetricCell(row.mtd_calls_used)}</td>
          <td class="num-col group-end-col">${formatMetricCell(row.ikon_used)}</td>
          ${attDayCells.join("")}
          ${callDayCells.join("")}
        </tr>`;
      }).join("");
    }

    renderManagementPrint(data);
    updateFrozenColumns();
    dom.statusBar.classList.remove("error");
    dom.statusText.textContent = `Showing ${totals.rows.toLocaleString()} of ${(data.totalSourceRows || totals.rows).toLocaleString()} ${teamLabel} record(s) · Reps: ${totals.reps.toLocaleString()} · Vacant: ${totals.vacant.toLocaleString()} · Unmapped: ${totals.unmapped.toLocaleString()}`;
  }

  async function loadSummary() {
    teamState.loading = true;
    dom.statusBar.classList.remove("error");
    dom.statusText.textContent = `Loading ${teamLabel} summary records from staff master and the shared call_attendence table…`;
    const days = getDaysInMonthFromIso(teamState.reportDate);
    updateTitles(formatDisplayDate(teamState.reportDate), formatMonthLabel(teamState.reportDate), days);
    renderColumnHeaders(days);
    try {
      const query = buildQueryParams();
      const summaryRequest = api(`${apiPath}?${query}`).then((response) => response.json());
      const shouldCheckSource = sourcePath && teamState.liveCallSource.status === "unknown";
      const sourceRequest = shouldCheckSource
        ? api(`${sourcePath}?limit=1`).then((response) => response.json())
        : Promise.resolve(null);
      const [summaryResult, sourceResult] = await Promise.allSettled([summaryRequest, sourceRequest]);
      if (summaryResult.status === "rejected") throw summaryResult.reason;
      const data = summaryResult.value;
      if (sourceResult.status === "fulfilled" && sourceResult.value) {
        teamState.liveCallSource = {
          status: "connected",
          table: sourceResult.value.table || "",
          total: sourceResult.value.total || 0,
          columns: sourceResult.value.columns || [],
        };
      } else if (shouldCheckSource) {
        teamState.liveCallSource = { status: "error", table: "", total: 0, columns: [] };
      }
      teamState.loaded = true;
      renderSummary(data);
    } catch (error) {
      dom.statusBar.classList.add("error");
      dom.statusText.textContent = error.message || `Failed to load ${teamLabel} summary records.`;
      dom.tbody.innerHTML = `<tr><td colspan="${getTotalColumns()}" class="fmcg-state-cell">${escapeHtml(error.message || `Unable to load ${teamLabel} summary data.`)}</td></tr>`;
    } finally {
      teamState.loading = false;
    }
  }

  let searchTimer = null;
  dom.searchInput.addEventListener("input", () => {
    teamState.search = dom.searchInput.value;
    clearTimeout(searchTimer);
    searchTimer = setTimeout(() => loadSummary(), 200);
  });

  dom.dateFilter.addEventListener("change", () => {
    teamState.reportDate = dom.dateFilter.value || "";
    loadSummary();
  });

  dom.resetButton.addEventListener("click", () => {
    resetTeamFilters(true);
  });

  dom.colHeaderRow.addEventListener("click", (event) => {
    const th = event.target.closest("th.sortable");
    if (!th) return;
    const key = th.dataset.sort;
    if (!key) return;
    if (teamState.sortBy === key) {
      teamState.sortDir = teamState.sortDir === "asc" ? "desc" : "asc";
    } else {
      teamState.sortBy = key;
      teamState.sortDir = "asc";
    }
    loadSummary();
  });

  dom.exportButton.addEventListener("click", () => {
    const query = buildQueryParams();
    const link = document.createElement("a");
    link.href = `${exportPath}?${query}`;
    link.click();
  });

  function setFullscreen(enable) {
    if (!dom.reportSheet) return;
    const isCurrent = dom.reportSheet.classList.contains("table-fullscreen-active");
    const next = typeof enable === "boolean" ? enable : !isCurrent;
    if (next === isCurrent) return;

    if (next) {
      $$(".fmcg-report-sheet.table-fullscreen-active").forEach((sheet) => {
        sheet.classList.remove("table-fullscreen-active");
      });
      $$(".table-fullscreen-btn").forEach((btn) => {
        btn.classList.remove("is-fullscreen");
        const openIcon = btn.querySelector(".fs-icon-open");
        const exitIcon = btn.querySelector(".fs-icon-exit");
        if (openIcon) openIcon.classList.remove("hidden");
        if (exitIcon) exitIcon.classList.add("hidden");
        btn.setAttribute("title", "Expand Table to Full Screen (⛶)");
        btn.setAttribute("aria-label", "Expand Table to Full Screen");
      });

      dom.reportSheet.classList.add("table-fullscreen-active");
      document.body.classList.add("has-table-fullscreen");
      syncFullscreenFilterInputs();
      if (dom.fullscreenBtn) {
        dom.fullscreenBtn.classList.add("is-fullscreen");
        const openIcon = dom.fullscreenBtn.querySelector(".fs-icon-open");
        const exitIcon = dom.fullscreenBtn.querySelector(".fs-icon-exit");
        if (openIcon) openIcon.classList.add("hidden");
        if (exitIcon) exitIcon.classList.remove("hidden");
        dom.fullscreenBtn.setAttribute("title", "Exit Full Screen (Esc)");
        dom.fullscreenBtn.setAttribute("aria-label", "Exit Full Screen");
      }
      if (dom.bannerSub) dom.bannerSub.classList.remove("hidden");
    } else {
      dom.reportSheet.classList.remove("table-fullscreen-active");
      document.body.classList.remove("has-table-fullscreen");
      if (dom.fullscreenBtn) {
        dom.fullscreenBtn.classList.remove("is-fullscreen");
        const openIcon = dom.fullscreenBtn.querySelector(".fs-icon-open");
        const exitIcon = dom.fullscreenBtn.querySelector(".fs-icon-exit");
        if (openIcon) openIcon.classList.remove("hidden");
        if (exitIcon) exitIcon.classList.add("hidden");
        dom.fullscreenBtn.setAttribute("title", "Expand Table to Full Screen (⛶)");
        dom.fullscreenBtn.setAttribute("aria-label", "Expand Table to Full Screen");
      }
      if (dom.bannerSub) dom.bannerSub.classList.add("hidden");
    }
    window.requestAnimationFrame(updateFrozenColumns);
  }

  if (dom.fullscreenBtn) {
    dom.fullscreenBtn.addEventListener("click", () => setFullscreen());
  }
  bindFullscreenFilters();

  const refreshFrozenColumns = () => {
    if (window.requestAnimationFrame) window.requestAnimationFrame(updateFrozenColumns);
    else updateFrozenColumns();
  };
  window.addEventListener("resize", refreshFrozenColumns);
  if (window.ResizeObserver && dom.table) {
    const freezeResizeObserver = new ResizeObserver(refreshFrozenColumns);
    freezeResizeObserver.observe(dom.table);
  }

  return {
    state: teamState,
    elements: dom,
    load: loadSummary,
    toggleFullscreen: setFullscreen,
  };
}

const fmcgSummary = createTeamIkonMtdSummary({
  prefix: "fmcg",
  teamLabel: "FMCG",
  regionLabel: "FMCG Region",
  apiPath: "/api/fmcg/ikon-mtd-summary",
  sourcePath: "/api/fmcg/ikon-mtd-call-source",
  exportPath: "/api/fmcg/ikon-mtd-summary/export",
});
const fmcgState = fmcgSummary.state;
const fmcgElements = fmcgSummary.elements;
const loadFmcgSummary = fmcgSummary.load;

function setManagementPrintFullscreen(summary, enable) {
  const { elements } = summary;
  const card = elements.managementPrintCard;
  const button = elements.managementPrintFullscreenBtn;
  if (!card || !button) return;

  const isCurrent = card.classList.contains("table-fullscreen-active");
  const next = typeof enable === "boolean" ? enable : !isCurrent;
  if (next === isCurrent) return;

  const setButtonState = (target, active) => {
    target.classList.toggle("is-fullscreen", active);
    const openIcon = target.querySelector(".fs-icon-open");
    const exitIcon = target.querySelector(".fs-icon-exit");
    if (openIcon) openIcon.classList.toggle("hidden", active);
    if (exitIcon) exitIcon.classList.toggle("hidden", !active);
    target.setAttribute("title", active ? "Exit Full Screen (Esc)" : "Expand Management Print to Full Screen (⛶)");
    target.setAttribute("aria-label", active ? "Exit Full Screen" : "Expand Management Print to Full Screen");
  };

  if (next) {
    $$(".fmcg-report-sheet.table-fullscreen-active").forEach((sheet) => sheet.classList.remove("table-fullscreen-active"));
    $$(".table-fullscreen-btn").forEach((target) => {
      target.classList.remove("is-fullscreen");
      const openIcon = target.querySelector(".fs-icon-open");
      const exitIcon = target.querySelector(".fs-icon-exit");
      if (openIcon) openIcon.classList.remove("hidden");
      if (exitIcon) exitIcon.classList.add("hidden");
    });
    card.classList.add("table-fullscreen-active");
    document.body.classList.add("has-table-fullscreen");
    elements.managementPrintBannerSub?.classList.remove("hidden");
    setButtonState(button, true);
  } else {
    card.classList.remove("table-fullscreen-active");
    document.body.classList.remove("has-table-fullscreen");
    elements.managementPrintBannerSub?.classList.add("hidden");
    setButtonState(button, false);
  }
}

fmcgElements.managementPrintFullscreenBtn?.addEventListener("click", () => setManagementPrintFullscreen(fmcgSummary));

const pharmaSummary = createTeamIkonMtdSummary({
  prefix: "pharma",
  teamLabel: "PHARMA",
  regionLabel: "PHARMA Region",
  apiPath: "/api/pharma/ikon-mtd-summary",
  sourcePath: "/api/pharma/ikon-mtd-call-source",
  exportPath: "/api/pharma/ikon-mtd-summary/export",
});
const pharmaState = pharmaSummary.state;
const pharmaElements = pharmaSummary.elements;
const loadPharmaSummary = pharmaSummary.load;

pharmaElements.managementPrintFullscreenBtn?.addEventListener("click", () => setManagementPrintFullscreen(pharmaSummary));

// ---------------------------------------------------------------------------
// FF Master Staff – PostgreSQL Database Management (CRUD)
// ---------------------------------------------------------------------------

const staffCrudState = {
  schemaLoaded: false,
  loading: false,
  saving: false,
  deleting: false,
  writeToken: "",
  primaryKey: "position_code",
  columns: [],
  options: {},
  rows: [],
  total: 0,
  page: 1,
  pageSize: 25,
  totalPages: 1,
  search: "",
  teamName: "",
  designationName: "",
  zoneName: "",
  regionName: "",
  baseTown: "",
  sortBy: "position_code",
  sortDir: "asc",
  modalMode: "add", // "add" | "edit"
  editingRecord: null,
  deletingRecord: null,
};

const staffElements = {
  notice: $("#staff-notice"),
  refreshBtn: $("#staff-refresh-btn"),
  addBtn: $("#staff-add-btn"),
  searchInput: $("#staff-search-input"),
  teamFilter: $("#staff-team-filter"),
  desgFilter: $("#staff-desg-filter"),
  zoneFilter: $("#staff-zone-filter"),
  regionFilter: $("#staff-region-filter"),
  townFilter: $("#staff-town-filter"),
  pageSizeSelect: $("#staff-page-size"),
  resetFiltersBtn: $("#staff-reset-filters"),
  statusBar: $("#staff-status-bar"),
  statusText: $("#staff-status-text"),
  schemaBadge: $("#staff-schema-badge"),
  totalBadge: $("#staff-total-badge"),
  headRow: $("#staff-table-head-row"),
  tbody: $("#staff-table-tbody"),
  pageInfo: $("#staff-page-info"),
  firstPageBtn: $("#staff-first-page"),
  prevPageBtn: $("#staff-prev-page"),
  pageNumbers: $("#staff-page-numbers"),
  nextPageBtn: $("#staff-next-page"),
  lastPageBtn: $("#staff-last-page"),
  // Add/Edit Modal
  modalBackdrop: $("#staff-modal-backdrop"),
  modalEyebrow: $("#staff-modal-eyebrow"),
  modalTitle: $("#staff-modal-title"),
  modalCloseBtn: $("#staff-modal-close"),
  modalForm: $("#staff-modal-form"),
  modalError: $("#staff-modal-error"),
  modalBody: $("#staff-modal-body"),
  modalCancelBtn: $("#staff-modal-cancel"),
  modalSaveBtn: $("#staff-modal-save"),
  // Delete Modal
  deleteBackdrop: $("#staff-delete-backdrop"),
  deleteCloseBtn: $("#staff-delete-close"),
  deleteError: $("#staff-delete-error"),
  deleteDetails: $("#staff-delete-details"),
  deleteCancelBtn: $("#staff-delete-cancel"),
  deleteConfirmBtn: $("#staff-delete-confirm"),
};

let staffNoticeTimer = null;
function showStaffNotice(message, isError = false) {
  clearTimeout(staffNoticeTimer);
  staffElements.notice.textContent = message;
  staffElements.notice.classList.toggle("error", isError);
  staffElements.notice.classList.remove("hidden");
  if (!isError) {
    staffNoticeTimer = setTimeout(() => {
      staffElements.notice.classList.add("hidden");
    }, 6000);
  }
}

async function loadStaffSchema(force = false) {
  if (staffCrudState.schemaLoaded && !force) return;
  const response = await api("/api/ff-master-staff/schema");
  const data = await response.json();
  staffCrudState.columns = data.columns || [];
  staffCrudState.options = data.options || {};
  staffCrudState.writeToken = data.writeToken || "";
  staffCrudState.primaryKey = data.primaryKey || "position_code";
  staffCrudState.schemaLoaded = true;

  staffElements.schemaBadge.textContent = `${staffCrudState.columns.length} Columns · Primary Identifier: ${staffCrudState.primaryKey}`;

  populateSelectOptions(staffElements.teamFilter, "All Teams", staffCrudState.options.team_name || [], staffCrudState.teamName);
  populateSelectOptions(staffElements.desgFilter, "All Designations", staffCrudState.options.designation_name || [], staffCrudState.designationName);
  populateSelectOptions(staffElements.zoneFilter, "All Zones", staffCrudState.options.zone_name || [], staffCrudState.zoneName);
  populateSelectOptions(staffElements.regionFilter, "All Regions", staffCrudState.options.region_name || [], staffCrudState.regionName);
  populateSelectOptions(staffElements.townFilter, "All Base Towns", staffCrudState.options.base_town || [], staffCrudState.baseTown);
}

function renderStaffTableHeaders() {
  const ths = ['<th class="staff-actions-th">Actions</th>'];
  for (const col of staffCrudState.columns) {
    const isSorted = staffCrudState.sortBy === col.name;
    const isNum = col.inputType === "number";
    const classes = ["sortable"];
    if (col.name === "position_code") classes.push("staff-poscode-th");
    if (isNum) classes.push("num-col");
    if (isSorted) classes.push("sorted");
    const icon = isSorted ? (staffCrudState.sortDir === "asc" ? "↑" : "↓") : "↕";
    const pkBadge = col.isPrimaryKey ? '<span class="staff-pk-badge" title="Unique Record Identifier">KEY</span>' : "";
    ths.push(
      `<th data-sort="${escapeHtml(col.name)}" class="${classes.join(" ")}" title="${escapeHtml(col.name)} (${escapeHtml(col.dataType)}${col.maxLength ? `(${col.maxLength})` : ""})">` +
      `<span>${escapeHtml(col.label)}</span>${pkBadge}<i class="sort-icon" aria-hidden="true">${icon}</i></th>`
    );
  }
  staffElements.headRow.innerHTML = ths.join("");
}

function renderStaffTableRows() {
  const colSpan = Math.max(staffCrudState.columns.length + 1, 2);
  if (!staffCrudState.rows.length) {
    staffElements.tbody.innerHTML = `<tr><td colspan="${colSpan}" class="fmcg-state-cell">No employee records in public.ff_master_staff matched your filter criteria.</td></tr>`;
    return;
  }

  staffElements.tbody.innerHTML = staffCrudState.rows.map((row, rowIdx) => {
    const cells = [
      `<td class="staff-actions-td">
        <div class="staff-row-actions">
          <button type="button" class="staff-action-btn staff-action-edit" data-row-index="${rowIdx}" title="Edit record">✎ Edit</button>
          <button type="button" class="staff-action-btn staff-action-delete" data-row-index="${rowIdx}" title="Delete record">✕ Del</button>
        </div>
      </td>`,
    ];

    for (const col of staffCrudState.columns) {
      const val = row[col.name];
      const isNum = col.inputType === "number";
      const tdClasses = [];
      if (col.name === "position_code") tdClasses.push("staff-poscode-td");
      if (isNum) tdClasses.push("num-col");

      if (val === null || val === undefined || val === "") {
        tdClasses.push("staff-null-cell");
        cells.push(`<td class="${tdClasses.join(" ")}" title="NULL">-</td>`);
      } else {
        const strVal = String(val);
        cells.push(`<td class="${tdClasses.join(" ")}" title="${escapeHtml(strVal)}">${escapeHtml(strVal)}</td>`);
      }
    }

    return `<tr>${cells.join("")}</tr>`;
  }).join("");
}

function renderStaffPagination() {
  const { page, pageSize, total, totalPages, rows } = staffCrudState;
  const startRow = total === 0 ? 0 : (page - 1) * pageSize + 1;
  const endRow = total === 0 ? 0 : Math.min(startRow + rows.length - 1, total);

  staffElements.totalBadge.textContent = `${total.toLocaleString()} Record${total === 1 ? "" : "s"}`;
  staffElements.pageInfo.textContent = total === 0
    ? "Showing 0–0 of 0 records | Page 1 of 1"
    : `Showing ${startRow.toLocaleString()}–${endRow.toLocaleString()} of ${total.toLocaleString()} records | Page ${page} of ${totalPages}`;

  staffElements.firstPageBtn.disabled = page <= 1 || staffCrudState.loading;
  staffElements.prevPageBtn.disabled = page <= 1 || staffCrudState.loading;
  staffElements.nextPageBtn.disabled = page >= totalPages || staffCrudState.loading;
  staffElements.lastPageBtn.disabled = page >= totalPages || staffCrudState.loading;

  const pageItems = [];
  const windowStart = Math.max(1, page - 2);
  const windowEnd = Math.min(totalPages, page + 2);

  if (windowStart > 1) {
    pageItems.push(`<button type="button" class="staff-page-num-btn${page === 1 ? " active" : ""}" data-page="1">1</button>`);
    if (windowStart > 2) {
      pageItems.push('<span class="staff-page-ellipsis" aria-hidden="true">…</span>');
    }
  }

  for (let p = windowStart; p <= windowEnd; p += 1) {
    pageItems.push(
      `<button type="button" class="staff-page-num-btn${p === page ? " active" : ""}" data-page="${p}">${p}</button>`
    );
  }

  if (windowEnd < totalPages) {
    if (windowEnd < totalPages - 1) {
      pageItems.push('<span class="staff-page-ellipsis" aria-hidden="true">…</span>');
    }
    pageItems.push(`<button type="button" class="staff-page-num-btn${page === totalPages ? " active" : ""}" data-page="${totalPages}">${totalPages}</button>`);
  }

  staffElements.pageNumbers.innerHTML = pageItems.join("");
}

function buildStaffQueryParams() {
  const params = new URLSearchParams();
  params.set("page", String(staffCrudState.page));
  params.set("page_size", String(staffCrudState.pageSize));
  if (staffCrudState.search.trim()) params.set("search", staffCrudState.search.trim());
  if (staffCrudState.teamName) params.set("team_name", staffCrudState.teamName);
  if (staffCrudState.designationName) params.set("designation_name", staffCrudState.designationName);
  if (staffCrudState.zoneName) params.set("zone_name", staffCrudState.zoneName);
  if (staffCrudState.regionName) params.set("region_name", staffCrudState.regionName);
  if (staffCrudState.baseTown) params.set("base_town", staffCrudState.baseTown);
  if (staffCrudState.sortBy) {
    params.set("sort_by", staffCrudState.sortBy);
    params.set("sort_dir", staffCrudState.sortDir);
  }
  return params.toString();
}

async function loadStaffRecords({ refreshSchema = false } = {}) {
  staffCrudState.loading = true;
  staffElements.statusBar.classList.remove("error");
  staffElements.statusText.textContent = "Loading records from PostgreSQL public.ff_master_staff…";
  staffElements.refreshBtn.disabled = true;

  try {
    await loadStaffSchema(refreshSchema);
    renderStaffTableHeaders();

    const query = buildStaffQueryParams();
    const response = await api(`/api/ff-master-staff?${query}`);
    const data = await response.json();

    staffCrudState.rows = data.rows || [];
    staffCrudState.total = Number(data.total || 0);
    staffCrudState.page = Number(data.page || 1);
    staffCrudState.pageSize = Number(data.pageSize || staffCrudState.pageSize);
    staffCrudState.totalPages = Number(data.totalPages || 1);

    renderStaffTableHeaders();
    renderStaffTableRows();
    renderStaffPagination();

    staffElements.statusText.textContent = `Connected to ff-management-db (public.ff_master_staff) · ${staffCrudState.total.toLocaleString()} matching record(s)`;
  } catch (error) {
    staffElements.statusBar.classList.add("error");
    staffElements.statusText.textContent = error.message || "Failed to load records from PostgreSQL.";
    const colSpan = Math.max(staffCrudState.columns.length + 1, 2);
    staffElements.tbody.innerHTML = `<tr><td colspan="${colSpan}" class="fmcg-state-cell">${escapeHtml(error.message || "Unable to load database records.")}</td></tr>`;
  } finally {
    staffCrudState.loading = false;
    staffElements.refreshBtn.disabled = false;
    renderStaffPagination();
  }
}

function buildStaffModalForm(initialValues = {}) {
  const sections = new Map();
  for (const col of staffCrudState.columns) {
    const secName = col.section || "Database Columns";
    if (!sections.has(secName)) sections.set(secName, []);
    sections.get(secName).push(col);
  }

  const sectionHtml = [];
  for (const [sectionTitle, cols] of sections.entries()) {
    const fieldsHtml = cols.map((col) => {
      const rawVal = initialValues[col.name];
      const val = rawVal === null || rawVal === undefined ? "" : String(rawVal);
      const listId = `staff-datalist-${col.name}`;
      const acOptions = staffCrudState.options[col.name] || [];
      const hasDatalist = acOptions.length > 0 && col.inputType !== "number";
      const datalistHtml = hasDatalist
        ? `<datalist id="${escapeHtml(listId)}">${acOptions.map((opt) => `<option value="${escapeHtml(opt)}"></option>`).join("")}</datalist>`
        : "";

      const typeHint = col.inputType === "number"
        ? "integer"
        : col.maxLength
          ? `varchar(${col.maxLength})`
          : "text";

      const placeholder = col.name === "doj"
        ? "e.g. 12-Oct-23 or YYYY-MM-DD"
        : col.isPrimaryKey
          ? "Required unique Position Code"
          : "NULL (empty)";

      return `<label class="staff-field-label">
        <span class="staff-field-meta">
          <span>${escapeHtml(col.label)}${col.required ? ' <abbr class="staff-required-star" title="Required">*</abbr>' : ""}</span>
          <small class="staff-type-hint">${escapeHtml(typeHint)}</small>
        </span>
        <input
          type="${col.inputType === "number" ? "number" : "text"}"
          name="${escapeHtml(col.name)}"
          value="${escapeHtml(val)}"
          ${col.inputType === "number" ? 'step="1"' : ""}
          ${col.maxLength ? `maxlength="${col.maxLength}"` : ""}
          ${col.required ? "required" : ""}
          ${hasDatalist ? `list="${escapeHtml(listId)}"` : ""}
          placeholder="${escapeHtml(placeholder)}"
          autocomplete="off"
        />
        ${datalistHtml}
      </label>`;
    }).join("");

    sectionHtml.push(`
      <div class="staff-form-section">
        <h3 class="staff-form-section-title">${escapeHtml(sectionTitle)}</h3>
        <div class="staff-form-grid">${fieldsHtml}</div>
      </div>
    `);
  }

  staffElements.modalBody.innerHTML = sectionHtml.join("");
}

function openStaffModal(mode, record = null) {
  staffCrudState.modalMode = mode;
  staffCrudState.editingRecord = record;
  staffElements.modalError.classList.add("hidden");
  staffElements.modalError.textContent = "";

  if (mode === "edit" && record) {
    staffElements.modalEyebrow.textContent = `EDIT RECORD · ${record.position_code || record._row_id}`;
    staffElements.modalTitle.textContent = `Edit Employee — ${record.employe_name || record.last_employee || record.position_code || "Record"}`;
    staffElements.modalSaveBtn.textContent = "Save Changes";
    buildStaffModalForm(record);
  } else {
    staffElements.modalEyebrow.textContent = "NEW DATABASE RECORD · PUBLIC.FF_MASTER_STAFF";
    staffElements.modalTitle.textContent = "Add New Employee";
    staffElements.modalSaveBtn.textContent = "Add Employee";
    buildStaffModalForm({});
  }

  staffElements.modalBackdrop.classList.remove("hidden");
  const firstInput = $("input", staffElements.modalBody);
  if (firstInput) setTimeout(() => firstInput.focus(), 30);
}

function closeStaffModal() {
  if (staffCrudState.saving) return;
  staffElements.modalBackdrop.classList.add("hidden");
  staffCrudState.editingRecord = null;
}

function collectAndValidateStaffForm() {
  const values = {};
  $$("input.input-invalid", staffElements.modalForm).forEach((el) => el.classList.remove("input-invalid"));

  for (const col of staffCrudState.columns) {
    const input = $(`input[name="${col.name}"]`, staffElements.modalForm);
    if (!input) continue;
    const raw = input.value.trim();

    if (col.required && !raw) {
      input.classList.add("input-invalid");
      input.focus();
      throw new Error(`'${col.label}' (${col.name}) is required.`);
    }

    if (raw === "") {
      values[col.name] = null;
      continue;
    }

    if (col.inputType === "number") {
      if (!/^-?\d+$/.test(raw)) {
        input.classList.add("input-invalid");
        input.focus();
        throw new Error(`'${col.label}' (${col.name}) must be a valid whole number.`);
      }
      values[col.name] = Number.parseInt(raw, 10);
    } else {
      if (col.maxLength && raw.length > col.maxLength) {
        input.classList.add("input-invalid");
        input.focus();
        throw new Error(`'${col.label}' (${col.name}) cannot exceed ${col.maxLength} characters.`);
      }
      values[col.name] = raw;
    }
  }
  return values;
}

async function submitStaffModalForm(event) {
  event.preventDefault();
  if (staffCrudState.saving) return;

  staffElements.modalError.classList.add("hidden");
  let values;
  try {
    values = collectAndValidateStaffForm();
  } catch (validationError) {
    staffElements.modalError.textContent = validationError.message;
    staffElements.modalError.classList.remove("hidden");
    return;
  }

  staffCrudState.saving = true;
  staffElements.modalSaveBtn.disabled = true;
  staffElements.modalCancelBtn.disabled = true;
  const originalBtnLabel = staffElements.modalSaveBtn.textContent;
  staffElements.modalSaveBtn.textContent = "Saving…";

  try {
    const headers = {
      "Content-Type": "application/json",
      "X-FF-Write-Token": staffCrudState.writeToken,
    };

    let response;
    if (staffCrudState.modalMode === "edit" && staffCrudState.editingRecord) {
      const targetId = encodeURIComponent(staffCrudState.editingRecord._row_id);
      response = await api(`/api/ff-master-staff/${targetId}`, {
        method: "PUT",
        headers,
        body: JSON.stringify({
          values,
          expectedVersion: staffCrudState.editingRecord._version,
        }),
      });
    } else {
      response = await api("/api/ff-master-staff", {
        method: "POST",
        headers,
        body: JSON.stringify({ values }),
      });
    }

    const result = await response.json();
    staffCrudState.saving = false;
    closeStaffModal();
    showStaffNotice(result.message || "Record saved successfully.");
    await loadStaffRecords({ refreshSchema: true });
    if (fmcgState.loaded) {
      fmcgState.filtersPopulated = false;
      loadFmcgSummary();
    }
    if (pharmaState.loaded) {
      pharmaState.filtersPopulated = false;
      loadPharmaSummary();
    }
  } catch (error) {
    staffElements.modalError.textContent = error.message || "Unable to save record.";
    staffElements.modalError.classList.remove("hidden");
  } finally {
    staffCrudState.saving = false;
    staffElements.modalSaveBtn.disabled = false;
    staffElements.modalCancelBtn.disabled = false;
    staffElements.modalSaveBtn.textContent = originalBtnLabel;
  }
}

function openDeleteStaffModal(record) {
  staffCrudState.deletingRecord = record;
  staffElements.deleteError.classList.add("hidden");
  staffElements.deleteError.textContent = "";

  const summaryFields = [
    ["Position Code", record.position_code || record._row_id],
    ["Employee Name", record.employe_name || record.last_employee || "-"],
    ["Emp. Code (E-Code)", record.e_code || "-"],
    ["SAS ID", record.sas_id || "-"],
    ["Designation", record.designation_name || "-"],
    ["Team", record.team_name || "-"],
    ["Zone / Region", [record.zone_name, record.region_name].filter(Boolean).join(" · ") || "-"],
    ["Base Town", record.base_town || "-"],
  ];

  staffElements.deleteDetails.innerHTML = summaryFields
    .map(
      ([label, val]) =>
        `<div class="delete-summary-item"><small>${escapeHtml(label)}</small><strong>${escapeHtml(String(val))}</strong></div>`
    )
    .join("");

  staffElements.deleteBackdrop.classList.remove("hidden");
}

function closeDeleteStaffModal() {
  if (staffCrudState.deleting) return;
  staffElements.deleteBackdrop.classList.add("hidden");
  staffCrudState.deletingRecord = null;
}

async function confirmDeleteStaffRecord() {
  const record = staffCrudState.deletingRecord;
  if (!record || staffCrudState.deleting) return;

  staffCrudState.deleting = true;
  staffElements.deleteConfirmBtn.disabled = true;
  staffElements.deleteCancelBtn.disabled = true;
  staffElements.deleteConfirmBtn.textContent = "Deleting…";
  staffElements.deleteError.classList.add("hidden");

  try {
    const targetId = encodeURIComponent(record._row_id);
    const versionParam = record._version ? `?expected_version=${encodeURIComponent(record._version)}` : "";
    const response = await api(`/api/ff-master-staff/${targetId}${versionParam}`, {
      method: "DELETE",
      headers: {
        "X-FF-Write-Token": staffCrudState.writeToken,
      },
    });
    const result = await response.json();
    staffCrudState.deleting = false;
    closeDeleteStaffModal();
    showStaffNotice(result.message || "Record deleted from public.ff_master_staff.");
    await loadStaffRecords({ refreshSchema: true });
    if (fmcgState.loaded) {
      fmcgState.filtersPopulated = false;
      loadFmcgSummary();
    }
    if (pharmaState.loaded) {
      pharmaState.filtersPopulated = false;
      loadPharmaSummary();
    }
  } catch (error) {
    staffElements.deleteError.textContent = error.message || "Unable to delete record.";
    staffElements.deleteError.classList.remove("hidden");
  } finally {
    staffCrudState.deleting = false;
    staffElements.deleteConfirmBtn.disabled = false;
    staffElements.deleteCancelBtn.disabled = false;
    staffElements.deleteConfirmBtn.textContent = "Delete Permanently";
  }
}

// Event Listeners for FF Master Staff CRUD
let staffSearchTimer = null;
staffElements.searchInput.addEventListener("input", () => {
  staffCrudState.search = staffElements.searchInput.value;
  staffCrudState.page = 1;
  clearTimeout(staffSearchTimer);
  staffSearchTimer = setTimeout(() => loadStaffRecords(), 220);
});

staffElements.teamFilter.addEventListener("change", () => {
  staffCrudState.teamName = staffElements.teamFilter.value;
  staffCrudState.page = 1;
  loadStaffRecords();
});

staffElements.desgFilter.addEventListener("change", () => {
  staffCrudState.designationName = staffElements.desgFilter.value;
  staffCrudState.page = 1;
  loadStaffRecords();
});

staffElements.zoneFilter.addEventListener("change", () => {
  staffCrudState.zoneName = staffElements.zoneFilter.value;
  staffCrudState.page = 1;
  loadStaffRecords();
});

staffElements.regionFilter.addEventListener("change", () => {
  staffCrudState.regionName = staffElements.regionFilter.value;
  staffCrudState.page = 1;
  loadStaffRecords();
});

staffElements.townFilter.addEventListener("change", () => {
  staffCrudState.baseTown = staffElements.townFilter.value;
  staffCrudState.page = 1;
  loadStaffRecords();
});

staffElements.pageSizeSelect.addEventListener("change", () => {
  staffCrudState.pageSize = Number(staffElements.pageSizeSelect.value) || 25;
  staffCrudState.page = 1;
  loadStaffRecords();
});

staffElements.resetFiltersBtn.addEventListener("click", () => {
  staffCrudState.search = "";
  staffCrudState.teamName = "";
  staffCrudState.designationName = "";
  staffCrudState.zoneName = "";
  staffCrudState.regionName = "";
  staffCrudState.baseTown = "";
  staffCrudState.sortBy = "position_code";
  staffCrudState.sortDir = "asc";
  staffCrudState.page = 1;
  staffElements.searchInput.value = "";
  staffElements.teamFilter.value = "";
  staffElements.desgFilter.value = "";
  staffElements.zoneFilter.value = "";
  staffElements.regionFilter.value = "";
  staffElements.townFilter.value = "";
  loadStaffRecords();
});

staffElements.refreshBtn.addEventListener("click", () => {
  loadStaffRecords({ refreshSchema: true });
});

staffElements.addBtn.addEventListener("click", async () => {
  try {
    await loadStaffSchema();
    openStaffModal("add");
  } catch (error) {
    showStaffNotice(error.message || "Failed to load table schema.", true);
  }
});

staffElements.headRow.addEventListener("click", (event) => {
  const th = event.target.closest("th.sortable");
  if (!th) return;
  const col = th.dataset.sort;
  if (!col) return;
  if (staffCrudState.sortBy === col) {
    staffCrudState.sortDir = staffCrudState.sortDir === "asc" ? "desc" : "asc";
  } else {
    staffCrudState.sortBy = col;
    staffCrudState.sortDir = "asc";
  }
  staffCrudState.page = 1;
  loadStaffRecords();
});

staffElements.tbody.addEventListener("click", (event) => {
  const editBtn = event.target.closest(".staff-action-edit");
  if (editBtn) {
    const idx = Number(editBtn.dataset.rowIndex);
    const record = staffCrudState.rows[idx];
    if (record) openStaffModal("edit", record);
    return;
  }
  const delBtn = event.target.closest(".staff-action-delete");
  if (delBtn) {
    const idx = Number(delBtn.dataset.rowIndex);
    const record = staffCrudState.rows[idx];
    if (record) openDeleteStaffModal(record);
  }
});

function goToStaffPage(targetPage) {
  if (targetPage < 1 || targetPage > staffCrudState.totalPages || targetPage === staffCrudState.page) return;
  staffCrudState.page = targetPage;
  const scrollWrap = $(".staff-table-scroll");
  if (scrollWrap) scrollWrap.scrollTop = 0;
  loadStaffRecords();
}

staffElements.firstPageBtn.addEventListener("click", () => {
  goToStaffPage(1);
});

staffElements.prevPageBtn.addEventListener("click", () => {
  goToStaffPage(staffCrudState.page - 1);
});

staffElements.nextPageBtn.addEventListener("click", () => {
  goToStaffPage(staffCrudState.page + 1);
});

staffElements.lastPageBtn.addEventListener("click", () => {
  goToStaffPage(staffCrudState.totalPages);
});

staffElements.pageNumbers.addEventListener("click", (event) => {
  const btn = event.target.closest(".staff-page-num-btn");
  if (!btn) return;
  const targetPage = Number(btn.dataset.page);
  if (targetPage) goToStaffPage(targetPage);
});

staffElements.modalCloseBtn.addEventListener("click", closeStaffModal);
staffElements.modalCancelBtn.addEventListener("click", closeStaffModal);
staffElements.modalForm.addEventListener("submit", submitStaffModalForm);

staffElements.deleteCloseBtn.addEventListener("click", closeDeleteStaffModal);
staffElements.deleteCancelBtn.addEventListener("click", closeDeleteStaffModal);
staffElements.deleteConfirmBtn.addEventListener("click", confirmDeleteStaffRecord);

window.addEventListener("keydown", (event) => {
  if (event.key === "Escape") {
    const activeFsSheet = $(".fmcg-report-sheet.table-fullscreen-active");
    if (activeFsSheet) {
      event.preventDefault();
      const fsBtn = activeFsSheet.querySelector(".table-fullscreen-btn");
      if (fsBtn) {
        fsBtn.click();
      } else {
        activeFsSheet.classList.remove("table-fullscreen-active");
        document.body.classList.remove("has-table-fullscreen");
      }
      return;
    }
    if (!staffElements.deleteBackdrop.classList.contains("hidden")) {
      closeDeleteStaffModal();
    } else if (!staffElements.modalBackdrop.classList.contains("hidden")) {
      closeStaffModal();
    }
  }
});

$$(".primary-tab").forEach((tab) => tab.addEventListener("click", () => {
  $$(".fmcg-report-sheet.table-fullscreen-active").forEach((sheet) => {
    sheet.classList.remove("table-fullscreen-active");
  });
  document.body.classList.remove("has-table-fullscreen");
  $$(".table-fullscreen-btn").forEach((btn) => {
    btn.classList.remove("is-fullscreen");
    const openIcon = btn.querySelector(".fs-icon-open");
    const exitIcon = btn.querySelector(".fs-icon-exit");
    if (openIcon) openIcon.classList.remove("hidden");
    if (exitIcon) exitIcon.classList.add("hidden");
    btn.setAttribute("title", "Expand Table to Full Screen (⛶)");
    btn.setAttribute("aria-label", "Expand Table to Full Screen");
  });

  $$(".primary-tab").forEach((item) => item.classList.toggle("active", item === tab));
  $("#attendance-page").classList.toggle("hidden", tab.dataset.page !== "attendance");
  $("#fmcg-ikon-mtd-page").classList.toggle("hidden", tab.dataset.page !== "fmcg-ikon-mtd");
  $("#pharma-ikon-mtd-page").classList.toggle("hidden", tab.dataset.page !== "pharma-ikon-mtd");
  $("#editor-page").classList.toggle("hidden", tab.dataset.page !== "editor");
  if (tab.dataset.page === "fmcg-ikon-mtd" && !fmcgState.loaded && !fmcgState.loading) {
    loadFmcgSummary();
  }
  if (tab.dataset.page === "pharma-ikon-mtd" && !pharmaState.loaded && !pharmaState.loading) {
    loadPharmaSummary();
  }
  if (tab.dataset.page === "editor" && !staffCrudState.schemaLoaded && !staffCrudState.loading) {
    loadStaffRecords();
  }
}));

loadFmcgSummary();
loadPharmaSummary();
loadStaffRecords();

// ---------------------------------------------------------------------------
// IKON Call Performance Dashboard (First Tab)
// Analytics + KPI Cards + Line/Bar/Pie Charts + Interactive DB Tables
// ---------------------------------------------------------------------------

const dashState = {
  loading: false,
  loaded: false,
  filtersPopulated: false,
  period: "monthly", // "daily" | "weekly" | "monthly"
  trendMode: "daily", // "daily" | "weekly" | "monthly"
  lineMetric: "calls_vs_att",
  empBarMetric: "period_comparison",
  empBarPage: 0,
  empBarPageSize: 18,
  groupDim: "team", // "team" | "manager" | "town"
  reportDate: "",
  team: "",
  manager: "",
  baseTown: "",
  callType: "",
  callNature: "",
  employee: "",
  focusedEmployeeKey: "",
  data: null,
  // Employee matrix state
  matrixSearch: "",
  matrixSortBy: "monthlyCalls",
  matrixSortDir: "desc",
  matrixPage: 1,
  matrixPageSize: 25,
  // Interactive call_attendence table state
  ca: {
    page: 1,
    pageSize: 25,
    totalPages: 1,
    total: 0,
    search: "",
    team: "",
    callType: "",
    dateFrom: "",
    dateTo: "",
    sortBy: "call_date",
    sortDir: "desc",
    rows: [],
    summary: {},
  },
  // Interactive live_call_details table state
  lcd: {
    page: 1,
    pageSize: 25,
    totalPages: 1,
    total: 0,
    search: "",
    team: "",
    callNature: "",
    callType: "",
    dateFrom: "",
    dateTo: "",
    sortBy: "call_date",
    sortDir: "desc",
    rows: [],
    summary: {},
  },
};

const dashDom = {
  headingDate: $("#dash-heading-date"),
  periodBtns: $$(".dash-period-btn"),
  reportDateInput: $("#dash-report-date"),
  refreshBtn: $("#dash-refresh-btn"),
  filterToggleBtn: $("#dash-filter-toggle-btn"),
  filterCard: $("#dash-filter-card"),
  toggleUploadsBtn: $("#dash-toggle-uploads-btn"),
  uploadsDrawer: $("#dash-uploads-drawer"),
  employeeFilter: $("#dash-employee-filter"),
  teamFilter: $("#dash-team-filter"),
  managerFilter: $("#dash-manager-filter"),
  townFilter: $("#dash-town-filter"),
  callTypeFilter: $("#dash-calltype-filter"),
  callNatureFilter: $("#dash-callnature-filter"),
  resetBtn: $("#dash-reset-btn"),
  statusBar: $("#dash-status-bar"),
  statusText: $("#dash-status-text"),
  activeEmpPill: $("#dash-active-employee-pill"),
  // KPI cards
  kpiCallsTitle: $("#kpi-calls-title"),
  kpiCallsValue: $("#kpi-calls-value"),
  kpiCallsSub: $("#kpi-calls-sub"),
  kpiEmpTitle: $("#kpi-emp-title"),
  kpiEmpValue: $("#kpi-emp-value"),
  kpiEmpSub: $("#kpi-emp-sub"),
  kpiAvgTitle: $("#kpi-avg-title"),
  kpiAvgValue: $("#kpi-avg-value"),
  kpiAvgSub: $("#kpi-avg-sub"),
  kpiAttTitle: $("#kpi-att-title"),
  kpiAttValue: $("#kpi-att-value"),
  kpiAttSub: $("#kpi-att-sub"),
  kpiLcdTitle: $("#kpi-lcd-title"),
  kpiLcdValue: $("#kpi-lcd-value"),
  kpiLcdSub: $("#kpi-lcd-sub"),
  kpiProdTitle: $("#kpi-prod-title"),
  kpiProdValue: $("#kpi-prod-value"),
  kpiProdSub: $("#kpi-prod-sub"),
  kpiExcTitle: $("#kpi-exc-title"),
  kpiExcValue: $("#kpi-exc-value"),
  kpiExcSub: $("#kpi-exc-sub"),
  kpiGeoValue: $("#kpi-geo-value"),
  kpiGeoSub: $("#kpi-geo-sub"),
  // Charts
  trendTabs: $$("#dash-trend-granularity .dash-mini-tab"),
  lineMetricSelect: $("#dash-line-metric"),
  lineTitle: $("#dash-line-chart-title"),
  lineSub: $("#dash-line-chart-sub"),
  lineLegend: $("#dash-line-legend"),
  lineContainer: $("#dash-line-chart-container"),
  pieTitle: $("#dash-pie-chart-title"),
  pieSub: $("#dash-pie-chart-sub"),
  pieCallType: $("#dash-pie-calltype"),
  pieCallTypeLegend: $("#dash-pie-calltype-legend"),
  pieCallNature: $("#dash-pie-callnature"),
  pieCallNatureLegend: $("#dash-pie-callnature-legend"),
  pieTeamShare: $("#dash-pie-teamshare"),
  pieTeamShareLegend: $("#dash-pie-teamshare-legend"),
  empBarTitle: $("#dash-emp-bar-title"),
  empBarMetricSelect: $("#dash-emp-bar-metric"),
  empBarPrev: $("#dash-emp-bar-prev"),
  empBarNext: $("#dash-emp-bar-next"),
  empBarPageLabel: $("#dash-emp-bar-page-label"),
  empBarLegend: $("#dash-emp-bar-legend"),
  empBarContainer: $("#dash-emp-bar-container"),
  groupDimTabs: $$("#dash-group-dimension .dash-mini-tab"),
  groupBarTitle: $("#dash-group-bar-title"),
  groupBarLegend: $("#dash-group-bar-legend"),
  groupBarContainer: $("#dash-group-bar-container"),
  // Employee Matrix
  empMatrixSheet: $("#dash-emp-matrix-sheet"),
  empMatrixSearch: $("#dash-emp-matrix-search"),
  empMatrixExport: $("#dash-emp-matrix-export"),
  empMatrixBadge: $("#dash-emp-matrix-badge"),
  empMatrixFs: $("#dash-emp-matrix-fs"),
  empMatrixHead: $("#dash-emp-matrix-head"),
  empMatrixBody: $("#dash-emp-matrix-tbody"),
  empMatrixFoot: $("#dash-emp-matrix-tfoot"),
  empMatrixPageInfo: $("#dash-emp-matrix-page-info"),
  empMatrixPageSize: $("#dash-emp-matrix-pagesize"),
  empMatrixPrev: $("#dash-emp-matrix-prev"),
  empMatrixNext: $("#dash-emp-matrix-next"),
  // DB Table Switcher
  dbTableTabs: $$("#dash-db-table-switcher .dash-mini-tab"),
  dbTablesGrid: $("#dash-db-tables-grid"),
  caCard: $("#dash-ca-card"),
  lcdCard: $("#dash-lcd-card"),
  // CA Table
  caSheet: $("#dash-ca-sheet"),
  caSumRows: $("#ca-sum-rows"),
  caSumCalls: $("#ca-sum-calls"),
  caSumExc: $("#ca-sum-exc"),
  caSumNoCalls: $("#ca-sum-nocalls"),
  caSumEmps: $("#ca-sum-emps"),
  caSearch: $("#ca-table-search"),
  caTeam: $("#ca-table-team"),
  caCallType: $("#ca-table-calltype"),
  caFrom: $("#ca-table-from"),
  caTo: $("#ca-table-to"),
  caReset: $("#ca-table-reset"),
  caExport: $("#ca-table-export"),
  caBadge: $("#ca-table-badge"),
  caFs: $("#ca-table-fs"),
  caHead: $("#ca-table-head"),
  caBody: $("#ca-table-tbody"),
  caPageInfo: $("#ca-table-page-info"),
  caPageSize: $("#ca-table-pagesize"),
  caFirst: $("#ca-table-first"),
  caPrev: $("#ca-table-prev"),
  caPageNums: $("#ca-table-pagenums"),
  caNext: $("#ca-table-next"),
  caLast: $("#ca-table-last"),
  // LCD Table
  lcdSheet: $("#dash-lcd-sheet"),
  lcdSumRows: $("#lcd-sum-rows"),
  lcdSumProvider: $("#lcd-sum-provider"),
  lcdSumAtt: $("#lcd-sum-att"),
  lcdSumExc: $("#lcd-sum-exc"),
  lcdSumProducts: $("#lcd-sum-products"),
  lcdSumDoctors: $("#lcd-sum-doctors"),
  lcdSumInRange: $("#lcd-sum-inrange"),
  lcdSearch: $("#lcd-table-search"),
  lcdTeam: $("#lcd-table-team"),
  lcdNature: $("#lcd-table-nature"),
  lcdCallType: $("#lcd-table-calltype"),
  lcdFrom: $("#lcd-table-from"),
  lcdTo: $("#lcd-table-to"),
  lcdReset: $("#lcd-table-reset"),
  lcdExport: $("#lcd-table-export"),
  lcdBadge: $("#lcd-table-badge"),
  lcdFs: $("#lcd-table-fs"),
  lcdHead: $("#lcd-table-head"),
  lcdBody: $("#lcd-table-tbody"),
  lcdPageInfo: $("#lcd-table-page-info"),
  lcdPageSize: $("#lcd-table-pagesize"),
  lcdFirst: $("#lcd-table-first"),
  lcdPrev: $("#lcd-table-prev"),
  lcdPageNums: $("#lcd-table-pagenums"),
  lcdNext: $("#lcd-table-next"),
  lcdLast: $("#lcd-table-last"),
};

const dashMultiFilters = {
  team: dashDom.teamFilter ? createExcelMultiSelect(dashDom.teamFilter, "All Teams", (val) => {
    dashState.team = val;
    dashState.empBarPage = 0;
    dashState.matrixPage = 1;
    loadDashboardAnalytics();
  }) : null,
  manager: dashDom.managerFilter ? createExcelMultiSelect(dashDom.managerFilter, "All Managers", (val) => {
    dashState.manager = val;
    dashState.empBarPage = 0;
    dashState.matrixPage = 1;
    loadDashboardAnalytics();
  }) : null,
  baseTown: dashDom.townFilter ? createExcelMultiSelect(dashDom.townFilter, "All Base Towns", (val) => {
    dashState.baseTown = val;
    dashState.empBarPage = 0;
    dashState.matrixPage = 1;
    loadDashboardAnalytics();
  }) : null,
  callType: dashDom.callTypeFilter ? createExcelMultiSelect(dashDom.callTypeFilter, "All Call Types", (val) => {
    dashState.callType = val;
    loadDashboardAnalytics();
  }) : null,
  callNature: dashDom.callNatureFilter ? createExcelMultiSelect(dashDom.callNatureFilter, "All Call Natures", (val) => {
    dashState.callNature = val;
    loadDashboardAnalytics();
  }) : null,
};

dashDom.employeeFilter?.addEventListener("change", () => {
  dashState.employee = dashDom.employeeFilter.value;
  dashState.empBarPage = 0;
  dashState.matrixPage = 1;
  loadDashboardAnalytics();
});

const CHART_PALETTE = [
  "#059669",
  "#2563eb",
  "#d97706",
  "#7c3aed",
  "#e11d48",
  "#0891b2",
  "#4f46e5",
  "#0d9488",
  "#ea580c",
  "#65a30d",
];

function toggleSheetFullscreen(sheetEl, btnEl) {
  if (!sheetEl || !btnEl) return;
  const isActive = sheetEl.classList.contains("table-fullscreen-active");
  $$(".fmcg-report-sheet.table-fullscreen-active").forEach((s) => s.classList.remove("table-fullscreen-active"));
  $$(".table-fullscreen-btn").forEach((b) => {
    b.classList.remove("is-fullscreen");
    b.querySelector(".fs-icon-open")?.classList.remove("hidden");
    b.querySelector(".fs-icon-exit")?.classList.add("hidden");
  });
  if (!isActive) {
    sheetEl.classList.add("table-fullscreen-active");
    document.body.classList.add("has-table-fullscreen");
    btnEl.classList.add("is-fullscreen");
    btnEl.querySelector(".fs-icon-open")?.classList.add("hidden");
    btnEl.querySelector(".fs-icon-exit")?.classList.remove("hidden");
  } else {
    document.body.classList.remove("has-table-fullscreen");
  }
}

function getFocusedEmployee() {
  if (!dashState.data || !dashState.focusedEmployeeKey) return null;
  return (dashState.data.employees || []).find((e) => e.employeeKey === dashState.focusedEmployeeKey) || null;
}

function renderDashboardKpis() {
  if (!dashState.data) return;
  const { kpis } = dashState.data;
  const period = dashState.period;
  const pData = kpis[period] || kpis.monthly;
  const dData = kpis.daily;
  const wData = kpis.weekly;
  const mData = kpis.monthly;
  const periodTitle = period === "daily" ? "Daily" : period === "weekly" ? "Weekly (7-Day)" : "Monthly (MTD)";

  const focusedEmp = getFocusedEmployee();
  if (focusedEmp) {
    const empCalls = period === "daily" ? focusedEmp.dailyCalls : period === "weekly" ? focusedEmp.weeklyCalls : focusedEmp.monthlyCalls;
    const empAtt = period === "daily" ? focusedEmp.dailyAtt : period === "weekly" ? focusedEmp.weeklyAttDays : focusedEmp.monthlyAttDays;
    const empAvg = period === "daily" ? focusedEmp.dailyCalls : period === "weekly" ? focusedEmp.weeklyAvgCalls : focusedEmp.monthlyAvgCalls;
    const empLcd = period === "daily" ? focusedEmp.dailyLcdCalls : period === "weekly" ? focusedEmp.weeklyLcdCalls : focusedEmp.monthlyLcdCalls;
    const empProd = period === "daily" ? focusedEmp.dailyProducts : period === "weekly" ? focusedEmp.weeklyProducts : focusedEmp.monthlyProducts;
    const empExc = (period === "daily" ? focusedEmp.dailyNoCalls : period === "weekly" ? focusedEmp.weeklyNoCallDays : focusedEmp.monthlyNoCallDays)
      + (period === "daily" ? focusedEmp.dailyLeave : period === "weekly" ? focusedEmp.weeklyLeaveDays : focusedEmp.monthlyLeaveDays);

    dashDom.kpiCallsTitle.textContent = `${periodTitle} Calls (${focusedEmp.ff_id})`;
    dashDom.kpiCallsValue.textContent = empCalls.toLocaleString();
    dashDom.kpiCallsSub.innerHTML = `<span>Daily: <strong>${focusedEmp.dailyCalls.toLocaleString()}</strong></span><span>Weekly: <strong>${focusedEmp.weeklyCalls.toLocaleString()}</strong></span><span>Monthly: <strong>${focusedEmp.monthlyCalls.toLocaleString()}</strong></span>`;

    dashDom.kpiEmpTitle.textContent = "Focused Employee";
    dashDom.kpiEmpValue.textContent = focusedEmp.ff_name;
    dashDom.kpiEmpSub.innerHTML = `<span>Team: <strong>${escapeHtml(focusedEmp.team)}</strong></span><span>Town: <strong>${escapeHtml(focusedEmp.base_town)}</strong></span>`;

    dashDom.kpiAvgTitle.textContent = `${periodTitle} Avg Calls / Day`;
    dashDom.kpiAvgValue.textContent = Number(empAvg).toFixed(1);
    dashDom.kpiAvgSub.innerHTML = `<span>Daily: <strong>${focusedEmp.dailyCalls}</strong></span><span>Wk Avg: <strong>${focusedEmp.weeklyAvgCalls}</strong></span><span>Mth Avg: <strong>${focusedEmp.monthlyAvgCalls}</strong></span>`;

    dashDom.kpiAttTitle.textContent = `${periodTitle} Attendance`;
    dashDom.kpiAttValue.textContent = `${empAtt.toLocaleString()} day(s)`;
    dashDom.kpiAttSub.innerHTML = `<span>Daily: <strong>${focusedEmp.dailyAtt}</strong></span><span>Weekly: <strong>${focusedEmp.weeklyAttDays}d</strong></span><span>Monthly: <strong>${focusedEmp.monthlyAttDays}d</strong></span>`;

    dashDom.kpiLcdTitle.textContent = `${periodTitle} Provider Calls`;
    dashDom.kpiLcdValue.textContent = empLcd.toLocaleString();
    dashDom.kpiLcdSub.innerHTML = `<span>Daily: <strong>${focusedEmp.dailyLcdCalls}</strong></span><span>Weekly: <strong>${focusedEmp.weeklyLcdCalls}</strong></span><span>Monthly: <strong>${focusedEmp.monthlyLcdCalls}</strong></span>`;

    dashDom.kpiProdTitle.textContent = `${periodTitle} Products`;
    dashDom.kpiProdValue.textContent = empProd.toLocaleString();
    dashDom.kpiProdSub.innerHTML = `<span>Daily: <strong>${focusedEmp.dailyProducts}</strong></span><span>Weekly: <strong>${focusedEmp.weeklyProducts}</strong></span><span>Monthly: <strong>${focusedEmp.monthlyProducts}</strong></span>`;

    dashDom.kpiExcTitle.textContent = "No Calls & Leave Days";
    dashDom.kpiExcValue.textContent = empExc.toLocaleString();
    dashDom.kpiExcSub.innerHTML = `<span>No Call: <strong>${focusedEmp.monthlyNoCallDays}d</strong></span><span>Leave: <strong>${focusedEmp.monthlyLeaveDays}d</strong></span><span>Doctors: <strong>${focusedEmp.uniqueDoctors}</strong></span>`;

    dashDom.kpiGeoValue.textContent = `${focusedEmp.inRangePct}%`;
    dashDom.kpiGeoSub.innerHTML = `<span>Manager: <strong>${escapeHtml(focusedEmp.manager)}</strong></span>`;
    return;
  }

  dashDom.kpiCallsTitle.textContent = `${periodTitle} Calls Done`;
  dashDom.kpiCallsValue.textContent = pData.callsDone.toLocaleString();
  dashDom.kpiCallsSub.innerHTML = `<span>Daily: <strong>${dData.callsDone.toLocaleString()}</strong></span><span>Weekly: <strong>${wData.callsDone.toLocaleString()}</strong></span><span>Monthly: <strong>${mData.callsDone.toLocaleString()}</strong></span>`;

  dashDom.kpiEmpTitle.textContent = `${periodTitle} Active Employees`;
  dashDom.kpiEmpValue.textContent = `${pData.activeEmployees.toLocaleString()} / ${kpis.totalEmployees.toLocaleString()}`;
  dashDom.kpiEmpSub.innerHTML = `<span>Daily: <strong>${dData.activeEmployees}</strong></span><span>Weekly: <strong>${wData.activeEmployees}</strong></span><span>Monthly: <strong>${mData.activeEmployees}</strong></span>`;

  dashDom.kpiAvgTitle.textContent = `${periodTitle} Avg Calls / Emp`;
  dashDom.kpiAvgValue.textContent = Number(pData.avgCallsPerEmp).toFixed(1);
  dashDom.kpiAvgSub.innerHTML = `<span>Daily Avg: <strong>${dData.avgCallsPerEmp}</strong></span><span>Wk Avg: <strong>${wData.avgCallsPerEmp}</strong></span><span>Mth Avg: <strong>${mData.avgCallsPerEmp}</strong></span>`;

  dashDom.kpiAttTitle.textContent = `${periodTitle} Attendance`;
  dashDom.kpiAttValue.textContent = pData.attendanceCount.toLocaleString();
  dashDom.kpiAttSub.innerHTML = `<span>Daily: <strong>${dData.attendanceCount.toLocaleString()}</strong></span><span>Weekly: <strong>${wData.attendanceCount.toLocaleString()}</strong></span><span>Monthly: <strong>${mData.attendanceCount.toLocaleString()}</strong></span>`;

  dashDom.kpiLcdTitle.textContent = `${periodTitle} Provider Calls`;
  dashDom.kpiLcdValue.textContent = pData.lcdProviderCalls.toLocaleString();
  dashDom.kpiLcdSub.innerHTML = `<span>Daily: <strong>${dData.lcdProviderCalls.toLocaleString()}</strong></span><span>Weekly: <strong>${wData.lcdProviderCalls.toLocaleString()}</strong></span><span>Monthly: <strong>${mData.lcdProviderCalls.toLocaleString()}</strong></span>`;

  dashDom.kpiProdTitle.textContent = `${periodTitle} Products Detailed`;
  dashDom.kpiProdValue.textContent = pData.lcdProducts.toLocaleString();
  dashDom.kpiProdSub.innerHTML = `<span>Daily: <strong>${dData.lcdProducts.toLocaleString()}</strong></span><span>Weekly: <strong>${wData.lcdProducts.toLocaleString()}</strong></span><span>Monthly: <strong>${mData.lcdProducts.toLocaleString()}</strong></span>`;

  const totalIssues = pData.noCallsCount + pData.leaveCount;
  dashDom.kpiExcTitle.textContent = `${periodTitle} No Calls & Leave`;
  dashDom.kpiExcValue.textContent = totalIssues.toLocaleString();
  dashDom.kpiExcSub.innerHTML = `<span>No Calls: <strong>${pData.noCallsCount.toLocaleString()}</strong></span><span>Leave: <strong>${pData.leaveCount.toLocaleString()}</strong></span><span>Other Exc: <strong>${pData.otherExcCount.toLocaleString()}</strong></span>`;

  dashDom.kpiGeoValue.textContent = `${mData.inRangePct}%`;
  dashDom.kpiGeoSub.innerHTML = `<span>Location: <strong>Within Geofence Range</strong></span>`;
}

// ---------------------------------------------------------------------------
// SVG Line Chart Renderer (Daily, Weekly, Monthly & Focused Employee Trend)
// ---------------------------------------------------------------------------
function renderDashboardLineChart() {
  if (!dashState.data) return;
  const { charts } = dashState.data;
  const focusedEmp = getFocusedEmployee();

  let points = [];
  if (focusedEmp) {
    // Build employee-specific trend points from dailyTrend dates
    if (dashState.trendMode === "weekly") {
      points = ["W1", "W2", "W3", "W4", "W5"].map((wKey) => ({
        label: wKey,
        callsDone: focusedEmp.weekBuckets?.[wKey] || 0,
        attendanceCount: 0,
        lcdProviderCalls: 0,
        lcdProducts: 0,
        avgCallsPerEmp: focusedEmp.weekBuckets?.[wKey] || 0,
        cumulativeCalls: focusedEmp.weekBuckets?.[wKey] || 0,
      })).filter((p, idx) => idx < 3 || p.callsDone > 0);
    } else {
      let running = 0;
      points = (charts.dailyTrend || []).map((d) => {
        const c = focusedEmp.dailySeries?.[d.date] || 0;
        const att = focusedEmp.dailyAttSeries?.[d.date] || 0;
        const lcd = focusedEmp.dailyLcdSeries?.[d.date] || 0;
        running += c;
        return {
          label: d.label,
          callsDone: c,
          attendanceCount: att,
          lcdProviderCalls: lcd,
          lcdProducts: lcd,
          avgCallsPerEmp: c,
          cumulativeCalls: running,
        };
      });
    }
    dashDom.lineTitle.textContent = `${focusedEmp.ff_name} (${focusedEmp.ff_id}) — ${dashState.trendMode.toUpperCase()} Call Trajectory`;
  } else {
    points = dashState.trendMode === "weekly"
      ? charts.weeklyTrend || []
      : dashState.trendMode === "monthly"
        ? charts.monthlyTrend || []
        : charts.dailyTrend || [];
    dashDom.lineTitle.textContent = `${dashState.trendMode === "daily" ? "Daily" : dashState.trendMode === "weekly" ? "Weekly" : "Monthly"} Call & Attendance Performance Trend`;
  }

  if (!points.length) {
    dashDom.lineContainer.innerHTML = '<div class="fmcg-state-cell">No time-series data available for selected filters.</div>';
    return;
  }

  const metricMode = dashDom.lineMetricSelect.value;
  let seriesDefs = [];
  if (metricMode === "lcd_vs_products") {
    seriesDefs = [
      { key: "lcdProviderCalls", label: "Provider Calls (live_call_details)", color: "#4f46e5" },
      { key: "lcdProducts", label: "Products Detailed (live_call_details)", color: "#7c3aed" },
    ];
  } else if (metricMode === "avg_calls") {
    seriesDefs = [
      { key: "avgCallsPerEmp", label: "Avg Calls / Active Employee", color: "#2563eb" },
      { key: "activeCallEmployees", label: "Active Calling Employees", color: "#0d9488" },
    ];
  } else if (metricMode === "cumulative") {
    seriesDefs = [
      { key: "cumulativeCalls", label: "Cumulative Calls Done", color: "#059669" },
      { key: "callsDone", label: "Period Calls Done", color: "#d97706" },
    ];
  } else {
    seriesDefs = [
      { key: "callsDone", label: "Calls Done (call_attendence)", color: "#059669" },
      { key: "attendanceCount", label: "Attendance Marked", color: "#d97706" },
      { key: "lcdProviderCalls", label: "Provider Calls (live_call_details)", color: "#2563eb" },
    ];
  }

  dashDom.lineLegend.innerHTML = seriesDefs
    .map((s) => `<span class="dash-legend-item"><i class="dash-legend-swatch" style="background:${s.color}"></i>${escapeHtml(s.label)}</span>`)
    .join("");

  const width = Math.max((dashDom.lineContainer?.clientWidth || 760) - 16, 680);
  const height = 245;
  const padLeft = 48;
  const padRight = 20;
  const padTop = 20;
  const padBottom = 38;
  const plotW = width - padLeft - padRight;
  const plotH = height - padTop - padBottom;

  const allVals = [];
  for (const p of points) {
    for (const s of seriesDefs) {
      allVals.push(Number(p[s.key] || 0));
    }
  }
  const maxVal = Math.max(...allVals, 10);

  const xAt = (idx) => (points.length === 1 ? padLeft + plotW / 2 : padLeft + (idx / (points.length - 1)) * plotW);
  const yAt = (val) => padTop + plotH - (Math.max(val, 0) / maxVal) * plotH;

  // Grid lines
  const gridTicks = 4;
  const gridLines = [];
  for (let i = 0; i <= gridTicks; i += 1) {
    const val = Math.round((maxVal / gridTicks) * i);
    const y = yAt(val).toFixed(1);
    gridLines.push(
      `<line x1="${padLeft}" y1="${y}" x2="${width - padRight}" y2="${y}" stroke="#e2efe9" stroke-dasharray="3,3" />`,
      `<text x="${padLeft - 6}" y="${Number(y) + 3}" text-anchor="end" font-size="9.5" font-weight="600" fill="#547267">${val.toLocaleString()}</text>`
    );
  }

  // X-axis labels
  const xLabels = points.map((p, idx) => {
    const showLabel = points.length <= 16 || idx % 2 === 0 || idx === points.length - 1;
    if (!showLabel) return "";
    const x = xAt(idx).toFixed(1);
    return `<text x="${x}" y="${height - 12}" text-anchor="middle" font-size="9.5" font-weight="700" fill="#064e3b">${escapeHtml(p.label)}</text>`;
  });

  // Lines & Dots
  const pathsHtml = seriesDefs.map((s, sIdx) => {
    const coords = points.map((p, idx) => [xAt(idx), yAt(Number(p[s.key] || 0))]);
    const pathD = coords.map(([x, y], i) => `${i === 0 ? "M" : "L"}${x.toFixed(1)},${y.toFixed(1)}`).join(" ");
    const areaD = sIdx === 0 && coords.length > 1
      ? `${pathD} L${coords[coords.length - 1][0].toFixed(1)},${(padTop + plotH).toFixed(1)} L${coords[0][0].toFixed(1)},${(padTop + plotH).toFixed(1)} Z`
      : "";
    const areaHtml = areaD ? `<path d="${areaD}" fill="${s.color}" fill-opacity="0.10" />` : "";
    const dotsHtml = coords.map(([x, y], idx) => {
      const val = Number(points[idx][s.key] || 0);
      return `<g>
        <circle cx="${x.toFixed(1)}" cy="${y.toFixed(1)}" r="3.8" fill="#ffffff" stroke="${s.color}" stroke-width="2.2">
          <title>${escapeHtml(points[idx].label)} · ${escapeHtml(s.label)}: ${val.toLocaleString()}</title>
        </circle>
        ${points.length <= 18 && sIdx === 0 && val > 0 ? `<text x="${x.toFixed(1)}" y="${(y - 7).toFixed(1)}" text-anchor="middle" font-size="8.5" font-weight="800" fill="${s.color}">${val.toLocaleString()}</text>` : ""}
      </g>`;
    }).join("");
    return `${areaHtml}<path d="${pathD}" fill="none" stroke="${s.color}" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round" />${dotsHtml}`;
  }).join("");

  dashDom.lineContainer.innerHTML = `
    <svg viewBox="0 0 ${width} ${height}" role="img" aria-label="Call performance line chart">
      ${gridLines.join("")}
      ${pathsHtml}
      ${xLabels.join("")}
    </svg>
  `;
}

// ---------------------------------------------------------------------------
// SVG Pie / Donut ("Par") Charts Renderer
// ---------------------------------------------------------------------------
function buildDonutSvg(slices, centerLabel, centerValue) {
  const total = slices.reduce((acc, s) => acc + Math.max(Number(s.value || 0), 0), 0);
  if (total <= 0) {
    return `<svg viewBox="0 0 140 140"><circle cx="70" cy="70" r="48" fill="none" stroke="#e2efe9" stroke-width="18" /><text x="70" y="73" text-anchor="middle" font-size="10" font-weight="700" fill="#547267">No Data</text></svg>`;
  }
  const cx = 70;
  const cy = 70;
  const r = 48;
  const circumference = 2 * Math.PI * r;
  let offset = 0;

  const circles = slices.map((s) => {
    const val = Math.max(Number(s.value || 0), 0);
    if (val === 0) return "";
    const pct = val / total;
    const dash = pct * circumference;
    const strokeOffset = -offset;
    offset += dash;
    return `<circle cx="${cx}" cy="${cy}" r="${r}" fill="none" stroke="${s.color}" stroke-width="19" stroke-dasharray="${dash.toFixed(2)} ${(circumference - dash).toFixed(2)}" stroke-dashoffset="${strokeOffset.toFixed(2)}" transform="rotate(-90 ${cx} ${cy})">
      <title>${escapeHtml(s.label)}: ${val.toLocaleString()} (${(pct * 100).toFixed(1)}%)</title>
    </circle>`;
  }).join("");

  return `
    <svg viewBox="0 0 140 140" role="img" aria-label="${escapeHtml(centerLabel)} pie chart">
      <circle cx="${cx}" cy="${cy}" r="${r}" fill="none" stroke="#f0f9f5" stroke-width="19" />
      ${circles}
      <text x="${cx}" y="${cy - 3}" text-anchor="middle" font-size="13" font-weight="800" fill="#022c22">${escapeHtml(centerValue)}</text>
      <text x="${cx}" y="${cy + 11}" text-anchor="middle" font-size="8" font-weight="700" fill="#547267">${escapeHtml(centerLabel)}</text>
    </svg>
  `;
}

function renderPieLegend(legendEl, slices) {
  const total = slices.reduce((acc, s) => acc + Math.max(Number(s.value || 0), 0), 0);
  legendEl.innerHTML = slices.map((s) => {
    const val = Math.max(Number(s.value || 0), 0);
    const pct = total > 0 ? ((val / total) * 100).toFixed(1) : "0.0";
    return `<div class="dash-pie-legend-row">
      <span><i class="dash-legend-swatch" style="background:${s.color}"></i>${escapeHtml(s.label)}</span>
      <strong>${val.toLocaleString()} (${pct}%)</strong>
    </div>`;
  }).join("");
}

function renderDashboardPieCharts() {
  if (!dashState.data) return;
  const { charts } = dashState.data;
  const period = dashState.period;
  const periodLabel = period === "daily" ? "Daily" : period === "weekly" ? "Weekly" : "Monthly";
  dashDom.pieTitle.textContent = `${periodLabel} Call Type, Call Nature & Team Share Breakdown`;

  // 1. call_attendence CallType pie
  const ctSource = charts.callTypePie?.[period] || charts.callTypePie?.monthly || {};
  const ctSlices = [
    { label: "CALL", value: ctSource.CALL?.done || 0, color: "#059669" },
    { label: "EXCEPTION", value: ctSource.EXCEPTION?.done || 0, color: "#d97706" },
    { label: "NO CALLS", value: ctSource["NO CALLS"]?.rows || 0, color: "#e11d48" },
  ];
  const ctTotal = ctSlices.reduce((a, b) => a + b.value, 0);
  dashDom.pieCallType.innerHTML = buildDonutSvg(ctSlices, `${periodLabel} Activity`, ctTotal.toLocaleString());
  renderPieLegend(dashDom.pieCallTypeLegend, ctSlices);

  // 2. live_call_details Call Nature pie
  const cnSource = charts.callNaturePie?.[period] || charts.callNaturePie?.monthly || {};
  const cnSlices = [
    { label: "Provider Calls", value: cnSource["Provider Calls"] || 0, color: "#2563eb" },
    { label: "Attendance", value: cnSource.Attendance || 0, color: "#059669" },
    { label: "Others Exc Calls", value: cnSource["Others Exc Calls"] || 0, color: "#d97706" },
    { label: "LEAVE", value: cnSource.LEAVE || 0, color: "#7c3aed" },
  ];
  const cnTotal = cnSlices.reduce((a, b) => a + b.value, 0);
  dashDom.pieCallNature.innerHTML = buildDonutSvg(cnSlices, `${periodLabel} LCD`, cnTotal.toLocaleString());
  renderPieLegend(dashDom.pieCallNatureLegend, cnSlices);

  // 3. Team Share of Calls pie
  const teamCallKey = period === "daily" ? "dailyCalls" : period === "weekly" ? "weeklyCalls" : "monthlyCalls";
  const teamSlices = (charts.teamPerformance || [])
    .map((t, idx) => ({
      label: t.name,
      value: Number(t[teamCallKey] || 0),
      color: CHART_PALETTE[idx % CHART_PALETTE.length],
    }))
    .filter((s) => s.value > 0)
    .slice(0, 6);
  const teamTotal = teamSlices.reduce((a, b) => a + b.value, 0);
  dashDom.pieTeamShare.innerHTML = buildDonutSvg(teamSlices, `${periodLabel} Calls`, teamTotal.toLocaleString());
  renderPieLegend(dashDom.pieTeamShareLegend, teamSlices);
}

// ---------------------------------------------------------------------------
// SVG Bar Chart 1: Every Employee Daily, Weekly, Monthly Call Performance
// ---------------------------------------------------------------------------
function renderEmployeeBarChart() {
  if (!dashState.data) return;
  const employees = dashState.data.employees || [];
  const metric = dashDom.empBarMetricSelect.value;

  const sorted = [...employees];
  if (metric !== "period_comparison") {
    sorted.sort((a, b) => Number(b[metric] || 0) - Number(a[metric] || 0));
  } else {
    const sortField = dashState.period === "daily" ? "dailyCalls" : dashState.period === "weekly" ? "weeklyCalls" : "monthlyCalls";
    sorted.sort((a, b) => Number(b[sortField] || 0) - Number(a[sortField] || 0));
  }

  const totalEmps = sorted.length;
  const pageSize = dashState.empBarPageSize;
  const maxPage = Math.max(Math.ceil(totalEmps / pageSize) - 1, 0);
  if (dashState.empBarPage > maxPage) dashState.empBarPage = maxPage;
  if (dashState.empBarPage < 0) dashState.empBarPage = 0;

  const startIdx = dashState.empBarPage * pageSize;
  const visible = sorted.slice(startIdx, startIdx + pageSize);

  dashDom.empBarPageLabel.textContent = totalEmps
    ? `${startIdx + 1}–${Math.min(startIdx + visible.length, totalEmps)} of ${totalEmps}`
    : "0 of 0";
  dashDom.empBarPrev.disabled = dashState.empBarPage <= 0;
  dashDom.empBarNext.disabled = dashState.empBarPage >= maxPage;

  if (!visible.length) {
    dashDom.empBarContainer.innerHTML = '<div class="fmcg-state-cell">No employee records match the selected filters.</div>';
    return;
  }

  let series = [];
  if (metric === "period_comparison") {
    series = [
      { key: "dailyCalls", label: "Daily Calls", color: "#2563eb" },
      { key: "weeklyCalls", label: "Weekly Calls", color: "#d97706" },
      { key: "monthlyCalls", label: "Monthly Calls", color: "#059669" },
    ];
  } else {
    const labelMap = {
      monthlyCalls: ["Monthly Calls Done", "#059669"],
      weeklyCalls: ["Weekly Calls Done", "#d97706"],
      dailyCalls: ["Daily Calls Done", "#2563eb"],
      monthlyLcdCalls: ["Provider Calls (live_call_details)", "#4f46e5"],
      monthlyProducts: ["Products Detailed", "#7c3aed"],
      monthlyAttDays: ["Monthly Attendance Days", "#0d9488"],
      monthlyAvgCalls: ["Monthly Avg Calls / Day", "#0891b2"],
    };
    const [lbl, clr] = labelMap[metric] || [metric, "#059669"];
    series = [{ key: metric, label: lbl, color: clr }];
  }

  dashDom.empBarLegend.innerHTML = series
    .map((s) => `<span class="dash-legend-item"><i class="dash-legend-swatch" style="background:${s.color}"></i>${escapeHtml(s.label)}</span>`)
    .join("");

  const width = Math.max((dashDom.empBarContainer?.clientWidth || 760) - 16, 680);
  const height = 255;
  const padLeft = 46;
  const padRight = 14;
  const padTop = 18;
  const padBottom = 64;
  const plotW = width - padLeft - padRight;
  const plotH = height - padTop - padBottom;

  const maxVal = Math.max(
    ...visible.flatMap((emp) => series.map((s) => Number(emp[s.key] || 0))),
    10
  );

  const yAt = (val) => padTop + plotH - (Math.max(val, 0) / maxVal) * plotH;
  const groupW = plotW / visible.length;
  const barW = Math.min(Math.max((groupW * 0.76) / series.length, 4), 24);

  const gridLines = [];
  for (let i = 0; i <= 4; i += 1) {
    const v = Math.round((maxVal / 4) * i);
    const y = yAt(v).toFixed(1);
    gridLines.push(
      `<line x1="${padLeft}" y1="${y}" x2="${width - padRight}" y2="${y}" stroke="#e2efe9" stroke-dasharray="3,3" />`,
      `<text x="${padLeft - 6}" y="${Number(y) + 3}" text-anchor="end" font-size="9" font-weight="600" fill="#547267">${v.toLocaleString()}</text>`
    );
  }

  const barsHtml = visible.map((emp, idx) => {
    const groupX = padLeft + idx * groupW;
    const totalBarsW = barW * series.length + (series.length - 1) * 1.5;
    const startX = groupX + (groupW - totalBarsW) / 2;
    const isFocused = dashState.focusedEmployeeKey === emp.employeeKey;
    const shortName = String(emp.ff_name || emp.ff_id).replace(/\s*\([^)]*\)/g, "").slice(0, 13);

    const rects = series.map((s, sIdx) => {
      const val = Number(emp[s.key] || 0);
      const bx = startX + sIdx * (barW + 1.5);
      const by = yAt(val);
      const bh = Math.max(padTop + plotH - by, val > 0 ? 2 : 0);
      return `<rect x="${bx.toFixed(1)}" y="${by.toFixed(1)}" width="${barW.toFixed(1)}" height="${bh.toFixed(1)}" rx="2" fill="${s.color}" opacity="${isFocused ? "1" : "0.88"}">
        <title>${escapeHtml(emp.ff_name)} (${escapeHtml(emp.ff_id)}) · ${escapeHtml(s.label)}: ${val.toLocaleString()} (Daily: ${emp.dailyCalls}, Weekly: ${emp.weeklyCalls}, Monthly: ${emp.monthlyCalls})</title>
      </rect>`;
    }).join("");

    const labelX = (groupX + groupW / 2).toFixed(1);
    const labelY = height - 46;
    return `<g class="dash-emp-bar-group" data-emp-key="${escapeHtml(emp.employeeKey)}" style="cursor:pointer">
      ${isFocused ? `<rect x="${(groupX + 1).toFixed(1)}" y="${padTop}" width="${(groupW - 2).toFixed(1)}" height="${plotH}" fill="#10b981" fill-opacity="0.14" rx="3" />` : ""}
      ${rects}
      <text x="${labelX}" y="${labelY}" text-anchor="end" transform="rotate(-36 ${labelX} ${labelY})" font-size="8.6" font-weight="${isFocused ? "800" : "600"}" fill="${isFocused ? "#022c22" : "#064e3b"}">${escapeHtml(shortName)}</text>
    </g>`;
  }).join("");

  dashDom.empBarContainer.innerHTML = `
    <svg viewBox="0 0 ${width} ${height}" role="img" aria-label="Employee call performance bar chart">
      ${gridLines.join("")}
      ${barsHtml}
    </svg>
  `;
}

// ---------------------------------------------------------------------------
// SVG Bar Chart 2: Team, Manager & Base Town Call Performance Comparison
// ---------------------------------------------------------------------------
function renderGroupBarChart() {
  if (!dashState.data) return;
  const { charts } = dashState.data;
  const dim = dashState.groupDim;

  const sourceList = dim === "manager"
    ? (charts.managerPerformance || []).slice(0, 12)
    : dim === "town"
      ? (charts.townPerformance || []).slice(0, 12)
      : (charts.teamPerformance || []).slice(0, 10);

  if (!sourceList.length) {
    dashDom.groupBarContainer.innerHTML = '<div class="fmcg-state-cell">No organizational breakdown data available.</div>';
    return;
  }

  const series = [
    { key: "dailyCalls", label: "Daily Calls", color: "#2563eb" },
    { key: "weeklyCalls", label: "Weekly Calls", color: "#d97706" },
    { key: "monthlyCalls", label: "Monthly Calls", color: "#059669" },
    { key: "lcdMonthlyCalls", label: "Provider Calls (LCD)", color: "#7c3aed" },
  ];

  dashDom.groupBarLegend.innerHTML = series
    .map((s) => `<span class="dash-legend-item"><i class="dash-legend-swatch" style="background:${s.color}"></i>${escapeHtml(s.label)}</span>`)
    .join("");

  const width = Math.max((dashDom.groupBarContainer?.clientWidth || 620) - 16, 580);
  const height = 255;
  const padLeft = 46;
  const padRight = 14;
  const padTop = 18;
  const padBottom = 64;
  const plotW = width - padLeft - padRight;
  const plotH = height - padTop - padBottom;

  const maxVal = Math.max(
    ...sourceList.flatMap((g) => series.map((s) => Number(g[s.key] || 0))),
    10
  );

  const yAt = (val) => padTop + plotH - (Math.max(val, 0) / maxVal) * plotH;
  const groupW = plotW / sourceList.length;
  const barW = Math.min(Math.max((groupW * 0.75) / series.length, 4), 18);

  const gridLines = [];
  for (let i = 0; i <= 4; i += 1) {
    const v = Math.round((maxVal / 4) * i);
    const y = yAt(v).toFixed(1);
    gridLines.push(
      `<line x1="${padLeft}" y1="${y}" x2="${width - padRight}" y2="${y}" stroke="#e2efe9" stroke-dasharray="3,3" />`,
      `<text x="${padLeft - 6}" y="${Number(y) + 3}" text-anchor="end" font-size="9" font-weight="600" fill="#547267">${v.toLocaleString()}</text>`
    );
  }

  const barsHtml = sourceList.map((item, idx) => {
    const groupX = padLeft + idx * groupW;
    const totalBarsW = barW * series.length + (series.length - 1) * 1.5;
    const startX = groupX + (groupW - totalBarsW) / 2;
    const shortLabel = String(item.name || "-").slice(0, 15);

    const rects = series.map((s, sIdx) => {
      const val = Number(item[s.key] || 0);
      const bx = startX + sIdx * (barW + 1.5);
      const by = yAt(val);
      const bh = Math.max(padTop + plotH - by, val > 0 ? 2 : 0);
      return `<rect x="${bx.toFixed(1)}" y="${by.toFixed(1)}" width="${barW.toFixed(1)}" height="${bh.toFixed(1)}" rx="2" fill="${s.color}">
        <title>${escapeHtml(item.name)} · ${escapeHtml(s.label)}: ${val.toLocaleString()} (Active Employees: ${item.monthlyActiveEmployees})</title>
      </rect>`;
    }).join("");

    const labelX = (groupX + groupW / 2).toFixed(1);
    const labelY = height - 46;
    return `<g>
      ${rects}
      <text x="${labelX}" y="${labelY}" text-anchor="end" transform="rotate(-34 ${labelX} ${labelY})" font-size="8.8" font-weight="700" fill="#064e3b">${escapeHtml(shortLabel)}</text>
    </g>`;
  }).join("");

  dashDom.groupBarContainer.innerHTML = `
    <svg viewBox="0 0 ${width} ${height}" role="img" aria-label="Organizational call performance bar chart">
      ${gridLines.join("")}
      ${barsHtml}
    </svg>
  `;
}

// ---------------------------------------------------------------------------
// Every Employee Daily, Weekly & Monthly Call Performance Matrix
// ---------------------------------------------------------------------------
function getFilteredSortedEmployees() {
  if (!dashState.data) return [];
  const q = dashState.matrixSearch.trim().toLowerCase();
  let list = (dashState.data.employees || []).map((e) => ({
    ...e,
    w1Calls: e.weekBuckets?.W1 || 0,
    w2Calls: e.weekBuckets?.W2 || 0,
    w3Calls: e.weekBuckets?.W3 || 0,
  }));

  if (q) {
    list = list.filter((e) =>
      `${e.ff_id} ${e.ff_name} ${e.team} ${e.manager} ${e.base_town}`.toLowerCase().includes(q)
    );
  }

  const sortKey = dashState.matrixSortBy;
  const reverse = dashState.matrixSortDir === "desc" ? -1 : 1;
  const textKeys = new Set(["ff_id", "ff_name", "team", "manager", "base_town"]);

  list.sort((a, b) => {
    if (textKeys.has(sortKey)) {
      return String(a[sortKey] || "").localeCompare(String(b[sortKey] || "")) * reverse;
    }
    return (Number(a[sortKey] || 0) - Number(b[sortKey] || 0)) * reverse;
  });
  return list;
}

function renderEmployeeMatrix() {
  const list = getFilteredSortedEmployees();
  const total = list.length;
  const pageSize = dashState.matrixPageSize;
  const totalPages = Math.max(Math.ceil(total / pageSize), 1);
  if (dashState.matrixPage > totalPages) dashState.matrixPage = totalPages;
  if (dashState.matrixPage < 1) dashState.matrixPage = 1;

  const startIdx = (dashState.matrixPage - 1) * pageSize;
  const pageRows = list.slice(startIdx, startIdx + pageSize);
  const maxMonthlyCalls = Math.max(...list.map((e) => e.monthlyCalls || 0), 1);

  dashDom.empMatrixBadge.textContent = `${total.toLocaleString()} Employee${total === 1 ? "" : "s"}`;
  dashDom.empMatrixPageInfo.textContent = total
    ? `Showing ${(startIdx + 1).toLocaleString()}–${Math.min(startIdx + pageRows.length, total).toLocaleString()} of ${total.toLocaleString()} employees | Page ${dashState.matrixPage} of ${totalPages}`
    : "Showing 0–0 of 0 employees";
  dashDom.empMatrixPrev.disabled = dashState.matrixPage <= 1;
  dashDom.empMatrixNext.disabled = dashState.matrixPage >= totalPages;

  $$("th[data-emp-sort]", dashDom.empMatrixHead).forEach((th) => {
    const isSorted = th.dataset.empSort === dashState.matrixSortBy;
    th.classList.toggle("sorted", isSorted);
    const icon = th.querySelector(".sort-icon");
    if (icon) icon.textContent = isSorted ? (dashState.matrixSortDir === "asc" ? "↑" : "↓") : "↕";
  });

  if (!pageRows.length) {
    dashDom.empMatrixBody.innerHTML = '<tr><td colspan="23" class="fmcg-state-cell">No matching employees found.</td></tr>';
    dashDom.empMatrixFoot.innerHTML = "";
    return;
  }

  dashDom.empMatrixBody.innerHTML = pageRows.map((emp) => {
    const isSelected = dashState.focusedEmployeeKey === emp.employeeKey;
    const barPct = Math.min(Math.round((emp.monthlyCalls / maxMonthlyCalls) * 100), 100);
    return `<tr class="dash-emp-row-clickable${isSelected ? " is-selected-emp" : ""}" data-emp-key="${escapeHtml(emp.employeeKey)}" title="Click to focus charts on ${escapeHtml(emp.ff_name)}">
      <td><strong>${escapeHtml(emp.ff_id)}</strong></td>
      <td title="${escapeHtml(emp.ff_name)}">${escapeHtml(emp.ff_name)}</td>
      <td>${escapeHtml(emp.team)}</td>
      <td>${escapeHtml(emp.manager)}</td>
      <td class="group-end-col">${escapeHtml(emp.base_town)}</td>
      <td class="num-col">${emp.dailyCalls.toLocaleString()}</td>
      <td class="num-col">${emp.dailyAtt ? '<span class="call-badge call-badge-call">Present</span>' : '<span class="missing-metric">-</span>'}</td>
      <td class="num-col">${emp.dailyLcdCalls.toLocaleString()}</td>
      <td class="num-col group-end-col">${emp.dailyProducts.toLocaleString()}</td>
      <td class="num-col"><strong>${emp.weeklyCalls.toLocaleString()}</strong></td>
      <td class="num-col">${emp.weeklyAttDays}</td>
      <td class="num-col">${emp.weeklyAvgCalls.toFixed(1)}</td>
      <td class="num-col">${emp.weeklyLcdCalls.toLocaleString()}</td>
      <td class="num-col group-end-col">${emp.weeklyProducts.toLocaleString()}</td>
      <td class="num-col">
        <div class="dash-inline-bar-cell">
          <strong>${emp.monthlyCalls.toLocaleString()}</strong>
          <span class="dash-inline-bar-track"><span class="dash-inline-bar-fill" style="width:${barPct}%"></span></span>
        </div>
      </td>
      <td class="num-col">${emp.monthlyAttDays}</td>
      <td class="num-col">${emp.monthlyAvgCalls.toFixed(1)}</td>
      <td class="day-col">${emp.w1Calls.toLocaleString()}</td>
      <td class="day-col">${emp.w2Calls.toLocaleString()}</td>
      <td class="day-col">${emp.w3Calls.toLocaleString()}</td>
      <td class="num-col">${emp.monthlyLcdCalls.toLocaleString()}</td>
      <td class="num-col">${emp.monthlyProducts.toLocaleString()}</td>
      <td class="num-col">${emp.uniqueDoctors.toLocaleString()}</td>
    </tr>`;
  }).join("");

  const sumCol = (k) => list.reduce((acc, e) => acc + Number(e[k] || 0), 0);
  dashDom.empMatrixFoot.innerHTML = `<tr>
    <td colspan="5" class="tfoot-label">Total Summary (${list.length.toLocaleString()} Employees)</td>
    <td class="num-col">${sumCol("dailyCalls").toLocaleString()}</td>
    <td class="num-col">${sumCol("dailyAtt").toLocaleString()}</td>
    <td class="num-col">${sumCol("dailyLcdCalls").toLocaleString()}</td>
    <td class="num-col group-end-col">${sumCol("dailyProducts").toLocaleString()}</td>
    <td class="num-col">${sumCol("weeklyCalls").toLocaleString()}</td>
    <td class="num-col">${sumCol("weeklyAttDays").toLocaleString()}</td>
    <td class="num-col">-</td>
    <td class="num-col">${sumCol("weeklyLcdCalls").toLocaleString()}</td>
    <td class="num-col group-end-col">${sumCol("weeklyProducts").toLocaleString()}</td>
    <td class="num-col">${sumCol("monthlyCalls").toLocaleString()}</td>
    <td class="num-col">${sumCol("monthlyAttDays").toLocaleString()}</td>
    <td class="num-col">-</td>
    <td class="day-col">${sumCol("w1Calls").toLocaleString()}</td>
    <td class="day-col">${sumCol("w2Calls").toLocaleString()}</td>
    <td class="day-col">${sumCol("w3Calls").toLocaleString()}</td>
    <td class="num-col">${sumCol("monthlyLcdCalls").toLocaleString()}</td>
    <td class="num-col">${sumCol("monthlyProducts").toLocaleString()}</td>
    <td class="num-col">${sumCol("uniqueDoctors").toLocaleString()}</td>
  </tr>`;
}

function setFocusedEmployee(empKey) {
  dashState.focusedEmployeeKey = dashState.focusedEmployeeKey === empKey ? "" : empKey;
  const focused = getFocusedEmployee();
  if (focused) {
    dashDom.activeEmpPill.classList.remove("hidden");
    dashDom.activeEmpPill.innerHTML = `<span>Focused Employee: ${escapeHtml(focused.ff_name)} (${escapeHtml(focused.ff_id)})</span><button type="button" id="dash-clear-focused-emp" title="Clear focused employee">×</button>`;
    $("#dash-clear-focused-emp")?.addEventListener("click", () => setFocusedEmployee(""));
  } else {
    dashDom.activeEmpPill.classList.add("hidden");
    dashDom.activeEmpPill.innerHTML = "";
  }
  renderDashboardKpis();
  renderDashboardLineChart();
  renderEmployeeBarChart();
  renderEmployeeMatrix();
}

function renderAllDashboardViews() {
  if (!dashState.data) return;
  dashDom.headingDate.textContent = dashState.data.formattedDate || "Latest Available";
  renderDashboardKpis();
  renderDashboardLineChart();
  renderDashboardPieCharts();
  renderEmployeeBarChart();
  renderGroupBarChart();
  renderEmployeeMatrix();
}

async function loadDashboardAnalytics() {
  dashState.loading = true;
  dashDom.statusBar?.classList.add("hidden");
  dashDom.statusBar?.classList.remove("error");

  const params = new URLSearchParams();
  if (dashState.reportDate) params.set("report_date", dashState.reportDate);
  if (dashState.team) params.set("team", dashState.team);
  if (dashState.manager) params.set("manager", dashState.manager);
  if (dashState.baseTown) params.set("base_town", dashState.baseTown);
  if (dashState.callType) params.set("call_type", dashState.callType);
  if (dashState.callNature) params.set("call_nature", dashState.callNature);
  if (dashState.employee) params.set("employee", dashState.employee);

  try {
    const response = await api(`/api/dashboard/performance?${params.toString()}`);
    const data = await response.json();
    dashState.data = data;
    dashState.loaded = true;

    if (!dashState.reportDate && data.selectedDate) {
      dashState.reportDate = data.selectedDate;
      dashDom.reportDateInput.value = data.selectedDate;
    }

    if (data.filters) {
      if (!dashState.filtersPopulated) {
        dashMultiFilters.team?.setOptions(data.filters.teams || []);
        dashMultiFilters.manager?.setOptions(data.filters.managers || []);
        dashMultiFilters.baseTown?.setOptions(data.filters.baseTowns || []);
        dashMultiFilters.callType?.setOptions(data.filters.callTypes || ["CALL", "EXCEPTION", "NO CALLS"]);
        dashMultiFilters.callNature?.setOptions(data.filters.callNatures || ["Provider Calls", "Attendance", "LEAVE", "Others Exc Calls"]);

        const empOptions = (data.filters.employees || []).map((emp) => {
          const id = String(emp.ff_id || "").trim();
          const name = String(emp.ff_name || "").trim();
          if (id && name) return `${id} · ${name}`;
          return id || name;
        }).filter(Boolean);
        populateSelectOptions(dashDom.employeeFilter, "All Employees", empOptions, dashState.employee);

        populateSelectOptions(dashDom.caTeam, "All Teams", data.filters.teams || [], dashState.ca.team);
        populateSelectOptions(dashDom.caCallType, "All Call Types", ["CALL", "EXCEPTION", "NO CALLS"], dashState.ca.callType);
        populateSelectOptions(dashDom.lcdTeam, "All Teams", data.filters.teams || [], dashState.lcd.team);
        populateSelectOptions(dashDom.lcdNature, "All Call Natures", data.filters.callNatures || [], dashState.lcd.callNature);
        populateSelectOptions(dashDom.lcdCallType, "All CallTypes", ["CALL", "EXCEPTION"], dashState.lcd.callType);

        dashState.filtersPopulated = true;
      }
      dashMultiFilters.team?.setValue(dashState.team);
      dashMultiFilters.manager?.setValue(dashState.manager);
      dashMultiFilters.baseTown?.setValue(dashState.baseTown);
      dashMultiFilters.callType?.setValue(dashState.callType);
      dashMultiFilters.callNature?.setValue(dashState.callNature);
      if (dashDom.employeeFilter) dashDom.employeeFilter.value = dashState.employee;
    }

    renderAllDashboardViews();
    dashDom.statusBar?.classList.add("hidden");
  } catch (error) {
    dashDom.statusBar?.classList.remove("hidden");
    dashDom.statusBar?.classList.add("error");
    if (dashDom.statusText) dashDom.statusText.textContent = error.message || "Unable to load dashboard analytics.";
  } finally {
    dashState.loading = false;
  }
}

// ---------------------------------------------------------------------------
// Interactive Table 1: public.call_attendence
// ---------------------------------------------------------------------------
function formatCallBadge(val) {
  const upper = String(val || "").toUpperCase();
  if (upper === "CALL" || upper === "PROVIDER CALLS") return `<span class="call-badge call-badge-call">${escapeHtml(val)}</span>`;
  if (upper === "EXCEPTION" || upper === "ATTENDANCE") return `<span class="call-badge call-badge-exc">${escapeHtml(val)}</span>`;
  if (upper === "NO CALLS") return `<span class="call-badge call-badge-nocall">${escapeHtml(val)}</span>`;
  if (upper === "LEAVE") return `<span class="call-badge call-badge-leave">${escapeHtml(val)}</span>`;
  return `<span class="call-badge call-badge-other">${escapeHtml(val || "-")}</span>`;
}

function buildCaTableParams() {
  const s = dashState.ca;
  const p = new URLSearchParams();
  p.set("page", String(s.page));
  p.set("page_size", String(s.pageSize));
  if (s.search.trim()) p.set("search", s.search.trim());
  if (s.team) p.set("team", s.team);
  if (s.callType) p.set("call_type", s.callType);
  if (s.dateFrom) p.set("date_from", s.dateFrom);
  if (s.dateTo) p.set("date_to", s.dateTo);
  p.set("sort_by", s.sortBy);
  p.set("sort_dir", s.sortDir);
  return p.toString();
}

function renderSimplePageNums(container, page, totalPages) {
  const btns = [];
  const start = Math.max(1, page - 2);
  const end = Math.min(totalPages, page + 2);
  for (let p = start; p <= end; p += 1) {
    btns.push(`<button type="button" class="staff-page-num-btn${p === page ? " active" : ""}" data-page="${p}">${p}</button>`);
  }
  container.innerHTML = btns.join("");
}

async function loadCallAttendenceInteractiveTable() {
  try {
    const response = await api(`/api/dashboard/call-attendence?${buildCaTableParams()}`);
    const data = await response.json();
    const s = dashState.ca;
    s.rows = data.rows || [];
    s.total = Number(data.total || 0);
    s.page = Number(data.page || 1);
    s.totalPages = Number(data.totalPages || 1);
    s.summary = data.summary || {};

    dashDom.caSumRows.textContent = (s.summary.totalRows || 0).toLocaleString();
    dashDom.caSumCalls.textContent = (s.summary.callDone || 0).toLocaleString();
    dashDom.caSumExc.textContent = (s.summary.exceptionDone || 0).toLocaleString();
    dashDom.caSumNoCalls.textContent = (s.summary.noCallsRows || 0).toLocaleString();
    dashDom.caSumEmps.textContent = (s.summary.distinctEmployees || 0).toLocaleString();
    dashDom.caBadge.textContent = `${s.total.toLocaleString()} Rows`;

    $$("th[data-ca-sort]", dashDom.caHead).forEach((th) => {
      const isSorted = th.dataset.caSort === s.sortBy;
      th.classList.toggle("sorted", isSorted);
      const icon = th.querySelector(".sort-icon");
      if (icon) icon.textContent = isSorted ? (s.sortDir === "asc" ? "↑" : "↓") : "↕";
    });

    if (!s.rows.length) {
      dashDom.caBody.innerHTML = '<tr><td colspan="9" class="fmcg-state-cell">No records in public.call_attendence matched the current filters.</td></tr>';
    } else {
      dashDom.caBody.innerHTML = s.rows.map((r) => `<tr>
        <td>${escapeHtml(r.call_date)}</td>
        <td>${escapeHtml(r.team)}</td>
        <td><strong>${escapeHtml(r.ff_id)}</strong></td>
        <td title="${escapeHtml(r.ff_name)}">${escapeHtml(r.ff_name)}</td>
        <td>${escapeHtml(r.manager)}</td>
        <td>${escapeHtml(r.base_town)}</td>
        <td>${formatCallBadge(r.call_type)}</td>
        <td class="num-col"><strong>${Number(r.total_sum_of_done || 0).toLocaleString()}</strong></td>
        <td class="num-col">${Number(r.total_distinct_count_of_ff_id || 0).toLocaleString()}</td>
      </tr>`).join("");
    }

    const startRow = s.total === 0 ? 0 : (s.page - 1) * s.pageSize + 1;
    const endRow = s.total === 0 ? 0 : Math.min(startRow + s.rows.length - 1, s.total);
    dashDom.caPageInfo.textContent = `Showing ${startRow.toLocaleString()}–${endRow.toLocaleString()} of ${s.total.toLocaleString()} records | Page ${s.page} of ${s.totalPages}`;
    dashDom.caFirst.disabled = s.page <= 1;
    dashDom.caPrev.disabled = s.page <= 1;
    dashDom.caNext.disabled = s.page >= s.totalPages;
    dashDom.caLast.disabled = s.page >= s.totalPages;
    renderSimplePageNums(dashDom.caPageNums, s.page, s.totalPages);
  } catch (error) {
    dashDom.caBody.innerHTML = `<tr><td colspan="9" class="fmcg-state-cell">${escapeHtml(error.message || "Unable to load public.call_attendence")}</td></tr>`;
  }
}

// ---------------------------------------------------------------------------
// Interactive Table 2: public.live_call_details
// ---------------------------------------------------------------------------
function buildLcdTableParams() {
  const s = dashState.lcd;
  const p = new URLSearchParams();
  p.set("page", String(s.page));
  p.set("page_size", String(s.pageSize));
  if (s.search.trim()) p.set("search", s.search.trim());
  if (s.team) p.set("team", s.team);
  if (s.callNature) p.set("call_nature", s.callNature);
  if (s.callType) p.set("calltype", s.callType);
  if (s.dateFrom) p.set("date_from", s.dateFrom);
  if (s.dateTo) p.set("date_to", s.dateTo);
  p.set("sort_by", s.sortBy);
  p.set("sort_dir", s.sortDir);
  return p.toString();
}

async function loadLiveCallDetailsInteractiveTable() {
  try {
    const response = await api(`/api/dashboard/live-call-details?${buildLcdTableParams()}`);
    const data = await response.json();
    const s = dashState.lcd;
    s.rows = data.rows || [];
    s.total = Number(data.total || 0);
    s.page = Number(data.page || 1);
    s.totalPages = Number(data.totalPages || 1);
    s.summary = data.summary || {};

    dashDom.lcdSumRows.textContent = (s.summary.totalRows || 0).toLocaleString();
    dashDom.lcdSumProvider.textContent = (s.summary.providerRows || 0).toLocaleString();
    dashDom.lcdSumAtt.textContent = (s.summary.attendanceRows || 0).toLocaleString();
    dashDom.lcdSumExc.textContent = `${(s.summary.leaveRows || 0).toLocaleString()} / ${(s.summary.otherExcRows || 0).toLocaleString()}`;
    dashDom.lcdSumProducts.textContent = (s.summary.totalProducts || 0).toLocaleString();
    dashDom.lcdSumDoctors.textContent = (s.summary.distinctDoctors || 0).toLocaleString();
    dashDom.lcdSumInRange.textContent = (s.summary.inRangeRows || 0).toLocaleString();
    dashDom.lcdBadge.textContent = `${s.total.toLocaleString()} Rows`;

    $$("th[data-lcd-sort]", dashDom.lcdHead).forEach((th) => {
      const isSorted = th.dataset.lcdSort === s.sortBy;
      th.classList.toggle("sorted", isSorted);
      const icon = th.querySelector(".sort-icon");
      if (icon) icon.textContent = isSorted ? (s.sortDir === "asc" ? "↑" : "↓") : "↕";
    });

    if (!s.rows.length) {
      dashDom.lcdBody.innerHTML = '<tr><td colspan="16" class="fmcg-state-cell">No records in public.live_call_details matched the current filters.</td></tr>';
    } else {
      dashDom.lcdBody.innerHTML = s.rows.map((r) => `<tr>
        <td>${escapeHtml(r.call_date)}</td>
        <td>${escapeHtml(r.team)}</td>
        <td><strong>${escapeHtml(r.ff_id)}</strong></td>
        <td title="${escapeHtml(r.ff_name)}">${escapeHtml(r.ff_name)}</td>
        <td>${escapeHtml(r.manager || "-")}</td>
        <td>${escapeHtml(r.base_town || "-")}</td>
        <td>${escapeHtml(r.pmdc || "-")}</td>
        <td title="${escapeHtml(r.doctorname)}">${escapeHtml(r.doctorname || "-")}</td>
        <td>${formatCallBadge(r.call_nature)}</td>
        <td>${formatCallBadge(r.calltype)}</td>
        <td>${escapeHtml(r.ww || "-")}</td>
        <td class="num-col"><strong>${Number(r.total_sum_noofcalls || 0).toLocaleString()}</strong></td>
        <td class="num-col">${Number(r.total_sum_noofproducts || 0).toLocaleString()}</td>
        <td>${escapeHtml(r.geometer_range || "-")}</td>
        <td class="num-col">${escapeHtml(r.lng || "-")}</td>
        <td class="num-col">${escapeHtml(r.lat || "-")}</td>
      </tr>`).join("");
    }

    const startRow = s.total === 0 ? 0 : (s.page - 1) * s.pageSize + 1;
    const endRow = s.total === 0 ? 0 : Math.min(startRow + s.rows.length - 1, s.total);
    dashDom.lcdPageInfo.textContent = `Showing ${startRow.toLocaleString()}–${endRow.toLocaleString()} of ${s.total.toLocaleString()} records | Page ${s.page} of ${s.totalPages}`;
    dashDom.lcdFirst.disabled = s.page <= 1;
    dashDom.lcdPrev.disabled = s.page <= 1;
    dashDom.lcdNext.disabled = s.page >= s.totalPages;
    dashDom.lcdLast.disabled = s.page >= s.totalPages;
    renderSimplePageNums(dashDom.lcdPageNums, s.page, s.totalPages);
  } catch (error) {
    dashDom.lcdBody.innerHTML = `<tr><td colspan="16" class="fmcg-state-cell">${escapeHtml(error.message || "Unable to load public.live_call_details")}</td></tr>`;
  }
}

// ---------------------------------------------------------------------------
// Dashboard Event Bindings
// ---------------------------------------------------------------------------
dashDom.periodBtns.forEach((btn) => {
  btn.addEventListener("click", () => {
    dashState.period = btn.dataset.dashPeriod || "monthly";
    dashDom.periodBtns.forEach((b) => {
      const active = b === btn;
      b.classList.toggle("active", active);
      b.setAttribute("aria-selected", String(active));
    });
    // Sync trend mode with period if desired
    dashState.trendMode = dashState.period;
    dashDom.trendTabs.forEach((t) => t.classList.toggle("active", t.dataset.trendMode === dashState.trendMode));
    renderAllDashboardViews();
  });
});

dashDom.trendTabs.forEach((tab) => {
  tab.addEventListener("click", () => {
    dashState.trendMode = tab.dataset.trendMode || "daily";
    dashDom.trendTabs.forEach((t) => t.classList.toggle("active", t === tab));
    renderDashboardLineChart();
  });
});

dashDom.lineMetricSelect?.addEventListener("change", renderDashboardLineChart);

dashDom.empBarMetricSelect?.addEventListener("change", () => {
  dashState.empBarPage = 0;
  renderEmployeeBarChart();
});

dashDom.empBarPrev?.addEventListener("click", () => {
  dashState.empBarPage = Math.max(0, dashState.empBarPage - 1);
  renderEmployeeBarChart();
});

dashDom.empBarNext?.addEventListener("click", () => {
  dashState.empBarPage += 1;
  renderEmployeeBarChart();
});

dashDom.empBarContainer?.addEventListener("click", (event) => {
  const grp = event.target.closest(".dash-emp-bar-group");
  if (grp && grp.dataset.empKey) {
    setFocusedEmployee(grp.dataset.empKey);
  }
});

dashDom.groupDimTabs.forEach((tab) => {
  tab.addEventListener("click", () => {
    dashState.groupDim = tab.dataset.groupDim || "team";
    dashDom.groupDimTabs.forEach((t) => t.classList.toggle("active", t === tab));
    renderGroupBarChart();
  });
});

dashDom.toggleUploadsBtn?.addEventListener("click", () => {
  const hidden = dashDom.uploadsDrawer.classList.toggle("hidden");
  dashDom.toggleUploadsBtn.textContent = hidden ? "↑ Upload Data Files" : "✕ Hide Upload Drawer";
});

dashDom.filterToggleBtn?.addEventListener("click", () => {
  const isHidden = dashDom.filterCard.classList.toggle("hidden");
  dashDom.filterToggleBtn.setAttribute("aria-expanded", String(!isHidden));
  const label = isHidden ? "Show filters" : "Hide filters";
  dashDom.filterToggleBtn.title = label;
  dashDom.filterToggleBtn.setAttribute("aria-label", label);
  dashDom.filterToggleBtn.classList.toggle("active", !isHidden);
});

dashDom.refreshBtn?.addEventListener("click", () => {
  loadDashboardAnalytics();
  loadCallAttendenceInteractiveTable();
  loadLiveCallDetailsInteractiveTable();
});

dashDom.reportDateInput?.addEventListener("change", () => {
  dashState.reportDate = dashDom.reportDateInput.value || "";
  loadDashboardAnalytics();
});

dashDom.resetBtn?.addEventListener("click", () => {
  dashState.team = "";
  dashState.manager = "";
  dashState.baseTown = "";
  dashState.callType = "";
  dashState.callNature = "";
  dashState.employee = "";
  dashState.focusedEmployeeKey = "";
  dashState.empBarPage = 0;
  dashState.matrixPage = 1;
  dashMultiFilters.team?.setValue("");
  dashMultiFilters.manager?.setValue("");
  dashMultiFilters.baseTown?.setValue("");
  dashMultiFilters.callType?.setValue("");
  dashMultiFilters.callNature?.setValue("");
  if (dashDom.employeeFilter) dashDom.employeeFilter.value = "";
  if (dashDom.teamFilter) dashDom.teamFilter.value = "";
  if (dashDom.managerFilter) dashDom.managerFilter.value = "";
  if (dashDom.townFilter) dashDom.townFilter.value = "";
  if (dashDom.callTypeFilter) dashDom.callTypeFilter.value = "";
  if (dashDom.callNatureFilter) dashDom.callNatureFilter.value = "";
  dashDom.activeEmpPill?.classList.add("hidden");
  if (dashDom.activeEmpPill) dashDom.activeEmpPill.innerHTML = "";
  loadDashboardAnalytics();
});

// Employee Matrix events
dashDom.empMatrixSearch?.addEventListener("input", () => {
  dashState.matrixSearch = dashDom.empMatrixSearch.value;
  dashState.matrixPage = 1;
  renderEmployeeMatrix();
});

dashDom.empMatrixHead?.addEventListener("click", (event) => {
  const th = event.target.closest("th[data-emp-sort]");
  if (!th) return;
  const col = th.dataset.empSort;
  if (dashState.matrixSortBy === col) {
    dashState.matrixSortDir = dashState.matrixSortDir === "asc" ? "desc" : "asc";
  } else {
    dashState.matrixSortBy = col;
    dashState.matrixSortDir = "desc";
  }
  dashState.matrixPage = 1;
  renderEmployeeMatrix();
});

dashDom.empMatrixBody?.addEventListener("click", (event) => {
  const row = event.target.closest("tr[data-emp-key]");
  if (row && row.dataset.empKey) {
    setFocusedEmployee(row.dataset.empKey);
  }
});

dashDom.empMatrixPageSize?.addEventListener("change", () => {
  dashState.matrixPageSize = Number(dashDom.empMatrixPageSize.value) || 25;
  dashState.matrixPage = 1;
  renderEmployeeMatrix();
});

dashDom.empMatrixPrev?.addEventListener("click", () => {
  dashState.matrixPage = Math.max(1, dashState.matrixPage - 1);
  renderEmployeeMatrix();
});

dashDom.empMatrixNext?.addEventListener("click", () => {
  dashState.matrixPage += 1;
  renderEmployeeMatrix();
});

dashDom.empMatrixFs?.addEventListener("click", () => {
  toggleSheetFullscreen(dashDom.empMatrixSheet, dashDom.empMatrixFs);
});

dashDom.empMatrixExport?.addEventListener("click", () => {
  const list = getFilteredSortedEmployees();
  if (!list.length) return;
  const headers = [
    "FF.Id", "Employee Name", "Team", "Manager", "Base Town",
    "Daily Calls", "Daily Attn", "Daily Provider Calls", "Daily Products",
    "Weekly Calls", "Weekly Attn Days", "Weekly Avg Calls", "Weekly Provider Calls", "Weekly Products",
    "Monthly Calls", "Monthly Attn Days", "Monthly Avg Calls", "W1 Calls", "W2 Calls", "W3 Calls",
    "Monthly Provider Calls", "Monthly Products", "Unique Doctors",
  ];
  const csvRows = [
    headers.join(","),
    ...list.map((e) => [
      `"${String(e.ff_id).replace(/"/g, '""')}"`,
      `"${String(e.ff_name).replace(/"/g, '""')}"`,
      `"${String(e.team).replace(/"/g, '""')}"`,
      `"${String(e.manager).replace(/"/g, '""')}"`,
      `"${String(e.base_town).replace(/"/g, '""')}"`,
      e.dailyCalls, e.dailyAtt, e.dailyLcdCalls, e.dailyProducts,
      e.weeklyCalls, e.weeklyAttDays, e.weeklyAvgCalls, e.weeklyLcdCalls, e.weeklyProducts,
      e.monthlyCalls, e.monthlyAttDays, e.monthlyAvgCalls, e.w1Calls, e.w2Calls, e.w3Calls,
      e.monthlyLcdCalls, e.monthlyProducts, e.uniqueDoctors,
    ].join(",")),
  ];
  const blob = new Blob([csvRows.join("\n")], { type: "text/csv;charset=utf-8;" });
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = `employee_call_performance_${dashState.reportDate || "latest"}.csv`;
  link.click();
});

// DB Table Switcher
dashDom.dbTableTabs.forEach((tab) => {
  tab.addEventListener("click", () => {
    const view = tab.dataset.dbView || "split";
    dashDom.dbTableTabs.forEach((t) => t.classList.toggle("active", t === tab));
    dashDom.dbTablesGrid?.classList.toggle("split-mode", view === "split");
    dashDom.caCard.classList.toggle("hidden", view === "lcd");
    dashDom.lcdCard.classList.toggle("hidden", view === "ca");
  });
});

let dashResizeTimer = null;
window.addEventListener("resize", () => {
  clearTimeout(dashResizeTimer);
  dashResizeTimer = setTimeout(() => {
    if (dashState.data) {
      renderDashboardLineChart();
      renderEmployeeBarChart();
      renderGroupBarChart();
    }
  }, 150);
});

// Interactive Table 1 (call_attendence) Events
let caSearchTimer = null;
dashDom.caSearch?.addEventListener("input", () => {
  dashState.ca.search = dashDom.caSearch.value;
  dashState.ca.page = 1;
  clearTimeout(caSearchTimer);
  caSearchTimer = setTimeout(loadCallAttendenceInteractiveTable, 220);
});
dashDom.caTeam?.addEventListener("change", () => {
  dashState.ca.team = dashDom.caTeam.value;
  dashState.ca.page = 1;
  loadCallAttendenceInteractiveTable();
});
dashDom.caCallType?.addEventListener("change", () => {
  dashState.ca.callType = dashDom.caCallType.value;
  dashState.ca.page = 1;
  loadCallAttendenceInteractiveTable();
});
dashDom.caFrom?.addEventListener("change", () => {
  dashState.ca.dateFrom = dashDom.caFrom.value;
  dashState.ca.page = 1;
  loadCallAttendenceInteractiveTable();
});
dashDom.caTo?.addEventListener("change", () => {
  dashState.ca.dateTo = dashDom.caTo.value;
  dashState.ca.page = 1;
  loadCallAttendenceInteractiveTable();
});
dashDom.caReset?.addEventListener("click", () => {
  dashState.ca = { ...dashState.ca, page: 1, search: "", team: "", callType: "", dateFrom: "", dateTo: "", sortBy: "call_date", sortDir: "desc" };
  dashDom.caSearch.value = "";
  dashDom.caTeam.value = "";
  dashDom.caCallType.value = "";
  dashDom.caFrom.value = "";
  dashDom.caTo.value = "";
  loadCallAttendenceInteractiveTable();
});
dashDom.caExport?.addEventListener("click", () => {
  const link = document.createElement("a");
  link.href = `/api/dashboard/call-attendence/export?${buildCaTableParams()}`;
  link.click();
});
dashDom.caFs?.addEventListener("click", () => toggleSheetFullscreen(dashDom.caSheet, dashDom.caFs));
dashDom.caHead?.addEventListener("click", (event) => {
  const th = event.target.closest("th[data-ca-sort]");
  if (!th) return;
  const col = th.dataset.caSort;
  if (dashState.ca.sortBy === col) {
    dashState.ca.sortDir = dashState.ca.sortDir === "asc" ? "desc" : "asc";
  } else {
    dashState.ca.sortBy = col;
    dashState.ca.sortDir = "desc";
  }
  dashState.ca.page = 1;
  loadCallAttendenceInteractiveTable();
});
dashDom.caPageSize?.addEventListener("change", () => {
  dashState.ca.pageSize = Number(dashDom.caPageSize.value) || 25;
  dashState.ca.page = 1;
  loadCallAttendenceInteractiveTable();
});
dashDom.caFirst?.addEventListener("click", () => { dashState.ca.page = 1; loadCallAttendenceInteractiveTable(); });
dashDom.caPrev?.addEventListener("click", () => { dashState.ca.page = Math.max(1, dashState.ca.page - 1); loadCallAttendenceInteractiveTable(); });
dashDom.caNext?.addEventListener("click", () => { dashState.ca.page = Math.min(dashState.ca.totalPages, dashState.ca.page + 1); loadCallAttendenceInteractiveTable(); });
dashDom.caLast?.addEventListener("click", () => { dashState.ca.page = dashState.ca.totalPages; loadCallAttendenceInteractiveTable(); });
dashDom.caPageNums?.addEventListener("click", (event) => {
  const btn = event.target.closest("button[data-page]");
  if (btn) {
    dashState.ca.page = Number(btn.dataset.page) || 1;
    loadCallAttendenceInteractiveTable();
  }
});

// Interactive Table 2 (live_call_details) Events
let lcdSearchTimer = null;
dashDom.lcdSearch?.addEventListener("input", () => {
  dashState.lcd.search = dashDom.lcdSearch.value;
  dashState.lcd.page = 1;
  clearTimeout(lcdSearchTimer);
  lcdSearchTimer = setTimeout(loadLiveCallDetailsInteractiveTable, 220);
});
dashDom.lcdTeam?.addEventListener("change", () => {
  dashState.lcd.team = dashDom.lcdTeam.value;
  dashState.lcd.page = 1;
  loadLiveCallDetailsInteractiveTable();
});
dashDom.lcdNature?.addEventListener("change", () => {
  dashState.lcd.callNature = dashDom.lcdNature.value;
  dashState.lcd.page = 1;
  loadLiveCallDetailsInteractiveTable();
});
dashDom.lcdCallType?.addEventListener("change", () => {
  dashState.lcd.callType = dashDom.lcdCallType.value;
  dashState.lcd.page = 1;
  loadLiveCallDetailsInteractiveTable();
});
dashDom.lcdFrom?.addEventListener("change", () => {
  dashState.lcd.dateFrom = dashDom.lcdFrom.value;
  dashState.lcd.page = 1;
  loadLiveCallDetailsInteractiveTable();
});
dashDom.lcdTo?.addEventListener("change", () => {
  dashState.lcd.dateTo = dashDom.lcdTo.value;
  dashState.lcd.page = 1;
  loadLiveCallDetailsInteractiveTable();
});
dashDom.lcdReset?.addEventListener("click", () => {
  dashState.lcd = { ...dashState.lcd, page: 1, search: "", team: "", callNature: "", callType: "", dateFrom: "", dateTo: "", sortBy: "call_date", sortDir: "desc" };
  dashDom.lcdSearch.value = "";
  dashDom.lcdTeam.value = "";
  dashDom.lcdNature.value = "";
  dashDom.lcdCallType.value = "";
  dashDom.lcdFrom.value = "";
  dashDom.lcdTo.value = "";
  loadLiveCallDetailsInteractiveTable();
});
dashDom.lcdExport?.addEventListener("click", () => {
  const link = document.createElement("a");
  link.href = `/api/dashboard/live-call-details/export?${buildLcdTableParams()}`;
  link.click();
});
dashDom.lcdFs?.addEventListener("click", () => toggleSheetFullscreen(dashDom.lcdSheet, dashDom.lcdFs));
dashDom.lcdHead?.addEventListener("click", (event) => {
  const th = event.target.closest("th[data-lcd-sort]");
  if (!th) return;
  const col = th.dataset.lcdSort;
  if (dashState.lcd.sortBy === col) {
    dashState.lcd.sortDir = dashState.lcd.sortDir === "asc" ? "desc" : "asc";
  } else {
    dashState.lcd.sortBy = col;
    dashState.lcd.sortDir = "desc";
  }
  dashState.lcd.page = 1;
  loadLiveCallDetailsInteractiveTable();
});
dashDom.lcdPageSize?.addEventListener("change", () => {
  dashState.lcd.pageSize = Number(dashDom.lcdPageSize.value) || 25;
  dashState.lcd.page = 1;
  loadLiveCallDetailsInteractiveTable();
});
dashDom.lcdFirst?.addEventListener("click", () => { dashState.lcd.page = 1; loadLiveCallDetailsInteractiveTable(); });
dashDom.lcdPrev?.addEventListener("click", () => { dashState.lcd.page = Math.max(1, dashState.lcd.page - 1); loadLiveCallDetailsInteractiveTable(); });
dashDom.lcdNext?.addEventListener("click", () => { dashState.lcd.page = Math.min(dashState.lcd.totalPages, dashState.lcd.page + 1); loadLiveCallDetailsInteractiveTable(); });
dashDom.lcdLast?.addEventListener("click", () => { dashState.lcd.page = dashState.lcd.totalPages; loadLiveCallDetailsInteractiveTable(); });
dashDom.lcdPageNums?.addEventListener("click", (event) => {
  const btn = event.target.closest("button[data-page]");
  if (btn) {
    dashState.lcd.page = Number(btn.dataset.page) || 1;
    loadLiveCallDetailsInteractiveTable();
  }
});

// Initialize Dashboard and Both Interactive Tables on Startup
loadDashboardAnalytics();
loadCallAttendenceInteractiveTable();
loadLiveCallDetailsInteractiveTable();
