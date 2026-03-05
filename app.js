(() => {
  const PROJECT_STORAGE_KEY = "billuqrProjectV2";
  const SCHEMA_STORAGE_KEY = "billuqrSchemaV1";
  const QR_SIZE = 300;

  const state = {
    step: 1,
    fieldCount: 1,
    fieldNames: [],
    currentProjectId: "",
    currentProjectName: "",
    inputMethod: "manual",
    manualRecords: [],
    excelRecords: [],
    activeRecords: [],
    payloads: [],
    savedProjects: [],
    savedProjectDraft: null
  };

  const els = {
    stepIndicator: document.getElementById("stepIndicator"),
    statusMessage: document.getElementById("statusMessage"),
    startOverBtn: document.getElementById("startOverBtn"),
    savedProjectPanel: document.getElementById("savedProjectPanel"),
    savedSchemaSelect: document.getElementById("savedSchemaSelect"),
    savedProjectSummary: document.getElementById("savedProjectSummary"),
    continueSavedBtn: document.getElementById("continueSavedBtn"),
    continueClearTuplesBtn: document.getElementById("continueClearTuplesBtn"),
    startFreshBtn: document.getElementById("startFreshBtn"),
    savedTupleSelect: document.getElementById("savedTupleSelect"),
    savedSelectionCount: document.getElementById("savedSelectionCount"),
    deleteSelectedSavedBtn: document.getElementById("deleteSelectedSavedBtn"),
    clearAllSavedTuplesBtn: document.getElementById("clearAllSavedTuplesBtn"),

    step1: document.getElementById("step1"),
    fieldCount: document.getElementById("fieldCount"),
    step1NextBtn: document.getElementById("step1NextBtn"),

    step2: document.getElementById("step2"),
    fieldNamesForm: document.getElementById("fieldNamesForm"),
    step2BackBtn: document.getElementById("step2BackBtn"),
    step2ContinueBtn: document.getElementById("step2ContinueBtn"),

    step3: document.getElementById("step3"),
    manualMethodBtn: document.getElementById("manualMethodBtn"),
    excelMethodBtn: document.getElementById("excelMethodBtn"),
    manualSection: document.getElementById("manualSection"),
    excelSection: document.getElementById("excelSection"),
    step3BackBtn: document.getElementById("step3BackBtn"),

    manualForm: document.getElementById("manualForm"),
    saveTupleBtn: document.getElementById("saveTupleBtn"),
    generateManualBtn: document.getElementById("generateManualBtn"),
    manualCount: document.getElementById("manualCount"),
    manualRecordSelect: document.getElementById("manualRecordSelect"),

    excelFileInput: document.getElementById("excelFileInput"),
    parseExcelBtn: document.getElementById("parseExcelBtn"),
    generateExcelBtn: document.getElementById("generateExcelBtn"),
    excelCount: document.getElementById("excelCount"),

    previewSection: document.getElementById("previewSection"),
    previewRecordSelect: document.getElementById("previewRecordSelect"),
    multiRecordSelect: document.getElementById("multiRecordSelect"),
    multiSelectCount: document.getElementById("multiSelectCount"),
    qrCanvas: document.getElementById("qrCanvas"),
    payloadView: document.getElementById("payloadView"),
    selectAllBtn: document.getElementById("selectAllBtn"),
    clearSelectionBtn: document.getElementById("clearSelectionBtn"),
    downloadSelectedBtn: document.getElementById("downloadSelectedBtn"),
    downloadAllBtn: document.getElementById("downloadAllBtn"),
    downloadBtn: document.getElementById("downloadBtn")
  };

  function setStatus(message, isError = false) {
    els.statusMessage.textContent = message;
    els.statusMessage.classList.toggle("error", isError);
  }

  function setSavedPanelVisible(isVisible) {
    els.savedProjectPanel.classList.toggle("hidden", !isVisible);
    els.stepIndicator.classList.toggle("hidden", isVisible);
    els.step1.classList.toggle("hidden", isVisible || state.step !== 1);
    els.step2.classList.toggle("hidden", isVisible || state.step !== 2);
    els.step3.classList.toggle("hidden", isVisible || state.step !== 3);
    els.previewSection.classList.toggle("hidden", isVisible || state.step !== 4);
  }

  function getQrOptions() {
    return {
      width: QR_SIZE,
      margin: 1,
      color: {
        dark: "#000000",
        light: "#ffffff"
      }
    };
  }

  function ensureCanvasSize(canvas) {
    if (canvas.width !== QR_SIZE) {
      canvas.width = QR_SIZE;
    }
    if (canvas.height !== QR_SIZE) {
      canvas.height = QR_SIZE;
    }
  }

  function toSolidWhitePngDataUrl(sourceCanvas) {
    const exportCanvas = document.createElement("canvas");
    exportCanvas.width = sourceCanvas.width;
    exportCanvas.height = sourceCanvas.height;
    const ctx = exportCanvas.getContext("2d");
    ctx.fillStyle = "#ffffff";
    ctx.fillRect(0, 0, exportCanvas.width, exportCanvas.height);
    ctx.drawImage(sourceCanvas, 0, 0);
    return exportCanvas.toDataURL("image/png");
  }

  async function renderQrToCanvas(canvas, payload) {
    ensureCanvasSize(canvas);
    let firstError = null;

    if (window.QRCode && typeof window.QRCode.toCanvas === "function") {
      try {
        await window.QRCode.toCanvas(canvas, payload, getQrOptions());
        return;
      } catch (error) {
        firstError = error;
      }
    }

    if (typeof window.QRious !== "undefined") {
      try {
        new window.QRious({
          element: canvas,
          value: payload,
          size: QR_SIZE,
          background: "#ffffff",
          foreground: "#000000",
          level: "L"
        });
        return;
      } catch (error) {
        if (!firstError) {
          firstError = error;
        }
      }
    }

    const suffix = firstError ? ` (${firstError.message})` : "";
    throw new Error(`QR library failed to render${suffix}`);
  }

  function validateFieldNameList(fieldNames) {
    if (!Array.isArray(fieldNames) || fieldNames.length < 1 || fieldNames.length > 50) {
      return null;
    }

    const cleaned = [];
    const seen = new Set();

    for (const rawName of fieldNames) {
      const name = String(rawName).trim();
      if (!name) {
        return null;
      }

      const normalized = name.toLowerCase();
      if (seen.has(normalized)) {
        return null;
      }

      seen.add(normalized);
      cleaned.push(name);
    }

    return cleaned;
  }

  function normalizeRecordArray(rawRecords, fieldNames) {
    if (!Array.isArray(rawRecords)) {
      return [];
    }

    const records = [];
    for (const rawRecord of rawRecords) {
      if (!rawRecord || typeof rawRecord !== "object") {
        continue;
      }

      const record = {};
      for (const field of fieldNames) {
        record[field] = String(rawRecord[field] ?? "").trim();
      }
      records.push(record);
    }
    return records;
  }

  function createProjectId() {
    return `schema_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`;
  }

  function defaultSchemaName(fieldNames, index = 0) {
    const fromFields = fieldNames.slice(0, 2).join(" + ");
    return fromFields || `Schema ${index + 1}`;
  }

  function createUniqueSchemaName(fieldNames, existingProjects) {
    const base = defaultSchemaName(fieldNames, existingProjects.length);
    const used = new Set(existingProjects.map((project) => project.name.toLowerCase()));
    if (!used.has(base.toLowerCase())) {
      return base;
    }

    let suffix = 2;
    while (used.has(`${base} (${suffix})`.toLowerCase())) {
      suffix += 1;
    }
    return `${base} (${suffix})`;
  }

  function normalizeProject(rawProject, index = 0) {
    if (!rawProject || typeof rawProject !== "object") {
      return null;
    }

    const fieldNames = validateFieldNameList(rawProject.fieldNames);
    if (!fieldNames) {
      return null;
    }

    const id =
      typeof rawProject.id === "string" && rawProject.id.trim()
        ? rawProject.id.trim()
        : createProjectId();
    const name =
      typeof rawProject.name === "string" && rawProject.name.trim()
        ? rawProject.name.trim()
        : defaultSchemaName(fieldNames, index);

    return {
      id,
      name,
      fieldNames,
      manualRecords: normalizeRecordArray(rawProject.manualRecords, fieldNames),
      excelRecords: normalizeRecordArray(rawProject.excelRecords, fieldNames),
      inputMethod: rawProject.inputMethod === "excel" ? "excel" : "manual"
    };
  }

  function cloneProject(project) {
    return JSON.parse(JSON.stringify(project));
  }

  function clearSavedProjectStorage() {
    try {
      localStorage.removeItem(PROJECT_STORAGE_KEY);
      localStorage.removeItem(SCHEMA_STORAGE_KEY);
    } catch (error) {
      // Non-blocking.
    }
  }

  function normalizeProjectList(rawValue) {
    let candidates = [];
    if (Array.isArray(rawValue)) {
      candidates = rawValue;
    } else if (rawValue && typeof rawValue === "object" && Array.isArray(rawValue.projects)) {
      candidates = rawValue.projects;
    } else if (rawValue && typeof rawValue === "object") {
      candidates = [rawValue];
    }

    const projects = [];
    for (let i = 0; i < candidates.length; i += 1) {
      const normalized = normalizeProject(candidates[i], i);
      if (normalized) {
        projects.push(normalized);
      }
    }

    return projects;
  }

  function readSavedProjects() {
    try {
      const raw = localStorage.getItem(PROJECT_STORAGE_KEY);
      if (raw) {
        const parsed = JSON.parse(raw);
        const normalizedList = normalizeProjectList(parsed);
        if (normalizedList.length) {
          return normalizedList;
        }
      }
    } catch (error) {
      // fallback to legacy schema below
    }

    try {
      const legacyRaw = localStorage.getItem(SCHEMA_STORAGE_KEY);
      if (!legacyRaw) {
        return [];
      }

      const legacyParsed = JSON.parse(legacyRaw);
      const legacyFieldNames = validateFieldNameList(legacyParsed.fieldNames);
      if (!legacyFieldNames) {
        return [];
      }

      return [
        {
          id: createProjectId(),
          name: defaultSchemaName(legacyFieldNames, 0),
          fieldNames: legacyFieldNames,
          manualRecords: [],
          excelRecords: [],
          inputMethod: "manual"
        }
      ];
    } catch (error) {
      return [];
    }
  }

  function writeSavedProjects(projects) {
    if (!projects.length) {
      clearSavedProjectStorage();
      return;
    }

    try {
      localStorage.setItem(PROJECT_STORAGE_KEY, JSON.stringify(projects));
      localStorage.setItem(SCHEMA_STORAGE_KEY, JSON.stringify({ fieldNames: projects[0].fieldNames }));
    } catch (error) {
      // Non-blocking.
    }
  }

  function persistProject(project) {
    const normalized = normalizeProject(project);
    if (!normalized) {
      clearSavedProjectStorage();
      return [];
    }

    const projects = readSavedProjects();
    const existingIndex = projects.findIndex((item) => item.id === normalized.id);
    if (existingIndex >= 0) {
      projects[existingIndex] = normalized;
    } else {
      projects.push(normalized);
    }

    writeSavedProjects(projects);
    return projects;
  }

  function persistProjectFromState() {
    if (!state.fieldNames.length) {
      return [];
    }

    if (!state.currentProjectId) {
      const existing = readSavedProjects();
      state.currentProjectId = createProjectId();
      state.currentProjectName = createUniqueSchemaName(state.fieldNames, existing);
    }

    const saved = persistProject({
      id: state.currentProjectId,
      name: state.currentProjectName || defaultSchemaName(state.fieldNames),
      fieldNames: state.fieldNames,
      manualRecords: state.manualRecords,
      excelRecords: state.excelRecords,
      inputMethod: state.inputMethod
    });

    state.savedProjects = saved;
    return saved;
  }

  function readSavedProject() {
    const projects = readSavedProjects();
    return projects.length ? projects[0] : null;
  }

  function showStep(step) {
    state.step = step;
    els.step1.classList.toggle("hidden", step !== 1);
    els.step2.classList.toggle("hidden", step !== 2);
    els.step3.classList.toggle("hidden", step !== 3);
    els.previewSection.classList.toggle("hidden", step !== 4);

    const labels = {
      1: "Step 1 of 3: Choose number of fields",
      2: "Step 2 of 3: Enter field names",
      3: "Step 3 of 3: Choose input method",
      4: "QR Preview"
    };
    els.stepIndicator.textContent = labels[step] || "BilluQR";
  }

  function makeSafeId(value, index) {
    const base = value.toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-|-$/g, "");
    return `field-${base || "name"}-${index + 1}`;
  }

  function renderFieldNameInputs(count, prefilledNames = []) {
    const html = [];
    for (let i = 0; i < count; i += 1) {
      html.push(
        `<div class="input-row">
          <label for="fieldName-${i + 1}">Field ${i + 1}</label>
          <input id="fieldName-${i + 1}" type="text" placeholder="Enter field name" />
        </div>`
      );
    }
    els.fieldNamesForm.innerHTML = html.join("");

    const inputs = els.fieldNamesForm.querySelectorAll("input");
    for (let i = 0; i < inputs.length; i += 1) {
      const value = prefilledNames[i];
      if (typeof value === "string") {
        inputs[i].value = value;
      }
    }
  }

  function validateFieldNames() {
    const inputs = els.fieldNamesForm.querySelectorAll("input");
    const names = [];
    const seen = new Set();

    for (const input of inputs) {
      const name = input.value.trim();
      if (!name) {
        return { ok: false, error: "Field names cannot be empty." };
      }
      const normalized = name.toLowerCase();
      if (seen.has(normalized)) {
        return { ok: false, error: `Field names must be unique. Duplicate: ${name}` };
      }
      seen.add(normalized);
      names.push(name);
    }

    return { ok: true, names };
  }

  function renderManualForm() {
    const html = state.fieldNames.map((field, index) => {
      const id = makeSafeId(field, index);
      return `<div class="input-row">
        <label for="${id}">${field}</label>
        <input id="${id}" data-field="${field}" type="text" placeholder="Enter ${field}" />
      </div>`;
    });
    els.manualForm.innerHTML = html.join("");
  }

  function readManualForm() {
    const inputs = els.manualForm.querySelectorAll("input[data-field]");
    const record = {};
    let hasValue = false;

    for (const input of inputs) {
      const field = input.getAttribute("data-field");
      const value = input.value.trim();
      if (value) {
        hasValue = true;
      }
      record[field] = value;
    }

    if (!hasValue) {
      return null;
    }

    return record;
  }

  function clearManualForm() {
    const inputs = els.manualForm.querySelectorAll("input[data-field]");
    for (const input of inputs) {
      input.value = "";
    }
    if (inputs[0]) {
      inputs[0].focus();
    }
  }

  function getRecordLabel(record, index) {
    const firstField = state.fieldNames[0];
    const firstValue = record[firstField] || "";
    return firstValue ? `Record ${index + 1} - ${firstValue}` : `Record ${index + 1}`;
  }

  function buildPlainTextPayload(record) {
    return state.fieldNames
      .map((field) => `${field}: ${record[field] ?? ""}`)
      .join("\n");
  }

  function refreshManualRecordsUI() {
    const count = state.manualRecords.length;
    els.manualCount.textContent = `Saved tuples: ${count}`;
    els.generateManualBtn.disabled = count === 0;

    if (!count) {
      els.manualRecordSelect.innerHTML = '<option value="">No records yet</option>';
      els.manualRecordSelect.disabled = true;
      return;
    }

    els.manualRecordSelect.disabled = false;
    const options = state.manualRecords.map((record, index) => {
      return `<option value="${index}">${getRecordLabel(record, index)}</option>`;
    });
    els.manualRecordSelect.innerHTML = options.join("");
  }

  function refreshExcelRecordsUI() {
    const count = state.excelRecords.length;
    els.excelCount.textContent = `Parsed tuples: ${count}`;
    els.generateExcelBtn.disabled = count === 0;
  }

  function switchInputMethod(method, shouldPersist = true) {
    state.inputMethod = method === "excel" ? "excel" : "manual";
    const isManual = state.inputMethod === "manual";
    els.manualMethodBtn.classList.toggle("active", isManual);
    els.excelMethodBtn.classList.toggle("active", !isManual);
    els.manualSection.classList.toggle("hidden", !isManual);
    els.excelSection.classList.toggle("hidden", isManual);
    if (shouldPersist && state.fieldNames.length) {
      persistProjectFromState();
    }
    setStatus("");
  }

  function clearPreviewCanvas() {
    ensureCanvasSize(els.qrCanvas);
    const context = els.qrCanvas.getContext("2d");
    context.clearRect(0, 0, els.qrCanvas.width, els.qrCanvas.height);
    context.fillStyle = "#ffffff";
    context.fillRect(0, 0, els.qrCanvas.width, els.qrCanvas.height);
  }

  function getSelectedRecordIndexes() {
    const values = [];
    const options = els.multiRecordSelect.selectedOptions;
    for (const option of options) {
      const index = Number.parseInt(option.value, 10);
      if (Number.isInteger(index)) {
        values.push(index);
      }
    }
    return values;
  }

  function refreshMultiSelectActions() {
    const selectedCount = getSelectedRecordIndexes().length;
    els.multiSelectCount.textContent = `Selected for ZIP: ${selectedCount}`;
    els.downloadSelectedBtn.disabled = selectedCount === 0;
    els.downloadAllBtn.disabled = state.activeRecords.length === 0;
  }

  function updatePreviewSelector() {
    const options = state.activeRecords.map((record, index) => {
      return `<option value="${index}">${getRecordLabel(record, index)}</option>`;
    });
    const html = options.join("");
    els.previewRecordSelect.innerHTML = html;
    els.multiRecordSelect.innerHTML = html;
    refreshMultiSelectActions();
  }

  async function renderSelectedRecord(index) {
    const record = state.activeRecords[index];
    const payload = state.payloads[index];

    if (!record || !payload) {
      clearPreviewCanvas();
      els.payloadView.value = "";
      return;
    }

    els.payloadView.value = payload;

    try {
      await renderQrToCanvas(els.qrCanvas, payload);
    } catch (error) {
      setStatus(`Failed to render QR code: ${error.message}`, true);
    }
  }

  async function generateQRCodes(records) {
    if (!records.length) {
      setStatus("No records available to generate QR codes.", true);
      return;
    }

    state.activeRecords = [...records];
    state.payloads = state.activeRecords.map((record) => buildPlainTextPayload(record));

    updatePreviewSelector();
    els.previewRecordSelect.value = "0";
    for (const option of els.multiRecordSelect.options) {
      option.selected = false;
    }
    refreshMultiSelectActions();
    showStep(4);
    await renderSelectedRecord(0);

    setStatus(`Generated ${records.length} QR code payload(s).`);
  }

  function parseExcelFile(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (event) => {
        try {
          const data = new Uint8Array(event.target.result);
          const workbook = XLSX.read(data, { type: "array" });
          const firstSheetName = workbook.SheetNames[0];
          if (!firstSheetName) {
            reject(new Error("The workbook does not contain any sheets."));
            return;
          }

          const sheet = workbook.Sheets[firstSheetName];
          const rows = XLSX.utils.sheet_to_json(sheet, {
            header: 1,
            defval: "",
            blankrows: false
          });

          if (!rows.length) {
            reject(new Error("The selected sheet is empty."));
            return;
          }

          const rawHeaders = rows[0].map((cell) => String(cell).trim());
          const missing = state.fieldNames.filter((required) => !rawHeaders.includes(required));
          if (missing.length) {
            reject(new Error(`Missing columns: ${missing.join(", ")}`));
            return;
          }

          const records = [];
          for (let i = 1; i < rows.length; i += 1) {
            const row = rows[i];
            const rowValues = rawHeaders.map((_, idx) => (row[idx] ?? ""));
            const isFullyEmpty = rowValues.every((value) => String(value).trim() === "");
            if (isFullyEmpty) {
              continue;
            }

            const record = {};
            for (const field of state.fieldNames) {
              const colIndex = rawHeaders.indexOf(field);
              record[field] = String(row[colIndex] ?? "").trim();
            }
            records.push(record);
          }

          resolve(records);
        } catch (error) {
          reject(new Error("Could not parse Excel file."));
        }
      };
      reader.onerror = () => reject(new Error("Could not read the selected file."));
      reader.readAsArrayBuffer(file);
    });
  }

  function summarizeProject(project) {
    const manualCount = project.manualRecords.length;
    const excelCount = project.excelRecords.length;
    const total = manualCount + excelCount;
    return `Schema: ${project.name}. Fields: ${project.fieldNames.length}. Saved tuples: ${total} (Manual ${manualCount}, Excel ${excelCount}).`;
  }

  function renderSavedSchemaSelect() {
    if (!state.savedProjects.length) {
      els.savedSchemaSelect.innerHTML = '<option value="">No saved schemas</option>';
      els.savedSchemaSelect.disabled = true;
      return;
    }

    const options = state.savedProjects.map((project) => {
      const tupleCount = project.manualRecords.length + project.excelRecords.length;
      return `<option value="${project.id}">${project.name} (${project.fieldNames.length} fields, ${tupleCount} tuples)</option>`;
    });

    els.savedSchemaSelect.innerHTML = options.join("");
    els.savedSchemaSelect.disabled = false;

    if (state.savedProjectDraft) {
      els.savedSchemaSelect.value = state.savedProjectDraft.id;
    }
  }

  function selectSavedProjectById(projectId) {
    const selected = state.savedProjects.find((project) => project.id === projectId);
    if (!selected) {
      return false;
    }

    state.savedProjectDraft = cloneProject(selected);
    renderSavedSchemaSelect();
    renderSavedProjectPanel();
    return true;
  }

  function getSavedTupleEntries(project) {
    const entries = [];

    for (let i = 0; i < project.manualRecords.length; i += 1) {
      entries.push({ source: "manual", index: i, record: project.manualRecords[i] });
    }

    for (let i = 0; i < project.excelRecords.length; i += 1) {
      entries.push({ source: "excel", index: i, record: project.excelRecords[i] });
    }

    return entries;
  }

  function refreshSavedSelectionActions() {
    let selectedCount = 0;
    for (const option of els.savedTupleSelect.options) {
      if (option.selected) {
        selectedCount += 1;
      }
    }

    els.savedSelectionCount.textContent = `Selected tuples: ${selectedCount}`;
    els.deleteSelectedSavedBtn.disabled = selectedCount === 0;
  }

  function renderSavedProjectPanel() {
    if (!state.savedProjectDraft) {
      return;
    }

    const project = state.savedProjectDraft;
    els.savedProjectSummary.textContent = summarizeProject(project);

    const entries = getSavedTupleEntries(project);
    if (!entries.length) {
      els.savedTupleSelect.innerHTML = '<option value="">No saved tuples</option>';
      els.savedTupleSelect.disabled = true;
      refreshSavedSelectionActions();
      return;
    }

    els.savedTupleSelect.disabled = false;
    const firstField = project.fieldNames[0];
    const options = entries.map((entry) => {
      const sourceName = entry.source === "manual" ? "Manual" : "Excel";
      const displayIndex = entry.index + 1;
      const firstValue = entry.record[firstField] || "";
      const label = firstValue
        ? `${sourceName} ${displayIndex} - ${firstValue}`
        : `${sourceName} ${displayIndex}`;
      return `<option value="${entry.source}:${entry.index}">${label}</option>`;
    });
    els.savedTupleSelect.innerHTML = options.join("");
    refreshSavedSelectionActions();
  }

  function applyProjectToState(project) {
    const normalized = normalizeProject(project);
    if (!normalized) {
      return false;
    }

    state.currentProjectId = normalized.id;
    state.currentProjectName = normalized.name;
    state.fieldNames = normalized.fieldNames;
    state.fieldCount = normalized.fieldNames.length;
    state.manualRecords = normalized.manualRecords;
    state.excelRecords = normalized.excelRecords;
    state.inputMethod = normalized.inputMethod;
    state.activeRecords = [];
    state.payloads = [];

    els.fieldCount.value = String(state.fieldCount);
    renderFieldNameInputs(state.fieldCount, state.fieldNames);
    renderManualForm();
    refreshManualRecordsUI();
    refreshExcelRecordsUI();
    switchInputMethod(state.inputMethod, false);

    els.previewRecordSelect.innerHTML = "";
    els.multiRecordSelect.innerHTML = "";
    els.payloadView.value = "";
    refreshMultiSelectActions();
    clearPreviewCanvas();
    return true;
  }

  function continueFromSavedProject(clearTuples) {
    if (!state.savedProjectDraft) {
      return;
    }

    const project = cloneProject(state.savedProjectDraft);
    if (clearTuples) {
      project.manualRecords = [];
      project.excelRecords = [];
    }

    if (!applyProjectToState(project)) {
      setStatus("Saved project could not be loaded. Please start fresh.", true);
      return;
    }

    state.savedProjects = persistProject(project);
    state.savedProjectDraft = null;
    setSavedPanelVisible(false);
    showStep(3);

    if (clearTuples) {
      setStatus("Loaded previous schema. All tuples were cleared.");
    } else {
      const total = state.manualRecords.length + state.excelRecords.length;
      setStatus(`Loaded previous schema and ${total} saved tuple(s).`);
    }
  }

  function onDeleteSelectedSavedTuples() {
    if (!state.savedProjectDraft) {
      return;
    }

    const selected = [];
    for (const option of els.savedTupleSelect.selectedOptions) {
      const [source, indexText] = option.value.split(":");
      const index = Number.parseInt(indexText, 10);
      if ((source === "manual" || source === "excel") && Number.isInteger(index)) {
        selected.push({ source, index });
      }
    }

    if (!selected.length) {
      return;
    }

    const manualIndexes = selected
      .filter((item) => item.source === "manual")
      .map((item) => item.index)
      .sort((a, b) => b - a);

    const excelIndexes = selected
      .filter((item) => item.source === "excel")
      .map((item) => item.index)
      .sort((a, b) => b - a);

    for (const index of manualIndexes) {
      if (index >= 0 && index < state.savedProjectDraft.manualRecords.length) {
        state.savedProjectDraft.manualRecords.splice(index, 1);
      }
    }

    for (const index of excelIndexes) {
      if (index >= 0 && index < state.savedProjectDraft.excelRecords.length) {
        state.savedProjectDraft.excelRecords.splice(index, 1);
      }
    }

    persistProject(state.savedProjectDraft);
    renderSavedProjectPanel();
    setStatus("Selected saved tuples were deleted.");
  }

  function onClearAllSavedTuples() {
    if (!state.savedProjectDraft) {
      return;
    }

    state.savedProjectDraft.manualRecords = [];
    state.savedProjectDraft.excelRecords = [];
    persistProject(state.savedProjectDraft);
    renderSavedProjectPanel();
    setStatus("All saved tuples were cleared. Schema is kept.");
  }

  function onStartFreshFromSavedPanel() {
    const confirmed = window.confirm("Clear saved schema and all saved tuples?");
    if (!confirmed) {
      return;
    }

    clearSavedProjectStorage();
    state.savedProjectDraft = null;
    state.fieldCount = 1;
    state.fieldNames = [];
    state.inputMethod = "manual";
    state.manualRecords = [];
    state.excelRecords = [];
    state.activeRecords = [];
    state.payloads = [];

    els.fieldCount.value = "1";
    renderFieldNameInputs(1);
    renderManualForm();
    refreshManualRecordsUI();
    refreshExcelRecordsUI();
    els.previewRecordSelect.innerHTML = "";
    els.multiRecordSelect.innerHTML = "";
    els.payloadView.value = "";
    clearPreviewCanvas();
    refreshMultiSelectActions();

    setSavedPanelVisible(false);
    showStep(1);
    setStatus("Started fresh. Set your new schema.");
  }

  function onStep1Next() {
    const value = Number.parseInt(els.fieldCount.value, 10);
    if (!Number.isInteger(value) || value < 1 || value > 50) {
      setStatus("Please enter a number between 1 and 50.", true);
      return;
    }

    state.fieldCount = value;
    renderFieldNameInputs(value);
    showStep(2);
    setStatus("");
  }

  function onStep2Continue() {
    const result = validateFieldNames();
    if (!result.ok) {
      setStatus(result.error, true);
      return;
    }

    state.fieldNames = result.names;
    state.manualRecords = [];
    state.excelRecords = [];
    state.activeRecords = [];
    state.payloads = [];
    state.inputMethod = "manual";

    renderManualForm();
    refreshManualRecordsUI();
    refreshExcelRecordsUI();
    els.excelFileInput.value = "";

    switchInputMethod("manual", false);
    persistProjectFromState();
    showStep(3);
    setStatus("Field names saved. Choose an input method.");
  }

  function onSaveTuple() {
    const record = readManualForm();
    if (!record) {
      setStatus("Enter at least one value before saving a tuple.", true);
      return;
    }

    state.manualRecords.push(record);
    refreshManualRecordsUI();
    persistProjectFromState();
    clearManualForm();
    setStatus(`Tuple saved. Total saved: ${state.manualRecords.length}.`);
  }

  async function onGenerateManual() {
    await generateQRCodes(state.manualRecords);
  }

  async function onParseExcel() {
    const file = els.excelFileInput.files && els.excelFileInput.files[0];
    if (!file) {
      setStatus("Please choose an Excel file first.", true);
      return;
    }

    try {
      const parsedRecords = await parseExcelFile(file);
      if (!parsedRecords.length) {
        setStatus("No usable rows found after parsing (empty rows are skipped).", true);
        state.excelRecords = [];
        refreshExcelRecordsUI();
        persistProjectFromState();
        return;
      }

      state.excelRecords = parsedRecords;
      refreshExcelRecordsUI();
      persistProjectFromState();
      setStatus(`Excel parsed successfully. Records loaded: ${parsedRecords.length}.`);
    } catch (error) {
      setStatus(error.message, true);
    }
  }

  async function onGenerateExcel() {
    await generateQRCodes(state.excelRecords);
  }

  async function onPreviewSelectChange() {
    const index = Number.parseInt(els.previewRecordSelect.value, 10);
    if (!Number.isInteger(index)) {
      return;
    }
    await renderSelectedRecord(index);
  }

  function onDownload() {
    const index = Number.parseInt(els.previewRecordSelect.value, 10);
    if (!Number.isInteger(index)) {
      setStatus("Select a record before downloading.", true);
      return;
    }

    const link = document.createElement("a");
    link.href = toSolidWhitePngDataUrl(els.qrCanvas);
    link.download = `BilluQR_Record-${index + 1}.png`;
    link.click();
    setStatus(`Downloaded BilluQR_Record-${index + 1}.png`);
  }

  function onSelectAll() {
    for (const option of els.multiRecordSelect.options) {
      option.selected = true;
    }
    refreshMultiSelectActions();
  }

  function onClearSelection() {
    for (const option of els.multiRecordSelect.options) {
      option.selected = false;
    }
    refreshMultiSelectActions();
  }

  function onMultiSelectionChange() {
    refreshMultiSelectActions();
  }

  async function renderPayloadToPngDataUrl(payload) {
    const canvas = document.createElement("canvas");
    await renderQrToCanvas(canvas, payload);
    return toSolidWhitePngDataUrl(canvas);
  }

  async function downloadRecordsAsZip(recordIndexes, zipFileName) {
    if (!recordIndexes.length) {
      setStatus("Please select one or more records for ZIP download.", true);
      return;
    }

    if (typeof JSZip === "undefined") {
      setStatus("ZIP download is unavailable because JSZip failed to load.", true);
      return;
    }

    els.downloadSelectedBtn.disabled = true;
    els.downloadAllBtn.disabled = true;
    setStatus("Preparing ZIP download...");

    try {
      const zip = new JSZip();

      for (const index of recordIndexes) {
        const payload = state.payloads[index];
        if (!payload) {
          continue;
        }

        const dataUrl = await renderPayloadToPngDataUrl(payload);
        const base64 = dataUrl.split(",")[1];
        zip.file(`BilluQR_Record-${index + 1}.png`, base64, { base64: true });
      }

      const zipBlob = await zip.generateAsync({ type: "blob" });
      const objectUrl = URL.createObjectURL(zipBlob);
      const link = document.createElement("a");
      link.href = objectUrl;
      link.download = zipFileName;
      document.body.appendChild(link);
      link.click();
      link.remove();
      setTimeout(() => URL.revokeObjectURL(objectUrl), 1500);

      setStatus(`Downloaded ${zipFileName}`);
    } catch (error) {
      setStatus(`Failed to create ZIP: ${error.message}`, true);
    } finally {
      refreshMultiSelectActions();
    }
  }

  async function onDownloadSelectedZip() {
    const selectedIndexes = getSelectedRecordIndexes();
    await downloadRecordsAsZip(selectedIndexes, "BilluQR_Selected_Records.zip");
  }

  async function onDownloadAllZip() {
    const allIndexes = state.activeRecords.map((_, index) => index);
    await downloadRecordsAsZip(allIndexes, "BilluQR_All_Records.zip");
  }

  function onStartOver() {
    const confirmed = window.confirm("Start over and clear your saved schema and tuples?");
    if (!confirmed) {
      return;
    }

    clearSavedProjectStorage();
    window.location.reload();
  }

  function showSavedProjectPanel(project) {
    state.savedProjectDraft = cloneProject(project);
    renderSavedProjectPanel();
    setSavedPanelVisible(true);
    setStatus("Previous schema and tuples found. Continue or clean up saved data.");
  }

  function bindEvents() {
    els.step1NextBtn.addEventListener("click", onStep1Next);

    els.step2BackBtn.addEventListener("click", () => {
      showStep(1);
      setStatus("");
    });
    els.step2ContinueBtn.addEventListener("click", onStep2Continue);

    els.manualMethodBtn.addEventListener("click", () => switchInputMethod("manual", true));
    els.excelMethodBtn.addEventListener("click", () => switchInputMethod("excel", true));
    els.step3BackBtn.addEventListener("click", () => {
      renderFieldNameInputs(state.fieldCount, state.fieldNames);
      showStep(2);
      setStatus("");
    });

    els.saveTupleBtn.addEventListener("click", onSaveTuple);
    els.generateManualBtn.addEventListener("click", onGenerateManual);

    els.parseExcelBtn.addEventListener("click", onParseExcel);
    els.generateExcelBtn.addEventListener("click", onGenerateExcel);

    els.previewRecordSelect.addEventListener("change", onPreviewSelectChange);
    els.multiRecordSelect.addEventListener("change", onMultiSelectionChange);
    els.selectAllBtn.addEventListener("click", onSelectAll);
    els.clearSelectionBtn.addEventListener("click", onClearSelection);
    els.downloadSelectedBtn.addEventListener("click", onDownloadSelectedZip);
    els.downloadAllBtn.addEventListener("click", onDownloadAllZip);
    els.downloadBtn.addEventListener("click", onDownload);
    els.startOverBtn.addEventListener("click", onStartOver);

    els.continueSavedBtn.addEventListener("click", () => continueFromSavedProject(false));
    els.continueClearTuplesBtn.addEventListener("click", () => continueFromSavedProject(true));
    els.startFreshBtn.addEventListener("click", onStartFreshFromSavedPanel);
    els.savedTupleSelect.addEventListener("change", refreshSavedSelectionActions);
    els.deleteSelectedSavedBtn.addEventListener("click", onDeleteSelectedSavedTuples);
    els.clearAllSavedTuplesBtn.addEventListener("click", onClearAllSavedTuples);
  }

  function init() {
    bindEvents();
    clearPreviewCanvas();
    renderFieldNameInputs(state.fieldCount);
    renderManualForm();
    refreshManualRecordsUI();
    refreshExcelRecordsUI();
    showStep(1);

    const savedProject = readSavedProject();
    if (savedProject) {
      showSavedProjectPanel(savedProject);
      return;
    }

    setSavedPanelVisible(false);
    setStatus("Set the number of fields to begin.");
  }

  init();
})();
