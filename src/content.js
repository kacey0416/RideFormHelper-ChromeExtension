(() => {
  if (window.__safeRideHelperInjected) {
    return;
  }
  window.__safeRideHelperInjected = true;

  const PROFILE_STORAGE_KEY = "saferideProfileV2";
  const LEGACY_PROFILE_STORAGE_KEY = "saferideProfileV1";
  const PROFILE_SAVED_AT_KEY = "saferideProfileSavedAtV1";
  const QUEUE_STORAGE_KEY = "saferideQueueV2";
  const QUEUE_SAVED_AT_KEY = "saferideQueueSavedAtV1";
  const CURRENT_RECEIPT_STORAGE_KEY = "saferideCurrentReceiptIdV2";
  const HELPER_ID = "saferide-helper-root";
  const STORAGE_TTL_MS = 3 * 24 * 60 * 60 * 1000;
  const MAX_RECEIPT_FILES = 30;
  const AVG_SCREENSHOT_IMAGE_BYTES = 2 * 1024 * 1024;
  const MAX_RECEIPT_FILE_BYTES = AVG_SCREENSHOT_IMAGE_BYTES * 2;
  const FORM_HOME_URL = "https://app.smartsheet.com/b/form/6b5489b0fe984763b2e96350fdcbd9a2";
  const STATUS_ICON_URLS = {
    complete: chrome.runtime.getURL("src/icons/check-circle.svg"),
    completeDone: chrome.runtime.getURL("src/icons/check-circle-gray.svg"),
    error: chrome.runtime.getURL("src/icons/cross-circle.svg"),
    progress: chrome.runtime.getURL("src/icons/loading.svg")
  };

  const FIELD_MAP = {
    partnerName: { key: "DzjPkPn", label: "Partner Name" },
    partnerNumber: { key: "ONGKyKm", label: "Partner Number" },
    email: { key: "dQ3K0wO", label: "Email Address" },
    dateOfRide: { key: "NORAKJN", label: "Date of Ride" },
    timeOfRide: { key: "E2logwY", label: "Time of Ride" },
    costOfRide: { key: "MDqgzJO", label: "Cost of Ride" },
    storeNumber: { key: "eKQPWJR", label: "Store Number" },
    province: { key: "JE1nW0Q", label: "Province or Territory" },
    certify: {
      key: "NOR5eyn",
      label:
        "I certify that the information submitted is true and correct to the best of my knowledge."
    },
    signature: { key: "GN9QKwQ", label: "Electronic Signature" },
    receipt: { key: "ATTACHMENT", label: "Receipt of Ride" }
  };

  const PROVINCES = [
    "Alberta",
    "British Columbia",
    "Manitoba",
    "New Brunswick",
    "Newfoundland & Labrador",
    "Northwest Territories",
    "Nova Scotia",
    "Nunavut",
    "Ontario",
    "Prince Edward Island",
    "Quebec",
    "Saskatchewan",
    "Yukon"
  ];

  const PROVINCE_ALIASES = {
    ab: "Alberta",
    bc: "British Columbia",
    mb: "Manitoba",
    nb: "New Brunswick",
    nl: "Newfoundland & Labrador",
    nt: "Northwest Territories",
    ns: "Nova Scotia",
    nu: "Nunavut",
    on: "Ontario",
    pe: "Prince Edward Island",
    pei: "Prince Edward Island",
    qc: "Quebec",
    sk: "Saskatchewan",
    yt: "Yukon"
  };

  const MONTHS = {
    jan: 1,
    january: 1,
    feb: 2,
    february: 2,
    mar: 3,
    march: 3,
    apr: 4,
    april: 4,
    may: 5,
    jun: 6,
    june: 6,
    jul: 7,
    july: 7,
    aug: 8,
    august: 8,
    sep: 9,
    sept: 9,
    september: 9,
    oct: 10,
    october: 10,
    nov: 11,
    november: 11,
    dec: 12,
    december: 12
  };

  const state = {
    profile: {
      partnerName: "",
      partnerNumber: "",
      email: "",
      storeNumber: "",
      province: "Ontario",
      signature: ""
    },
    receipts: [],
    currentReceiptId: null,
    ocrWorker: null,
    ocrLoading: false,
    currentOcrReceiptId: null,
    currentOcrReceiptRef: null,
    ocrRecognitionActive: false,
    ocrProgressLastRenderAt: 0,
    fillInFlightReceiptId: null,
    previewReceiptId: null,
    submitClickedAt: 0,
    showReturnHomeBtn: false,
    formDefinition: null
  };

  const els = {};
  let queueSaveTail = Promise.resolve();

  injectHelperUI();
  wireUIEvents();
  attachSubmitHints();
  observeConfirmationMessage();
  void bootstrap();

  async function bootstrap() {
    await loadProfile();
    await loadQueueFromStorage();
    const profileComplete = syncProfileCompletionUI({ forceOpenWhenIncomplete: true });
    setSectionCollapsed(els.profileSection, profileComplete);
    setSectionCollapsed(els.receiptSection, els.receiptSection.dataset.collapsed === "true");
    renderQueue();
    setStatus("SafeRide Helper loaded. Add receipt screenshots to start.");
  }

  function injectHelperUI() {
    const root = document.createElement("aside");
    root.id = HELPER_ID;
    root.innerHTML = `
      <div class="saferide-helper-body">
        <section class="saferide-helper-section" id="srhProfileSection" data-collapsed="false">
          <div class="saferide-helper-section-head">
            <div class="saferide-helper-section-head-row">
              <span class="saferide-helper-section-title">Partner Info</span>
              <div class="saferide-helper-section-head-actions">
                <span class="saferide-helper-section-check is-warning" id="srhProfileWarning" hidden>!</span>
                <span class="saferide-helper-section-check" id="srhProfileCheck" hidden>OK</span>
                <button type="button" class="saferide-helper-section-collapse-btn" id="srhProfileToggle" aria-label="Collapse section" aria-expanded="true"></button>
              </div>
            </div>
          </div>
          <div class="saferide-helper-section-content">
            <div class="saferide-helper-grid saferide-helper-partner-grid">
              <div class="saferide-helper-field">
                <label for="srhPartnerName">Partner Name</label>
                <input id="srhPartnerName" type="text" autocomplete="off" />
              </div>
              <div class="saferide-helper-field">
                <label for="srhPartnerNumber">Partner Number</label>
                <input id="srhPartnerNumber" type="text" inputmode="numeric" pattern="[0-9]*" autocomplete="off" />
              </div>
              <div class="saferide-helper-field">
                <label for="srhStoreNumber">Store Number</label>
                <input id="srhStoreNumber" type="text" inputmode="numeric" pattern="[0-9]*" autocomplete="off" />
              </div>
              <div class="saferide-helper-field">
                <label for="srhProvince">Province</label>
                <select id="srhProvince">
                  ${PROVINCES.map((p) => `<option value="${escapeHtml(p)}">${escapeHtml(p)}</option>`).join("")}
                </select>
              </div>
              <div class="saferide-helper-field full">
                <label for="srhEmail">Email Address</label>
                <input id="srhEmail" type="email" autocomplete="off" />
              </div>
              <div class="saferide-helper-field full">
                <label for="srhSignature">Signature</label>
                <input id="srhSignature" type="text" autocomplete="off" />
              </div>
            </div>
          </div>
          <div class="saferide-helper-section-footer saferide-helper-row">
            <button type="button" class="saferide-helper-clear-text-btn" id="srhClearProfile">Clear</button>
            <button type="button" id="srhSaveProfile">Save</button>
          </div>
        </section>

        <section class="saferide-helper-section" id="srhReceiptSection" data-collapsed="false">
          <div class="saferide-helper-section-head">
            <div class="saferide-helper-section-head-row">
              <span class="saferide-helper-section-title">Receipts</span>
              <div class="saferide-helper-section-head-actions">
                <button type="button" class="saferide-helper-return-home-btn" id="srhReturnHomeBtn" hidden>Back to Form</button>
                <button type="button" class="saferide-helper-section-collapse-btn" id="srhReceiptToggle" aria-label="Collapse section" aria-expanded="true"></button>
              </div>
            </div>
          </div>
          <div class="saferide-helper-section-content">
            <div class="saferide-helper-queue-wrap" id="srhQueueWrap" hidden>
              <div class="saferide-helper-queue-count" id="srhQueueCount"></div>
              <ul class="saferide-helper-queue" id="srhQueue"></ul>
            </div>
            <div class="saferide-helper-dropzone" id="srhDropzone" aria-label="Drop receipt images here">
              <div class="saferide-helper-dropzone-title">Drop receipt images here</div>
              <div class="saferide-helper-dropzone-subtitle">or use Browse Files</div>
              <button type="button" class="saferide-helper-queue-show" id="srhBrowseBtn">Browse Files</button>
              <input class="saferide-helper-file-input srh-hidden-file-input" id="srhFileInput" type="file" accept="image/*" multiple />
            </div>
          </div>
        </section>
      </div>
    `;

    document.body.appendChild(root);

    els.root = root;
    els.partnerName = root.querySelector("#srhPartnerName");
    els.partnerNumber = root.querySelector("#srhPartnerNumber");
    els.email = root.querySelector("#srhEmail");
    els.storeNumber = root.querySelector("#srhStoreNumber");
    els.province = root.querySelector("#srhProvince");
    els.signature = root.querySelector("#srhSignature");
    els.profileSection = root.querySelector("#srhProfileSection");
    els.profileToggle = root.querySelector("#srhProfileToggle");
    els.profileWarning = root.querySelector("#srhProfileWarning");
    els.profileCheck = root.querySelector("#srhProfileCheck");
    els.receiptSection = root.querySelector("#srhReceiptSection");
    els.receiptToggle = root.querySelector("#srhReceiptToggle");
    els.returnHomeBtn = root.querySelector("#srhReturnHomeBtn");
    els.saveProfile = root.querySelector("#srhSaveProfile");
    els.clearProfile = root.querySelector("#srhClearProfile");
    els.dropzone = root.querySelector("#srhDropzone");
    els.browseBtn = root.querySelector("#srhBrowseBtn");
    els.fileInput = root.querySelector("#srhFileInput");
    els.queueWrap = root.querySelector("#srhQueueWrap");
    els.queueCount = root.querySelector("#srhQueueCount");
    els.queue = root.querySelector("#srhQueue");
  }

  function wireUIEvents() {
    els.saveProfile.addEventListener("click", async () => {
      readProfileFromUI();
      await saveProfile();
      const complete = syncProfileCompletionUI({ forceOpenWhenIncomplete: true });
      if (complete) {
        setSectionCollapsed(els.profileSection, true);
        setSectionCollapsed(els.receiptSection, false);
        setStatus("Profile saved. Partner info collapsed and receipts opened.");
      } else {
        setSectionCollapsed(els.profileSection, false);
        scrollSectionIntoView(els.profileSection);
        setStatus("Profile saved, but some required partner info is still missing.");
      }
    });

    els.clearProfile.addEventListener("click", async () => {
      state.profile = {
        partnerName: "",
        partnerNumber: "",
        email: "",
        storeNumber: "",
        province: "Ontario",
        signature: ""
      };
      renderProfile();
      syncProfileCompletionUI({ forceOpenWhenIncomplete: true });
      setSectionCollapsed(els.profileSection, false);
      scrollSectionIntoView(els.profileSection);
      await saveProfile();
      setStatus("Profile cleared.");
    });

    els.profileToggle.addEventListener("click", () => {
      const willCollapse = els.profileSection.dataset.collapsed !== "true";
      setSectionCollapsed(els.profileSection, willCollapse);
      if (!willCollapse) {
        scrollSectionIntoView(els.profileSection);
      }
    });

    els.receiptToggle.addEventListener("click", () => {
      const willCollapse = els.receiptSection.dataset.collapsed !== "true";
      setSectionCollapsed(els.receiptSection, willCollapse);
    });

    if (els.returnHomeBtn) {
      els.returnHomeBtn.addEventListener("click", () => {
        window.location.href = FORM_HOME_URL;
      });
    }

    const profileInputs = [
      els.partnerName,
      els.partnerNumber,
      els.email,
      els.storeNumber,
      els.province,
      els.signature
    ];
    for (const input of profileInputs) {
      input.addEventListener("input", () => {
        readProfileFromUI();
        syncProfileCompletionUI();
      });
      input.addEventListener("change", () => {
        readProfileFromUI();
        syncProfileCompletionUI();
      });
    }

    els.fileInput.addEventListener("change", async (event) => {
      await handleIncomingFiles(Array.from(event.target.files || []));
      event.target.value = "";
    });

    els.browseBtn.addEventListener("click", (event) => {
      event.preventDefault();
      event.stopPropagation();
      els.fileInput.click();
    });

    const dragEvents = ["dragenter", "dragover"];
    for (const eventName of dragEvents) {
      els.dropzone.addEventListener(eventName, (event) => {
        event.preventDefault();
        event.stopPropagation();
        els.dropzone.classList.add("drag-over");
      });
    }

    const leaveEvents = ["dragleave", "dragend", "drop"];
    for (const eventName of leaveEvents) {
      els.dropzone.addEventListener(eventName, (event) => {
        event.preventDefault();
        event.stopPropagation();
        els.dropzone.classList.remove("drag-over");
      });
    }

    els.dropzone.addEventListener("drop", (event) => {
      const files = Array.from(event.dataTransfer?.files || []);
      void handleIncomingFiles(files);
    });

    els.queue.addEventListener("click", async (event) => {
      const deleteBtn = event.target.closest(".saferide-helper-queue-delete");
      if (deleteBtn) {
        const deleteTarget = deleteBtn.closest("li[data-id]");
        if (deleteTarget?.dataset?.id) {
          await deleteReceiptById(deleteTarget.dataset.id);
        }
        return;
      }

      const showBtn = event.target.closest(".saferide-helper-queue-show");
      if (showBtn) {
        const showTarget = showBtn.closest("li[data-id]");
        if (!showTarget?.dataset?.id) {
          return;
        }
        const targetId = showTarget.dataset.id;
        const changedSelection = state.currentReceiptId !== targetId;
        state.currentReceiptId = targetId;
        state.previewReceiptId = state.previewReceiptId === targetId ? null : targetId;
        if (state.previewReceiptId) {
          setSectionCollapsed(els.profileSection, true);
          setSectionCollapsed(els.receiptSection, false);
        }
        renderQueue();
        if (changedSelection) {
          void maybeAutoFillSelectedReceipt(targetId);
        }
        return;
      }

      const li = event.target.closest("li[data-id]");
      if (!li) {
        return;
      }
      const selectedId = li.dataset.id;
      const changedSelection = state.currentReceiptId !== selectedId;
      state.currentReceiptId = selectedId;
      if (changedSelection) {
        state.previewReceiptId = null;
      }
      renderQueue();
      const item = getCurrentReceipt();
      if (item) {
        setStatus(`Selected: ${getReceiptDisplayName(item)}`);
      }
      if (changedSelection) {
        void maybeAutoFillSelectedReceipt(selectedId);
      }
    });
  }

  function readProfileFromUI() {
    state.profile.partnerName = els.partnerName.value.trim();
    state.profile.partnerNumber = normalizeDigits(els.partnerNumber.value);
    state.profile.email = els.email.value.trim();
    state.profile.storeNumber = normalizeDigits(els.storeNumber.value);
    state.profile.province = els.province.value;
    state.profile.signature = els.signature.value.trim();

    if (els.partnerNumber && els.partnerNumber.value !== state.profile.partnerNumber) {
      els.partnerNumber.value = state.profile.partnerNumber;
    }
    if (els.storeNumber && els.storeNumber.value !== state.profile.storeNumber) {
      els.storeNumber.value = state.profile.storeNumber;
    }
  }

  function renderProfile() {
    els.partnerName.value = state.profile.partnerName;
    els.partnerNumber.value = normalizeDigits(state.profile.partnerNumber);
    els.email.value = state.profile.email;
    els.storeNumber.value = normalizeDigits(state.profile.storeNumber);
    els.province.value = state.profile.province || "Ontario";
    els.signature.value = state.profile.signature;
    state.profile.partnerNumber = els.partnerNumber.value;
    state.profile.storeNumber = els.storeNumber.value;
    syncProfileCompletionUI();
  }

  function setSectionCollapsed(sectionEl, collapsed) {
    if (!sectionEl) {
      return;
    }
    sectionEl.dataset.collapsed = collapsed ? "true" : "false";
    const toggle = sectionEl.querySelector(".saferide-helper-section-collapse-btn");
    if (toggle) {
      toggle.dataset.collapsed = collapsed ? "true" : "false";
      toggle.setAttribute("aria-expanded", collapsed ? "false" : "true");
      toggle.setAttribute("aria-label", collapsed ? "Expand section" : "Collapse section");
    }
  }

  function scrollSectionIntoView(sectionEl) {
    if (!sectionEl || typeof sectionEl.scrollIntoView !== "function") {
      return;
    }
    sectionEl.scrollIntoView({ block: "start", behavior: "smooth" });
  }

  function isProfileComplete(profile = state.profile) {
    const partnerNumber = String(profile.partnerNumber || "").trim();
    const email = String(profile.email || "").trim();
    const storeNumber = String(profile.storeNumber || "").trim();
    const required = [
      profile.partnerName,
      partnerNumber,
      email,
      storeNumber,
      profile.province,
      profile.signature
    ];
    if (!required.every((value) => String(value || "").trim().length > 0)) {
      return false;
    }
    if (!isDigitsOnly(partnerNumber) || !isDigitsOnly(storeNumber)) {
      return false;
    }
    if (!isValidEmail(email)) {
      return false;
    }
    return true;
  }

  function syncProfileFieldValidationUI(profile = state.profile) {
    const partnerNumber = String(profile.partnerNumber || "").trim();
    const email = String(profile.email || "").trim();
    const storeNumber = String(profile.storeNumber || "").trim();
    const fieldState = {
      partnerName: String(profile.partnerName || "").trim().length > 0,
      partnerNumber: isDigitsOnly(partnerNumber),
      email: isValidEmail(email),
      storeNumber: isDigitsOnly(storeNumber),
      province: String(profile.province || "").trim().length > 0,
      signature: String(profile.signature || "").trim().length > 0
    };

    const controls = {
      partnerName: els.partnerName,
      partnerNumber: els.partnerNumber,
      email: els.email,
      storeNumber: els.storeNumber,
      province: els.province,
      signature: els.signature
    };

    let hasMissing = false;
    for (const [key, control] of Object.entries(controls)) {
      if (!control) {
        continue;
      }
      const valid = Boolean(fieldState[key]);
      const missing = !valid;
      const wrapper = control.closest(".saferide-helper-field");
      if (wrapper) {
        wrapper.classList.toggle("is-missing", missing);
      }
      control.setAttribute("aria-invalid", missing ? "true" : "false");
      if (missing) {
        hasMissing = true;
      }
    }

    return !hasMissing;
  }

  function syncProfileCompletionUI(options = {}) {
    const forceOpenWhenIncomplete = Boolean(options.forceOpenWhenIncomplete);
    syncProfileFieldValidationUI();
    const complete = isProfileComplete();
    if (els.profileCheck) {
      els.profileCheck.hidden = !complete;
      els.profileCheck.style.display = complete ? "inline-flex" : "none";
    }
    if (els.profileWarning) {
      const missing = !complete;
      els.profileWarning.hidden = !missing;
      els.profileWarning.style.display = missing ? "inline-flex" : "none";
    }

    if (!complete && forceOpenWhenIncomplete) {
      setSectionCollapsed(els.profileSection, false);
      scrollSectionIntoView(els.profileSection);
    }
    return complete;
  }

  async function loadProfile() {
    try {
      const result = await chrome.storage.local.get([
        PROFILE_STORAGE_KEY,
        LEGACY_PROFILE_STORAGE_KEY,
        PROFILE_SAVED_AT_KEY
      ]);
      const profileSavedAt = toTimestamp(result?.[PROFILE_SAVED_AT_KEY]);
      if (isStorageExpired(profileSavedAt)) {
        await chrome.storage.local.remove([
          PROFILE_STORAGE_KEY,
          LEGACY_PROFILE_STORAGE_KEY,
          PROFILE_SAVED_AT_KEY
        ]);
        setStatus("Stored profile expired after 3 days and was cleared.");
        renderProfile();
        return;
      }

      const savedProfile = result?.[PROFILE_STORAGE_KEY] || result?.[LEGACY_PROFILE_STORAGE_KEY];
      if (savedProfile && typeof savedProfile === "object") {
        state.profile = {
          ...state.profile,
          ...savedProfile
        };
        if (!profileSavedAt) {
          await chrome.storage.local.set({
            [PROFILE_SAVED_AT_KEY]: Date.now()
          });
        }
      }
    } catch (error) {
      setStatus(`Profile load warning: ${error.message}`);
    }
    renderProfile();
  }

  async function saveProfile() {
    try {
      await chrome.storage.local.set({
        [PROFILE_STORAGE_KEY]: state.profile,
        [PROFILE_SAVED_AT_KEY]: Date.now()
      });
    } catch (error) {
      setStatus(`Profile save failed: ${error.message}`);
    }
  }

  async function loadQueueFromStorage() {
    try {
      const queueFromSession = readSessionJson(QUEUE_STORAGE_KEY, []);
      const savedQueue = Array.isArray(queueFromSession) ? queueFromSession : [];
      const currentIdValue = readSessionJson(CURRENT_RECEIPT_STORAGE_KEY, "");
      const currentId = typeof currentIdValue === "string" ? currentIdValue : null;
      const storedQueueSavedAt = toTimestamp(readSessionJson(QUEUE_SAVED_AT_KEY, 0));
      const fallbackQueueSavedAt = getLatestQueueTimestamp(savedQueue);
      const queueSavedAt = storedQueueSavedAt || fallbackQueueSavedAt;

      if (isStorageExpired(queueSavedAt)) {
        removeSessionKeys([
          QUEUE_STORAGE_KEY,
          CURRENT_RECEIPT_STORAGE_KEY,
          QUEUE_SAVED_AT_KEY
        ]);
        state.receipts = [];
        state.currentReceiptId = null;
        setStatus("Stored receipt queue expired after 3 days and was cleared.");
        return;
      }

      state.receipts = savedQueue
        .map((item) => normalizeSavedReceipt(item))
        .filter(Boolean);

      if (currentId && state.receipts.some((item) => item.id === currentId)) {
        state.currentReceiptId = currentId;
      } else {
        state.currentReceiptId = state.receipts[0]?.id || null;
      }

      if (!storedQueueSavedAt && state.receipts.length) {
        writeSessionJson(QUEUE_SAVED_AT_KEY, Date.now());
      }

    } catch (error) {
      setStatus(`Queue load warning: ${error.message}`);
    }
  }

  function buildQueueStoragePayload() {
    return {
      [QUEUE_STORAGE_KEY]: state.receipts.map((item) => ({
        id: item.id,
        fileName: item.fileName,
        mimeType: item.mimeType,
        fileSize: item.fileSize,
        dataUrl: item.dataUrl,
        status: item.status,
        progress: Number(item.progress || 0),
        parsed: item.parsed || null,
        error: item.error || "",
        createdAt: item.createdAt || Date.now(),
        updatedAt: item.updatedAt || Date.now()
      })),
      [CURRENT_RECEIPT_STORAGE_KEY]: state.currentReceiptId || "",
      [QUEUE_SAVED_AT_KEY]: Date.now()
    };
  }

  function saveQueueToStorage() {
    queueSaveTail = queueSaveTail
      .catch(() => undefined)
      .then(async () => {
        try {
          const payload = buildQueueStoragePayload();
          for (const [key, value] of Object.entries(payload)) {
            writeSessionJson(key, value);
          }
        } catch (error) {
          setStatus(`Queue save warning: ${error.message}`);
        }
      });
    return queueSaveTail;
  }

  function normalizeSavedReceipt(item) {
    if (!item || typeof item !== "object") {
      return null;
    }
    if (!item.id || !item.fileName || !item.dataUrl) {
      return null;
    }

    const savedStatus = String(item.status || "pending");
    const normalizedStatus = savedStatus === "parsing" ? "pending" : savedStatus;

    return {
      id: String(item.id),
      fileName: String(item.fileName),
      mimeType: String(item.mimeType || "image/png"),
      fileSize: Number(item.fileSize || 0),
      dataUrl: String(item.dataUrl),
      status: normalizedStatus,
      progress: Number(item.progress || 0),
      parsed: item.parsed && typeof item.parsed === "object" ? item.parsed : null,
      error: String(item.error || ""),
      createdAt: Number(item.createdAt || Date.now()),
      updatedAt: Number(item.updatedAt || Date.now())
    };
  }

  function getLatestQueueTimestamp(queueItems) {
    if (!Array.isArray(queueItems) || !queueItems.length) {
      return 0;
    }

    let latest = 0;
    for (const item of queueItems) {
      if (!item || typeof item !== "object") {
        continue;
      }
      const updatedAt = toTimestamp(item.updatedAt);
      const createdAt = toTimestamp(item.createdAt);
      latest = Math.max(latest, updatedAt, createdAt);
    }
    return latest;
  }

  function isStorageExpired(savedAt) {
    if (!savedAt) {
      return false;
    }
    return Date.now() - savedAt > STORAGE_TTL_MS;
  }

  function toTimestamp(value) {
    const n = Number(value);
    if (!Number.isFinite(n) || n <= 0) {
      return 0;
    }
    return n;
  }

  function readSessionJson(key, fallback = null) {
    try {
      const raw = window.sessionStorage.getItem(key);
      if (raw == null || raw === "") {
        return fallback;
      }
      return JSON.parse(raw);
    } catch (error) {
      setStatus(`Session read warning (${key}): ${error.message}`);
      return fallback;
    }
  }

  function writeSessionJson(key, value) {
    try {
      window.sessionStorage.setItem(key, JSON.stringify(value));
      return true;
    } catch (error) {
      setStatus(`Session save warning (${key}): ${error.message}`);
      return false;
    }
  }

  function removeSessionKeys(keys) {
    if (!Array.isArray(keys) || !keys.length) {
      return;
    }
    for (const key of keys) {
      try {
        window.sessionStorage.removeItem(key);
      } catch (error) {
        setStatus(`Session remove warning (${key}): ${error.message}`);
      }
    }
  }

  async function getReceiptFile(item) {
    if (!item || !item.dataUrl) {
      throw new Error("Receipt image not available");
    }
    return dataUrlToFile(item.dataUrl, item.fileName || "receipt.png", item.mimeType || "image/png");
  }

  async function handleIncomingFiles(files) {
    if (!Array.isArray(files) || !files.length) {
      return;
    }
    try {
      await addReceipts(files);
    } catch (error) {
      setStatus(`File load failed: ${error.message}`);
    }
  }

  async function addReceipts(files) {
    const maxCountAlertMessage = `최대 ${MAX_RECEIPT_FILES}개까지만 업로드할 수 있습니다.`;
    const imageFiles = files.filter((file) => file.type.startsWith("image/"));
    if (!imageFiles.length) {
      setStatus("No image files detected.");
      return;
    }

    const remainingSlots = Math.max(0, MAX_RECEIPT_FILES - state.receipts.length);
    if (remainingSlots <= 0) {
      window.alert(maxCountAlertMessage);
      setStatus(`Queue is full. Max ${MAX_RECEIPT_FILES} receipt(s) allowed.`);
      return;
    }

    const countAcceptedFiles = imageFiles.slice(0, remainingSlots);
    const skippedByCount = imageFiles.length - countAcceptedFiles.length;
    if (skippedByCount > 0) {
      window.alert(maxCountAlertMessage);
    }
    const oversizedFiles = countAcceptedFiles.filter(
      (file) => Number(file?.size || 0) > MAX_RECEIPT_FILE_BYTES
    );
    const eligibleFiles = countAcceptedFiles.filter(
      (file) => Number(file?.size || 0) <= MAX_RECEIPT_FILE_BYTES
    );

    if (!eligibleFiles.length) {
      const notes = [];
      if (oversizedFiles.length) {
        notes.push(
          `${oversizedFiles.length} file(s) exceed ${formatFileSizeMB(MAX_RECEIPT_FILE_BYTES)} each.`
        );
      }
      if (skippedByCount > 0) {
        notes.push(`${skippedByCount} file(s) skipped by queue limit (${MAX_RECEIPT_FILES}).`);
      }
      setStatus(notes.join(" ") || "No files could be added.");
      return;
    }

    const newItems = [];
    for (const file of eligibleFiles) {
      // eslint-disable-next-line no-await-in-loop
      const dataUrl = await fileToDataUrl(file);
      newItems.push({
        id: makeId(),
        fileName: file.name,
        mimeType: file.type,
        fileSize: file.size,
        dataUrl,
        status: "pending",
        progress: 0,
        parsed: null,
        error: "",
        createdAt: Date.now(),
        updatedAt: Date.now()
      });
    }

    state.receipts.push(...newItems);

    if (!state.currentReceiptId) {
      state.currentReceiptId = newItems[0].id;
    }

    renderQueue();
    await saveQueueToStorage();

    const statusNotes = [];
    if (oversizedFiles.length) {
      statusNotes.push(
        `${oversizedFiles.length} file(s) skipped (over ${formatFileSizeMB(MAX_RECEIPT_FILE_BYTES)}).`
      );
    }
    if (skippedByCount > 0) {
      statusNotes.push(`${skippedByCount} file(s) skipped (max ${MAX_RECEIPT_FILES} receipts).`);
    }
    const noteSuffix = statusNotes.length ? ` ${statusNotes.join(" ")}` : "";
    setStatus(`${newItems.length} receipt(s) added. Auto-parsing started...${noteSuffix}`);

    await parseReceiptsByIds(newItems.map((item) => item.id));
  }

  function getCurrentReceipt() {
    const found = state.receipts.find((item) => item.id === state.currentReceiptId);
    if (found) {
      return found;
    }
    const first = state.receipts[0] || null;
    if (first) {
      state.currentReceiptId = first.id;
    }
    return first;
  }

  function renderQueue() {
    els.queue.innerHTML = "";

    if (!state.receipts.length) {
      if (els.queueWrap) {
        els.queueWrap.hidden = true;
      }
      if (els.queueCount) {
        els.queueCount.textContent = "";
      }
      if (els.receiptSection) {
        els.receiptSection.dataset.hasPreview = "false";
      }
      if (els.root) {
        els.root.dataset.hasPreview = "false";
      }
      syncReturnHomeButton();
      return;
    }

    if (els.queueWrap) {
      els.queueWrap.hidden = false;
    }
    if (els.queueCount) {
      const total = state.receipts.length;
      const errorCount = state.receipts.filter((item) => String(item.status || "") === "error").length;
      const okCount = Math.max(0, total - errorCount);
      const totalText = `${okCount} receipt${okCount === 1 ? "" : "s"}`;
      if (errorCount > 0) {
        const errorText = `${errorCount} error${errorCount === 1 ? "" : "s"}`;
        els.queueCount.innerHTML = `<span class="saferide-helper-queue-count-total">${escapeHtml(
          totalText
        )}</span><span class="saferide-helper-queue-count-error">${escapeHtml(errorText)}</span>`;
      } else {
        els.queueCount.textContent = totalText;
      }
    }
    if (state.previewReceiptId) {
      const previewItem = state.receipts.find((item) => item.id === state.previewReceiptId) || null;
      if (!canShowPreview(previewItem)) {
        state.previewReceiptId = null;
      }
    }
    if (els.receiptSection) {
      els.receiptSection.dataset.hasPreview = state.previewReceiptId ? "true" : "false";
    }
    if (els.root) {
      els.root.dataset.hasPreview = state.previewReceiptId ? "true" : "false";
    }
    syncReturnHomeButton();

    const orderedReceipts = state.receipts
      .map((item, index) => ({ item, index }))
      .sort((a, b) => {
        const rankDiff = getQueueDisplayRank(a.item) - getQueueDisplayRank(b.item);
        if (rankDiff !== 0) {
          return rankDiff;
        }
        return a.index - b.index;
      })
      .map(({ item }) => item);

    for (const item of orderedReceipts) {
      const li = document.createElement("li");
      li.dataset.id = item.id;
      li.dataset.status = item.status || "";
      if (item.id === state.currentReceiptId) {
        li.classList.add("active");
      }
      if (item.status === "submitted") {
        li.classList.add("done");
      }
      const displayName = getReceiptDisplayName(item);
      const statusVisual = getQueueStatusVisual(item.status);
      const progress = clampProgress(item.progress);
      const statusMarkup =
        item.status === "parsing"
          ? `<div class="saferide-helper-queue-progress" aria-label="OCR progress ${progress}%">
              <div class="saferide-helper-queue-progress-track">
                <div class="saferide-helper-queue-progress-fill" style="width:${progress}%"></div>
              </div>
              <span class="saferide-helper-queue-progress-label">${progress}%</span>
            </div>`
          : "";
      const showButtonMarkup =
        item.id === state.currentReceiptId && canShowPreview(item)
          ? `<button type="button" class="saferide-helper-queue-show">${
              state.previewReceiptId === item.id ? "Hide" : "Show"
            }</button>`
          : "";
      li.innerHTML = `
        <div class="saferide-helper-queue-main">
          <div class="saferide-helper-queue-left">
            <span
              class="saferide-helper-queue-status ${statusVisual.className}"
              aria-label="${statusVisual.label}"
              title="${statusVisual.label}"
            >
              <img
                class="saferide-helper-queue-status-icon"
                src="${escapeHtml(statusVisual.iconUrl)}"
                alt=""
                aria-hidden="true"
              />
            </span>
            <span class="saferide-helper-queue-name" title="${escapeHtml(displayName)}">${escapeHtml(
              shortenFileName(displayName, 34)
            )}</span>
            ${showButtonMarkup}
          </div>
          <div class="saferide-helper-queue-right">
            ${statusMarkup}
            <button type="button" class="saferide-helper-queue-delete" aria-label="Delete ${escapeHtml(
              displayName
            )}" title="Delete">
              <span aria-hidden="true">x</span>
            </button>
          </div>
        </div>
      `;
      els.queue.appendChild(li);

      if (state.previewReceiptId === item.id && canShowPreview(item)) {
        const previewLi = document.createElement("li");
        previewLi.className = "saferide-helper-queue-preview-item";
        previewLi.innerHTML = `
          <div class="saferide-helper-queue-inline-preview">
            <img src="${item.dataUrl}" alt="Receipt preview: ${escapeHtml(
          getReceiptDisplayName(item) || "selected receipt"
        )}" />
          </div>
        `;
        els.queue.appendChild(previewLi);
      }
    }
  }

  function updateQueueProgressUI(itemId, progress) {
    if (!els.queue || !itemId) {
      return false;
    }

    const row = els.queue.querySelector(`li[data-id="${itemId}"]`);
    if (!row) {
      return false;
    }

    const fill = row.querySelector(".saferide-helper-queue-progress-fill");
    const label = row.querySelector(".saferide-helper-queue-progress-label");
    if (!fill || !label) {
      return false;
    }

    const pct = clampProgress(progress);
    fill.style.width = `${pct}%`;
    label.textContent = `${pct}%`;
    return true;
  }

  async function parseReceiptsByIds(ids) {
    if (!Array.isArray(ids) || !ids.length) {
      return;
    }

    for (const id of ids) {
      // eslint-disable-next-line no-await-in-loop
      await parseCurrentReceipt(id);
    }
  }

  async function parseCurrentReceipt(targetReceiptId = null) {
    const item = targetReceiptId
      ? state.receipts.find((entry) => entry.id === targetReceiptId)
      : getCurrentReceipt();
    if (!item) {
      setStatus(targetReceiptId ? "Cannot find selected receipt for OCR." : "Select a receipt first.");
      return;
    }

    if (item.status === "parsing") {
      setStatus("OCR already running for this receipt.");
      return;
    }

    item.status = "parsing";
    item.progress = 0;
    item.error = "";
    item.updatedAt = Date.now();
    renderQueue();
    await saveQueueToStorage();

    try {
      state.currentOcrReceiptId = item.id;
      state.currentOcrReceiptRef = item;
      state.ocrProgressLastRenderAt = 0;
      setStatus(`Running OCR for ${item.fileName}...`);
      const file = await getReceiptFile(item);
      const rawText = await extractTextFromImage(file);
      const parsed = parseReceiptText(rawText);

      item.parsed = {
        ...parsed
      };

      if (!parsed.dateForField || !parsed.timeForField || !parsed.cost) {
        item.status = "error";
        item.error = "Missing one or more values";
        item.progress = 0;
        setStatus("OCR finished with low confidence. Date/Time/Cost could not be fully extracted.");
      } else {
        item.status = "ready";
        item.error = "";
        item.progress = 100;
        const msg = `OCR parsed: ${parsed.dateForField}, ${parsed.timeForField}, $${parsed.cost}`;
        setStatus(msg);
      }

      item.updatedAt = Date.now();
      renderQueue();
      await saveQueueToStorage();
      if (item.status === "ready") {
        try {
          await maybeAutoFillSelectedReceipt(item.id);
        } catch (autoFillError) {
          setStatus(`Auto-fill failed: ${autoFillError.message}`);
        }
      }
    } catch (error) {
      item.status = "error";
      item.error = error.message;
      item.progress = 0;
      item.updatedAt = Date.now();
      renderQueue();
      await saveQueueToStorage();
      setStatus(`OCR failed: ${error.message}`);
    } finally {
      if (state.currentOcrReceiptId === item.id) {
        state.currentOcrReceiptId = null;
        state.currentOcrReceiptRef = null;
        state.ocrRecognitionActive = false;
        state.ocrProgressLastRenderAt = 0;
      }
    }
  }

  async function maybeAutoFillSelectedReceipt(receiptId) {
    const targetId = String(receiptId || "");
    if (!targetId || state.currentReceiptId !== targetId) {
      return false;
    }

    const selected = state.receipts.find((entry) => entry.id === targetId);
    if (!selected) {
      return false;
    }

    const canAutoFillStatuses = new Set(["ready", "filled", "submitted"]);
    if (!canAutoFillStatuses.has(String(selected.status || ""))) {
      return false;
    }

    const parsed = selected.parsed && typeof selected.parsed === "object" ? selected.parsed : null;
    const hasCoreParsedFields = Boolean(parsed?.dateForField && parsed?.timeForField && parsed?.cost);
    if (!hasCoreParsedFields) {
      return false;
    }

    return fillFormForReceiptId(targetId, { autoTriggered: true });
  }

  async function fillFormForReceiptId(receiptId, options = {}) {
    readProfileFromUI();
    const item = state.receipts.find((entry) => entry.id === receiptId);
    if (!item) {
      setStatus("Select a receipt first.");
      return false;
    }

    if (state.fillInFlightReceiptId) {
      if (!options.autoTriggered) {
        setStatus("Form fill is already running. Please wait.");
      }
      return false;
    }

    state.fillInFlightReceiptId = item.id;
    try {
      const parsed = item.parsed && typeof item.parsed === "object" ? item.parsed : {};
      const review = {
        dateForField: String(parsed.dateForField || "").trim(),
        timeForField: normalizeTimeString(String(parsed.timeForField || "").trim()),
        cost: normalizeCost(String(parsed.cost || "").trim())
      };

      if (!isProfileComplete(state.profile)) {
        syncProfileFieldValidationUI(state.profile);
        setSectionCollapsed(els.profileSection, false);
        setStatus(
          "Partner Info를 확인하세요. Partner Number/Store Number는 숫자만, Email은 올바른 형식이어야 합니다."
        );
        return false;
      }

      if (!review.dateForField || !review.timeForField || !review.cost) {
        setStatus("Date/Time/Cost are missing from OCR. Upload a clearer receipt image and try again.");
        return false;
      }

      const eligible = isEligibleRideTime(review.timeForField);
      if (!eligible) {
        setStatus(
          "Warning: time is not in reimbursement window (< 6:00 AM or >= 6:00 PM). You can still submit if this is intended."
        );
      }

      const failures = [];
      const previousStatus = String(item.status || "pending");

      await runFieldFillStep("Partner Name", failures, () =>
        setTextField(FIELD_MAP.partnerName, state.profile.partnerName)
      );
      await runFieldFillStep("Partner Number", failures, () =>
        setTextField(FIELD_MAP.partnerNumber, state.profile.partnerNumber)
      );
      await runFieldFillStep("Email Address", failures, () => setTextField(FIELD_MAP.email, state.profile.email));
      await runFieldFillStep("Date of Ride", failures, () => setDateField(FIELD_MAP.dateOfRide, review.dateForField));
      await runFieldFillStep("Time of Ride", failures, () => setSelectField(FIELD_MAP.timeOfRide, review.timeForField));
      await runFieldFillStep("Cost of Ride", failures, () => setTextField(FIELD_MAP.costOfRide, review.cost));
      await runFieldFillStep("Store Number", failures, () =>
        setSelectField(FIELD_MAP.storeNumber, state.profile.storeNumber)
      );
      await runFieldFillStep("Province", failures, () => setSelectField(FIELD_MAP.province, state.profile.province));
      await runFieldFillStep("Certification", failures, () => setCheckboxField(FIELD_MAP.certify, true));
      await runFieldFillStep("Signature", failures, () => setTextField(FIELD_MAP.signature, state.profile.signature));
      let receiptFile = null;
      try {
        receiptFile = await getReceiptFile(item);
      } catch (error) {
        failures.push(`Receipt Attachment: ${error.message}`);
      }
      if (receiptFile) {
        await runFieldFillStep("Receipt Attachment", failures, () => setFileField(FIELD_MAP.receipt, receiptFile));
      }

      if (!failures.length) {
        item.status = previousStatus === "submitted" ? "submitted" : "filled";
        item.parsed = {
          ...(item.parsed || {}),
          ...review
        };
        item.error = "";
        item.updatedAt = Date.now();
        renderQueue();
        await saveQueueToStorage();
        setStatus(
          options.autoTriggered
            ? "Form auto-filled for selected receipt. Please verify and click Smartsheet Submit manually."
            : "Form filled. Please verify and click Smartsheet Submit manually."
        );
        return true;
      }

      const keepCompletedStatus =
        previousStatus === "submitted" || previousStatus === "filled" || previousStatus === "ready";
      item.status = keepCompletedStatus ? previousStatus : "error";
      item.error = failures.join(" | ");
      item.updatedAt = Date.now();
      renderQueue();
      await saveQueueToStorage();
      setStatus(`Form fill finished with issues: ${failures.join(" | ")}`);
      return false;
    } finally {
      if (state.fillInFlightReceiptId === item.id) {
        state.fillInFlightReceiptId = null;
      }
    }
  }

  function moveToNextPending() {
    const next = state.receipts.find((item) => item.status !== "submitted");
    if (next) {
      state.currentReceiptId = next.id;
      state.previewReceiptId = null;
      renderQueue();
      setStatus(`Moved to next receipt: ${getReceiptDisplayName(next)}`);
      void saveQueueToStorage();
      return;
    }

    void saveQueueToStorage();
  }

  function attachSubmitHints() {
    document.addEventListener(
      "click",
      (event) => {
        const button = event.target.closest("button, input[type='submit']");
        if (!button) {
          return;
        }
        const text = `${button.textContent || ""} ${button.value || ""}`.toLowerCase();
        if (!text.includes("submit")) {
          return;
        }

        const current = getCurrentReceipt();
        if (!current) {
          return;
        }

        state.submitClickedAt = Date.now();
        setStatus("Submit clicked. Waiting for confirmation... if success is detected, this receipt will be marked submitted.");
      },
      true
    );
  }

  function observeConfirmationMessage() {
    const observer = new MutationObserver(() => {
      if (!state.submitClickedAt || Date.now() - state.submitClickedAt > 60000) {
        return;
      }

      const bodyText = document.body?.innerText || "";
      if (!/thank you for your submission/i.test(bodyText)) {
        return;
      }

      const current = getCurrentReceipt();
      if (!current || current.status === "submitted") {
        return;
      }

      current.status = "submitted";
      current.updatedAt = Date.now();
      state.showReturnHomeBtn = true;
      renderQueue();
      void saveQueueToStorage();
      setStatus(`Submission confirmed for: ${getReceiptDisplayName(current)}`);
      state.submitClickedAt = 0;
      moveToNextPending();
    });

    observer.observe(document.documentElement, {
      subtree: true,
      childList: true,
      characterData: true
    });
  }

  async function extractTextFromImage(file) {
    const worker = await getOcrWorker();
    state.ocrRecognitionActive = true;
    try {
      const result = await worker.recognize(file);
      const text = result?.data?.text || "";
      if (!text.trim()) {
        throw new Error("OCR returned empty text");
      }
      return text;
    } finally {
      state.ocrRecognitionActive = false;
    }
  }

  async function getOcrWorker() {
    if (state.ocrWorker) {
      return state.ocrWorker;
    }

    if (state.ocrLoading) {
      while (state.ocrLoading) {
        // eslint-disable-next-line no-await-in-loop
        await wait(150);
      }
      if (state.ocrWorker) {
        return state.ocrWorker;
      }
    }

    state.ocrLoading = true;
    try {
      if (!window.Tesseract || !window.Tesseract.createWorker) {
        throw new Error("Tesseract library missing. Reload extension and page.");
      }

      const worker = await window.Tesseract.createWorker("eng", 1, {
        workerPath: chrome.runtime.getURL("vendor/tesseract/worker.min.js"),
        corePath: chrome.runtime.getURL("vendor/tesseract-core/tesseract-core.wasm.js"),
        langPath: "https://tessdata.projectnaptha.com/4.0.0_fast",
        logger: (message) => {
          if (!message || typeof message.progress !== "number") {
            return;
          }
          if (!state.ocrRecognitionActive) {
            return;
          }
          const statusText = String(message.status || "").toLowerCase();
          if (statusText && !statusText.includes("recognizing")) {
            return;
          }
          const currentId = state.currentOcrReceiptId;
          if (!currentId) {
            return;
          }
          const item =
            state.currentOcrReceiptRef && state.currentOcrReceiptRef.id === currentId
              ? state.currentOcrReceiptRef
              : state.receipts.find((receipt) => receipt.id === currentId);
          if (!item || item.status !== "parsing") {
            return;
          }

          const pct = clampProgress(Math.round(message.progress * 100));
          const now = Date.now();
          const isFinal = pct >= 100;
          const enoughTimePassed = now - state.ocrProgressLastRenderAt >= 1500;

          if (!isFinal && !enoughTimePassed) {
            return;
          }

          if (item.progress !== pct) {
            item.progress = pct;
            updateQueueProgressUI(item.id, pct);
          }
          state.ocrProgressLastRenderAt = now;
        }
      });

      state.ocrWorker = worker;
      return worker;
    } catch (error) {
      throw new Error(`OCR init failed: ${error.message}`);
    } finally {
      state.ocrLoading = false;
    }
  }

  function parseReceiptText(rawText) {
    const normalizedText = normalizeOcrText(rawText);

    const combined = extractCombinedDateTime(normalizedText);
    const dateOnly = combined?.date || extractDate(normalizedText);
    const timeOnly = combined?.time || extractTime(normalizedText);
    const cost = extractCost(normalizedText);

    const parsed = {
      dateForField: dateOnly ? formatDateForField(dateOnly) : "",
      timeForField: timeOnly ? formatTimeForField(timeOnly) : "",
      cost: cost ? cost.toFixed(2) : "",
      confidence: 0
    };

    let confidence = 0;
    if (combined) {
      confidence += 0.45;
    }
    if (parsed.dateForField) {
      confidence += 0.2;
    }
    if (parsed.timeForField) {
      confidence += 0.2;
    }
    if (parsed.cost) {
      confidence += 0.15;
    }
    parsed.confidence = Math.min(1, confidence);

    return parsed;
  }

  function normalizeOcrText(text) {
    const normalized = String(text || "")
      .replace(/\u00A0/g, " ")
      .replace(/[•·]/g, " ")
      .replace(/a\.?m\.?/gi, " AM")
      .replace(/p\.?m\.?/gi, " PM")
      .replace(/\r/g, "");

    return normalized
      .split("\n")
      .map((line) => line.replace(/\s+/g, " ").trim())
      .filter(Boolean)
      .join("\n");
  }

  function extractCombinedDateTime(text) {
    let match = text.match(/\b(20\d{2})[-\/.](\d{1,2})[-\/.](\d{1,2})[^\d]{0,10}(\d{1,2}):(\d{2})\s*(AM|PM)\b/i);
    if (match) {
      const date = toDateParts(match[1], match[2], match[3]);
      const time = toTimeParts(match[4], match[5], match[6]);
      if (date && time) {
        return { date, time };
      }
    }

    match = text.match(
      /\b(Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|Jul(?:y)?|Aug(?:ust)?|Sep(?:t|tember)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)\s+(\d{1,2})(?:,\s*(20\d{2}))?[^\d]{0,10}(\d{1,2}):(\d{2})\s*(AM|PM)\b/i
    );
    if (!match) {
      return null;
    }

    const month = MONTHS[match[1].toLowerCase()];
    const day = Number(match[2]);
    const year = match[3] ? Number(match[3]) : new Date().getFullYear();

    const date = toDateParts(year, month, day);
    const time = toTimeParts(match[4], match[5], match[6]);

    if (!date || !time) {
      return null;
    }

    return { date, time };
  }

  function extractDate(text) {
    let match = text.match(/\b(20\d{2})[-\/.](\d{1,2})[-\/.](\d{1,2})\b/);
    if (match) {
      return toDateParts(match[1], match[2], match[3]);
    }

    match = text.match(
      /\b(Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|Jul(?:y)?|Aug(?:ust)?|Sep(?:t|tember)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)\s+(\d{1,2})(?:,\s*(20\d{2}))?/i
    );
    if (!match) {
      return null;
    }

    const month = MONTHS[match[1].toLowerCase()];
    const day = Number(match[2]);
    const year = match[3] ? Number(match[3]) : new Date().getFullYear();

    return toDateParts(year, month, day);
  }

  function extractTime(text) {
    const lines = text.split(/\n+/).map((line) => line.trim()).filter(Boolean);
    const fallbackLines = lines.length ? lines : [text];

    const candidates = [];
    for (let i = 0; i < fallbackLines.length; i += 1) {
      const line = fallbackLines[i];
      const regex = /(\d{1,2}):(\d{2})\s*(AM|PM)\b/gi;
      let match;
      while ((match = regex.exec(line))) {
        const score = scoreTimeLine(line, i);
        const time = toTimeParts(match[1], match[2], match[3]);
        if (time) {
          candidates.push({ time, score });
        }
      }
    }

    if (!candidates.length) {
      return null;
    }

    candidates.sort((a, b) => b.score - a.score);
    return candidates[0].time;
  }

  function scoreTimeLine(line, index) {
    const lower = line.toLowerCase();
    let score = 0.4;

    if (/ride|trip|uber|lyft/.test(lower)) {
      score += 0.35;
    }
    if (/pickup|pick-up/.test(lower)) {
      score += 0.2;
    }
    if (/drop|drop-off/.test(lower)) {
      score -= 0.08;
    }
    if (/payment|visa|hst|promo|tip|total/.test(lower)) {
      score -= 0.2;
    }
    if (index < 6) {
      score += 0.1;
    }

    return score;
  }

  function extractCost(text) {
    let match = text.match(/total\s*charge[^\d$-]{0,20}\$?\s*([0-9]+(?:\.[0-9]{2}))/i);
    if (match) {
      return Number(match[1]);
    }

    match = text.match(/(?:total|charged|amount\s*paid)[^\d$-]{0,20}\$?\s*([0-9]+(?:\.[0-9]{2}))/i);
    if (match) {
      return Number(match[1]);
    }

    const values = [];
    const regex = /\$\s*([0-9]+(?:\.[0-9]{2}))/g;
    while ((match = regex.exec(text))) {
      const value = Number(match[1]);
      if (Number.isFinite(value) && value > 0) {
        values.push(value);
      }
    }

    if (!values.length) {
      return null;
    }

    return values[values.length - 1];
  }

  function toDateParts(year, month, day) {
    const y = Number(year);
    const m = Number(month);
    const d = Number(day);

    if (!Number.isInteger(y) || !Number.isInteger(m) || !Number.isInteger(d)) {
      return null;
    }
    if (y < 2000 || y > 2100 || m < 1 || m > 12 || d < 1 || d > 31) {
      return null;
    }

    return { year: y, month: m, day: d };
  }

  function toTimeParts(hour, minute, meridiem) {
    const h = Number(hour);
    const m = Number(minute);
    const mer = String(meridiem || "").toUpperCase();

    if (!Number.isInteger(h) || !Number.isInteger(m)) {
      return null;
    }
    if (h < 1 || h > 12 || m < 0 || m > 59) {
      return null;
    }
    if (mer !== "AM" && mer !== "PM") {
      return null;
    }

    return {
      hour: h,
      minute: m,
      meridiem: mer
    };
  }

  function formatDateForField(parts) {
    return `${String(parts.month).padStart(2, "0")}/${String(parts.day).padStart(2, "0")}/${parts.year}`;
  }

  function formatTimeForField(parts) {
    return `${parts.hour}:${String(parts.minute).padStart(2, "0")} ${parts.meridiem}`;
  }

  function normalizeTimeString(value) {
    const text = String(value || "")
      .trim()
      .replace(/a\.?m\.?/i, "AM")
      .replace(/p\.?m\.?/i, "PM");
    const match = text.match(/^(\d{1,2}):(\d{2})\s*(AM|PM)$/i);
    if (!match) {
      return "";
    }

    const parts = toTimeParts(match[1], match[2], match[3]);
    if (!parts) {
      return "";
    }
    return formatTimeForField(parts);
  }

  function normalizeCost(value) {
    const cleaned = String(value || "").replace(/[^0-9.]/g, "");
    const n = Number(cleaned);
    if (!Number.isFinite(n)) {
      return "";
    }
    return n.toFixed(2);
  }

  function isEligibleRideTime(timeStr) {
    const match = timeStr.match(/^(\d{1,2}):(\d{2})\s*(AM|PM)$/i);
    if (!match) {
      return false;
    }
    let hour = Number(match[1]);
    const minute = Number(match[2]);
    const mer = match[3].toUpperCase();

    if (mer === "PM" && hour < 12) {
      hour += 12;
    }
    if (mer === "AM" && hour === 12) {
      hour = 0;
    }

    const totalMinutes = hour * 60 + minute;
    return totalMinutes < 360 || totalMinutes >= 1080;
  }

  async function setTextField(field, value) {
    const input = findFieldInput(field);
    if (!input) {
      throw new Error(`Cannot find field: ${field.label}`);
    }
    setNativeValue(input, value);
    await wait(60);
  }

  async function setDateField(field, dateValue) {
    const input = findFieldInput(field);
    if (!input) {
      throw new Error(`Cannot find date field: ${field.label}`);
    }

    if (input.type === "date") {
      const parts = dateValue.split("/");
      if (parts.length === 3) {
        setNativeValue(input, `${parts[2]}-${parts[0].padStart(2, "0")}-${parts[1].padStart(2, "0")}`);
      } else {
        setNativeValue(input, dateValue);
      }
    } else {
      setNativeValue(input, dateValue);
    }

    await wait(60);
  }

  async function setCheckboxField(field, checked) {
    let input = document.querySelector(`input[type='checkbox'][name='${field.key}']`);

    if (!input) {
      const container = findFieldContainer(field.label);
      input = container ? container.querySelector("input[type='checkbox']") : null;
    }

    if (!input) {
      throw new Error(`Cannot find checkbox field: ${field.label}`);
    }

    if (input.checked !== checked) {
      input.click();
      await wait(50);
    }
  }

  async function setSelectField(field, value) {
    const normalizedValue = normalizeSelectValue(field, value);
    if (!normalizedValue) {
      throw new Error(`Empty value for ${field.label}`);
    }

    const lodestarInput = findLodestarComboboxInput(field);
    if (lodestarInput) {
      await setLodestarComboboxValue(field, lodestarInput, normalizedValue);
      return;
    }

    const nativeSelect = document.querySelector(`select[name='${field.key}']`);
    if (nativeSelect) {
      setNativeValue(nativeSelect, normalizedValue);
      await wait(60);
      if (normalizeText(nativeSelect.value) === normalizeText(normalizedValue)) {
        return;
      }
    }

    const hiddenOrTextInput = document.querySelector(`input[name='${field.key}'], textarea[name='${field.key}']`);
    if (hiddenOrTextInput) {
      setNativeValue(hiddenOrTextInput, normalizedValue);
      await wait(80);
      if (isSelectFieldSet(field, normalizedValue)) {
        return;
      }
    }

    const container = findFieldContainer(field.label);
    const combobox =
      (container &&
        container.querySelector(
          "[role='combobox'], input[role='combobox'], button[aria-haspopup='listbox'], div[aria-haspopup='listbox']"
        )) ||
      findFieldComboboxByHeuristic(field);

    if (!combobox) {
      throw new Error(`Cannot find select field: ${field.label}`);
    }

    combobox.scrollIntoView({ behavior: "smooth", block: "center" });
    clickElement(combobox);
    combobox.dispatchEvent(new KeyboardEvent("keydown", { key: "ArrowDown", bubbles: true }));
    await wait(140);

    const activeInput = findOpenSelectInput(container);
    if (activeInput) {
      setNativeValueNoBlur(activeInput, normalizedValue);
      activeInput.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", bubbles: true }));
      activeInput.dispatchEvent(new KeyboardEvent("keyup", { key: "Enter", bubbles: true }));
      await wait(160);
      if (isSelectFieldSet(field, normalizedValue)) {
        return;
      }
    }

    const option = findVisibleOptionByText(normalizedValue);
    if (option) {
      clickElement(option);
      await wait(180);
      if (isSelectFieldSet(field, normalizedValue)) {
        return;
      }
    }

    const looseOption = findVisibleClickableByText(normalizedValue);
    if (looseOption) {
      clickElement(looseOption);
      await wait(180);
      if (isSelectFieldSet(field, normalizedValue)) {
        return;
      }
    }

    if (activeInput) {
      activeInput.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", bubbles: true }));
      activeInput.dispatchEvent(new KeyboardEvent("keyup", { key: "Enter", bubbles: true }));
      await wait(120);
      if (isSelectFieldSet(field, normalizedValue)) {
        return;
      }
    }

    throw new Error(`Option not found for ${field.label}: ${normalizedValue}`);
  }

  async function setLodestarComboboxValue(field, input, value) {
    const wrapper = input.closest("[role='combobox']");
    const listboxId = input.getAttribute("aria-controls") || "";

    if (wrapper) {
      wrapper.scrollIntoView({ behavior: "smooth", block: "center" });
    }

    const toggle =
      (wrapper && wrapper.querySelector("button[aria-label='Toggle menu'], button[id$='--toggle-button']")) ||
      null;

    clickElement(toggle || wrapper || input);
    await wait(120);

    // Typing narrows options in Smartsheet comboboxes.
    setNativeValueNoBlur(input, value);
    await wait(80);

    const optionIndex = getFieldOptionIndex(field.key, value);
    const indexedOption =
      optionIndex >= 0 ? findVisibleOptionByIdSuffixInListbox(listboxId, `-item-${optionIndex}`) : null;

    const listboxOption = indexedOption || findVisibleOptionInListbox(listboxId, value);
    if (listboxOption) {
      clickElement(listboxOption);
      await wait(180);
    } else {
      input.dispatchEvent(new KeyboardEvent("keydown", { key: "ArrowDown", bubbles: true }));
      await wait(80);
      input.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", bubbles: true }));
      input.dispatchEvent(new KeyboardEvent("keyup", { key: "Enter", bubbles: true }));
      await wait(180);
    }

    const selected = await waitForLodestarSelection(field, value, 1200);
    if (!selected) {
      throw new Error(`Value not selected for ${field.label}: ${value}`);
    }

    input.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape", bubbles: true }));
    await wait(60);
  }

  async function clearExistingFileAttachments(field) {
    const container = findFieldContainer(field.label);
    if (!container) {
      return;
    }

    const removeSelector =
      "button.sds-file-list-item-remove-button, button[aria-label^='Remove this file'], button[aria-label*='Remove this file']";

    for (let attempt = 0; attempt < 10; attempt += 1) {
      const removeButtons = Array.from(container.querySelectorAll(removeSelector)).filter(
        (button) => button.getAttribute("aria-disabled") !== "true"
      );

      if (!removeButtons.length) {
        return;
      }

      for (const button of removeButtons) {
        clickElement(button);
        // Wait for Smartsheet to remove file chips from React state.
        // eslint-disable-next-line no-await-in-loop
        await wait(140);
      }

      // eslint-disable-next-line no-await-in-loop
      await wait(180);
    }

    const remaining = Array.from(container.querySelectorAll(removeSelector)).filter(
      (button) => button.getAttribute("aria-disabled") !== "true"
    );
    if (remaining.length) {
      throw new Error("Could not remove existing attached files");
    }
  }

  async function setFileField(field, file) {
    let input = document.querySelector(`input[type='file'][name='${field.key}']`);
    if (!input) {
      const container = findFieldContainer(field.label);
      input = container ? container.querySelector("input[type='file']") : null;
    }

    if (!input) {
      throw new Error(`Cannot find file upload field: ${field.label}`);
    }

    await clearExistingFileAttachments(field);

    const transfer = new DataTransfer();
    transfer.items.add(file);
    input.files = transfer.files;
    input.dispatchEvent(new Event("input", { bubbles: true }));
    input.dispatchEvent(new Event("change", { bubbles: true }));
    await wait(120);
  }

  function findFieldInput(field) {
    let input = document.querySelector(`input[name='${field.key}'], textarea[name='${field.key}']`);

    if (input) {
      return input;
    }

    const byDataKey = document.querySelector(`[data-key='${field.key}'] input, [data-key='${field.key}'] textarea`);
    if (byDataKey) {
      return byDataKey;
    }

    const container = findFieldContainer(field.label);
    if (!container) {
      return null;
    }

    return container.querySelector("input, textarea") || null;
  }

  function findFieldContainer(labelText) {
    const normalizedTarget = normalizeText(labelText);
    const labels = Array.from(document.querySelectorAll("label, h3, h4, span, div, p"));

    for (const label of labels) {
      if (els.root && els.root.contains(label)) {
        continue;
      }

      const text = normalizeText(label.textContent || "");
      if (!text) {
        continue;
      }
      if (!text.includes(normalizedTarget)) {
        continue;
      }

      const root =
        label.closest("[data-key]") ||
        label.closest(".ss-form-item") ||
        label.closest(".form-field") ||
        label.closest("fieldset") ||
        label.parentElement;

      if (root && els.root && els.root.contains(root)) {
        continue;
      }

      if (root && root.querySelector("input, textarea, select, [role='combobox'], input[type='file']")) {
        return root;
      }
    }

    return null;
  }

  function findFieldComboboxByHeuristic(field) {
    const target = normalizeText(field.label);
    const candidates = Array.from(
      document.querySelectorAll(
        "[role='combobox'], input[role='combobox'], button[aria-haspopup='listbox'], div[aria-haspopup='listbox']"
      )
    );
    for (const candidate of candidates) {
      if (els.root && els.root.contains(candidate)) {
        continue;
      }
      const ctx = normalizeText(candidate.closest("div, section, fieldset")?.textContent || "");
      if (ctx.includes(target)) {
        return candidate;
      }
    }
    return null;
  }

  function findLodestarComboboxInput(field) {
    const inputs = Array.from(
      document.querySelectorAll("input.sds-select-combobox-input[aria-controls], input[aria-autocomplete='list'][aria-controls]")
    );
    const target = normalizeText(field.label);

    for (const input of inputs) {
      if (els.root && els.root.contains(input)) {
        continue;
      }

      const inputLabel = normalizeText(input.getAttribute("aria-label") || "");
      if (inputLabel && (inputLabel === target || inputLabel.includes(target) || target.includes(inputLabel))) {
        return input;
      }

      const labelledBy = (input.getAttribute("aria-labelledby") || "").split(/\s+/).filter(Boolean);
      const labelText = labelledBy
        .map((id) => document.getElementById(id)?.textContent || "")
        .join(" ");
      const normalizedLabelText = normalizeText(labelText);
      if (
        normalizedLabelText &&
        (normalizedLabelText.includes(target) || target.includes(normalizedLabelText))
      ) {
        return input;
      }

      const wrapper = input.closest("[role='combobox']");
      const wrapperText = normalizeText(wrapper?.textContent || "");
      if (wrapperText && wrapperText.includes(target)) {
        return input;
      }
    }

    return null;
  }

  function findVisibleOptionInListbox(listboxId, targetText) {
    const target = normalizeText(targetText);
    const listbox = listboxId
      ? document.getElementById(listboxId)
      : document.querySelector("[role='listbox'][id]");

    if (!listbox || !isElementVisible(listbox)) {
      return null;
    }

    const options = Array.from(listbox.querySelectorAll("[role='option']"));
    for (const option of options) {
      if (!isElementVisible(option)) {
        continue;
      }
      const text = normalizeText(option.textContent || "");
      if (text === target) {
        return option;
      }
    }

    for (const option of options) {
      if (!isElementVisible(option)) {
        continue;
      }
      const text = normalizeText(option.textContent || "");
      if (text.includes(target) || target.includes(text)) {
        return option;
      }
    }

    return null;
  }

  function findVisibleOptionByIdSuffixInListbox(listboxId, idSuffix) {
    const listbox = listboxId
      ? document.getElementById(listboxId)
      : document.querySelector("[role='listbox'][id]");

    if (!listbox || !isElementVisible(listbox)) {
      return null;
    }

    const options = Array.from(listbox.querySelectorAll("[role='option'][id]"));
    for (const option of options) {
      if (!isElementVisible(option)) {
        continue;
      }
      if (option.id.endsWith(idSuffix)) {
        return option;
      }
    }

    return null;
  }

  async function waitForLodestarSelection(field, expectedValue, timeoutMs = 1000) {
    const startedAt = Date.now();
    while (Date.now() - startedAt < timeoutMs) {
      const freshInput = findLodestarComboboxInput(field);
      if (freshInput) {
        const freshWrapper = freshInput.closest("[role='combobox']");
        if (isComboboxValueSelected(freshInput, freshWrapper, expectedValue)) {
          return true;
        }
      }
      // eslint-disable-next-line no-await-in-loop
      await wait(80);
    }
    return false;
  }

  function isComboboxValueSelected(input, wrapper, expectedValue) {
    const expected = normalizeOptionValue(expectedValue);
    if (!expected) {
      return false;
    }

    const inputValue = normalizeOptionValue(input?.value || "");
    if (inputValue && inputValue === expected) {
      return true;
    }

    const displayedValue = normalizeOptionValue(getComboboxDisplayedValue(wrapper));
    if (displayedValue && displayedValue === expected) {
      return true;
    }

    return false;
  }

  function getComboboxDisplayedValue(wrapper) {
    if (!wrapper) {
      return "";
    }

    const explicitValue =
      wrapper.querySelector(".css-4wnbzs") ||
      wrapper.querySelector("[data-client-id='selected-value']") ||
      wrapper.querySelector("[class*='select-value']");
    if (explicitValue && explicitValue.textContent) {
      return explicitValue.textContent.trim();
    }

    const input = wrapper.querySelector("input");
    if (input && input.value) {
      return input.value.trim();
    }

    return "";
  }

  function findVisibleOptionByText(targetText) {
    const target = normalizeText(targetText);
    const options = Array.from(document.querySelectorAll("[role='option'], li[role='option'], div[role='option']"));

    for (const option of options) {
      if (!isElementVisible(option)) {
        continue;
      }
      if (els.root && els.root.contains(option)) {
        continue;
      }
      const text = normalizeText(option.textContent || "");
      if (text === target) {
        return option;
      }
    }

    for (const option of options) {
      if (!isElementVisible(option)) {
        continue;
      }
      if (els.root && els.root.contains(option)) {
        continue;
      }
      const text = normalizeText(option.textContent || "");
      if (text.includes(target) || target.includes(text)) {
        return option;
      }
    }

    return null;
  }

  function findVisibleClickableByText(targetText) {
    const target = normalizeText(targetText);
    const selector = "li, div, span, button";
    const candidates = Array.from(document.querySelectorAll(selector));

    for (const el of candidates) {
      if (!isElementVisible(el)) {
        continue;
      }
      if (els.root && els.root.contains(el)) {
        continue;
      }
      const text = normalizeText(el.textContent || "");
      if (!text || text !== target) {
        continue;
      }
      if (
        el.matches("[role='option'], [data-value], button") ||
        /option|menu|list|item|select|dropdown/i.test(el.className || "")
      ) {
        return el;
      }
    }

    return null;
  }

  function findOpenSelectInput(container) {
    const active = document.activeElement;
    if (
      active &&
      active instanceof HTMLInputElement &&
      active.type !== "file" &&
      !(els.root && els.root.contains(active))
    ) {
      return active;
    }

    if (container) {
      const scoped = container.querySelector(
        "input[role='combobox'], input[aria-autocomplete='list'], input[type='text']"
      );
      if (scoped && !(els.root && els.root.contains(scoped))) {
        return scoped;
      }
    }

    const globalInput = document.querySelector(
      "input[role='combobox'], input[aria-autocomplete='list'], input[aria-controls][type='text']"
    );
    if (globalInput && !(els.root && els.root.contains(globalInput))) {
      return globalInput;
    }

    return null;
  }

  function isSelectFieldSet(field, expectedValue) {
    const expected = normalizeOptionValue(expectedValue);
    if (!expected) {
      return false;
    }

    const named = document.querySelector(`select[name='${field.key}'], input[name='${field.key}']`);
    if (named && normalizeOptionValue(named.value) === expected) {
      return true;
    }

    const lodestarInput = findLodestarComboboxInput(field);
    if (lodestarInput) {
      const wrapper = lodestarInput.closest("[role='combobox']");
      if (isComboboxValueSelected(lodestarInput, wrapper, expectedValue)) {
        return true;
      }
    }

    const container = findFieldContainer(field.label);
    if (!container) {
      return false;
    }

    const text = normalizeOptionValue(container.textContent || "");
    const targetLabel = normalizeOptionValue(field.label);
    const compactText = text.replace(targetLabel, " ");
    return compactText.includes(expected);
  }

  function normalizeSelectValue(field, rawValue) {
    const input = String(rawValue || "").trim();
    if (!input) {
      return "";
    }

    let value = input;
    if (field.key === FIELD_MAP.province.key) {
      const aliasKey = normalizeOptionValue(input);
      if (PROVINCE_ALIASES[aliasKey]) {
        value = PROVINCE_ALIASES[aliasKey];
      }
    }

    const options = getFieldOptions(field.key);
    if (!options.length) {
      return value;
    }

    const direct = options.find((opt) => opt === value);
    if (direct) {
      return direct;
    }

    const normalized = normalizeOptionValue(value);
    const normalizedMatch = options.find((opt) => normalizeOptionValue(opt) === normalized);
    if (normalizedMatch) {
      return normalizedMatch;
    }

    if (field.key === FIELD_MAP.storeNumber.key) {
      const digitInput = value.replace(/\D/g, "");
      if (digitInput) {
        const digitMatch = options.find((opt) => opt.replace(/\D/g, "") === digitInput);
        if (digitMatch) {
          return digitMatch;
        }
      }
    }

    if (field.key === FIELD_MAP.province.key) {
      const partialMatch = options.find((opt) => {
        const n = normalizeOptionValue(opt);
        return n.includes(normalized) || normalized.includes(n);
      });
      if (partialMatch) {
        return partialMatch;
      }
    }

    return value;
  }

  function getFieldOptions(key) {
    const definition = getFormDefinition();
    const components = definition?.components;
    if (!Array.isArray(components)) {
      return [];
    }

    const component = components.find((entry) => entry && entry.key === key);
    if (!component || !Array.isArray(component.options)) {
      return [];
    }

    return component.options
      .map((opt) => String(opt?.value || "").trim())
      .filter(Boolean);
  }

  function getFormDefinition() {
    if (state.formDefinition) {
      return state.formDefinition;
    }

    const scripts = Array.from(document.querySelectorAll("script"));
    for (const script of scripts) {
      const text = script.textContent || "";
      if (!text.includes("window.formDefinition")) {
        continue;
      }
      const match = text.match(/window\.formDefinition\s*=\s*"([^"]+)";/);
      if (!match || !match[1]) {
        continue;
      }

      try {
        const decoded = atob(match[1]);
        state.formDefinition = JSON.parse(decoded);
        return state.formDefinition;
      } catch (error) {
        setStatus(`Form definition parse warning: ${error.message}`);
        return null;
      }
    }

    return null;
  }

  function getFieldOptionIndex(key, value) {
    const options = getFieldOptions(key);
    if (!options.length) {
      return -1;
    }

    const normalized = normalizeOptionValue(value);
    for (let i = 0; i < options.length; i += 1) {
      if (normalizeOptionValue(options[i]) === normalized) {
        return i;
      }
    }
    return -1;
  }

  function normalizeOptionValue(value) {
    return String(value || "")
      .toLowerCase()
      .replace(/&/g, " and ")
      .replace(/[^a-z0-9]+/g, "")
      .trim();
  }

  function setNativeValue(element, value) {
    const descriptor = Object.getOwnPropertyDescriptor(element.constructor.prototype, "value");
    if (descriptor && descriptor.set) {
      descriptor.set.call(element, value);
    } else {
      element.value = value;
    }
    element.dispatchEvent(new Event("input", { bubbles: true }));
    element.dispatchEvent(new Event("change", { bubbles: true }));
    element.dispatchEvent(new Event("blur", { bubbles: true }));
  }

  function setNativeValueNoBlur(element, value) {
    const descriptor = Object.getOwnPropertyDescriptor(element.constructor.prototype, "value");
    if (descriptor && descriptor.set) {
      descriptor.set.call(element, value);
    } else {
      element.value = value;
    }
    element.dispatchEvent(new Event("input", { bubbles: true }));
    element.dispatchEvent(new Event("change", { bubbles: true }));
  }

  function clickElement(element) {
    if (!element) {
      return;
    }
    element.dispatchEvent(new MouseEvent("mousedown", { bubbles: true }));
    element.dispatchEvent(new MouseEvent("mouseup", { bubbles: true }));
    element.click();
  }

  async function runFieldFillStep(name, failures, action) {
    try {
      await action();
    } catch (error) {
      failures.push(`${name}: ${error.message}`);
    }
  }

  function normalizeText(value) {
    return String(value || "")
      .toLowerCase()
      .replace(/\s+/g, " ")
      .trim();
  }

  function syncReturnHomeButton() {
    if (!els.returnHomeBtn) {
      return;
    }
    const visible = Boolean(state.showReturnHomeBtn);
    els.returnHomeBtn.hidden = !visible;
    els.returnHomeBtn.style.display = visible ? "inline-flex" : "none";
  }

  function normalizeDigits(value) {
    return String(value || "").replace(/\D+/g, "");
  }

  function isDigitsOnly(value) {
    return /^\d+$/.test(String(value || "").trim());
  }

  function isValidEmail(value) {
    const email = String(value || "").trim();
    if (!email) {
      return false;
    }
    return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
  }

  function canShowPreview(item) {
    if (!item || !item.dataUrl) {
      return false;
    }
    const status = String(item.status || "");
    return status === "ready" || status === "filled" || status === "submitted";
  }

  function getQueueStatusVisual(status) {
    if (status === "pending") {
      return {
        className: "is-pending",
        iconUrl: STATUS_ICON_URLS.progress,
        label: "Pending"
      };
    }
    if (status === "parsing") {
      return {
        className: "is-progress",
        iconUrl: STATUS_ICON_URLS.progress,
        label: "Parsing"
      };
    }
    if (status === "error") {
      return {
        className: "is-error",
        iconUrl: STATUS_ICON_URLS.error,
        label: "Error"
      };
    }
    if (status === "submitted") {
      return {
        className: "is-complete",
        iconUrl: STATUS_ICON_URLS.completeDone,
        label: "Complete"
      };
    }
    if (status === "ready" || status === "filled") {
      return {
        className: "is-complete",
        iconUrl: STATUS_ICON_URLS.complete,
        label: "Complete"
      };
    }
    return {
      className: "is-pending",
      iconUrl: STATUS_ICON_URLS.progress,
      label: "Pending"
    };
  }

  function getQueueDisplayRank(item) {
    const status = String(item?.status || "");
    if (status === "submitted") {
      return 2;
    }
    if (status === "error") {
      return 1;
    }
    return 0;
  }

  async function deleteReceiptById(itemId) {
    const index = state.receipts.findIndex((item) => item.id === itemId);
    if (index < 0) {
      return;
    }

    const target = state.receipts[index];
    if (!target) {
      return;
    }

    if (target.status === "parsing" || state.currentOcrReceiptId === itemId) {
      window.alert("이 항목은 현재 OCR 처리 중이라 삭제할 수 없습니다. 잠시 후 다시 시도해주세요.");
      setStatus("Cannot delete a receipt while OCR is running.");
      return;
    }

    const deletedName = getReceiptDisplayName(target);
    state.receipts.splice(index, 1);

    if (state.currentReceiptId === itemId) {
      const sameIndexItem = state.receipts[index] || null;
      const previousItem = state.receipts[index - 1] || null;
      const fallbackItem = sameIndexItem || previousItem || state.receipts[0] || null;
      state.currentReceiptId = fallbackItem ? fallbackItem.id : null;
    }
    if (state.previewReceiptId === itemId) {
      state.previewReceiptId = null;
    }

    renderQueue();
    await saveQueueToStorage();
    setStatus(`Deleted: ${deletedName}`);
  }

  function clampProgress(value) {
    const n = Number(value);
    if (!Number.isFinite(n)) {
      return 0;
    }
    return Math.max(0, Math.min(100, Math.round(n)));
  }

  function getReceiptDisplayName(item) {
    if (!item) {
      return "";
    }

    const mdy = toMdyFromFieldDate(item.parsed?.dateForField || "");
    const cost = normalizeCost(item.parsed?.cost || "");
    if (mdy && cost) {
      return `${mdy} - $${cost}`;
    }
    return item.fileName || "Receipt";
  }

  function toMdyFromFieldDate(value) {
    const match = String(value || "").match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (!match) {
      return "";
    }
    const month = String(Number(match[1])).padStart(2, "0");
    const day = String(Number(match[2])).padStart(2, "0");
    const year = match[3];
    return `${month}/${day}/${year}`;
  }

  function escapeHtml(value) {
    return String(value || "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#39;");
  }

  function shortenFileName(name, max) {
    if (name.length <= max) {
      return name;
    }
    return `${name.slice(0, max - 3)}...`;
  }

  function isElementVisible(element) {
    const style = window.getComputedStyle(element);
    if (style.display === "none" || style.visibility === "hidden" || style.opacity === "0") {
      return false;
    }
    const rect = element.getBoundingClientRect();
    return rect.width > 0 && rect.height > 0;
  }

  function setStatus(message) {
    if (!message) {
      return;
    }
    console.debug(`[SafeRide Helper] ${message}`);
  }

  function fileToDataUrl(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(String(reader.result || ""));
      reader.onerror = () => reject(new Error("Failed to read image file"));
      reader.readAsDataURL(file);
    });
  }

  function dataUrlToFile(dataUrl, fileName, mimeType) {
    const parts = String(dataUrl).split(",");
    if (parts.length < 2) {
      throw new Error("Invalid image data");
    }

    const meta = parts[0];
    const base64 = parts.slice(1).join(",");
    const detectedMime = (meta.match(/data:([^;]+);base64/) || [])[1] || mimeType || "image/png";
    const binary = atob(base64);
    const bytes = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i += 1) {
      bytes[i] = binary.charCodeAt(i);
    }
    return new File([bytes], fileName, { type: detectedMime });
  }

  function formatFileSizeMB(bytes) {
    const mb = Number(bytes) / (1024 * 1024);
    return `${mb.toFixed(1)} MB`;
  }

  function makeId() {
    return `${Date.now()}-${Math.random().toString(16).slice(2, 9)}`;
  }

  function wait(ms) {
    return new Promise((resolve) => {
      setTimeout(resolve, ms);
    });
  }
})();
