// ─────────────────────────────────────────────────────────────────────
// Copyright (c) 2025-2026 Thothica Private Limited, Delhi, India.
// All rights reserved.  Proprietary and confidential.
// Unauthorized copying or distribution is strictly prohibited.
// ─────────────────────────────────────────────────────────────────────
// ── PDF Toolkit — Frontend Logic ─────────────────────────────────

(function () {
  "use strict";

  // ── License Gate ──────────────────────────────────────────────────

  const licenseModal = document.getElementById("license-modal");
  const licenseKeyInput = document.getElementById("license-key-input");
  const licenseActivateBtn = document.getElementById("license-activate-btn");
  const licenseError = document.getElementById("license-error");
  const licenseMessage = document.getElementById("license-message");

  async function checkLicense() {
    try {
      const resp = await fetch("/api/license/status");
      const data = await resp.json();
      if (data.valid) {
        licenseModal.hidden = true;
        return;
      }
      // Show license modal with the message from the server
      licenseMessage.textContent =
        data.message || "Enter your license key to get started.";
      licenseModal.hidden = false;
      licenseKeyInput.focus();
    } catch {
      // If the status check itself fails, show the modal
      licenseModal.hidden = false;
      licenseKeyInput.focus();
    }
  }

  async function activateLicense() {
    const key = licenseKeyInput.value.trim();
    if (!key) {
      showLicenseError("Please enter a license key.");
      return;
    }

    licenseActivateBtn.disabled = true;
    licenseActivateBtn.textContent = "Activating...";
    licenseError.hidden = true;

    try {
      const resp = await fetch("/api/license/activate", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ key }),
      });
      const data = await resp.json();

      if (data.success) {
        licenseModal.hidden = true;
        licenseError.hidden = true;
      } else {
        showLicenseError(data.message || "Activation failed.");
      }
    } catch {
      showLicenseError("Cannot reach the license server.");
    } finally {
      licenseActivateBtn.disabled = false;
      licenseActivateBtn.textContent = "Activate";
    }
  }

  function showLicenseError(msg) {
    licenseError.textContent = msg;
    licenseError.hidden = false;
  }

  licenseActivateBtn.addEventListener("click", activateLicense);
  licenseKeyInput.addEventListener("keydown", (e) => {
    if (e.key === "Enter") activateLicense();
  });

  // Auto-format key input as XXXX-XXXX-XXXX-XXXX
  licenseKeyInput.addEventListener("input", () => {
    let raw = licenseKeyInput.value.replace(/[^A-Za-z0-9]/g, "").toUpperCase();
    let parts = [];
    for (let i = 0; i < raw.length && parts.length < 4; i += 4) {
      parts.push(raw.slice(i, i + 4));
    }
    licenseKeyInput.value = parts.join("-");
  });

  // Run license check on startup
  checkLicense();

  // ── DOM refs ────────────────────────────────────────────────────

  // Top bar
  const dirPath = document.getElementById("dir-path");
  const browseFolderBtn = document.getElementById("browse-folder-btn");
  const fileCountLabel = document.getElementById("file-count");
  const h1Only = document.getElementById("h1-only");
  const countToc = document.getElementById("count-toc");

  // Sidebar
  const fileList = document.getElementById("file-list");
  const sidebarEmpty = document.getElementById("sidebar-empty");
  const remainingCount = document.getElementById("remaining-count");
  const skipAllBtn = document.getElementById("skip-all-btn");
  const tocBatchBtn = document.getElementById("toc-batch-btn");
  const docxBatchBtn = document.getElementById("docx-batch-btn");

  // Editor
  const editorFilename = document.getElementById("editor-filename");
  const selectionCount = document.getElementById("selection-count");
  const deleteBtn = document.getElementById("delete-btn");
  const saveBtn = document.getElementById("save-btn");
  const finishBtn = document.getElementById("finish-btn");
  const skipCurrentBtn = document.getElementById("skip-current-btn");
  const pageGrid = document.getElementById("page-grid");
  const editorEmpty = document.getElementById("editor-empty");

  // Status bar
  const statusText = document.getElementById("status-text");
  const statusProgress = document.getElementById("status-progress");

  // Preview modal
  const previewModal = document.getElementById("preview-modal");
  const previewImg = document.getElementById("preview-img");
  const previewPageLabel = document.getElementById("preview-page-label");
  const modalClose = previewModal.querySelector(".modal-close");
  const modalPrev = previewModal.querySelector(".modal-prev");
  const modalNext = previewModal.querySelector(".modal-next");
  const modalBackdrop = previewModal.querySelector(".modal-backdrop");
  const imgContainer = document.querySelector(".preview-img-container");

  // Region selection
  const regionSelectBtn = document.getElementById("region-select-btn");
  const regionOverlay = document.getElementById("region-overlay");
  const redactActions = document.getElementById("redact-actions");
  const redactPageBtn = document.getElementById("redact-page-btn");
  const redactAllBtn = document.getElementById("redact-all-btn");
  const redactAllBooksBtn = document.getElementById("redact-all-books-btn");
  const redactCancelBtn = document.getElementById("redact-cancel-btn");

  // Page numbers
  const pageNumBtn = document.getElementById("pagenum-btn");
  const pageNumConfig = document.getElementById("pagenumber-config");
  const pnFontsize = document.getElementById("pn-fontsize");
  const pnStart = document.getElementById("pn-start");
  const pnFont = document.getElementById("pn-font");
  const pnHint = document.getElementById("pn-hint");
  const pnApplyBtn = document.getElementById("pn-apply-btn");
  const pnApplyAllBtn = document.getElementById("pn-apply-all-btn");
  const pnCancelBtn = document.getElementById("pn-cancel-btn");
  const placementMarker = document.getElementById("placement-marker");

  // DOCX modal
  const docxModal = document.getElementById("docx-modal");
  const docxModalBackdrop = document.getElementById("docx-modal-backdrop");
  const docxCloseBtn = document.getElementById("docx-close-btn");
  const batchProgress = document.getElementById("batch-progress");
  const batchProgressLabel = document.getElementById("batch-progress-label");
  const logOutput = document.getElementById("log-output");
  const stopBtn = document.getElementById("stop-btn");

  // ── State ───────────────────────────────────────────────────────

  // Folder state
  let folderPath = "";
  let folderFiles = [];       // [{name, path, size, size_str}, ...]
  let activeFileName = null;  // currently open file's name

  // Editor state
  let editorPageCount = 0;
  let editorPath = "";
  let editorModified = false;
  let pageSizes = [];         // [{w, h}, ...] from backend
  let selectedPages = new Set();
  let lastClickedPage = -1;
  let previewPage = -1;
  let observer = null;

  // Region selection state
  let regionMode = false;
  let regionDrawing = false;
  let regionStart = null;
  let regionRect = null;

  // Page number placement state
  let pageNumMode = false;
  let pageNumPoint = null;

  // Batch state
  let eventSource = null;
  let batchRunning = false;

  // Cache-busting counter for thumbnails after modifications
  let thumbCacheBust = Date.now();

  function invalidateThumbs() {
    thumbCacheBust = Date.now();
    invalidateThumbQueue();
  }

  // ── Concurrent Thumbnail Limiter ──────────────────────────────
  // Limits in-flight thumbnail requests to avoid lock contention on the server.

  const THUMB_CONCURRENCY = 6;
  let thumbActive = 0;
  let thumbQueue = [];
  let thumbGeneration = 0;

  function invalidateThumbQueue() {
    thumbQueue = [];
    thumbActive = 0;
    thumbGeneration++;
  }

  function queueThumb(url, wrapper, page) {
    const gen = thumbGeneration;
    if (thumbActive < THUMB_CONCURRENCY) {
      startThumbLoad(url, wrapper, page, gen);
    } else {
      thumbQueue.push({ url, wrapper, page, gen });
    }
  }

  function startThumbLoad(url, wrapper, page, gen) {
    if (gen !== thumbGeneration) return; // stale — skip
    thumbActive++;

    // Create and append img only when we're ready to load
    const img = document.createElement("img");
    img.alt = "Page " + (page + 1);
    img.src = url;
    wrapper.appendChild(img);

    const done = () => {
      thumbActive = Math.max(0, thumbActive - 1);
      drainThumbQueue();
    };

    img.onload = () => {
      if (gen !== thumbGeneration) {
        img.src = "";  // release stale image memory
        img.remove();
        done();
        return;
      }
      const ph = wrapper.querySelector(".thumb-placeholder");
      if (ph) ph.remove();
      done();
    };

    img.onerror = () => {
      // Gracefully handle errors (doc closed, page deleted, etc.)
      img.remove();
      done();
    };
  }

  function drainThumbQueue() {
    while (thumbQueue.length > 0 && thumbActive < THUMB_CONCURRENCY) {
      const item = thumbQueue.shift();
      if (item.gen === thumbGeneration) {
        startThumbLoad(item.url, item.wrapper, item.page, item.gen);
      }
    }
  }

  // ── Browse Folder ─────────────────────────────────────────────

  browseFolderBtn.addEventListener("click", async () => {
    if (editorModified) {
      if (!confirm("You have unsaved changes. Load a new folder anyway?")) return;
    }
    const res = await fetch("/api/browse", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ mode: "folder", initial: dirPath.value || "" }),
    });
    const data = await res.json();
    if (data.path) {
      await loadFolder(data.path);
    }
  });

  async function loadFolder(path) {
    setStatus("Loading folder...");
    try {
      const res = await fetch("/api/folder/load", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ path }),
      });
      const data = await res.json();
      if (!res.ok) {
        alert("Error: " + (data.error || "Unknown"));
        setStatus("Ready");
        return;
      }

      folderPath = path;
      folderFiles = data.files;
      dirPath.value = path;
      fileCountLabel.textContent = data.count + " PDF" + (data.count !== 1 ? "s" : "") + " found";

      // Clear editor state
      editorPageCount = 0;
      editorPath = "";
      editorModified = false;
      pageSizes = [];
      selectedPages.clear();
      lastClickedPage = -1;
      activeFileName = null;

      renderFileList();
      clearEditor();
      updateEditorButtons();
      updateSidebarButtons();
      setStatus("Ready");
    } catch (e) {
      alert("Failed to load folder: " + e.message);
      setStatus("Ready");
    }
  }

  // ── File List Rendering ────────────────────────────────────────

  function renderFileList() {
    // Clear existing items (but not the empty state)
    const items = fileList.querySelectorAll(".file-item");
    items.forEach((el) => el.remove());

    if (folderFiles.length === 0) {
      sidebarEmpty.style.display = "";
      if (folderPath) {
        sidebarEmpty.querySelector("p").innerHTML =
          "No PDF files found in this folder.";
      } else {
        sidebarEmpty.querySelector("p").innerHTML =
          'No folder loaded.<br>Click <strong>Browse</strong> to start.';
      }
      remainingCount.textContent = "";
    } else {
      sidebarEmpty.style.display = "none";
      remainingCount.textContent = folderFiles.length + " remaining";

      folderFiles.forEach((f) => {
        const item = document.createElement("div");
        item.className = "file-item";
        if (f.name === activeFileName) item.classList.add("active");
        item.dataset.name = f.name;

        const info = document.createElement("div");
        info.className = "file-item-info";

        const nameEl = document.createElement("div");
        nameEl.className = "file-item-name";
        nameEl.textContent = f.name;
        nameEl.title = f.name;

        const sizeEl = document.createElement("div");
        sizeEl.className = "file-item-size";
        sizeEl.textContent = f.size_str;

        info.appendChild(nameEl);
        info.appendChild(sizeEl);

        const skipBtn = document.createElement("button");
        skipBtn.className = "file-item-skip";
        skipBtn.title = "Skip — add TOC only";
        skipBtn.textContent = "\u00d7";
        skipBtn.addEventListener("click", (e) => {
          e.stopPropagation();
          skipFile(f.name);
        });

        item.appendChild(info);
        item.appendChild(skipBtn);

        item.addEventListener("click", () => selectFile(f.name));

        fileList.appendChild(item);
      });
    }

    updateSidebarButtons();
  }

  function updateSidebarButtons() {
    const hasFiles = folderFiles.length > 0;
    skipAllBtn.disabled = !hasFiles || batchRunning;
    tocBatchBtn.disabled = !folderPath || batchRunning;
    docxBatchBtn.disabled = !folderPath || batchRunning;
  }

  // ── Select File ────────────────────────────────────────────────

  async function selectFile(name) {
    if (name === activeFileName) return; // already open

    // Check unsaved changes
    if (editorModified) {
      if (!confirm("Discard unsaved changes to " + activeFileName + "?")) return;
    }

    setStatus("Opening " + name + "...");

    try {
      const res = await fetch("/api/folder/select", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ name }),
      });
      const data = await res.json();

      if (res.status === 409) {
        // Unsaved changes — backend safety net
        if (!confirm("Discard unsaved changes to " + (data.current_file || "current file") + "?")) {
          setStatus("Ready");
          return;
        }
        // Close and retry
        await fetch("/api/editor/close", { method: "POST" });
        return selectFile(name);
      }

      if (!res.ok) {
        alert("Error: " + (data.error || "Unknown"));
        setStatus("Ready");
        return;
      }

      activeFileName = name;
      editorPageCount = data.page_count;
      editorPath = data.path;
      editorModified = false;
      pageSizes = data.page_sizes || [];
      selectedPages.clear();
      lastClickedPage = -1;

      editorFilename.textContent = name + " (" + editorPageCount + " pages)";
      editorEmpty.style.display = "none";

      // Highlight in sidebar
      fileList.querySelectorAll(".file-item").forEach((el) => {
        el.classList.toggle("active", el.dataset.name === name);
      });

      invalidateThumbs();
      updateEditorButtons();
      buildPageGrid();
      setStatus("Ready");
    } catch (e) {
      alert("Failed to open: " + e.message);
      setStatus("Ready");
    }
  }

  // ── Skip File ──────────────────────────────────────────────────

  async function skipFile(name) {
    // Check unsaved changes if this is the active file
    if (name === activeFileName && editorModified) {
      if (!confirm("Discard unsaved changes to " + name + "?")) return;
    }

    // Mark as processing in sidebar
    const item = fileList.querySelector('.file-item[data-name="' + CSS.escape(name) + '"]');
    if (item) {
      item.classList.add("processing");
      const skipBtn = item.querySelector(".file-item-skip");
      if (skipBtn) {
        skipBtn.outerHTML = '<div class="spinner"></div>';
      }
    }

    setStatus("Adding TOC to " + name + "...");

    try {
      const res = await fetch("/api/folder/skip", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          name,
          h1_only: h1Only.checked,
          count_toc: countToc.checked,
        }),
      });
      const data = await res.json();
      if (!res.ok) {
        alert("Error: " + (data.error || "Unknown"));
        // Re-enable item
        if (item) item.classList.remove("processing");
        setStatus("Ready");
        return;
      }

      // Remove from local list
      folderFiles = folderFiles.filter((f) => f.name !== name);

      // If was the active file, clear editor
      if (name === activeFileName) {
        clearEditor();
      }

      renderFileList();

      if (data.toc_result === "pending") {
        setStatus(name + " — TOC processing...");
      } else if (data.toc_result === "success") {
        setStatus(name + " — TOC added");
      } else {
        setStatus(name + " — TOC " + data.toc_result);
      }
    } catch (e) {
      alert("Skip failed: " + e.message);
      if (item) item.classList.remove("processing");
      setStatus("Ready");
    }
  }

  // ── Save (intermediate) ────────────────────────────────────────

  saveBtn.addEventListener("click", async () => {
    if (!editorPageCount) return;

    saveBtn.disabled = true;
    setStatus("Saving " + activeFileName + "...");

    try {
      const res = await fetch("/api/editor/save", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({}),
      });
      const data = await res.json();
      if (!res.ok) {
        alert("Save failed: " + (data.error || "Unknown"));
        setStatus("Ready");
        updateEditorButtons();
        return;
      }

      editorModified = false;
      invalidateThumbs();
      updateEditorButtons();
      buildPageGrid();
      setStatus("Saved " + activeFileName);
    } catch (e) {
      alert("Save failed: " + e.message);
      setStatus("Ready");
      updateEditorButtons();
    }
  });

  // ── Save & Finish ──────────────────────────────────────────────

  finishBtn.addEventListener("click", saveAndFinish);

  async function saveAndFinish() {
    if (!editorPageCount && !activeFileName) return;

    finishBtn.disabled = true;
    saveBtn.disabled = true;
    setStatus("Saving " + activeFileName + "...");

    try {
      const res = await fetch("/api/folder/save-and-finish", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          h1_only: h1Only.checked,
          count_toc: countToc.checked,
        }),
      });
      const data = await res.json();
      if (!res.ok) {
        alert("Error: " + (data.error || "Unknown"));
        setStatus("Ready");
        updateEditorButtons();
        return;
      }

      const savedName = activeFileName;

      // Remove from local list
      folderFiles = folderFiles.filter((f) => f.name !== savedName);
      clearEditor();
      renderFileList();

      if (data.toc_result === "pending") {
        setStatus(savedName + " — saved, TOC processing...");
      } else if (data.toc_result === "success") {
        setStatus(savedName + " — saved with TOC");
      } else {
        setStatus(savedName + " — saved (TOC: " + data.toc_result + ")");
      }

      // Auto-advance to next file
      if (data.next_file) {
        setTimeout(() => selectFile(data.next_file), 300);
      } else if (folderFiles.length === 0) {
        showAllDone();
      }
    } catch (e) {
      alert("Save & Finish failed: " + e.message);
      setStatus("Ready");
      updateEditorButtons();
    }
  }

  // ── Skip Current ───────────────────────────────────────────────

  skipCurrentBtn.addEventListener("click", () => {
    if (activeFileName) {
      skipFile(activeFileName);
    }
  });

  // ── Clear Editor ───────────────────────────────────────────────

  function clearEditor() {
    editorPageCount = 0;
    editorPath = "";
    editorModified = false;
    pageSizes = [];
    selectedPages.clear();
    lastClickedPage = -1;
    activeFileName = null;

    if (observer) {
      observer.disconnect();
      observer = null;
    }

    // Release decoded image memory before removing DOM nodes
    pageGrid.querySelectorAll("img").forEach((img) => {
      img.src = "";
    });
    invalidateThumbQueue();

    pageGrid.innerHTML = "";
    pageGrid.appendChild(editorEmpty);
    editorEmpty.style.display = "";
    editorFilename.textContent = "";
    selectionCount.textContent = "";

    // Remove active highlight from sidebar
    fileList.querySelectorAll(".file-item").forEach((el) => {
      el.classList.remove("active");
    });

    updateEditorButtons();
  }

  function showAllDone() {
    editorEmpty.innerHTML =
      '<div class="empty-icon">&#10004;</div><p>All files processed!</p>';
    editorEmpty.classList.add("all-done");
    editorEmpty.style.display = "";
  }

  // ── Update Editor Buttons ──────────────────────────────────────

  function updateEditorButtons() {
    const hasDoc = editorPageCount > 0;
    const hasSel = selectedPages.size > 0;

    deleteBtn.disabled = !hasSel;
    saveBtn.disabled = !hasDoc || !editorModified;
    finishBtn.disabled = !hasDoc;
    skipCurrentBtn.disabled = !hasDoc;

    selectionCount.textContent = hasSel
      ? selectedPages.size + " of " + editorPageCount + " selected"
      : hasDoc
        ? editorPageCount + " pages"
        : "";
  }

  // ── Status Bar ─────────────────────────────────────────────────

  function setStatus(text) {
    statusText.textContent = text;
  }

  // ── Page Grid ──────────────────────────────────────────────────

  function buildPageGrid() {
    // Cleanup previous observer
    if (observer) {
      observer.disconnect();
      observer = null;
    }

    // Release decoded image memory from previous grid
    pageGrid.querySelectorAll("img").forEach((img) => {
      img.src = "";
    });

    // Clear grid but keep the empty state element
    const emptyEl = editorEmpty;
    pageGrid.innerHTML = "";

    for (let i = 0; i < editorPageCount; i++) {
      const cell = document.createElement("div");
      cell.className = "page-cell";
      cell.dataset.page = i;

      const wrapper = document.createElement("div");
      wrapper.className = "thumb-wrapper";

      const placeholder = document.createElement("span");
      placeholder.className = "thumb-placeholder";
      placeholder.textContent = i + 1;
      wrapper.appendChild(placeholder);

      const label = document.createElement("span");
      label.className = "page-label";
      label.textContent = "Page " + (i + 1);

      cell.appendChild(wrapper);
      cell.appendChild(label);

      // Click handlers
      cell.addEventListener("click", (e) => handlePageClick(e, i));
      cell.addEventListener("dblclick", (e) => {
        e.preventDefault();
        showPreview(i);
      });

      pageGrid.appendChild(cell);
    }

    // Re-append the empty state (hidden)
    pageGrid.appendChild(emptyEl);
    emptyEl.style.display = "none";

    // Wait for DOM layout before setting up observer
    requestAnimationFrame(() => {
      requestAnimationFrame(() => {
        setupIntersectionObserver();
      });
    });
  }

  function setupIntersectionObserver() {
    observer = new IntersectionObserver(
      (entries) => {
        entries.forEach((entry) => {
          if (!entry.isIntersecting) return;

          const cell = entry.target;
          const page = parseInt(cell.dataset.page);
          const wrapper = cell.querySelector(".thumb-wrapper");

          // Already loaded or queued?
          if (wrapper.querySelector("img")) return;

          // Queue the load through the concurrency limiter
          const url =
            "/api/editor/thumb/" + page + "?w=180&t=" + thumbCacheBust;
          queueThumb(url, wrapper, page);

          observer.unobserve(cell);
        });
      },
      { root: pageGrid, rootMargin: "300px" }
    );

    pageGrid.querySelectorAll(".page-cell").forEach((cell) => {
      observer.observe(cell);
    });
  }

  // ── Selection ──────────────────────────────────────────────────

  function handlePageClick(e, pageIndex) {
    if (e.detail >= 2) return; // let dblclick handle it

    if (e.ctrlKey || e.metaKey) {
      // Toggle
      if (selectedPages.has(pageIndex)) {
        selectedPages.delete(pageIndex);
      } else {
        selectedPages.add(pageIndex);
      }
    } else if (e.shiftKey && lastClickedPage >= 0) {
      // Range
      const from = Math.min(lastClickedPage, pageIndex);
      const to = Math.max(lastClickedPage, pageIndex);
      for (let i = from; i <= to; i++) {
        selectedPages.add(i);
      }
    } else {
      // Single
      selectedPages.clear();
      selectedPages.add(pageIndex);
    }

    lastClickedPage = pageIndex;
    applySelectionVisuals();
    updateEditorButtons();
  }

  function applySelectionVisuals() {
    pageGrid.querySelectorAll(".page-cell").forEach((cell) => {
      const idx = parseInt(cell.dataset.page);
      cell.classList.toggle("selected", selectedPages.has(idx));
    });
  }

  // ── Delete ─────────────────────────────────────────────────────

  deleteBtn.addEventListener("click", deleteSelected);

  async function deleteSelected() {
    if (selectedPages.size === 0) return;

    const count = selectedPages.size;
    if (!confirm("Delete " + count + " page" + (count > 1 ? "s" : "") + "?"))
      return;

    const pages = Array.from(selectedPages);
    deleteBtn.disabled = true;

    try {
      const res = await fetch("/api/editor/delete", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ pages: pages }),
      });
      const data = await res.json();
      if (!res.ok) {
        alert("Error: " + (data.error || "Unknown"));
        return;
      }

      editorPageCount = data.page_count;
      editorModified = true;
      selectedPages.clear();
      lastClickedPage = -1;
      invalidateThumbs();

      editorFilename.textContent =
        activeFileName + " (" + editorPageCount + " pages)";

      updateEditorButtons();
      buildPageGrid();
    } catch (e) {
      alert("Delete failed: " + e.message);
    }
  }

  // ── Preview Modal ──────────────────────────────────────────────

  function showPreview(page) {
    previewPage = page;
    previewImg.src = "/api/editor/preview/" + page + "?dpi=150";
    previewPageLabel.textContent =
      "Page " + (page + 1) + " of " + editorPageCount;
    previewModal.hidden = false;
    modalPrev.disabled = page <= 0;
    modalNext.disabled = page >= editorPageCount - 1;
    document.body.style.overflow = "hidden";

    // Reset modal modes
    exitRegionMode();
    exitPageNumMode();
  }

  function hidePreview() {
    previewModal.hidden = true;
    previewPage = -1;
    previewImg.src = "";  // release high-res preview image memory
    document.body.style.overflow = "";
    exitRegionMode();
    exitPageNumMode();
  }

  function navigatePreview(delta) {
    if (regionMode || pageNumMode) return;
    const next = previewPage + delta;
    if (next < 0 || next >= editorPageCount) return;
    showPreview(next);
  }

  modalClose.addEventListener("click", hidePreview);
  modalBackdrop.addEventListener("click", hidePreview);
  modalPrev.addEventListener("click", () => navigatePreview(-1));
  modalNext.addEventListener("click", () => navigatePreview(1));

  document.addEventListener("keydown", (e) => {
    if (previewModal.hidden) return;
    if (e.key === "Escape") {
      if (regionMode) {
        exitRegionMode();
      } else if (pageNumMode) {
        exitPageNumMode();
      } else {
        hidePreview();
      }
    } else if (e.key === "ArrowLeft") navigatePreview(-1);
    else if (e.key === "ArrowRight") navigatePreview(1);
  });

  // ── Region Selection ──────────────────────────────────────────

  regionSelectBtn.addEventListener("click", () => {
    if (regionMode) {
      exitRegionMode();
    } else {
      enterRegionMode();
    }
  });

  function enterRegionMode() {
    exitPageNumMode(); // mutually exclusive
    regionMode = true;
    regionSelectBtn.classList.add("active");
    imgContainer.classList.add("region-mode");
    regionOverlay.hidden = true;
    redactActions.hidden = true;
    regionRect = null;
  }

  function exitRegionMode() {
    regionMode = false;
    regionDrawing = false;
    regionStart = null;
    regionRect = null;
    regionSelectBtn.classList.remove("active");
    imgContainer.classList.remove("region-mode");
    regionOverlay.hidden = true;
    redactActions.hidden = true;
  }

  // Mouse handlers on the image container for drawing selection
  imgContainer.addEventListener("mousedown", (e) => {
    if (!regionMode || e.button !== 0) return;
    e.preventDefault();

    const imgRect = previewImg.getBoundingClientRect();
    regionStart = {
      x: e.clientX - imgRect.left,
      y: e.clientY - imgRect.top,
    };
    regionDrawing = true;

    // Reset overlay
    regionOverlay.hidden = false;
    regionOverlay.style.left = regionStart.x + "px";
    regionOverlay.style.top = regionStart.y + "px";
    regionOverlay.style.width = "0px";
    regionOverlay.style.height = "0px";
    redactActions.hidden = true;
  });

  document.addEventListener("mousemove", (e) => {
    if (!regionDrawing) return;
    e.preventDefault();

    const imgRect = previewImg.getBoundingClientRect();
    const curX = Math.max(0, Math.min(e.clientX - imgRect.left, imgRect.width));
    const curY = Math.max(
      0,
      Math.min(e.clientY - imgRect.top, imgRect.height)
    );

    const x0 = Math.min(regionStart.x, curX);
    const y0 = Math.min(regionStart.y, curY);
    const x1 = Math.max(regionStart.x, curX);
    const y1 = Math.max(regionStart.y, curY);

    regionOverlay.style.left = x0 + "px";
    regionOverlay.style.top = y0 + "px";
    regionOverlay.style.width = x1 - x0 + "px";
    regionOverlay.style.height = y1 - y0 + "px";
  });

  document.addEventListener("mouseup", (e) => {
    if (!regionDrawing) return;
    regionDrawing = false;

    const imgRect = previewImg.getBoundingClientRect();
    const curX = Math.max(0, Math.min(e.clientX - imgRect.left, imgRect.width));
    const curY = Math.max(
      0,
      Math.min(e.clientY - imgRect.top, imgRect.height)
    );

    const x0 = Math.min(regionStart.x, curX);
    const y0 = Math.min(regionStart.y, curY);
    const x1 = Math.max(regionStart.x, curX);
    const y1 = Math.max(regionStart.y, curY);

    // Minimum size check (at least 5px)
    if (x1 - x0 < 5 || y1 - y0 < 5) {
      regionOverlay.hidden = true;
      regionRect = null;
      return;
    }

    regionRect = { x0, y0, x1, y1 };
    redactActions.hidden = false;
    redactAllBooksBtn.disabled = !folderPath;
  });

  // Convert screen coords to PDF coords
  function screenToPdfRect(screenRect) {
    if (!pageSizes[previewPage]) return null;

    const imgRect = previewImg.getBoundingClientRect();
    const pdfW = pageSizes[previewPage].w;
    const pdfH = pageSizes[previewPage].h;
    const scaleX = pdfW / imgRect.width;
    const scaleY = pdfH / imgRect.height;

    return {
      x0: screenRect.x0 * scaleX,
      y0: screenRect.y0 * scaleY,
      x1: screenRect.x1 * scaleX,
      y1: screenRect.y1 * scaleY,
    };
  }

  async function doRedact(scope) {
    if (!regionRect) return;

    const pdfRect = screenToPdfRect(regionRect);
    if (!pdfRect) return;

    const label =
      scope === "all"
        ? "Remove this region from ALL " + editorPageCount + " pages?"
        : "Remove this region from page " + (previewPage + 1) + "?";
    if (!confirm(label)) return;

    redactPageBtn.disabled = true;
    redactAllBtn.disabled = true;
    redactAllBooksBtn.disabled = true;

    try {
      const res = await fetch("/api/editor/redact", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          rect: pdfRect,
          page: previewPage,
          scope: scope,
        }),
      });
      const data = await res.json();
      if (!res.ok) {
        alert("Error: " + (data.error || "Unknown"));
        return;
      }

      editorModified = true;
      invalidateThumbs();
      updateEditorButtons();

      // Refresh the preview image
      previewImg.src =
        "/api/editor/preview/" + previewPage + "?dpi=150&t=" + Date.now();

      exitRegionMode();

      // Rebuild thumbnail grid to show changes
      buildPageGrid();
    } catch (e) {
      alert("Redact failed: " + e.message);
    } finally {
      redactPageBtn.disabled = false;
      redactAllBtn.disabled = false;
      redactAllBooksBtn.disabled = !folderPath;
    }
  }

  redactPageBtn.addEventListener("click", () => doRedact("page"));
  redactAllBtn.addEventListener("click", () => doRedact("all"));
  redactAllBooksBtn.addEventListener("click", redactAllBooks);
  redactCancelBtn.addEventListener("click", exitRegionMode);

  async function redactAllBooks() {
    if (!regionRect || !folderPath) return;

    const pdfRect = screenToPdfRect(regionRect);
    if (!pdfRect) return;

    if (!confirm("Remove this region from ALL pages of EVERY PDF in the folder? This modifies the original files and cannot be undone.")) return;

    redactPageBtn.disabled = true;
    redactAllBtn.disabled = true;
    redactAllBooksBtn.disabled = true;
    exitRegionMode();
    hidePreview();

    // Open batch modal
    docxModal.hidden = false;
    document.querySelector(".docx-modal-header h3").textContent = "Removing Region from All Books";
    logOutput.textContent = "";
    batchProgress.value = 0;
    batchProgressLabel.textContent = "";
    batchRunning = true;
    stopBtn.disabled = false;

    try {
      const res = await fetch("/api/folder/redact-all", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ rect: pdfRect }),
      });
      const data = await res.json();
      if (!res.ok) {
        appendLog("ERROR: " + (data.error || "Unknown error") + "\n");
        batchRunning = false;
        stopBtn.disabled = true;
        return;
      }

      batchProgress.max = data.total;
      clearEditor();
      connectDocxSSE();
    } catch (e) {
      appendLog("ERROR: " + e.message + "\n");
      batchRunning = false;
      stopBtn.disabled = true;
    }
  }

  // ── Page Number Placement ─────────────────────────────────────

  pageNumBtn.addEventListener("click", () => {
    if (pageNumMode) {
      exitPageNumMode();
    } else {
      enterPageNumMode();
    }
  });

  function enterPageNumMode() {
    exitRegionMode(); // mutually exclusive
    pageNumMode = true;
    pageNumPoint = null;
    pageNumBtn.classList.add("active");
    imgContainer.classList.add("region-mode"); // reuse crosshair cursor
    pageNumConfig.hidden = false;
    placementMarker.hidden = true;
    pnApplyBtn.disabled = true;
    pnApplyAllBtn.disabled = true;
    pnHint.textContent = "Click on the page to place";
  }

  function exitPageNumMode() {
    pageNumMode = false;
    pageNumPoint = null;
    pageNumBtn.classList.remove("active");
    if (!regionMode) {
      imgContainer.classList.remove("region-mode");
    }
    pageNumConfig.hidden = true;
    placementMarker.hidden = true;
  }

  // Click handler for placing the page number position
  imgContainer.addEventListener("click", (e) => {
    if (!pageNumMode || regionMode) return;
    // Ignore if it was a drag (region selection mouseup)
    if (regionDrawing) return;

    const imgRect = previewImg.getBoundingClientRect();
    const x = e.clientX - imgRect.left;
    const y = e.clientY - imgRect.top;

    // Clamp within image
    if (x < 0 || y < 0 || x > imgRect.width || y > imgRect.height) return;

    pageNumPoint = { x, y };

    // Show marker
    placementMarker.hidden = false;
    placementMarker.style.left = x + "px";
    placementMarker.style.top = y + "px";

    pnApplyBtn.disabled = false;
    updateApplyAllBtn();
    pnHint.textContent = "Position set - click Apply";
  });

  // Enable "Apply to all books" only when start == 1 and a point is placed
  function updateApplyAllBtn() {
    const startVal = parseInt(pnStart.value) || 1;
    pnApplyAllBtn.disabled = !pageNumPoint || startVal !== 1 || !folderPath;
  }

  pnStart.addEventListener("input", updateApplyAllBtn);

  function screenToPdfPoint(screenX, screenY) {
    if (!pageSizes[previewPage]) return null;
    const imgRect = previewImg.getBoundingClientRect();
    const pdfW = pageSizes[previewPage].w;
    const pdfH = pageSizes[previewPage].h;
    return {
      x: screenX * (pdfW / imgRect.width),
      y: screenY * (pdfH / imgRect.height),
    };
  }

  pnApplyBtn.addEventListener("click", applyPageNumbers);
  pnApplyAllBtn.addEventListener("click", applyPageNumbersAllBooks);
  pnCancelBtn.addEventListener("click", exitPageNumMode);

  async function applyPageNumbers() {
    if (!pageNumPoint) return;

    const pdfPoint = screenToPdfPoint(pageNumPoint.x, pageNumPoint.y);
    if (!pdfPoint) return;

    pnApplyBtn.disabled = true;

    try {
      const res = await fetch("/api/editor/page-numbers", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          point: pdfPoint,
          fontsize: parseInt(pnFontsize.value) || 10,
          font: pnFont.value,
          start: parseInt(pnStart.value) || 1,
          color: [0, 0, 0],
        }),
      });
      const data = await res.json();
      if (!res.ok) {
        alert("Error: " + (data.error || "Unknown"));
        return;
      }

      editorModified = true;
      invalidateThumbs();
      updateEditorButtons();

      // Refresh preview
      previewImg.src =
        "/api/editor/preview/" + previewPage + "?dpi=150&t=" + Date.now();

      exitPageNumMode();
      buildPageGrid();
    } catch (e) {
      alert("Failed: " + e.message);
    } finally {
      pnApplyBtn.disabled = false;
    }
  }

  async function applyPageNumbersAllBooks() {
    if (!pageNumPoint || !folderPath) return;

    const pdfPoint = screenToPdfPoint(pageNumPoint.x, pageNumPoint.y);
    if (!pdfPoint) return;

    if (!confirm("Add page numbers to ALL PDFs in the folder? Each book will be numbered starting from page 1. This modifies the original files.")) return;

    pnApplyAllBtn.disabled = true;
    pnApplyBtn.disabled = true;
    exitPageNumMode();
    hidePreview();

    // Open batch modal to show progress
    docxModal.hidden = false;
    document.querySelector(".docx-modal-header h3").textContent = "Adding Page Numbers to All Books";
    logOutput.textContent = "";
    batchProgress.value = 0;
    batchProgressLabel.textContent = "";
    batchRunning = true;
    stopBtn.disabled = false;

    try {
      const res = await fetch("/api/folder/page-numbers-all", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          point: pdfPoint,
          fontsize: parseInt(pnFontsize.value) || 10,
          font: pnFont.value,
          color: [0, 0, 0],
        }),
      });
      const data = await res.json();
      if (!res.ok) {
        appendLog("ERROR: " + (data.error || "Unknown error") + "\n");
        batchRunning = false;
        stopBtn.disabled = true;
        return;
      }

      batchProgress.max = data.total;

      // Clear editor since backend closed it
      clearEditor();
      connectDocxSSE();
    } catch (e) {
      appendLog("ERROR: " + e.message + "\n");
      batchRunning = false;
      stopBtn.disabled = true;
    }
  }

  // ── Skip All Remaining ─────────────────────────────────────────

  skipAllBtn.addEventListener("click", skipAllRemaining);

  async function skipAllRemaining() {
    const n = folderFiles.length;
    if (n === 0) return;

    if (!confirm("Add TOC to all " + n + " remaining files and remove them from the list?")) return;

    // Close editor if open
    if (editorPageCount > 0) {
      if (editorModified) {
        if (!confirm("Discard unsaved changes to " + activeFileName + "?")) return;
      }
    }

    skipAllBtn.disabled = true;
    setStatus("Processing all remaining files...");
    statusProgress.hidden = false;
    statusProgress.value = 0;
    statusProgress.max = n;

    try {
      const res = await fetch("/api/folder/skip-all", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          h1_only: h1Only.checked,
          count_toc: countToc.checked,
        }),
      });
      const data = await res.json();
      if (!res.ok) {
        alert("Error: " + (data.error || "Unknown"));
        setStatus("Ready");
        statusProgress.hidden = true;
        skipAllBtn.disabled = false;
        return;
      }

      batchRunning = true;
      updateSidebarButtons();
      clearEditor();

      // Connect SSE
      connectSkipAllSSE();
    } catch (e) {
      alert("Failed: " + e.message);
      setStatus("Ready");
      statusProgress.hidden = true;
      skipAllBtn.disabled = false;
    }
  }

  function connectSkipAllSSE() {
    if (eventSource) eventSource.close();
    eventSource = new EventSource("/api/batch/events");

    eventSource.addEventListener("progress", (e) => {
      const d = JSON.parse(e.data);
      statusProgress.value = d.current;
      setStatus(d.current + "/" + d.total + "  " + d.file + "  (" + d.size + ")");
    });

    eventSource.addEventListener("processed", (e) => {
      const name = JSON.parse(e.data);
      // Remove from local list
      folderFiles = folderFiles.filter((f) => f.name !== name);
      renderFileList();
    });

    eventSource.addEventListener("log", () => {
      // We don't display log output in skip-all mode
    });

    eventSource.addEventListener("done", (e) => {
      const d = JSON.parse(e.data);
      batchRunning = false;
      eventSource.close();
      eventSource = null;
      statusProgress.hidden = true;

      if (d.success !== undefined) {
        const parts = [];
        if (d.success) parts.push(d.success + " succeeded");
        if (d.skipped) parts.push(d.skipped + " skipped");
        if (d.crashed) parts.push(d.crashed + " crashed");
        if (d.failed) parts.push(d.failed + " failed");
        setStatus("Done: " + parts.join(" | "));
      } else {
        setStatus("Done");
      }

      renderFileList();
      updateSidebarButtons();

      if (folderFiles.length === 0) {
        showAllDone();
      }

      // Reclaim backend native memory after bulk operation
      fetch("/api/cleanup", { method: "POST" }).catch(() => {});
    });

    eventSource.onerror = () => {
      batchRunning = false;
      eventSource.close();
      eventSource = null;
      statusProgress.hidden = true;
      setStatus("Connection lost");
      updateSidebarButtons();
    };
  }

  // ── Batch Modal (shared by DOCX, TOC, Page Numbers) ────────────

  tocBatchBtn.addEventListener("click", openTocBatchModal);
  docxBatchBtn.addEventListener("click", openDocxModal);
  docxCloseBtn.addEventListener("click", closeDocxModal);
  docxModalBackdrop.addEventListener("click", () => {
    if (!batchRunning) closeDocxModal();
  });
  stopBtn.addEventListener("click", stopBatch);

  function openTocBatchModal() {
    if (!folderPath) return;
    docxModal.hidden = false;
    document.querySelector(".docx-modal-header h3").textContent = "Add TOC to All PDFs";
    logOutput.textContent = "";
    batchProgress.value = 0;
    batchProgressLabel.textContent = "";
    startBatch("toc");
  }

  function openDocxModal() {
    if (!folderPath) return;
    docxModal.hidden = false;
    document.querySelector(".docx-modal-header h3").textContent = "Convert All PDFs to DOCX";
    logOutput.textContent = "";
    batchProgress.value = 0;
    batchProgressLabel.textContent = "";
    startBatch("docx");
  }

  function closeDocxModal() {
    if (batchRunning) {
      if (!confirm("A batch is still running. Close anyway?")) return;
      stopBatch();
    }
    docxModal.hidden = true;
  }

  async function startBatch(mode) {
    logOutput.textContent = "";
    batchRunning = true;
    stopBtn.disabled = false;

    try {
      const res = await fetch("/api/batch/start", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          dir: folderPath,
          mode: mode,
          h1_only: h1Only.checked,
          count_toc: countToc.checked,
        }),
      });

      const data = await res.json();
      if (!res.ok) {
        appendLog("ERROR: " + (data.error || "Unknown error") + "\n");
        batchRunning = false;
        stopBtn.disabled = true;
        return;
      }

      batchProgress.max = data.total;
      connectDocxSSE();
    } catch (e) {
      appendLog("ERROR: " + e.message + "\n");
      batchRunning = false;
      stopBtn.disabled = true;
    }
  }

  function connectDocxSSE() {
    if (eventSource) eventSource.close();

    eventSource = new EventSource("/api/batch/events");

    eventSource.addEventListener("log", (e) => {
      appendLog(JSON.parse(e.data));
    });

    eventSource.addEventListener("progress", (e) => {
      const d = JSON.parse(e.data);
      batchProgress.value = d.current;
      batchProgressLabel.textContent =
        d.current + "/" + d.total + "  " + d.file + "  (" + d.size + ")";
    });

    eventSource.addEventListener("done", (e) => {
      const d = JSON.parse(e.data);
      if (d.success !== undefined) {
        const parts = [];
        if (d.success) parts.push(d.success + " succeeded");
        if (d.skipped) parts.push(d.skipped + " skipped");
        if (d.crashed) parts.push(d.crashed + " crashed");
        if (d.failed) parts.push(d.failed + " failed");
        appendLog(
          "\n" + "=".repeat(50) + "\nDone.  " + parts.join(" | ") + "\n"
        );
        if (d.output_dir) appendLog("Output: " + d.output_dir + "\n");
      }
      batchRunning = false;
      stopBtn.disabled = true;
      batchProgressLabel.textContent = "";
      eventSource.close();
      eventSource = null;

      // Reclaim backend native memory after bulk operation
      fetch("/api/cleanup", { method: "POST" }).catch(() => {});
    });

    eventSource.onerror = () => {
      batchRunning = false;
      stopBtn.disabled = true;
      eventSource.close();
      eventSource = null;
    };
  }

  async function stopBatch() {
    await fetch("/api/batch/stop", { method: "POST" });
  }

  function appendLog(text) {
    logOutput.textContent += text;
    // Prevent log from growing unbounded (keep last 80KB)
    if (logOutput.textContent.length > 100000) {
      logOutput.textContent =
        "...(truncated)...\n" + logOutput.textContent.slice(-80000);
    }
    logOutput.scrollTop = logOutput.scrollHeight;
  }

  // ── Unsaved changes warning ────────────────────────────────────

  window.addEventListener("beforeunload", (e) => {
    if (editorModified) {
      e.preventDefault();
      e.returnValue = "";
    }
  });

  // ── Utility ────────────────────────────────────────────────────

  function debounce(fn, ms) {
    let timer;
    return function (...args) {
      clearTimeout(timer);
      timer = setTimeout(() => fn.apply(this, args), ms);
    };
  }
})();
