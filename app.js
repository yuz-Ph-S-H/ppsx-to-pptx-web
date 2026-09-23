const PPSX_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.presentationml.slideshow.main+xml";
const PPTX_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml";
const PPTX_MIME =
  "application/vnd.openxmlformats-officedocument.presentationml.presentation";

const state = {
  directoryHandle: null,
  files: [],
  existingPptx: new Set(),
  running: false,
};

const els = {
  selectFolderBtn: document.querySelector("#selectFolderBtn"),
  convertBtn: document.querySelector("#convertBtn"),
  rescanBtn: document.querySelector("#rescanBtn"),
  folderName: document.querySelector("#folderName"),
  folderHint: document.querySelector("#folderHint"),
  fileCount: document.querySelector("#fileCount"),
  existingCount: document.querySelector("#existingCount"),
  statusText: document.querySelector("#statusText"),
  fileList: document.querySelector("#fileList"),
  emptyState: document.querySelector("#emptyState"),
  progressWrap: document.querySelector("#progressWrap"),
  currentFile: document.querySelector("#currentFile"),
  progressText: document.querySelector("#progressText"),
  progressBar: document.querySelector("#progressBar"),
  compatibilityNote: document.querySelector("#compatibilityNote"),
  confirmDialog: document.querySelector("#confirmDialog"),
  confirmMessage: document.querySelector("#confirmMessage"),
  resultDialog: document.querySelector("#resultDialog"),
  resultIcon: document.querySelector("#resultIcon"),
  resultTitle: document.querySelector("#resultTitle"),
  resultMessage: document.querySelector("#resultMessage"),
  errorDetails: document.querySelector("#errorDetails"),
};

function setStatus(text, kind = "idle") {
  els.statusText.textContent = text;
  els.statusText.className = `status-${kind}`;
}

function formatBytes(bytes) {
  if (bytes === 0) return "0 B";
  const units = ["B", "KB", "MB", "GB"];
  const i = Math.min(Math.floor(Math.log(bytes) / Math.log(1024)), units.length - 1);
  const value = bytes / 1024 ** i;
  return `${value.toFixed(i === 0 ? 0 : value >= 10 ? 1 : 2)} ${units[i]}`;
}

function outputName(name) {
  return name.replace(/\.ppsx$/i, ".pptx");
}

function renderFiles() {
  els.fileList.innerHTML = "";
  const hasFiles = state.files.length > 0;
  els.emptyState.hidden = hasFiles;
  els.fileList.hidden = !hasFiles;

  for (const entry of state.files) {
    const row = document.createElement("div");
    row.className = "file-item";
    row.dataset.name = entry.name;

    const main = document.createElement("div");
    main.className = "file-main";

    const name = document.createElement("div");
    name.className = "file-name";
    name.textContent = entry.name;

    const meta = document.createElement("div");
    meta.className = "file-meta";
    meta.textContent = `${formatBytes(entry.size)} → ${outputName(entry.name)}${
      state.existingPptx.has(outputName(entry.name).toLowerCase()) ? "（将覆盖）" : ""
    }`;

    const badge = document.createElement("span");
    badge.className = "file-state pending";
    badge.textContent = "等待";

    main.append(name, meta);
    row.append(main, badge);
    els.fileList.append(row);
  }
}

function setFileState(name, kind, text) {
  const rows = [...els.fileList.querySelectorAll(".file-item")];
  const row = rows.find((item) => item.dataset.name === name);
  if (!row) return;
  const badge = row.querySelector(".file-state");
  badge.className = `file-state ${kind}`;
  badge.textContent = text;
}

function resetFileStates() {
  for (const file of state.files) setFileState(file.name, "pending", "等待");
}

function updateProgress(percent, label) {
  const safe = Math.max(0, Math.min(100, percent));
  els.progressWrap.hidden = false;
  els.progressBar.style.width = `${safe}%`;
  els.progressText.textContent = `${Math.round(safe)}%`;
  if (label) els.currentFile.textContent = label;
}

async function scanDirectory() {
  if (!state.directoryHandle) return;

  const files = [];
  const existingPptx = new Set();

  for await (const [name, handle] of state.directoryHandle.entries()) {
    if (handle.kind !== "file") continue;
    if (/\.pptx$/i.test(name)) existingPptx.add(name.toLowerCase());
    if (!/\.ppsx$/i.test(name)) continue;

    const file = await handle.getFile();
    files.push({
      name,
      size: file.size,
      lastModified: file.lastModified,
      handle,
    });
  }

  files.sort((a, b) => a.name.localeCompare(b.name, "zh-CN", { numeric: true }));
  state.files = files;
  state.existingPptx = existingPptx;

  const overwriteCount = files.filter((f) =>
    existingPptx.has(outputName(f.name).toLowerCase())
  ).length;

  els.fileCount.textContent = String(files.length);
  els.existingCount.textContent = String(overwriteCount);
  els.convertBtn.disabled = files.length === 0 || state.running;
  els.rescanBtn.disabled = state.running;

  if (files.length) {
    els.folderHint.textContent = `已扫描完成，共 ${files.length} 个 PPSX 文件`;
    setStatus("可转换", "ready");
  } else {
    els.folderHint.textContent = "这个文件夹最外层没有找到 PPSX 文件";
    setStatus("无文件", "idle");
  }

  renderFiles();
}

async function chooseDirectory() {
  if (!("showDirectoryPicker" in window)) {
    showResult({
      ok: false,
      title: "当前浏览器不支持文件夹读写",
      message: "请使用最新版 Microsoft Edge 或 Google Chrome，并通过 http://localhost 打开本项目。",
    });
    return;
  }

  try {
    const handle = await window.showDirectoryPicker({
      id: "ppsx-pptx-converter",
      mode: "readwrite",
      startIn: "documents",
    });

    state.directoryHandle = handle;
    els.folderName.textContent = handle.name;
    els.folderHint.textContent = "正在扫描…";
    els.rescanBtn.disabled = false;
    els.convertBtn.disabled = true;
    setStatus("扫描中", "working");
    await scanDirectory();
  } catch (error) {
    if (error?.name !== "AbortError") {
      showResult({
        ok: false,
        title: "无法打开文件夹",
        message: error?.message || String(error),
      });
    }
  }
}

async function convertFile(fileEntry, fileIndex, totalFiles) {
  const sourceFile = await fileEntry.handle.getFile();
  const arrayBuffer = await sourceFile.arrayBuffer();
  const zip = await JSZip.loadAsync(arrayBuffer);
  const contentTypesEntry = zip.file("[Content_Types].xml");

  if (!contentTypesEntry) {
    throw new Error("文件中缺少 [Content_Types].xml，无法识别为有效 OOXML 演示文稿。 ");
  }

  const contentTypes = await contentTypesEntry.async("string");
  if (!contentTypes.includes(PPSX_CONTENT_TYPE)) {
    if (contentTypes.includes(PPTX_CONTENT_TYPE)) {
      throw new Error("文件内部已经标记为 PPTX，不需要转换。 ");
    }
    throw new Error("未找到标准 PPSX 主文档 Content-Type，暂不支持这个文件。 ");
  }

  zip.file(
    "[Content_Types].xml",
    contentTypes.replace(PPSX_CONTENT_TYPE, PPTX_CONTENT_TYPE)
  );

  const basePercent = (fileIndex / totalFiles) * 100;
  const oneFileShare = 1 / totalFiles;
  const blob = await zip.generateAsync(
    {
      type: "blob",
      mimeType: PPTX_MIME,
      compression: "DEFLATE",
      compressionOptions: { level: 6 },
    },
    (metadata) => {
      const overall = basePercent + metadata.percent * oneFileShare;
      updateProgress(overall, `正在转换：${fileEntry.name}`);
    }
  );

  const targetName = outputName(fileEntry.name);
  const targetHandle = await state.directoryHandle.getFileHandle(targetName, { create: true });
  const writable = await targetHandle.createWritable();
  try {
    await writable.write(blob);
  } finally {
    await writable.close();
  }

  return targetName;
}

async function askForConfirmation() {
  const overwriteCount = state.files.filter((f) =>
    state.existingPptx.has(outputName(f.name).toLowerCase())
  ).length;

  const overwriteLine = overwriteCount
    ? `\n其中 ${overwriteCount} 个同名 PPTX 已存在，将被覆盖。`
    : "";

  els.confirmMessage.textContent = `即将转换 ${state.files.length} 个 PPSX 文件。${overwriteLine}\n原 PPSX 文件会保留。是否继续？`;
  els.confirmDialog.showModal();

  return new Promise((resolve) => {
    els.confirmDialog.addEventListener(
      "close",
      () => resolve(els.confirmDialog.returnValue === "confirm"),
      { once: true }
    );
  });
}

async function startConversion() {
  if (state.running || !state.directoryHandle || state.files.length === 0) return;
  const confirmed = await askForConfirmation();
  if (!confirmed) return;

  state.running = true;
  els.selectFolderBtn.disabled = true;
  els.convertBtn.disabled = true;
  els.rescanBtn.disabled = true;
  resetFileStates();
  setStatus("转换中", "working");
  updateProgress(0, "准备转换…");

  let success = 0;
  const failures = [];

  for (let i = 0; i < state.files.length; i += 1) {
    const entry = state.files[i];
    setFileState(entry.name, "processing", "处理中");

    try {
      await convertFile(entry, i, state.files.length);
      success += 1;
      setFileState(entry.name, "success", "完成");
    } catch (error) {
      const message = error?.message || String(error);
      failures.push(`${entry.name}: ${message}`);
      setFileState(entry.name, "error", "失败");
    }
  }

  updateProgress(100, `处理完成：${success}/${state.files.length}`);
  state.running = false;
  els.selectFolderBtn.disabled = false;
  els.rescanBtn.disabled = false;
  setStatus(failures.length ? "部分完成" : "已完成", failures.length ? "error" : "done");

  await scanDirectory();
  showResult({
    ok: failures.length === 0,
    title: failures.length ? "转换已完成，但有文件失败" : "转换完成",
    message: `成功 ${success} 个，失败 ${failures.length} 个。\n生成的 PPTX 已写回“${state.directoryHandle.name}”文件夹。`,
    details: failures,
  });
}

function showResult({ ok, title, message, details = [] }) {
  els.resultTitle.textContent = title;
  els.resultMessage.textContent = message;
  els.resultIcon.textContent = ok ? "✓" : "!";
  els.resultIcon.className = `dialog-icon ${ok ? "success" : "error"}`;

  if (details.length) {
    els.errorDetails.hidden = false;
    els.errorDetails.textContent = details.join("\n\n");
  } else {
    els.errorDetails.hidden = true;
    els.errorDetails.textContent = "";
  }

  els.resultDialog.showModal();
}

function checkEnvironment() {
  const isLocalhost = ["localhost", "127.0.0.1", "[::1]"].includes(location.hostname);
  const pickerSupported = "showDirectoryPicker" in window;

  if (!window.isSecureContext) {
    els.compatibilityNote.textContent =
      "当前页面不是安全上下文，浏览器会禁止文件夹读写。请通过 http://localhost 启动本项目。";
    setStatus("环境受限", "error");
  } else if (!pickerSupported) {
    els.compatibilityNote.textContent =
      "当前浏览器不支持所需的文件夹选择接口。请改用最新版 Microsoft Edge 或 Google Chrome。";
    setStatus("浏览器不支持", "error");
  } else if (!isLocalhost && location.protocol !== "https:") {
    els.compatibilityNote.textContent =
      "部署到公网时请使用 HTTPS，否则浏览器可能禁止文件夹读写。";
  }
}

els.selectFolderBtn.addEventListener("click", chooseDirectory);
els.rescanBtn.addEventListener("click", scanDirectory);
els.convertBtn.addEventListener("click", startConversion);
checkEnvironment();
