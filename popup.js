const scanButton = document.getElementById("scanButton");
const exportButton = document.getElementById("exportButton");
const invertButton = document.getElementById("invertButton");
const checkAllBox = document.getElementById("checkAllBox");
const statusText = document.getElementById("status");
const selectionSummary = document.getElementById("selectionSummary");
const previewList = document.getElementById("previewList");

let sourceTabId = null;
let previewItems = [];
let selectedIds = new Set();

init();

scanButton.addEventListener("click", scanCart);
exportButton.addEventListener("click", exportSelected);

checkAllBox.addEventListener("change", () => {
  selectedIds = checkAllBox.checked
    ? new Set(previewItems.map((item) => item.id))
    : new Set();
  renderPreview();
  saveState();
});

invertButton.addEventListener("click", () => {
  selectedIds = new Set(previewItems.filter((item) => !selectedIds.has(item.id)).map((item) => item.id));
  renderPreview();
  saveState();
});

previewList.addEventListener("change", (event) => {
  const checkbox = event.target.closest("input[data-item-id]");
  if (!checkbox) {
    return;
  }

  if (checkbox.checked) {
    selectedIds.add(checkbox.dataset.itemId);
  } else {
    selectedIds.delete(checkbox.dataset.itemId);
  }
  updateControls();
  saveState();
});

async function init() {
  const tab = await getActiveTab();
  sourceTabId = tab?.id || null;

  const state = await chrome.storage.session.get([
    "taobaoCartSourceTabId",
    "taobaoCartPreviewItems",
    "taobaoCartSelectedIds"
  ]);

  if (state.taobaoCartSourceTabId === sourceTabId && Array.isArray(state.taobaoCartPreviewItems)) {
    previewItems = state.taobaoCartPreviewItems;
    selectedIds = new Set(state.taobaoCartSelectedIds || previewItems.map((item) => item.id));
    renderPreview();
    if (previewItems.length) {
      setStatus(`已恢复 ${previewItems.length} 件预览商品。`);
    }
  }
}

async function scanCart() {
  setBusy(true);
  setStatus("正在扫描购物车，请稍等...");

  try {
    const tab = await getTaobaoTab();
    const result = await chrome.tabs.sendMessage(tab.id, {
      type: "PREVIEW_TAOBAO_CART"
    });

    if (!result?.ok) {
      throw new Error(result?.error || "扫描失败，请刷新购物车页面后重试。");
    }

    sourceTabId = tab.id;
    previewItems = result.items || [];
    selectedIds = new Set(previewItems.map((item) => item.id));
    renderPreview();
    await saveState();
    setStatus(`扫描到 ${previewItems.length} 件商品，请确认后导出。`);
  } catch (error) {
    setStatus(error.message || "扫描失败，请刷新页面后重试。");
  } finally {
    setBusy(false);
  }
}

async function exportSelected() {
  const items = previewItems.filter((item) => selectedIds.has(item.id));
  if (!items.length) {
    setStatus("请至少勾选 1 件商品。");
    return;
  }

  setBusy(true);
  setStatus("正在生成表格并嵌入图片...");

  try {
    const tab = await getTaobaoTab();
    const result = await chrome.tabs.sendMessage(tab.id, {
      type: "EXPORT_TAOBAO_CART",
      items
    });

    if (!result?.ok) {
      throw new Error(result?.error || "导出失败，请刷新购物车页面后重试。");
    }

    setStatus(`已导出 ${result.count} 件商品。`);
  } catch (error) {
    setStatus(error.message || "导出失败，请刷新页面后重试。");
  } finally {
    setBusy(false);
  }
}

async function getActiveTab() {
  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
  return tab;
}

async function getTaobaoTab() {
  const tab = await getActiveTab();
  if (!tab?.id || !/^https?:\/\/.*(taobao|tmall)\.com\//i.test(tab.url || "")) {
    if (sourceTabId) {
      return { id: sourceTabId };
    }
    throw new Error("请先切换到淘宝或天猫购物车页面。");
  }
  await ensureContentScript(tab.id);
  sourceTabId = tab.id;
  return tab;
}

async function ensureContentScript(tabId) {
  try {
    await chrome.tabs.sendMessage(tabId, { type: "PING_TAOBAO_CART_EXPORTER" });
  } catch {
    await chrome.scripting.executeScript({
      target: { tabId },
      files: ["content.js"]
    });
  }
}

function renderPreview() {
  previewList.innerHTML = previewItems.map(renderPreviewItem).join("");
  updateControls();
}

function renderPreviewItem(item) {
  const checked = selectedIds.has(item.id) ? "checked" : "";
  const image = item.image
    ? `<img class="itemImage" src="${escapeHtml(item.image)}" alt="">`
    : `<div class="itemImage placeholder"></div>`;

  return `
    <label class="previewItem">
      <input class="itemCheck" type="checkbox" data-item-id="${escapeHtml(item.id)}" ${checked}>
      ${image}
      <span class="itemInfo">
        <strong title="${escapeHtml(item.title)}">${escapeHtml(item.title)}</strong>
        <span class="meta">${escapeHtml(item.shop || "店铺未知")}</span>
        <span class="meta">${escapeHtml(item.sku || "规格未知")}</span>
        <span class="price">¥${escapeHtml(item.unitPrice || "")} × ${escapeHtml(item.quantity || "1")}</span>
      </span>
    </label>
  `;
}

function updateControls() {
  const total = previewItems.length;
  const selected = selectedIds.size;

  selectionSummary.textContent = total ? `已选 ${selected}/${total}` : "未扫描";
  exportButton.disabled = !selected;
  invertButton.disabled = !total;
  checkAllBox.disabled = !total;
  checkAllBox.checked = total > 0 && selected === total;
  checkAllBox.indeterminate = selected > 0 && selected < total;
}

function setBusy(isBusy) {
  scanButton.disabled = isBusy;
  exportButton.disabled = isBusy || !selectedIds.size;
  invertButton.disabled = isBusy || !previewItems.length;
  checkAllBox.disabled = isBusy || !previewItems.length;
}

function setStatus(message) {
  statusText.textContent = message;
}

async function saveState() {
  await chrome.storage.session.set({
    taobaoCartSourceTabId: sourceTabId,
    taobaoCartPreviewItems: previewItems,
    taobaoCartSelectedIds: [...selectedIds]
  });
}

function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}
