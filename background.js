chrome.runtime.onMessage.addListener((message, _sender, sendResponse) => {
  if (message?.type !== "FETCH_IMAGE_DATA_URL") {
    return false;
  }

  fetchImageAsDataUrl(message.url)
    .then((dataUrl) => sendResponse({ ok: true, dataUrl }))
    .catch((error) => sendResponse({ ok: false, error: error.message }));

  return true;
});

async function fetchImageAsDataUrl(url) {
  if (!/^https?:\/\//i.test(url || "")) {
    throw new Error("图片地址无效。");
  }

  const response = await fetch(url, { credentials: "omit" });
  if (!response.ok) {
    throw new Error(`图片读取失败：${response.status}`);
  }

  const blob = await response.blob();
  return await blobToDataUrl(blob);
}

function blobToDataUrl(blob) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = () => reject(new Error("图片转换失败。"));
    reader.readAsDataURL(blob);
  });
}
