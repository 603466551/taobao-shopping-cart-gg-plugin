(() => {
  if (window.__taobaoCartExporterLoaded) {
    return;
  }
  window.__taobaoCartExporterLoaded = true;

  const PRODUCT_LINK_PATTERN = /(item\.taobao\.com\/item\.htm|detail\.tmall\.com\/item\.htm|detail\.tmall\.hk\/hk\/item\.htm)/i;
  const PRICE_PATTERN = /(?:¥|￥|RMB\s*)\s*\d+(?:\.\d{1,2})?/i;
  const QUANTITY_PATTERN = /(?:数量|x|×)\s*(\d{1,5})/i;
  const CART_CONTAINER_SELECTOR = ".trade-cart-item-info, [class*='trade-cart-item-info'], [class*='cartItemInfoContainer'], [class*='cartItemInfo'], .item, .cart-item, [class*='item']";
  const TITLE_SELECTOR = "[class*='itemTitle'], .title, .item-title, [class*='title'], a[href*='item.taobao.com'], a[href*='detail.tmall.com']";
  const PRICE_SELECTOR = ".price, [class*='price'], .cellprice";
  const SKU_SELECTOR = ".sku, [class*='sku'], .spec, .skuCon, [class*='spec']";

  chrome.runtime.onMessage.addListener((message, _sender, sendResponse) => {
    if (message?.type === "PING_TAOBAO_CART_EXPORTER") {
      sendResponse({ ok: true });
      return false;
    }

    if (message?.type === "PREVIEW_TAOBAO_CART") {
      previewCart()
        .then(sendResponse)
        .catch((error) => sendResponse({ ok: false, error: error.message }));
      return true;
    }

    if (message?.type === "EXPORT_TAOBAO_CART") {
      exportCart(message.items)
        .then(sendResponse)
        .catch((error) => sendResponse({ ok: false, error: error.message }));
      return true;
    }

    return false;
  });

  async function previewCart() {
    const items = await collectCartItemsWithScrolling();

    if (!items.length) {
      throw new Error("没有识别到商品。请确认当前页面是已登录的淘宝购物车页面，并等待商品加载完成。");
    }

    return { ok: true, count: items.length, items: preparePreviewItems(items) };
  }

  async function exportCart(selectedItems) {
    const items = Array.isArray(selectedItems) ? selectedItems : [];

    if (!items.length) {
      throw new Error("请先选择至少 1 件要导出的商品。");
    }

    await embedItemImages(items);
    downloadXls(items);
    return { ok: true, count: items.length };
  }

  function preparePreviewItems(items) {
    return items.map((item, index) => ({
      id: getProductKey(item.url) || normalizeUrl(item.url) || `item-${index}`,
      shop: item.shop,
      title: item.title,
      sku: item.sku,
      unitPrice: item.unitPrice,
      quantity: item.quantity,
      subtotal: item.subtotal,
      url: item.url,
      image: item.image,
      selected: item.selected,
      capturedAt: item.capturedAt
    }));
  }

  async function collectCartItemsWithScrolling() {
    const originalX = window.scrollX;
    const originalY = window.scrollY;
    const itemsByKey = new Map();
    let lastHeight = 0;

    window.scrollTo({ top: 0, behavior: "auto" });
    await delay(350);

    for (let index = 0; index < 16; index += 1) {
      mergeItems(itemsByKey, extractCartItems());
      if (itemsByKey.size > 0 && isRecommendBoundaryVisible()) {
        break;
      }

      const nextTop = Math.min(
        document.documentElement.scrollHeight,
        window.scrollY + Math.max(window.innerHeight * 0.85, 500)
      );
      window.scrollTo({ top: nextTop, behavior: "auto" });
      await delay(450);

      const nextHeight = document.documentElement.scrollHeight;
      const reachedBottom = window.scrollY + window.innerHeight >= nextHeight - 8;
      if ((itemsByKey.size > 0 && isRecommendBoundaryVisible()) || (reachedBottom && nextHeight === lastHeight)) {
        break;
      }
      lastHeight = nextHeight;
    }

    mergeItems(itemsByKey, extractCartItems());
    window.scrollTo({ top: originalY, left: originalX, behavior: "auto" });
    await delay(250);
    return [...itemsByKey.values()];
  }

  function isRecommendBoundaryVisible() {
    const boundaryPatterns = /^(猜你喜欢|为你推荐|相似好物|掌柜热卖)$/;
    const candidates = [
      ...document.querySelectorAll("h1, h2, h3, h4, [class*='title'], [class*='Title']")
    ];

    return candidates.some((element) => {
      const text = cleanText(element.innerText || element.textContent);
      if (!boundaryPatterns.test(text)) {
        return false;
      }

      const rect = element.getBoundingClientRect();
      return rect.width > 0 &&
        rect.height > 0 &&
        rect.top >= -20 &&
        rect.top <= window.innerHeight + 180;
    });
  }

  function mergeItems(itemsByKey, items) {
    for (const item of items) {
      const key = [normalizeUrl(item.url), item.title, item.sku].join("|");
      if (!itemsByKey.has(key)) {
        itemsByKey.set(key, item);
      }
    }
  }

  function extractCartItems() {
    const imageContext = buildImageContext();
    const containerItems = extractCartItemsFromContainers(imageContext);
    if (containerItems.length) {
      return dedupeItems(containerItems);
    }

    const links = getCartRoots()
      .flatMap((root) => [...root.querySelectorAll("a[href]")])
      .filter((link) => PRODUCT_LINK_PATTERN.test(link.href));

    return dedupeItems(links.map((link) => buildItemFromLink(link, imageContext)).filter(Boolean));
  }

  function extractCartItemsFromContainers(imageContext) {
    const seenContainers = new Set();
    const containers = getCartRoots()
      .flatMap((root) => [...root.querySelectorAll(CART_CONTAINER_SELECTOR)])
      .filter((container) => {
        if (seenContainers.has(container) || !isLikelyCartItemContainer(container)) {
          return false;
        }
        seenContainers.add(container);
        return true;
      });

    return containers.map((container, index) => buildItemFromContainer(container, imageContext, index)).filter(Boolean);
  }

  function buildItemFromContainer(container, imageContext, index) {
    const link = findProductLink(container);
    const cartItemContainer = getCartItemInfoContainer(link) || getCartItemInfoContainer(container) || container;
    const titleElement = findTitleElement(container, link);
    const priceElement = container.querySelector(PRICE_SELECTOR);
    const title = cleanText(titleElement?.getAttribute?.("title") || titleElement?.getAttribute?.("aria-label") || titleElement?.innerText);
    const text = cleanText(cartItemContainer.innerText || container.innerText);

    if (!link || !title || !PRICE_PATTERN.test(text)) {
      return null;
    }

    const prices = extractPrices(cleanText(priceElement?.innerText) || text);
    const allPrices = prices.length ? prices : extractPrices(text);
    const quantity = extractQuantity(container, text);
    const unitPrice = allPrices[0] || "";
    const subtotal = allPrices.length > 1 ? allPrices[allPrices.length - 1] : calculateSubtotal(unitPrice, quantity);
    const productId = getProductKey(link.href);
    const shop = findShopName(cartItemContainer);
    const sku = extractSkuFromContainer(cartItemContainer, title, allPrices);
    const image = findBestImageForItem(cartItemContainer, imageContext, productId, index);
    const selected = detectSelected(cartItemContainer);

    return {
      shop,
      title,
      sku,
      unitPrice,
      quantity,
      subtotal,
      url: link.href,
      image,
      selected,
      capturedAt: formatDateTime(new Date())
    };
  }

  function dedupeItems(items) {
    const itemsByKey = new Map();
    for (const item of items) {
      const key = getProductKey(item.url) || normalizeUrl(item.url);
      if (!key) {
        continue;
      }

      const existing = itemsByKey.get(key);
      itemsByKey.set(key, existing ? mergeDuplicateItem(existing, item) : item);
    }
    return [...itemsByKey.values()].filter((item) => item.title && !isMarketingTitle(item.title));
  }

  function mergeDuplicateItem(left, right) {
    return {
      shop: chooseBetterText(left.shop, right.shop),
      title: chooseBetterTitle(left.title, right.title),
      sku: chooseBetterText(left.sku, right.sku),
      unitPrice: chooseBetterPrice(left.unitPrice, right.unitPrice),
      quantity: chooseBetterText(left.quantity, right.quantity) || "1",
      subtotal: chooseBetterPrice(left.subtotal, right.subtotal),
      url: left.url || right.url,
      image: chooseBetterImage(left.image, right.image),
      selected: left.selected || right.selected,
      capturedAt: left.capturedAt || right.capturedAt
    };
  }

  function chooseBetterTitle(left, right) {
    const candidates = [left, right].map(cleanText).filter(Boolean);
    candidates.sort((a, b) => scoreTitle(b) - scoreTitle(a));
    return candidates[0] || "";
  }

  function scoreTitle(value) {
    const text = cleanText(value);
    if (!text) {
      return -100;
    }
    let score = Math.min(text.length, 120);
    if (isMarketingTitle(text)) {
      score -= 200;
    }
    if (/[【\[]|旗舰店|官方|正品|新款|男女|儿童|家用|套装|颜色|规格|型号|尺寸/.test(text)) {
      score += 15;
    }
    if (text.length < 8) {
      score -= 60;
    }
    return score;
  }

  function isMarketingTitle(value) {
    return /^(全网低价|五一狂欢|限时优惠|优惠券|券后|跨店满减|满减|促销|活动|热卖|爆款|猜你喜欢|为你推荐)$/i.test(cleanText(value));
  }

  function chooseBetterText(left, right) {
    const candidates = [left, right].map(cleanText).filter(Boolean);
    candidates.sort((a, b) => b.length - a.length);
    return candidates[0] || "";
  }

  function chooseBetterPrice(left, right) {
    return normalizePrice(left) || normalizePrice(right) || "";
  }

  function chooseBetterImage(left, right) {
    const candidates = [left, right].map(normalizeImageUrl).filter(isProductImageUrl);
    return candidates[0] || "";
  }

  function buildItemFromLink(link, imageContext) {
    const container = findProductContainer(link);
    if (!container) {
      return null;
    }
    const cartItemContainer = getCartItemInfoContainer(link) || getCartItemInfoContainer(container) || container;

    const title = cleanText(link.getAttribute("title") || link.getAttribute("aria-label") || link.innerText);
    const text = cleanText(cartItemContainer.innerText || container.innerText);
    if (!title || !PRICE_PATTERN.test(text)) {
      return null;
    }

    const prices = extractPrices(text);
    const quantity = extractQuantity(container, text);
    const unitPrice = prices[0] || "";
    const subtotal = prices.length > 1 ? prices[prices.length - 1] : calculateSubtotal(unitPrice, quantity);
    const productId = getProductKey(link.href);
    const shop = findShopName(cartItemContainer);
    const sku = extractSku(cartItemContainer, title, prices);
    const image = findBestImageForItem(cartItemContainer, imageContext, productId, 0);
    const selected = detectSelected(cartItemContainer);

    return {
      shop,
      title,
      sku,
      unitPrice,
      quantity,
      subtotal,
      url: link.href,
      image,
      selected,
      capturedAt: formatDateTime(new Date())
    };
  }

  function getCartRoots() {
    const roots = [
      document.getElementById("tbpc-trade-cart"),
      document.getElementById("ice-container")
    ].filter(Boolean);

    return roots.some((root) => cleanText(root.innerText))
      ? roots.filter((root) => cleanText(root.innerText))
      : [document.body];
  }

  function isLikelyCartItemContainer(container) {
    const text = cleanText(container.innerText);
    if (!text || text.length < 8 || text.length > 1800) {
      return false;
    }
    if (isInsideRecommendArea(container)) {
      return false;
    }
    return Boolean(findProductLink(container) && findTitleElement(container) && PRICE_PATTERN.test(text));
  }

  function isInsideRecommendArea(element) {
    let current = element;
    for (let depth = 0; depth < 6 && current && current !== document.body; depth += 1) {
      const marker = cleanText(`${current.id || ""} ${current.className || ""} ${current.getAttribute?.("data-spm") || ""}`);
      const text = cleanText(current.innerText);
      if (/(recommend|guess|feed|ad-|advert|market|mkt|promotion|promo|猜你喜欢|推荐|广告|热卖|相似好物)/i.test(marker)) {
        return true;
      }
      if (/猜你喜欢|为你推荐|相似好物|掌柜热卖|广告/.test(text.slice(0, 160))) {
        return true;
      }
      current = current.parentElement;
    }
    return false;
  }

  function findProductLink(container) {
    return [...container.querySelectorAll("a[href]")].find((link) => PRODUCT_LINK_PATTERN.test(link.href)) || null;
  }

  function getCartItemInfoContainer(element) {
    return element?.closest?.("[class*='trade-cart-item-info'], .trade-cart-item-info") || null;
  }

  function findTitleElement(container, link = findProductLink(container)) {
    const titleElement = container.querySelector(TITLE_SELECTOR);
    if (titleElement && cleanText(titleElement.innerText || titleElement.getAttribute?.("title"))) {
      return titleElement;
    }
    return link;
  }

  function findProductContainer(link) {
    let current = link;
    let best = null;

    for (let depth = 0; depth < 8 && current?.parentElement; depth += 1) {
      current = current.parentElement;
      const text = cleanText(current.innerText);
      const linkCount = current.querySelectorAll("a[href]").length;
      const hasQuantityInput = Boolean(current.querySelector("input[value], input[aria-valuenow], input[type='number']"));
      const hasPrice = PRICE_PATTERN.test(text);
      const isReasonableSize = text.length >= 10 && text.length <= 2500;

      if (hasPrice && isReasonableSize && (hasQuantityInput || linkCount <= 12)) {
        best = current;
        if (text.length < 900) {
          break;
        }
      }
    }

    return best;
  }

  function extractPrices(text) {
    return [...text.matchAll(new RegExp(PRICE_PATTERN, "gi"))]
      .map((match) => normalizePrice(match[0]))
      .filter(Boolean);
  }

  function normalizePrice(value) {
    const match = value.match(/\d+(?:\.\d{1,2})?/);
    return match ? match[0] : "";
  }

  function extractQuantity(container, text) {
    const input = [...container.querySelectorAll("input")]
      .find((element) => {
        const value = element.value || element.getAttribute("aria-valuenow") || element.getAttribute("value");
        return /^\d{1,5}$/.test(value || "");
      });

    if (input) {
      return input.value || input.getAttribute("aria-valuenow") || input.getAttribute("value") || "1";
    }

    const quantityText = text.match(QUANTITY_PATTERN);
    return quantityText?.[1] || "1";
  }

  function calculateSubtotal(unitPrice, quantity) {
    const priceNumber = Number(unitPrice);
    const quantityNumber = Number(quantity);
    if (!Number.isFinite(priceNumber) || !Number.isFinite(quantityNumber)) {
      return "";
    }
    return (priceNumber * quantityNumber).toFixed(2);
  }

  function findShopName(container) {
    const shopSelectors = [
      "a[href*='shop']",
      "a[href*='seller']",
      "[class*='shop']",
      "[class*='seller']",
      "[data-spm*='shop']"
    ];

    for (const selector of shopSelectors) {
      const element = container.querySelector(selector) || findNearbyPrevious(container, selector);
      const text = cleanText(element?.innerText || element?.getAttribute?.("title"));
      if (text && text !== "店铺" && text.length <= 80) {
        return text;
      }
    }

    return "";
  }

  function findNearbyPrevious(container, selector) {
    let current = container;
    for (let depth = 0; depth < 4 && current; depth += 1) {
      let sibling = current.previousElementSibling;
      while (sibling) {
        const match = sibling.matches?.(selector) ? sibling : sibling.querySelector?.(selector);
        if (match) {
          return match;
        }
        sibling = sibling.previousElementSibling;
      }
      current = current.parentElement;
    }
    return null;
  }

  function extractSku(container, title, prices) {
    const fragments = (container.innerText || "")
      .split(/\n| {2,}|\t/)
      .map(cleanText)
      .filter(Boolean)
      .filter((part) => {
        if (part === title || PRICE_PATTERN.test(part) || prices.includes(normalizePrice(part))) {
          return false;
        }
        if (/^(加入|移入|删除|收藏|结算|全选|已选|优惠|券|店铺)$/.test(part)) {
          return false;
        }
        return /(颜色|尺码|尺寸|规格|型号|款式|分类|套餐|容量|版本|净含量|适合|口味|编号|码数)/.test(part);
      });

    return fragments.slice(0, 4).join("；");
  }

  function extractSkuFromContainer(container, title, prices) {
    const skuElement = container.querySelector(SKU_SELECTOR);
    const skuText = cleanText(skuElement?.innerText);
    if (skuText && skuText !== title && !PRICE_PATTERN.test(skuText)) {
      return skuText.replace(/^(规格|SKU|属性)[:：]?\s*/i, "");
    }
    return extractSku(container, title, prices);
  }

  function extractImage(container) {
    const cartItemContainer = getCartItemInfoContainer(container) || container;
    const preferred = extractPreferredCartImage(cartItemContainer);
    if (preferred) {
      return preferred;
    }

    const candidates = [...cartItemContainer.querySelectorAll("img[src], img[data-src], img[data-ks-lazyload]")]
      .map((image) => ({
        src: normalizeImageUrl(image.src || image.getAttribute("data-src") || image.getAttribute("data-ks-lazyload")),
        score: scoreImageElement(image)
      }))
      .filter((candidate) => isProductImageUrl(candidate.src))
      .sort((a, b) => b.score - a.score);

    return candidates[0]?.src || "";
  }

  function buildImageContext() {
    const roots = getCartRoots();
    const imageByProductId = new Map();
    const seenRows = new Set();

    for (const root of roots) {
      for (const container of root.querySelectorAll(CART_CONTAINER_SELECTOR)) {
        if (!isLikelyCartItemContainer(container)) {
          continue;
        }
        const cartItemContainer = getCartItemInfoContainer(container) || container;
        if (seenRows.has(cartItemContainer)) {
          continue;
        }
        seenRows.add(cartItemContainer);

        const link = findProductLink(cartItemContainer);
        const productId = getProductKey(link?.href || "");
        const image = extractImage(cartItemContainer);
        if (productId && isProductImageUrl(image) && !imageByProductId.has(productId)) {
          imageByProductId.set(productId, normalizeImageUrl(image));
        }
      }
    }

    return { imageByProductId };
  }

  function findBestImageForItem(container, imageContext, productId, _index) {
    const directImage = normalizeImageUrl(extractImage(container));
    if (isProductImageUrl(directImage)) {
      return directImage;
    }
    if (productId && imageContext.imageByProductId.has(productId)) {
      return imageContext.imageByProductId.get(productId);
    }
    return "";
  }

  function scoreImageElement(image) {
    const src = normalizeImageUrl(image.src || image.getAttribute("data-src") || image.getAttribute("data-ks-lazyload"));
    let score = 0;
    const rect = image.getBoundingClientRect();
    const width = image.naturalWidth || Number(image.getAttribute("width")) || rect.width || 0;
    const height = image.naturalHeight || Number(image.getAttribute("height")) || rect.height || 0;
    const classText = `${image.className || ""} ${image.parentElement?.className || ""}`;

    if (/item|pic|img|photo|thumb/i.test(classText)) {
      score += 30;
    }
    if (width >= 50 && height >= 50) {
      score += 25;
    }
    if (width && height) {
      const ratio = width / height;
      if (ratio >= 0.65 && ratio <= 1.55) {
        score += 35;
      }
      if (ratio > 2.2 || ratio < 0.45) {
        score -= 80;
      }
    }
    if (/tps-|banner|promo|activity|act|coupon|logo|sprite|icon/i.test(src + classText)) {
      score -= 120;
    }
    return score;
  }

  function extractPreferredCartImage(container) {
    const selectors = [
      "a[class*='imageContainer'] img[class*='image']",
      "a[class*='imageContainer'] img[src]",
      "[class*='trade-cart-item-image'] a[class*='imageContainer'] img",
      "[class*='cartImage'] a[class*='imageContainer'] img",
      "[class*='cartItemImage'] img[class*='image']",
      "[class*='cartItemImage'] img[src]"
    ];

    for (const selector of selectors) {
      const image = container.querySelector(selector);
      const src = normalizeImageUrl(image?.src || image?.getAttribute?.("data-src") || image?.getAttribute?.("data-ks-lazyload"));
      if (isProductImageUrl(src)) {
        return src;
      }
    }

    return "";
  }

  function normalizeImageUrl(src) {
    if (!src) {
      return "";
    }
    const fullUrl = src.startsWith("//") ? `${location.protocol}${src}` : src;
    return fullUrl
      .replace(/_\d+x\d+(?:q\d+)?\.(jpg|jpeg|png|webp)$/i, "_400x400.$1")
      .replace(/_\d+x\d+(?:q\d+)?(?=(?:\?|$))/i, "_400x400");
  }

  function isProductImageUrl(src) {
    if (!src || !/(alicdn|taobao|tbcdn)/i.test(src)) {
      return false;
    }
    return !/(_s\.png|\.gif(?:\?|$)|tps-|sprite|icon|logo|loading|avatar|shop|ww)/i.test(src);
  }

  function getProductId(url) {
    const idFromText = String(url || "").match(/[?&]id=(\d+)/);
    return idFromText?.[1] || "";
  }

  function getProductKey(url) {
    try {
      const parsedUrl = new URL(url, location.href);
      const id = parsedUrl.searchParams.get("id") || "";
      const skuId = parsedUrl.searchParams.get("skuId") || parsedUrl.searchParams.get("sku_id") || "";
      const cartId = parsedUrl.searchParams.get("cartId") || parsedUrl.searchParams.get("cart_id") || "";
      const miId = parsedUrl.searchParams.get("mi_id") || "";
      return [id, skuId || cartId || miId].filter(Boolean).join("-");
    } catch {
      const id = getProductId(url);
      const skuMatch = String(url || "").match(/[?&](?:skuId|sku_id|cartId|cart_id|mi_id)=([^&#]+)/);
      return [id, skuMatch?.[1] || ""].filter(Boolean).join("-");
    }
  }

  async function embedItemImages(items) {
    await Promise.all(items.map(async (item) => {
      if (!item.image) {
        item.embeddedImage = "";
        return;
      }

      try {
        const response = await chrome.runtime.sendMessage({
          type: "FETCH_IMAGE_DATA_URL",
          url: item.image
        });
        item.embeddedImage = response?.ok ? response.dataUrl : item.image;
      } catch {
        item.embeddedImage = item.image;
      }
    }));
  }

  function detectSelected(container) {
    const checkbox = container.querySelector("input[type='checkbox']");
    if (checkbox) {
      return checkbox.checked ? "是" : "否";
    }

    const ariaChecked = container.querySelector("[aria-checked]");
    if (ariaChecked) {
      return ariaChecked.getAttribute("aria-checked") === "true" ? "是" : "否";
    }

    return "";
  }

  function cleanText(value) {
    return (value || "").replace(/\u00a0/g, " ").replace(/\s+/g, " ").trim();
  }

  function normalizeUrl(value) {
    try {
      const url = new URL(value);
      return `${url.origin}${url.pathname}?${getProductKey(value) || `id=${url.searchParams.get("id") || ""}`}`;
    } catch {
      return value;
    }
  }

  function downloadXls(items) {
    const headers = ["店铺", "商品名称", "规格", "型号", "床垫尺寸", "单价", "数量", "小计", "商品链接", "图片", "是否选中", "抓取时间"];
    const rows = items.map((item) => [
      item.shop,
      item.title,
      item.sku,
      extractModel(item),
      extractMattressSize(item),
      item.unitPrice,
      item.quantity,
      item.subtotal,
      item.url,
      item.embeddedImage || item.image,
      item.selected,
      item.capturedAt
    ]);

    const html = [
      "<!doctype html>",
      "<html><head><meta charset=\"utf-8\"></head><body>",
      "<table border=\"1\" style=\"border-collapse:collapse;table-layout:fixed;\">",
      `<thead><tr>${headers.map((header) => `<th>${escapeHtml(header)}</th>`).join("")}</tr></thead>`,
      `<tbody>${rows.map((row) => renderExcelRow(row)).join("")}</tbody>`,
      "</table>",
      "</body></html>"
    ].join("");

    const blob = new Blob([html], { type: "application/vnd.ms-excel;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = `淘宝购物车_${formatFileDate(new Date())}.xls`;
    document.body.appendChild(anchor);
    anchor.click();
    anchor.remove();
    URL.revokeObjectURL(url);
  }

  function extractModel(item) {
    return extractModelFromSku(item.sku) || extractModelFromSku(item.title);
  }

  function extractModelFromSku(value) {
    const text = cleanText(value).replace(/[()（）【】[\]{}]/g, " ");
    if (!text) {
      return "";
    }

    const parts = text
      .split(/[\/｜|；;,，\s]+/)
      .map(cleanText)
      .filter(Boolean);

    for (let index = parts.length - 1; index >= 0; index -= 1) {
      const model = normalizeModel(parts[index]);
      if (model) {
        return model;
      }
    }

    const explicitModel = text.match(/(?:型号|款号|编号)[:：\s]*([A-Za-z]{1,4}\d[A-Za-z0-9]{0,12})/i);
    return normalizeModel(explicitModel?.[1] || "");
  }

  function normalizeModel(value) {
    const text = cleanText(value).toUpperCase();
    if (!text || /MM|CM|M\*/i.test(text)) {
      return "";
    }

    const matches = [...text.matchAll(/\b([A-Z]{1,4}\d[A-Z0-9]{0,12})\b/g)];
    if (!matches.length) {
      return "";
    }

    const model = matches[matches.length - 1][1];
    if (/^\d+$/.test(model) || /^(RMB|SKU)$/i.test(model)) {
      return "";
    }
    return model;
  }

  function extractMattressSize(item) {
    const text = cleanText(`${item.sku || ""} ${item.title || ""}`)
      .replace(/[×xX]/g, "*")
      .replace(/\s+/g, "");
    const match = text.match(/(\d{3,4})(?:mm|毫米)?\*(\d{3,4})(?:mm|毫米)?/i);
    if (!match) {
      return "";
    }

    return `${match[1]}mm*${match[2]}mm`;
  }

  function renderExcelRow(row) {
    return `<tr style="height:86px;">${row.map((cell, index) => renderExcelCell(cell, index)).join("")}</tr>`;
  }

  function renderExcelCell(cell, index) {
    if (index === 9) {
      return `<td style="width:92px;text-align:center;vertical-align:middle;">${renderImageCell(cell)}</td>`;
    }
    return `<td style="vertical-align:middle;">${escapeHtml(cell)}</td>`;
  }

  function renderImageCell(src) {
    if (!src) {
      return "";
    }
    return `<img src="${escapeHtml(src)}" width="72" height="72" style="width:72px;height:72px;object-fit:contain;">`;
  }

  function escapeHtml(value) {
    return String(value ?? "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#39;");
  }

  function formatDateTime(date) {
    const pad = (number) => String(number).padStart(2, "0");
    return `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())} ${pad(date.getHours())}:${pad(date.getMinutes())}:${pad(date.getSeconds())}`;
  }

  function formatFileDate(date) {
    const pad = (number) => String(number).padStart(2, "0");
    return `${date.getFullYear()}${pad(date.getMonth() + 1)}${pad(date.getDate())}_${pad(date.getHours())}${pad(date.getMinutes())}${pad(date.getSeconds())}`;
  }

  function delay(ms) {
    return new Promise((resolve) => setTimeout(resolve, ms));
  }
})();
