const SHEET_API_URL = "https://script.google.com/macros/s/AKfycbzSTYSBsd2umh4G__iIzVoWDeEDTHvHLX--VkS4JIbduGIo_vyQ8LcSTzbJJHhyHl2NIQ/exec";
const SITE_TITLE = "沖繩行程地圖";
const DAY_COLORS = ["#3b82f6", "#f97316", "#10b981", "#8b5cf6", "#ec4899", "#14b8a6", "#f59e0b", "#ef4444"];

const state = {
  allItems: [],
  groupedItems: {},
  orderedDays: [],
  currentDay: null,
  selectedId: null,
  map: null,
  markersLayer: null,
  routeLayer: null,
  markerMap: new Map(),
};

const defaultCenter = [26.2124, 127.6809];

const mapEl = document.getElementById("map");
const dayTabsEl = document.getElementById("day-tabs");
const itineraryListEl = document.getElementById("itinerary-list");
const currentDayTitleEl = document.getElementById("current-day-title");
const heroDateEl = document.getElementById("hero-date");
const heroTitleEl = document.getElementById("hero-title");
const fitDayBtn = document.getElementById("fit-day-btn");
const cardTemplate = document.getElementById("itinerary-card-template");
const addModalEl = document.getElementById("add-modal");
const addFormEl = document.getElementById("add-form");
const addFormMessageEl = document.getElementById("add-form-message");
const openAddModalBtn = document.getElementById("open-add-modal-btn");
const closeAddModalBtn = document.getElementById("close-add-modal-btn");
const cancelAddBtn = document.getElementById("cancel-add-btn");
const submitAddBtn = document.getElementById("submit-add-btn");

bootstrap();

async function bootstrap() {
  initMap();
  bindEvents();

  try {
    const remoteItems = await loadRemoteSheetData();
    state.allItems = normalizeItems(remoteItems);
    state.groupedItems = groupItemsByDay(state.allItems);
    state.orderedDays = sortDays(Object.keys(state.groupedItems));

    if (!state.orderedDays.length) {
      renderNoData();
      return;
    }

    state.currentDay = state.orderedDays[0];
    state.selectedId = getDefaultSelectedId(state.currentDay);
    renderDayTabs();
    renderCurrentDay({ fitBounds: true });
  } catch (error) {
    console.error(error);
    renderLoadError(error);
  }
}

function initMap() {
  state.map = L.map(mapEl, { zoomControl: false, preferCanvas: true }).setView(defaultCenter, 10);
  L.tileLayer("https://tile.openstreetmap.org/{z}/{x}/{y}.png", {
    maxZoom: 19,
    attribution: '&copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a>',
  }).addTo(state.map);
  L.control.zoom({ position: "bottomright" }).addTo(state.map);
  state.markersLayer = L.layerGroup().addTo(state.map);
  state.routeLayer = L.layerGroup().addTo(state.map);
}

function bindEvents() {
  fitDayBtn?.addEventListener("click", () => fitCurrentDayBounds({ animate: true }));
  window.addEventListener("resize", () => state.map?.invalidateSize());

  openAddModalBtn?.addEventListener("click", openAddModal);
  closeAddModalBtn?.addEventListener("click", closeAddModal);
  cancelAddBtn?.addEventListener("click", closeAddModal);
  addFormEl?.addEventListener("submit", handleAddSubmit);
  addModalEl?.addEventListener("click", (event) => {
    if (event.target === addModalEl) closeAddModal();
  });
}

function openAddModal() {
  if (!addModalEl || !addFormEl) return;
  clearFormMessage();
  presetFormDefaults();
  addModalEl.showModal();
}

function closeAddModal() {
  addModalEl?.close();
}

function presetFormDefaults() {
  if (!addFormEl) return;
  const today = state.currentDay || "day1";
  const items = getItemsForDay(today);
  const nextOrder = (items.reduce((max, item) => Math.max(max, Number(item.order) || 0), 0) || 0) + 1;
  const firstItemDate = items[0]?.date || formatDateInput(new Date());
  addFormEl.elements.day.value = today;
  addFormEl.elements.order.value = String(nextOrder);
  addFormEl.elements.date.value = firstItemDate;
  if (!addFormEl.elements.status.value) addFormEl.elements.status.value = "planned";
}

function requestJsonp(params = {}, { timeout = 12000 } = {}) {
  return new Promise((resolve, reject) => {
    const callbackName = `__sheetCallback_${Date.now()}_${Math.random().toString(36).slice(2)}`;
    const timeoutId = setTimeout(() => {
      cleanup();
      reject(new Error("Google Sheets 連線逾時"));
    }, timeout);

    const url = new URL(SHEET_API_URL);
    Object.entries(params).forEach(([key, value]) => {
      if (value === undefined || value === null || value === "") return;
      url.searchParams.set(key, String(value));
    });
    url.searchParams.set("prefix", callbackName);

    const script = document.createElement("script");
    script.src = url.toString();
    script.async = true;
    script.onerror = () => {
      cleanup();
      reject(new Error("無法載入 Google Sheets JSONP"));
    };

    function cleanup() {
      clearTimeout(timeoutId);
      delete window[callbackName];
      script.remove();
    }

    window[callbackName] = (payload) => {
      cleanup();
      resolve(payload);
    };

    document.body.appendChild(script);
  });
}

async function loadRemoteSheetData() {
  const payload = await requestJsonp();
  if (!payload || payload.ok !== true || !Array.isArray(payload.items)) {
    throw new Error(payload?.message || "Google Sheets 回傳格式不正確");
  }
  return payload.items;
}

function normalizeItems(items) {
  return items.map((item, index) => {
    const parsed = parseCoords(item.coords, item.lat, item.lng);
    return {
      id: item.id || `${item.day || "day0"}-${item.order || index + 1}`,
      date: item.date || "",
      day: item.day || "day0",
      order: Number(item.order) || index + 1,
      place_name: item.place_name || "",
      address: item.address || "",
      coords: item.coords || "",
      lat: parsed.lat,
      lng: parsed.lng,
      start_time: item.start_time || "",
      end_time: item.end_time || "",
      note: item.note || "",
      category: item.category || inferCategory(item),
      google_maps_url: item.google_maps_url || buildGoogleMapsUrl(item, parsed),
      image_url: item.image_url || "",
      status: item.status || "",
      hasCoordinates: Number.isFinite(parsed.lat) && Number.isFinite(parsed.lng),
    };
  }).sort((a, b) => {
    if (a.day !== b.day) return extractDayNumber(a.day) - extractDayNumber(b.day);
    return a.order - b.order;
  });
}

function parseCoords(coords, latRaw, lngRaw) {
  if (typeof coords === "string" && coords.includes(",")) {
    const [latStr, lngStr] = coords.split(",").map((v) => v.trim());
    const lat = parseFloat(latStr);
    const lng = parseFloat(lngStr);
    return { lat, lng };
  }
  return { lat: parseFloat(latRaw), lng: parseFloat(lngRaw) };
}

function inferCategory(item) {
  const text = `${item.place_name || ""} ${item.note || ""}`;
  if (/機場|航班|車站|交通|單軌/.test(text)) return "交通";
  if (/飯店|ホテル|hotel|リッチモンド/i.test(text)) return "住宿";
  if (/咖啡|燒肉|麵|うどん|壽司|餐|吃|冰淇淋/.test(text)) return "餐飲";
  if (/海洋|水族館|公園|植物園/.test(text)) return "景點";
  return item.status || "待確認";
}

function buildGoogleMapsUrl(item, parsed) {
  if (Number.isFinite(parsed.lat) && Number.isFinite(parsed.lng)) {
    return `https://www.google.com/maps/search/?api=1&query=${parsed.lat},${parsed.lng}`;
  }
  const q = [item.place_name, item.address].filter(Boolean).join(" ").trim();
  return q ? `https://www.google.com/maps/search/?api=1&query=${encodeURIComponent(q)}` : "";
}

function groupItemsByDay(items) {
  return items.reduce((acc, item) => {
    (acc[item.day] ||= []).push(item);
    return acc;
  }, {});
}

function sortDays(days) {
  return [...days].sort((a, b) => extractDayNumber(a) - extractDayNumber(b));
}

function extractDayNumber(day) {
  const match = String(day).match(/(\d+)/);
  return match ? Number(match[1]) : Number.MAX_SAFE_INTEGER;
}

function getItemsForDay(day) {
  return state.groupedItems[day] || [];
}

function getDefaultSelectedId(day) {
  const items = getItemsForDay(day);
  const firstMapped = items.find((item) => item.hasCoordinates);
  return firstMapped?.id || items[0]?.id || null;
}

function renderDayTabs() {
  dayTabsEl.innerHTML = "";
  state.orderedDays.forEach((day) => {
    const items = getItemsForDay(day);
    const mappedCount = items.filter((item) => item.hasCoordinates).length;
    const button = document.createElement("button");
    button.type = "button";
    button.className = `day-tab${day === state.currentDay ? " active" : ""}`;
    button.innerHTML = `
      <span class="day-tab-label">Day ${extractDayNumber(day)}</span>
      <span class="day-tab-meta">
        <span>${escapeHtml(items[0]?.date || "—")}</span>
        <span>・</span>
        <span>${mappedCount}/${items.length}</span>
      </span>
    `;
    button.addEventListener("click", () => {
      state.currentDay = day;
      state.selectedId = getDefaultSelectedId(day);
      renderDayTabs();
      renderCurrentDay({ fitBounds: true });
    });
    dayTabsEl.appendChild(button);
  });
}

function renderCurrentDay({ fitBounds = false } = {}) {
  const items = getItemsForDay(state.currentDay);
  const dayNumber = extractDayNumber(state.currentDay);
  currentDayTitleEl.textContent = `Day ${dayNumber} 行程`;
  heroDateEl.textContent = items[0]?.date || "";
  heroTitleEl.textContent = `Day ${dayNumber}`;
  renderList(items);
  renderMap(items, { fitBounds });
}

function renderList(items) {
  itineraryListEl.innerHTML = "";
  if (!items.length) {
    const empty = document.createElement("div");
    empty.className = "empty-state-card";
    empty.textContent = "這一天目前還沒有行程。";
    itineraryListEl.appendChild(empty);
    return;
  }

  items.forEach((item) => {
    const fragment = cardTemplate.content.cloneNode(true);
    const card = fragment.querySelector(".itinerary-card");
    const cardOrder = fragment.querySelector(".card-order");
    const cardTime = fragment.querySelector(".card-time");
    const cardCategory = fragment.querySelector(".card-category");
    const cardMapState = fragment.querySelector(".card-map-state");
    const cardTitle = fragment.querySelector(".card-title");
    const cardAddress = fragment.querySelector(".card-address");
    const cardNote = fragment.querySelector(".card-note");
    const focusBtn = fragment.querySelector(".focus-btn");
    const mapsBtn = fragment.querySelector(".maps-btn");
    const media = fragment.querySelector(".card-media");
    const img = fragment.querySelector(".card-image");

    card.dataset.id = item.id;
    if (item.id === state.selectedId) card.classList.add("active");
    cardOrder.textContent = item.order;
    cardTime.textContent = formatTimeRange(item.start_time, item.end_time) || "未排時間";
    cardCategory.textContent = formatCategory(item.category);
    cardMapState.textContent = item.hasCoordinates ? "已定位" : "待補座標";
    if (!item.hasCoordinates) cardMapState.classList.add("muted");
    cardTitle.textContent = item.place_name || "未命名地點";

    const sameAddress = normalizeCompareText(item.place_name) && normalizeCompareText(item.place_name) === normalizeCompareText(item.address);
    if (item.address && !sameAddress) {
      cardAddress.textContent = item.address;
      cardAddress.hidden = false;
    } else {
      cardAddress.hidden = true;
    }

    if (item.note) {
      cardNote.textContent = item.note;
      cardNote.hidden = false;
    } else {
      cardNote.hidden = true;
    }

    if (item.image_url) {
      media.hidden = false;
      img.src = item.image_url;
      img.alt = item.place_name ? `${item.place_name} 圖片` : "行程圖片";
      img.addEventListener("error", () => {
        media.hidden = true;
        card.classList.remove("has-image");
      }, { once: true });
      card.classList.add("has-image");
    } else {
      media.hidden = true;
      img.removeAttribute("src");
      card.classList.remove("has-image");
    }

    focusBtn.disabled = !item.hasCoordinates;
    mapsBtn.href = item.google_maps_url || "#";
    mapsBtn.setAttribute("aria-disabled", item.google_maps_url ? "false" : "true");
    if (!item.google_maps_url) mapsBtn.removeAttribute("href");

    card.addEventListener("click", () => selectItem(item.id, { panTo: true }));
    card.addEventListener("keydown", (event) => {
      if (event.key === "Enter" || event.key === " ") {
        event.preventDefault();
        selectItem(item.id, { panTo: true });
      }
    });
    focusBtn.addEventListener("click", (event) => {
      event.stopPropagation();
      selectItem(item.id, { panTo: true, zoom: 15 });
    });

    itineraryListEl.appendChild(fragment);
  });
}

function renderMap(items, { fitBounds = false } = {}) {
  state.markersLayer.clearLayers();
  state.routeLayer.clearLayers();
  state.markerMap.clear();

  const mapped = items.filter((item) => item.hasCoordinates);
  if (!mapped.length) {
    state.map.setView(defaultCenter, 10);
    return;
  }

  const color = getDayColor(state.currentDay);
  mapped.forEach((item) => {
    const marker = createMarker(item, color);
    marker.addTo(state.markersLayer);
    state.markerMap.set(item.id, marker);
  });

  if (mapped.length >= 2) {
    const coords = mapped.map((item) => [item.lat, item.lng]);
    L.polyline(coords, {
      color,
      weight: 4,
      opacity: 0.8,
      dashArray: "1 11",
      lineCap: "round",
    }).addTo(state.routeLayer);
  }

  highlightSelectedMarker();
  if (fitBounds) fitCurrentDayBounds({ animate: false });
}

function createMarker(item, color) {
  const iconHtml = `
    <div class="marker-shell" style="--marker-color:${color}">
      <div class="marker-badge">${item.order}</div>
    </div>
  `;
  const marker = L.marker([item.lat, item.lng], {
    icon: L.divIcon({
      html: iconHtml,
      className: "custom-marker-icon",
      iconSize: [38, 50],
      iconAnchor: [19, 44],
      popupAnchor: [0, -40],
    })
  });

  const sameAddress = normalizeCompareText(item.place_name) && normalizeCompareText(item.place_name) === normalizeCompareText(item.address);
  const addressLine = item.address && !sameAddress ? `<div>${escapeHtml(item.address)}</div>` : "";
  const noteLine = item.note ? `<div>${escapeHtml(item.note)}</div>` : "";
  const mapsLink = item.google_maps_url ? `<div style="margin-top:8px;"><a href="${escapeAttribute(item.google_maps_url)}" target="_blank" rel="noreferrer noopener">在 Google Maps 開啟</a></div>` : "";

  marker.bindPopup(`
    <div>
      <strong>${escapeHtml(item.place_name || "未命名地點")}</strong>
      <div>${escapeHtml(formatTimeRange(item.start_time, item.end_time) || "未排時間")}</div>
      ${addressLine}
      ${noteLine}
      ${mapsLink}
    </div>
  `);

  marker.on("click", () => selectItem(item.id, { panTo: false }));
  return marker;
}

function selectItem(id, { panTo = false, zoom = null } = {}) {
  state.selectedId = id;
  const items = getItemsForDay(state.currentDay);
  const selected = items.find((item) => item.id === id);
  renderList(items);
  highlightSelectedMarker();
  if (selected && panTo && selected.hasCoordinates) {
    state.map.flyTo([selected.lat, selected.lng], zoom || Math.max(state.map.getZoom(), 14), { duration: 0.6 });
    const marker = state.markerMap.get(id);
    marker?.openPopup();
  }
}

function highlightSelectedMarker() {
  state.markerMap.forEach((marker, id) => {
    const markerShell = marker.getElement()?.querySelector(".marker-shell");
    markerShell?.classList.toggle("is-active", id === state.selectedId);
  });
}

function fitCurrentDayBounds({ animate = true } = {}) {
  const items = getItemsForDay(state.currentDay).filter((item) => item.hasCoordinates);
  if (!items.length) return;
  const bounds = L.latLngBounds(items.map((item) => [item.lat, item.lng]));
  state.map.fitBounds(bounds, { padding: [48, 48], animate, maxZoom: 14 });
}

function renderNoData() {
  heroDateEl.textContent = "";
  heroTitleEl.textContent = SITE_TITLE;
  currentDayTitleEl.textContent = "目前沒有行程資料";
  itineraryListEl.innerHTML = '<div class="empty-state-card">目前 Google Sheets 沒有可用資料。</div>';
}

function renderLoadError(error) {
  heroDateEl.textContent = "";
  heroTitleEl.textContent = SITE_TITLE;
  currentDayTitleEl.textContent = "資料載入失敗";
  itineraryListEl.innerHTML = `<div class="empty-state-card">${escapeHtml(error.message || String(error))}</div>`;
}

async function handleAddSubmit(event) {
  event.preventDefault();
  clearFormMessage();
  const formData = new FormData(addFormEl);
  const payload = Object.fromEntries(formData.entries());

  const validationError = validateFormPayload(payload);
  if (validationError) {
    showFormMessage(validationError, "error");
    return;
  }

  submitAddBtn.disabled = true;
  showFormMessage("正在送出到 Google Sheets…", "info");

  const cleanItem = {
    date: payload.date.trim(),
    day: payload.day.trim(),
    order: String(payload.order).trim(),
    place_name: payload.place_name.trim(),
    address: payload.address.trim(),
    coords: payload.coords.trim(),
    start_time: payload.start_time.trim(),
    end_time: payload.end_time.trim(),
    note: payload.note.trim(),
    category: payload.category.trim(),
    google_maps_url: payload.google_maps_url.trim(),
    image_url: payload.image_url.trim(),
    status: payload.status.trim() || "planned",
  };

  const requestPayload = {
    action: "add_item",
    admin_key: payload.admin_key,
    item: cleanItem,
  };

  try {
    const payload = await requestJsonp({
      action: "add_item",
      admin_key: requestPayload.admin_key,
      date: cleanItem.date,
      day: cleanItem.day,
      order: cleanItem.order,
      place_name: cleanItem.place_name,
      address: cleanItem.address,
      coords: cleanItem.coords,
      start_time: cleanItem.start_time,
      end_time: cleanItem.end_time,
      note: cleanItem.note,
      category: cleanItem.category,
      google_maps_url: cleanItem.google_maps_url,
      image_url: cleanItem.image_url,
      status: cleanItem.status,
    }, { timeout: 15000 });

    if (!payload || payload.ok !== true) {
      throw new Error(payload?.message || "新增失敗，請稍後再試。");
    }

    showFormMessage(payload.message || "已新增到 Google Sheets。", "success");

    await refreshDataAfterAdd(cleanItem.day);
    submitAddBtn.disabled = false;
    addFormEl.reset();
    setTimeout(() => closeAddModal(), 700);
  } catch (error) {
    console.error(error);
    showFormMessage(error.message || "送出失敗，請檢查 Apps Script 權限設定。", "error");
    submitAddBtn.disabled = false;
  }
}

async function refreshDataAfterAdd(preferredDay) {
  const remoteItems = await loadRemoteSheetData();
  state.allItems = normalizeItems(remoteItems);
  state.groupedItems = groupItemsByDay(state.allItems);
  state.orderedDays = sortDays(Object.keys(state.groupedItems));
  state.currentDay = state.groupedItems[preferredDay] ? preferredDay : state.orderedDays[0] || null;
  state.selectedId = getDefaultSelectedId(state.currentDay);
  renderDayTabs();
  renderCurrentDay({ fitBounds: true });
}

function validateFormPayload(payload) {
  if (!payload.date) return "請填日期。";
  if (!/^\d{4}-\d{2}-\d{2}$/.test(payload.date.trim())) return "日期格式需為 YYYY-MM-DD。";
  if (!payload.day || !/^day\d+$/i.test(payload.day.trim())) return "day 欄請填像 day1、day2。";
  if (!payload.order || !/^\d+$/.test(String(payload.order).trim())) return "順序必須是正整數。";
  if (!payload.place_name || !payload.place_name.trim()) return "請填地點名稱。";
  if (payload.coords && !/^\s*-?\d+(\.\d+)?\s*,\s*-?\d+(\.\d+)?\s*$/.test(payload.coords)) return "coords 格式請填「緯度, 經度」。";
  if (payload.start_time && !/^\d{2}:\d{2}$/.test(payload.start_time)) return "開始時間格式需為 HH:MM。";
  if (payload.end_time && !/^\d{2}:\d{2}$/.test(payload.end_time)) return "結束時間格式需為 HH:MM。";
  if (payload.google_maps_url && !/^https?:\/\//i.test(payload.google_maps_url.trim())) return "Google Maps 連結必須是 http 或 https 開頭。";
  if (payload.image_url && !/^https?:\/\//i.test(payload.image_url.trim())) return "圖片網址必須是 http 或 https 開頭。";
  if (!payload.admin_key || !payload.admin_key.trim()) return "請輸入管理密碼。";
  return "";
}

function showFormMessage(message, type = "info") {
  addFormMessageEl.textContent = message;
  addFormMessageEl.className = `form-message ${type}`;
  addFormMessageEl.classList.remove("is-hidden");
}

function clearFormMessage() {
  addFormMessageEl.textContent = "";
  addFormMessageEl.className = "form-message is-hidden";
}

function formatTimeRange(start, end) {
  const s = start?.trim();
  const e = end?.trim();
  if (s && e) return `${s}–${e}`;
  return s || e || "";
}

function formatCategory(category) {
  const value = String(category || "").trim();
  if (!value) return "待確認";
  return value;
}

function getDayColor(day) {
  const index = Math.max(0, extractDayNumber(day) - 1) % DAY_COLORS.length;
  return DAY_COLORS[index];
}

function formatDateInput(dateObj) {
  const y = dateObj.getFullYear();
  const m = String(dateObj.getMonth() + 1).padStart(2, "0");
  const d = String(dateObj.getDate()).padStart(2, "0");
  return `${y}-${m}-${d}`;
}

function normalizeCompareText(value) {
  return String(value || "").trim().toLowerCase();
}

function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function escapeAttribute(value) {
  return escapeHtml(value).replace(/`/g, "&#96;");
}
