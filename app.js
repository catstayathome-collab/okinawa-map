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
  sheetVh: 54,
  isEditMode: false,
};

const defaultCenter = [26.2124, 127.6809];
const mapEl = document.getElementById("map");
const dayTabsEl = document.getElementById("day-tabs");
const itineraryListEl = document.getElementById("itinerary-list");
const currentDayTitleEl = document.getElementById("current-day-title");
const currentDayMetaEl = document.getElementById("current-day-meta");
const heroDateEl = document.getElementById("hero-date");
const heroTitleEl = document.getElementById("hero-title");
const heroSummaryEl = document.getElementById("hero-summary");
const fitDayBtn = document.getElementById("fit-day-btn");
const mapEmptyStateEl = document.getElementById("map-empty-state");
const sheetHandleArea = document.getElementById("sheet-handle-area");
const cardTemplate = document.getElementById("itinerary-card-template");
const openAddModalBtn = document.getElementById("open-add-modal-btn");
const toggleEditModeBtn = document.getElementById("toggle-edit-mode-btn");
const itemModalEl = document.getElementById("item-modal");
const itemFormEl = document.getElementById("item-form");
const itemFormMessageEl = document.getElementById("item-form-message");
const closeItemModalBtn = document.getElementById("close-item-modal-btn");
const cancelItemBtn = document.getElementById("cancel-item-btn");
const submitItemBtn = document.getElementById("submit-item-btn");
const formKickerEl = document.getElementById("form-kicker");
const formTitleEl = document.getElementById("form-title");
const autofillPlaceBtn = document.getElementById("autofill-place-btn");

bootstrap();

async function bootstrap() {
  initMap();
  bindEvents();
  setupBottomSheet();

  try {
    await reloadFromRemote();
  } catch (error) {
    console.error(error);
    renderLoadError(error);
  }
}

function initMap() {
  state.map = L.map(mapEl, {
    zoomControl: false,
    preferCanvas: true,
  }).setView(defaultCenter, 12);

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

  openAddModalBtn?.addEventListener("click", () => openItemModal("add"));
  toggleEditModeBtn?.addEventListener("click", toggleEditMode);
  closeItemModalBtn?.addEventListener("click", closeItemModal);
  cancelItemBtn?.addEventListener("click", closeItemModal);
  itemFormEl?.addEventListener("submit", handleItemSubmit);
  autofillPlaceBtn?.addEventListener("click", handleAutofillPlace);
  itemModalEl?.addEventListener("click", (event) => {
    if (event.target === itemModalEl) closeItemModal();
  });
}

function requestJsonp(params = {}, { timeout = 12000, includeEmpty = false } = {}) {
  return new Promise((resolve, reject) => {
    const callbackName = `__sheetCallback_${Date.now()}_${Math.random().toString(36).slice(2)}`;
    const timeoutId = setTimeout(() => {
      cleanup();
      reject(new Error("Google Sheets 連線逾時"));
    }, timeout);

    const url = new URL(SHEET_API_URL);
    Object.entries(params).forEach(([key, value]) => {
      if (!includeEmpty && (value === undefined || value === null || value === "")) return;
      url.searchParams.set(key, value == null ? "" : String(value));
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
  const payload = await requestJsonp({}, { timeout: 12000 });
  if (!payload || payload.ok !== true || !Array.isArray(payload.items)) {
    throw new Error(payload?.message || "Google Sheets 回傳格式不正確");
  }
  return payload.items;
}

async function reloadFromRemote(preferredDay = null, preferredId = null, fitBounds = false) {
  const items = await loadRemoteSheetData();
  state.allItems = normalizeItems(items);
  state.groupedItems = groupItemsByDay(state.allItems);
  state.orderedDays = sortDays(Object.keys(state.groupedItems));

  if (!state.orderedDays.length) {
    renderNoData();
    return;
  }

  const nextDay = preferredDay && state.groupedItems[preferredDay] ? preferredDay : state.currentDay;
  state.currentDay = nextDay && state.groupedItems[nextDay] ? nextDay : state.orderedDays[0];

  const itemsForDay = getCurrentItems();
  state.selectedId = preferredId && itemsForDay.some((item) => item.id === preferredId)
    ? preferredId
    : getDefaultSelectedId(state.currentDay);

  renderDayTabs();
  renderCurrentDay({ fitBounds });
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
      coords: item.coords || formatCoords(parsed.lat, parsed.lng),
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
    return { lat: parseFloat(latStr), lng: parseFloat(lngStr) };
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
  if (item.google_maps_url) return item.google_maps_url;
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
  const match = String(day || "").match(/day(\d+)/i);
  return match ? Number(match[1]) : Number.MAX_SAFE_INTEGER;
}

function getCurrentItems() {
  return state.groupedItems[state.currentDay] || [];
}

function getSelectedItem() {
  return getCurrentItems().find((item) => item.id === state.selectedId) || null;
}

function getDefaultSelectedId(day) {
  const items = state.groupedItems[day] || [];
  const firstMapped = items.find((item) => item.hasCoordinates);
  return firstMapped?.id || items[0]?.id || null;
}

function getDayColor(dayKey) {
  const index = Math.max(0, extractDayNumber(dayKey) - 1) % DAY_COLORS.length;
  return DAY_COLORS[index];
}

function renderDayTabs() {
  dayTabsEl.innerHTML = "";
  const count = state.orderedDays.length;
  dayTabsEl.classList.toggle("is-scroll", count > 3);

  state.orderedDays.forEach((dayKey) => {
    const items = state.groupedItems[dayKey] || [];
    const mappedCount = items.filter((item) => item.hasCoordinates).length;
    const button = document.createElement("button");
    button.type = "button";
    button.className = `day-tab ${dayKey === state.currentDay ? "active" : ""}`;

    if (count <= 3) {
      const gap = 8;
      button.style.flex = `0 0 calc((100% - ${(count - 1) * gap}px) / ${count})`;
      button.style.maxWidth = "none";
      button.style.minWidth = "0";
    } else {
      button.style.flex = "0 0 132px";
      button.style.maxWidth = "132px";
    }

    button.innerHTML = `
      <span class="day-tab-label">Day ${extractDayNumber(dayKey)}</span>
      <span class="day-tab-meta">
        <span class="day-tab-date">${escapeHtml(items[0]?.date || "")}</span>
        <span class="day-tab-divider">・</span>
        <span class="day-tab-ratio">${mappedCount}/${items.length}</span>
      </span>
    `;

    button.addEventListener("click", () => {
      if (state.currentDay === dayKey) return;
      state.currentDay = dayKey;
      state.selectedId = getDefaultSelectedId(dayKey);
      renderDayTabs();
      renderCurrentDay({ fitBounds: true });
      closeAnyOpenPopup();
    });

    dayTabsEl.appendChild(button);
  });

  const activeBtn = dayTabsEl.querySelector(".day-tab.active");
  activeBtn?.scrollIntoView({ block: "nearest", inline: "nearest", behavior: "smooth" });
}

function renderCurrentDay({ fitBounds = false } = {}) {
  const items = getCurrentItems();
  if (!items.length) return;

  const selected = getSelectedItem() || items[0];
  if (!state.selectedId && selected) state.selectedId = selected.id;

  const color = getDayColor(state.currentDay);
  document.documentElement.style.setProperty("--day-color", color);
  document.documentElement.style.setProperty("--day-color-soft", hexToRgba(color, 0.14));

  if (currentDayTitleEl) currentDayTitleEl.textContent = `Day ${extractDayNumber(state.currentDay)} 行程`;
  if (currentDayMetaEl) currentDayMetaEl.textContent = "";
  if (heroDateEl) heroDateEl.textContent = items[0]?.date || "";
  if (heroTitleEl) heroTitleEl.textContent = `Day ${extractDayNumber(state.currentDay)}`;
  if (heroSummaryEl) heroSummaryEl.textContent = "";

  syncEditModeUi();
  renderList(items);
  renderMap(items, { fitBounds });
}

function normalizeCompareText(value) {
  return String(value || "").replace(/\s+/g, "").trim().toLowerCase();
}

function getDisplayAddress(item) {
  const title = item.place_name || "";
  const address = item.address || "";
  if (!address) return "";
  return normalizeCompareText(title) === normalizeCompareText(address) ? "" : address;
}

function renderList(items) {
  itineraryListEl.innerHTML = "";

  if (!items.length) {
    itineraryListEl.innerHTML = `<div class="empty-state-card">這一天還沒有行程資料。</div>`;
    return;
  }

  const fragment = document.createDocumentFragment();

  items.forEach((item) => {
    const node = cardTemplate.content.firstElementChild.cloneNode(true);
    node.dataset.id = item.id;

    const orderEl = node.querySelector(".card-order");
    const timeEl = node.querySelector(".card-time");
    const categoryEl = node.querySelector(".card-category");
    const mapStateEl = node.querySelector(".card-map-state");
    const titleEl = node.querySelector(".card-title");
    const addressEl = node.querySelector(".card-address");
    const noteEl = node.querySelector(".card-note");
    const focusBtn = node.querySelector(".focus-btn");
    const mapsBtn = node.querySelector(".maps-btn");
    const mediaEl = node.querySelector(".card-media");
    const imageEl = node.querySelector(".card-image");

    orderEl.textContent = item.order;
    timeEl.textContent = formatTimeRange(item.start_time, item.end_time) || "時間未定";
    categoryEl.textContent = formatCategory(item.category);
    mapStateEl.textContent = item.hasCoordinates ? "已定位" : "待補座標";
    if (!item.hasCoordinates) mapStateEl.classList.add("muted");

    const title = item.place_name || "未命名地點";
    const displayAddress = getDisplayAddress(item);
    titleEl.textContent = title;
    if (displayAddress) {
      addressEl.textContent = displayAddress;
      addressEl.hidden = false;
    } else {
      addressEl.hidden = true;
      addressEl.textContent = "";
    }
    noteEl.textContent = item.note || "尚未填寫備註";

    const imageUrl = String(item.image_url || "").trim();
    if (imageUrl) {
      node.classList.add("has-image");
      node.classList.remove("no-image");
      mediaEl.hidden = false;
      imageEl.src = imageUrl;
      imageEl.alt = `${title} 圖片`;
      imageEl.onerror = () => {
        mediaEl.hidden = true;
        node.classList.remove("has-image");
        node.classList.add("no-image");
        imageEl.removeAttribute("src");
      };
    } else {
      mediaEl.hidden = true;
      node.classList.remove("has-image");
      node.classList.add("no-image");
      imageEl.removeAttribute("src");
      imageEl.alt = "";
    }

    if (state.isEditMode) {
      focusBtn.textContent = "編輯";
      focusBtn.classList.add("editing");
      focusBtn.addEventListener("click", (event) => {
        event.stopPropagation();
        openItemModal("edit", item);
      });

      mapsBtn.textContent = "刪除";
      mapsBtn.classList.add("danger");
      mapsBtn.removeAttribute("href");
      mapsBtn.setAttribute("role", "button");
      mapsBtn.addEventListener("click", async (event) => {
        event.preventDefault();
        event.stopPropagation();
        await handleDelete(item);
      });
    } else {
      if (item.hasCoordinates) {
        focusBtn.textContent = "定位";
        focusBtn.classList.remove("editing");
        focusBtn.disabled = false;
        focusBtn.addEventListener("click", (event) => {
          event.stopPropagation();
          focusItem(item.id, { openPopup: true, scrollCard: false, flyTo: true });
        });
      } else {
        focusBtn.textContent = "未定位";
        focusBtn.disabled = true;
      }

      mapsBtn.textContent = item.google_maps_url ? "Google Maps" : "無連結";
      mapsBtn.classList.remove("danger");
      if (item.google_maps_url) {
        mapsBtn.href = item.google_maps_url;
        mapsBtn.removeAttribute("aria-disabled");
      } else {
        mapsBtn.removeAttribute("href");
        mapsBtn.setAttribute("aria-disabled", "true");
      }
    }

    node.addEventListener("click", () => {
      if (state.isEditMode) return;
      focusItem(item.id, { openPopup: true, scrollCard: false, flyTo: true });
    });

    node.addEventListener("keydown", (event) => {
      if (event.key === "Enter" || event.key === " ") {
        event.preventDefault();
        if (state.isEditMode) return;
        focusItem(item.id, { openPopup: true, scrollCard: false, flyTo: true });
      }
    });

    if (item.id === state.selectedId) node.classList.add("active");
    fragment.appendChild(node);
  });

  itineraryListEl.appendChild(fragment);
}

function renderMap(items, { fitBounds = false } = {}) {
  state.markersLayer.clearLayers();
  state.routeLayer.clearLayers();
  state.markerMap.clear();

  const mappable = items.filter((item) => item.hasCoordinates);
  const dayColor = getDayColor(state.currentDay);

  if (!mappable.length) {
    if (mapEmptyStateEl) mapEmptyStateEl.classList.add("is-hidden");
    state.map.setView(defaultCenter, 10);
    return;
  }

  if (mappable.length >= 2) {
    const routePoints = mappable.map((item) => [item.lat, item.lng]);
    L.polyline(routePoints, {
      color: "#ffffff",
      weight: 10,
      opacity: 0.58,
      lineCap: "round",
      lineJoin: "round",
    }).addTo(state.routeLayer);

    L.polyline(routePoints, {
      color: dayColor,
      weight: 5.5,
      opacity: 0.96,
      lineCap: "round",
      lineJoin: "round",
      dashArray: "1 10",
    }).addTo(state.routeLayer);
  }

  mappable.forEach((item) => {
    const marker = L.marker([item.lat, item.lng], {
      icon: createNumberedIcon(item.order, dayColor),
      keyboard: true,
      title: item.place_name || `景點 ${item.order}`,
    });

    marker.bindPopup(buildPopupHtml(item), { closeButton: false, autoPanPadding: [24, 160] });
    marker.on("click", () => {
      focusItem(item.id, { openPopup: false, scrollCard: true, flyTo: false });
      marker.openPopup();
    });

    marker.addTo(state.markersLayer);
    state.markerMap.set(item.id, marker);
  });

  syncMarkerActiveState();
  if (fitBounds) fitCurrentDayBounds({ animate: false });
}

function buildPopupHtml(item) {
  const title = escapeHtml(item.place_name || `景點 ${item.order}`);
  const address = getDisplayAddress(item);
  const addressHtml = address ? `<br><span>${escapeHtml(address)}</span>` : "";
  const note = escapeHtml(item.note || "尚未填寫備註");
  const time = escapeHtml(formatTimeRange(item.start_time, item.end_time) || "時間未定");
  const maps = item.google_maps_url
    ? `<div style="margin-top:10px;"><a href="${item.google_maps_url}" target="_blank" rel="noreferrer noopener">在 Google Maps 開啟</a></div>`
    : "";
  return `<div><strong>${title}</strong><br><span>${time}</span>${addressHtml}<br><span>${note}</span>${maps}</div>`;
}

function createNumberedIcon(order, color) {
  return L.divIcon({
    className: "custom-div-icon",
    html: `<div class="marker-shell" style="--marker-color:${color}"><div class="marker-badge">${escapeHtml(String(order))}</div></div>`,
    iconSize: [38, 50],
    iconAnchor: [19, 38],
    popupAnchor: [0, -32],
  });
}

function syncMarkerActiveState() {
  state.markerMap.forEach((marker, id) => {
    const shell = marker.getElement()?.querySelector(".marker-shell");
    if (shell) shell.classList.toggle("is-active", id === state.selectedId);
  });
}

function focusItem(itemId, options = {}) {
  const item = getCurrentItems().find((entry) => entry.id === itemId);
  if (!item) return;

  state.selectedId = item.id;
  renderList(getCurrentItems());
  syncMarkerActiveState();

  if (item.hasCoordinates && options.flyTo !== false) {
    state.map.flyTo([item.lat, item.lng], Math.max(state.map.getZoom(), 14), {
      animate: true,
      duration: 0.65,
    });
  }

  if (item.hasCoordinates && options.openPopup) {
    const marker = state.markerMap.get(item.id);
    if (marker) setTimeout(() => marker.openPopup(), 200);
  }

  if (options.scrollCard) {
    const card = itineraryListEl.querySelector(`[data-id="${CSS.escape(item.id)}"]`);
    card?.scrollIntoView({ behavior: "smooth", block: "center" });
  }
}

function fitCurrentDayBounds({ animate = true } = {}) {
  const mappable = getCurrentItems().filter((item) => item.hasCoordinates);
  if (!mappable.length) {
    state.map.flyTo(defaultCenter, 10, { animate });
    return;
  }
  if (mappable.length === 1) {
    state.map.flyTo([mappable[0].lat, mappable[0].lng], 14, { animate });
    return;
  }
  const bounds = L.latLngBounds(mappable.map((item) => [item.lat, item.lng]));
  state.map.fitBounds(bounds.pad(0.2), { animate, paddingTopLeft: [20, 100], paddingBottomRight: [20, 170] });
}

function closeAnyOpenPopup() {
  state.map.closePopup();
}

function syncEditModeUi() {
  if (!toggleEditModeBtn) return;
  toggleEditModeBtn.classList.toggle("is-active", state.isEditMode);
  toggleEditModeBtn.setAttribute("aria-pressed", state.isEditMode ? "true" : "false");
  toggleEditModeBtn.textContent = state.isEditMode ? "完成" : "編輯";
}

function toggleEditMode() {
  state.isEditMode = !state.isEditMode;
  syncEditModeUi();
  renderList(getCurrentItems());
}

function openItemModal(mode = "add", item = null) {
  if (!itemModalEl || !itemFormEl) return;
  clearFormMessage();
  itemFormEl.reset();
  itemFormEl.elements.mode.value = mode;
  itemFormEl.elements.id.value = item?.id || "";

  if (mode === "edit" && item) {
    formKickerEl.textContent = "編輯旅遊地點";
    formTitleEl.textContent = "直接更新 Google Sheets";
    submitItemBtn.textContent = "更新這筆行程";
    fillFormWithItem(item);
  } else {
    formKickerEl.textContent = "新增旅遊地點";
    formTitleEl.textContent = "直接寫入 Google Sheets";
    submitItemBtn.textContent = "新增到 Google Sheets";
    presetFormDefaults();
  }

  itemModalEl.showModal();
}

function fillFormWithItem(item) {
  itemFormEl.elements.date.value = item.date || "";
  itemFormEl.elements.day.value = item.day || "";
  itemFormEl.elements.order.value = String(item.order || "");
  itemFormEl.elements.place_name.value = item.place_name || "";
  itemFormEl.elements.address.value = item.address || "";
  itemFormEl.elements.coords.value = item.coords || formatCoords(item.lat, item.lng);
  itemFormEl.elements.start_time.value = item.start_time || "";
  itemFormEl.elements.end_time.value = item.end_time || "";
  itemFormEl.elements.note.value = item.note || "";
  itemFormEl.elements.category.value = item.category || "";
  itemFormEl.elements.google_maps_url.value = item.google_maps_url || "";
  itemFormEl.elements.image_url.value = item.image_url || "";
  itemFormEl.elements.status.value = item.status || "planned";
}

function presetFormDefaults() {
  const today = state.currentDay || "day1";
  const items = getCurrentItems();
  const nextOrder = (items.reduce((max, item) => Math.max(max, Number(item.order) || 0), 0) || 0) + 1;
  const firstItemDate = items[0]?.date || formatDateInput(new Date());
  itemFormEl.elements.day.value = today;
  itemFormEl.elements.order.value = String(nextOrder);
  itemFormEl.elements.date.value = firstItemDate;
  itemFormEl.elements.status.value = "planned";
}

function closeItemModal() {
  itemModalEl?.close();
}


async function handleAutofillPlace() {
  if (!itemFormEl) return;
  clearFormMessage();

  const googleMapsUrl = String(itemFormEl.elements.google_maps_url?.value || "").trim();
  const adminKey = String(itemFormEl.elements.admin_key?.value || "").trim();

  if (!googleMapsUrl) {
    showFormMessage("請先貼上 Google Maps 連結。", "error");
    itemFormEl.elements.google_maps_url?.focus();
    return;
  }

  if (!/^https?:\/\//i.test(googleMapsUrl)) {
    showFormMessage("Google Maps 連結必須是 http 或 https 開頭。", "error");
    itemFormEl.elements.google_maps_url?.focus();
    return;
  }

  if (!adminKey) {
    showFormMessage("請先輸入管理密碼，才能使用自動帶入。", "error");
    itemFormEl.elements.admin_key?.focus();
    return;
  }

  autofillPlaceBtn.disabled = true;
  showFormMessage("正在解析 Google Maps 連結並帶入地點名稱與座標…", "info");

  try {
    const result = await requestJsonp(
      {
        action: "resolve_place",
        admin_key: adminKey,
        google_maps_url: googleMapsUrl,
      },
      { includeEmpty: true, timeout: 20000 }
    );

    if (!result || result.ok !== true) {
      throw new Error(result?.message || "無法自動帶入資料");
    }

    if (result.place_name) {
      itemFormEl.elements.place_name.value = result.place_name;
    }
    if (result.coords) {
      itemFormEl.elements.coords.value = result.coords;
    }
    if (result.google_maps_url) {
      itemFormEl.elements.google_maps_url.value = result.google_maps_url;
    }

    const parts = [];
    if (result.place_name) parts.push(`名稱：${result.place_name}`);
    if (result.coords) parts.push(`座標：${result.coords}`);
    showFormMessage(parts.length ? `已自動帶入。${parts.join("｜")}` : "已自動帶入可辨識資料。", "success");
  } catch (error) {
    console.error(error);
    showFormMessage(error.message || String(error), "error");
  } finally {
    autofillPlaceBtn.disabled = false;
  }
}

async function handleItemSubmit(event) {
  event.preventDefault();
  clearFormMessage();

  const formData = new FormData(itemFormEl);
  const payload = Object.fromEntries(formData.entries());
  const mode = payload.mode === "edit" ? "edit" : "add";
  const validationError = validateFormPayload(payload, mode);
  if (validationError) {
    showFormMessage(validationError, "error");
    return;
  }

  const action = mode === "edit" ? "update_item" : "add_item";
  const params = {
    action,
    admin_key: payload.admin_key.trim(),
    id: payload.id || "",
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

  submitItemBtn.disabled = true;
  showFormMessage(mode === "edit" ? "正在更新 Google Sheets…" : "正在新增到 Google Sheets…", "info");

  try {
    const result = await requestJsonp(params, { includeEmpty: true, timeout: 15000 });
    if (!result || result.ok !== true) {
      throw new Error(result?.message || "無法完成請求");
    }

    showFormMessage(result.message || (mode === "edit" ? "已更新。" : "已新增。"), "success");
    const preferredDay = params.day;
    const preferredId = mode === "edit" ? params.id : null;
    await reloadFromRemote(preferredDay, preferredId, true);
    setTimeout(() => {
      submitItemBtn.disabled = false;
      closeItemModal();
    }, 300);
  } catch (error) {
    console.error(error);
    showFormMessage(error.message || String(error), "error");
    submitItemBtn.disabled = false;
  }
}

async function handleDelete(item) {
  const adminKey = window.prompt(`刪除「${item.place_name || "未命名地點"}」\n請輸入管理密碼：`);
  if (adminKey === null) return;
  if (!adminKey.trim()) {
    window.alert("未輸入管理密碼。");
    return;
  }
  if (!window.confirm(`確定要刪除「${item.place_name || "未命名地點"}」嗎？`)) return;

  try {
    const result = await requestJsonp(
      { action: "delete_item", admin_key: adminKey.trim(), id: item.id },
      { includeEmpty: true, timeout: 15000 }
    );
    if (!result || result.ok !== true) {
      throw new Error(result?.message || "刪除失敗");
    }
    const fallbackId = getCurrentItems().find((x) => x.id !== item.id)?.id || null;
    await reloadFromRemote(state.currentDay, fallbackId, true);
  } catch (error) {
    window.alert(error.message || String(error));
  }
}

function validateFormPayload(payload, mode) {
  if (!payload.date) return "請填日期。";
  if (!/^\d{4}-\d{2}-\d{2}$/.test(payload.date.trim())) return "日期格式需為 YYYY-MM-DD。";
  if (!payload.day || !/^day\d+$/i.test(payload.day.trim())) return "day 欄請填像 day1、day2。";
  if (!payload.order || !/^\d+$/.test(String(payload.order).trim())) return "順序必須是正整數。";
  if (!payload.place_name || !payload.place_name.trim()) return "請填地點名稱。";
  if (!payload.coords || !/^\s*-?\d+(\.\d+)?\s*,\s*-?\d+(\.\d+)?\s*$/.test(payload.coords)) return "coords 格式請填「緯度, 經度」。";
  if (payload.start_time && !/^\d{2}:\d{2}$/.test(payload.start_time)) return "開始時間格式需為 HH:MM。";
  if (payload.end_time && !/^\d{2}:\d{2}$/.test(payload.end_time)) return "結束時間格式需為 HH:MM。";
  if (payload.google_maps_url && !/^https?:\/\//i.test(payload.google_maps_url.trim())) return "Google Maps 連結必須是 http 或 https 開頭。";
  if (payload.image_url && !/^https?:\/\//i.test(payload.image_url.trim())) return "圖片網址必須是 http 或 https 開頭。";
  if (!payload.admin_key || !payload.admin_key.trim()) return "請輸入管理密碼。";
  if (mode === "edit" && !payload.id) return "找不到這筆資料的 id。";
  return "";
}

function showFormMessage(message, type = "info") {
  itemFormMessageEl.textContent = message;
  itemFormMessageEl.className = `form-message ${type}`;
  itemFormMessageEl.classList.remove("is-hidden");
}

function clearFormMessage() {
  itemFormMessageEl.textContent = "";
  itemFormMessageEl.className = "form-message is-hidden";
}

function formatTimeRange(start, end) {
  const s = start?.trim();
  const e = end?.trim();
  if (s && e) return `${s}–${e}`;
  return s || e || "";
}

function formatCategory(category) {
  const value = String(category || "").trim();
  return value || "待確認";
}

function formatCoords(lat, lng) {
  if (Number.isFinite(lat) && Number.isFinite(lng)) return `${lat}, ${lng}`;
  return "";
}

function formatDateInput(dateObj) {
  const y = dateObj.getFullYear();
  const m = String(dateObj.getMonth() + 1).padStart(2, "0");
  const d = String(dateObj.getDate()).padStart(2, "0");
  return `${y}-${m}-${d}`;
}

function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function hexToRgba(hex, alpha) {
  const clean = hex.replace("#", "");
  const bigint = parseInt(clean.length === 3 ? clean.split("").map((c) => c + c).join("") : clean, 16);
  const r = (bigint >> 16) & 255;
  const g = (bigint >> 8) & 255;
  const b = bigint & 255;
  return `rgba(${r}, ${g}, ${b}, ${alpha})`;
}

function renderNoData() {
  itineraryListEl.innerHTML = `<div class="empty-state-card">找不到行程資料。</div>`;
}

function renderLoadError(error) {
  itineraryListEl.innerHTML = `<div class="empty-state-card">無法讀取 itinerary 資料：${escapeHtml(error.message || "Load failed")}</div>`;
}

function setupBottomSheet() {
  const snapPoints = [34, 54, 82];
  let dragging = false;
  let startY = 0;
  let startVh = state.sheetVh;

  applySheetHeight(state.sheetVh);

  const onPointerMove = (event) => {
    if (!dragging) return;
    const delta = startY - event.clientY;
    const vhDelta = (delta / window.innerHeight) * 100;
    const next = clamp(startVh + vhDelta, 28, 88);
    applySheetHeight(next);
  };

  const onPointerUp = () => {
    if (!dragging) return;
    dragging = false;
    const nearest = snapPoints.reduce((best, point) => Math.abs(point - state.sheetVh) < Math.abs(best - state.sheetVh) ? point : best, snapPoints[0]);
    applySheetHeight(nearest);
    document.body.style.userSelect = "";
    window.removeEventListener("pointermove", onPointerMove);
    window.removeEventListener("pointerup", onPointerUp);
  };

  sheetHandleArea?.addEventListener("pointerdown", (event) => {
    dragging = true;
    startY = event.clientY;
    startVh = state.sheetVh;
    document.body.style.userSelect = "none";
    window.addEventListener("pointermove", onPointerMove);
    window.addEventListener("pointerup", onPointerUp);
  });
}

function applySheetHeight(vh) {
  state.sheetVh = clamp(vh, 28, 88);
  document.documentElement.style.setProperty("--sheet-height", `${state.sheetVh}vh`);
  requestAnimationFrame(() => state.map?.invalidateSize());
}

function clamp(value, min, max) {
  return Math.max(min, Math.min(max, value));
}
