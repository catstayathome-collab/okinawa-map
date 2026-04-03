const SHEET_NAME = '';
// 留空 = 自動讀第一個工作表
// 如果你想指定工作表，就改成例如：const SHEET_NAME = '工作表1';

// 在「專案設定 → 指令碼屬性」新增：
// ADMIN_KEY = 你的管理密碼
// PLACES_API_KEY = 你的 Google Maps Platform API key

function getTargetSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (SHEET_NAME) {
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) throw new Error('找不到工作表：' + SHEET_NAME);
    return sheet;
  }
  return ss.getSheets()[0];
}

function getHeaders_() {
  const sheet = getTargetSheet_();
  const values = sheet.getDataRange().getDisplayValues();
  if (!values.length) return [];
  return values[0].map(h => String(h).trim());
}

function getAdminKey_() {
  return PropertiesService.getScriptProperties().getProperty('ADMIN_KEY') || '';
}

function getPlacesApiKey_() {
  return PropertiesService.getScriptProperties().getProperty('PLACES_API_KEY') || '';
}

function response_(payload, prefix) {
  const json = JSON.stringify(payload);
  if (prefix) {
    return ContentService
      .createTextOutput(prefix + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

function parseCoords_(coords) {
  if (!coords) return { lat: '', lng: '' };
  const parts = String(coords).split(',').map(s => s.trim());
  if (parts.length < 2) return { lat: '', lng: '' };
  return { lat: parts[0] || '', lng: parts[1] || '' };
}

function formatCoords_(lat, lng) {
  if (lat === '' || lng === '') return '';
  return String(lat) + ', ' + String(lng);
}

function rowToObject_(headers, row) {
  const obj = {};
  headers.forEach((header, i) => { obj[header] = row[i] ?? ''; });
  if (obj.coords && String(obj.coords).includes(',')) {
    const parsed = parseCoords_(obj.coords);
    obj.lat = parsed.lat;
    obj.lng = parsed.lng;
  }
  return obj;
}

function sheetToObjects_() {
  const sheet = getTargetSheet_();
  const values = sheet.getDataRange().getDisplayValues();
  if (values.length < 2) return [];
  const headers = values[0].map(h => String(h).trim());
  return values
    .slice(1)
    .filter(row => row.some(cell => String(cell).trim() !== ''))
    .map(row => rowToObject_(headers, row));
}

function requireAdmin_(params) {
  const adminKey = getAdminKey_();
  if (!adminKey) return { ok: false, message: '尚未設定 ADMIN_KEY。' };
  if (!params.admin_key || params.admin_key !== adminKey) {
    return { ok: false, message: '管理密碼錯誤。' };
  }
  return { ok: true };
}

function validateItemFields_(item) {
  if (!item.date || !/^\d{4}-\d{2}-\d{2}$/.test(String(item.date))) {
    return { ok: false, message: 'date 格式需為 YYYY-MM-DD。' };
  }
  if (!item.day || !/^day\d+$/i.test(String(item.day))) {
    return { ok: false, message: 'day 格式需為 day1、day2 這類。' };
  }
  if (!item.order || !/^\d+$/.test(String(item.order))) {
    return { ok: false, message: 'order 必須是正整數。' };
  }
  if (!item.place_name || !String(item.place_name).trim()) {
    return { ok: false, message: 'place_name 不可空白。' };
  }
  if (!item.coords || !String(item.coords).includes(',')) {
    return { ok: false, message: 'coords 格式請填「緯度, 經度」。' };
  }
  const parsed = parseCoords_(item.coords);
  if (!parsed.lat || !parsed.lng || isNaN(Number(parsed.lat)) || isNaN(Number(parsed.lng))) {
    return { ok: false, message: 'coords 不是有效的經緯度。' };
  }
  if (item.google_maps_url && !/^https?:\/\//i.test(String(item.google_maps_url))) {
    return { ok: false, message: 'google_maps_url 必須是 http 或 https 開頭。' };
  }
  if (item.image_url && !/^https?:\/\//i.test(String(item.image_url))) {
    return { ok: false, message: 'image_url 必須是 http 或 https 開頭。' };
  }
  return { ok: true };
}

function buildRowFromObject_(obj, headers) {
  return headers.map(h => obj[h] ?? '');
}

function findRowById_(id) {
  const sheet = getTargetSheet_();
  const values = sheet.getDataRange().getDisplayValues();
  if (values.length < 2) return null;
  const headers = values[0].map(h => String(h).trim());
  for (let i = 1; i < values.length; i++) {
    const rowObj = rowToObject_(headers, values[i]);
    if (String(rowObj.id) === String(id)) {
      return {
        rowNumber: i + 1,
        headers: headers,
        rowValues: values[i],
        rowObject: rowObj
      };
    }
  }
  return null;
}

function addItemFromRequest_(params) {
  const adminCheck = requireAdmin_(params);
  if (!adminCheck.ok) return adminCheck;

  const headers = getHeaders_();
  if (!headers.length) return { ok: false, message: '找不到表頭，無法寫入。' };

  const day = String(params.day || '').trim();
  const order = String(params.order || '').trim();

  const rowObject = {
    id: params.id && String(params.id).trim() ? String(params.id).trim() : day + '-' + order + '-' + Date.now(),
    date: params.date || '',
    day: params.day || '',
    order: params.order || '',
    place_name: params.place_name || '',
    address: params.address || '',
    coords: params.coords || '',
    start_time: params.start_time || '',
    end_time: params.end_time || '',
    note: params.note || '',
    category: params.category || '',
    google_maps_url: params.google_maps_url || '',
    image_url: params.image_url || '',
    status: params.status || 'planned'
  };

  const validation = validateItemFields_(rowObject);
  if (!validation.ok) return validation;

  const row = buildRowFromObject_(rowObject, headers);
  getTargetSheet_().appendRow(row);
  return { ok: true, message: '已新增到 Google Sheets。' };
}

function updateItemFromRequest_(params) {
  const adminCheck = requireAdmin_(params);
  if (!adminCheck.ok) return adminCheck;
  if (!params.id) return { ok: false, message: '缺少 id，無法更新。' };

  const found = findRowById_(params.id);
  if (!found) return { ok: false, message: '找不到要更新的資料列。' };

  const current = found.rowObject;
  const headers = found.headers;
  const editableFields = [
    'date','day','order','place_name','address','coords',
    'start_time','end_time','note','category',
    'google_maps_url','image_url','status'
  ];

  const updated = Object.assign({}, current);
  editableFields.forEach(field => {
    if (Object.prototype.hasOwnProperty.call(params, field)) {
      updated[field] = params[field];
    }
  });

  const validation = validateItemFields_(updated);
  if (!validation.ok) return validation;

  const newRow = buildRowFromObject_(updated, headers);
  getTargetSheet_().getRange(found.rowNumber, 1, 1, headers.length).setValues([newRow]);

  return { ok: true, message: '已更新 Google Sheets。' };
}

function deleteItemFromRequest_(params) {
  const adminCheck = requireAdmin_(params);
  if (!adminCheck.ok) return adminCheck;
  if (!params.id) return { ok: false, message: '缺少 id，無法刪除。' };

  const found = findRowById_(params.id);
  if (!found) return { ok: false, message: '找不到要刪除的資料列。' };

  getTargetSheet_().deleteRow(found.rowNumber);
  return { ok: true, message: '已刪除該筆資料。' };
}

function expandGoogleMapsUrl_(inputUrl) {
  let current = String(inputUrl || '').trim();
  if (!current) throw new Error('Google Maps 連結不可空白。');

  for (var i = 0; i < 5; i++) {
    const resp = UrlFetchApp.fetch(current, {
      method: 'get',
      followRedirects: false,
      muteHttpExceptions: true,
      headers: { 'User-Agent': 'Mozilla/5.0' }
    });
    const code = resp.getResponseCode();
    const headers = resp.getHeaders();
    const location = headers.Location || headers.location;
    if (location && [301,302,303,307,308].indexOf(code) !== -1) {
      current = location;
      continue;
    }
    break;
  }
  return current;
}

function extractPlaceQueryFromMapsUrl_(mapsUrl) {
  const expanded = expandGoogleMapsUrl_(mapsUrl);
  const raw = String(expanded || '').trim();

  if (!raw) {
    throw new Error('Google Maps 連結不可空白。');
  }

  let placeName = '';
  let lat = '';
  let lng = '';

  const placeMatch = raw.match(/\/maps\/place\/([^/]+)\//i);
  if (placeMatch && placeMatch[1]) {
    placeName = decodeURIComponent(placeMatch[1]).replace(/\+/g, ' ').trim();
  }

  // 先抓更接近 place pin 的 !3d !4d
  const pinMatch = raw.match(/!3d(-?\d+(?:\.\d+)?)!4d(-?\d+(?:\.\d+)?)/);
  if (pinMatch) {
    lat = pinMatch[1];
    lng = pinMatch[2];
  } else {
    // 沒有 pin 座標時，才退回抓 @lat,lng
    const centerMatch = raw.match(/@(-?\d+(?:\.\d+)?),(-?\d+(?:\.\d+)?)/);
    if (centerMatch) {
      lat = centerMatch[1];
      lng = centerMatch[2];
    }
  }

  let query = '';
  const queryMatch = raw.match(/[?&]query=([^&]+)/i);
  const qMatch = raw.match(/[?&]q=([^&]+)/i);

  if (queryMatch && queryMatch[1]) {
    query = decodeURIComponent(queryMatch[1]).replace(/\+/g, ' ').trim();
  } else if (qMatch && qMatch[1]) {
    query = decodeURIComponent(qMatch[1]).replace(/\+/g, ' ').trim();
  }

  if (!query && placeName) {
    query = placeName;
  }

  if (!query && !placeName) {
    throw new Error('無法從 Google Maps 連結辨識地點關鍵字，請改貼較完整的地圖網址。');
  }

  return {
    expandedUrl: expanded,
    query: query || placeName,
    directPlaceName: placeName || '',
    directLat: lat || '',
    directLng: lng || ''
  };
}

function fetchJson_(url, options) {
  const resp = UrlFetchApp.fetch(url, options);
  const code = resp.getResponseCode();
  const text = resp.getContentText();
  if (code < 200 || code >= 300) {
    throw new Error('Places API 呼叫失敗：' + code + '｜' + text);
  }
  return JSON.parse(text);
}

function resolvePlaceFromGoogleMapsUrl_(params) {
  const adminCheck = requireAdmin_(params);
  if (!adminCheck.ok) return adminCheck;

  const mapsUrl = String(params.google_maps_url || '').trim();
  if (!/^https?:\/\//i.test(mapsUrl)) {
    return { ok: false, message: 'Google Maps 連結必須是 http 或 https 開頭。' };
  }

  const apiKey = getPlacesApiKey_();
  if (!apiKey) {
    return { ok: false, message: '尚未設定 PLACES_API_KEY。' };
  }

  try {
    const parsed = extractPlaceQueryFromMapsUrl_(mapsUrl);

    // 一律優先用 Places API 拿「名稱」
    // directPlaceName 只當備援，不直接當最終答案
    const textSearchUrl = 'https://places.googleapis.com/v1/places:searchText';

    const textSearchPayload = {
      textQuery: parsed.query,
      languageCode: 'ja'
    };

    // 如果網址本身有解析出座標，就用來幫搜尋做 locationBias
    if (parsed.directLat && parsed.directLng) {
      textSearchPayload.locationBias = {
        circle: {
          center: {
            latitude: Number(parsed.directLat),
            longitude: Number(parsed.directLng)
          },
          radius: 80.0
        }
      };
    }

    const searchResult = fetchJson_(textSearchUrl, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(textSearchPayload),
      headers: {
        'X-Goog-Api-Key': apiKey,
        'X-Goog-FieldMask': 'places.id'
      },
      muteHttpExceptions: true
    });

    if (!searchResult.places || !searchResult.places.length || !searchResult.places[0].id) {
      // 找不到時才退回網址本身解析出的名稱與座標
      if (parsed.directPlaceName && parsed.directLat && parsed.directLng) {
        return {
          ok: true,
          message: 'Places API 找不到更精確結果，已改用 Google Maps 網址資訊。',
          place_name: parsed.directPlaceName,
          coords: formatCoords_(parsed.directLat, parsed.directLng),
          lat: parsed.directLat,
          lng: parsed.directLng,
          google_maps_url: parsed.expandedUrl,
          query_used: parsed.query
        };
      }

      return { ok: false, message: '找不到符合的地點，請改貼更完整的 Google Maps 連結。' };
    }

    const placeId = searchResult.places[0].id;

    const detailsUrl =
      'https://places.googleapis.com/v1/places/' +
      encodeURIComponent(placeId) +
      '?languageCode=ja';

    const details = fetchJson_(detailsUrl, {
      method: 'get',
      headers: {
        'X-Goog-Api-Key': apiKey,
        'X-Goog-FieldMask': 'displayName,location'
      },
      muteHttpExceptions: true
    });

    const displayName =
      details.displayName && details.displayName.text
        ? details.displayName.text
        : '';

    const detailsLat =
      details.location && details.location.latitude != null
        ? details.location.latitude
        : '';

    const detailsLng =
      details.location && details.location.longitude != null
        ? details.location.longitude
        : '';

    // 名稱優先用 Places 的日文化 displayName
    const finalName = displayName || parsed.directPlaceName || '';

    // 座標優先保留你從 Google Maps 網址抓到的 pin 座標
    // 如果網址沒抓到，才退回 Places location
    const finalLat = parsed.directLat || detailsLat;
    const finalLng = parsed.directLng || detailsLng;

    if (!finalName || finalLat === '' || finalLng === '') {
      return { ok: false, message: '已找到地點，但無法完整取得名稱或座標。' };
    }

    return {
      ok: true,
      message: '已自動帶入地點名稱與座標。',
      place_name: finalName,
      coords: formatCoords_(finalLat, finalLng),
      lat: finalLat,
      lng: finalLng,
      google_maps_url: parsed.expandedUrl,
      place_id: placeId,
      query_used: parsed.query
    };
  } catch (error) {
    return { ok: false, message: error.message || String(error) };
  }
}

function handleAction_(params) {
  const action = params.action || '';

  if (action === 'add_item') return addItemFromRequest_(params);
  if (action === 'update_item') return updateItemFromRequest_(params);
  if (action === 'delete_item') return deleteItemFromRequest_(params);
  if (action === 'resolve_place') return resolvePlaceFromGoogleMapsUrl_(params);

  return null;
}

function doGet(e) {
  const params = (e && e.parameter) ? e.parameter : {};
  const actionResult = handleAction_(params);
  if (actionResult) {
    return response_(actionResult, params.prefix || '');
  }

  const items = sheetToObjects_();
  const payload = {
    ok: true,
    updatedAt: new Date().toISOString(),
    count: items.length,
    items: items
  };
  return response_(payload, params.prefix || '');
}

function doPost(e) {
  try {
    let params = {};
    if (e && e.postData && e.postData.contents) {
      try {
        params = JSON.parse(e.postData.contents);
      } catch (err) {
        params = e.parameter || {};
      }
    } else {
      params = (e && e.parameter) ? e.parameter : {};
    }

    const actionResult = handleAction_(params);
    if (actionResult) {
      return response_(actionResult, params.prefix || '');
    }

    return response_({ ok: false, message: '未知的 action。' }, params.prefix || '');
  } catch (error) {
    return response_({ ok: false, message: error.message || String(error) }, '');
  }
}

function testDoGet() {
  const output = doGet({ parameter: {} });
  Logger.log(output.getContent());
}

function testResolvePlace() {
  const result = resolvePlaceFromGoogleMapsUrl_({
    admin_key: 'okinawa2026',
    google_maps_url: 'https://www.google.com/maps/place/%E9%82%A3%E9%9C%B8%E6%A9%9F%E5%A0%B4/@26.2001329,127.6054502,13z/data=!4m6!3m5!1s0x34e569c48e8b2c8d:0x504cceaa3756fe90!8m2!3d26.2001297!4d127.6466452!16zL20vMDNibmM3?authuser=0&entry=ttu&g_ep=EgoyMDI2MDMyNC4wIKXMDSoASAFQAw%3D%3D'
  });

  Logger.log(JSON.stringify(result));
}