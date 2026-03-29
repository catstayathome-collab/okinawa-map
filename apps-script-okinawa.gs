const SHEET_NAME = '';

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

function jsonResponse_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function parseCoords_(coords) {
  if (!coords) return { lat: '', lng: '' };
  const parts = String(coords).split(',').map(s => s.trim());
  if (parts.length < 2) return { lat: '', lng: '' };
  return { lat: parts[0] || '', lng: parts[1] || '' };
}

function sheetToObjects_() {
  const sheet = getTargetSheet_();
  const values = sheet.getDataRange().getDisplayValues();
  if (values.length < 2) return [];

  const headers = values[0].map(h => String(h).trim());

  return values
    .slice(1)
    .filter(row => row.some(cell => String(cell).trim() !== ''))
    .map(row => {
      const obj = {};
      headers.forEach((header, i) => {
        obj[header] = row[i] ?? '';
      });

      if (obj.coords && String(obj.coords).includes(',')) {
        const parsed = parseCoords_(obj.coords);
        obj.lat = parsed.lat;
        obj.lng = parsed.lng;
      }

      return obj;
    });
}

function doGet(e) {
  try {
    const action = e && e.parameter ? e.parameter.action : '';

    if (action === 'add_item') {
      return response_(addItemFromRequest_(e.parameter), e.parameter.prefix);
    }
    if (action === 'update_item') {
      return response_(updateItemFromRequest_(e.parameter), e.parameter.prefix);
    }
    if (action === 'delete_item') {
      return response_(deleteItemFromRequest_(e.parameter), e.parameter.prefix);
    }

    const items = sheetToObjects_();
    return response_({
      ok: true,
      updatedAt: new Date().toISOString(),
      count: items.length,
      items: items,
    }, e && e.parameter ? e.parameter.prefix : '');
  } catch (error) {
    return response_({ ok: false, message: error.message || String(error) }, e && e.parameter ? e.parameter.prefix : '');
  }
}

function doPost(e) {
  try {
    const body = parsePostBody_(e) || {};
    if (body.action === 'add_item') return jsonResponse_(addItemFromRequest_(body));
    if (body.action === 'update_item') return jsonResponse_(updateItemFromRequest_(body));
    if (body.action === 'delete_item') return jsonResponse_(deleteItemFromRequest_(body));
    return jsonResponse_({ ok: false, message: '未知動作。' });
  } catch (error) {
    return jsonResponse_({ ok: false, message: error.message || String(error) });
  }
}

function parsePostBody_(e) {
  if (!e || !e.postData || !e.postData.contents) return null;
  try {
    return JSON.parse(e.postData.contents);
  } catch (err) {
    return null;
  }
}

function requireAdminKey_(params) {
  const expectedKey = getAdminKey_();
  if (!expectedKey) {
    return { ok: false, message: '尚未設定 ADMIN_KEY。' };
  }
  if (!params.admin_key || params.admin_key !== expectedKey) {
    return { ok: false, message: '管理密碼錯誤。' };
  }
  return { ok: true, message: '驗證通過。' };
}

function validateItem_(item, requireId) {
  if (requireId && (!item.id || !String(item.id).trim())) return '缺少 id。';
  if (!item.date || !/^\d{4}-\d{2}-\d{2}$/.test(String(item.date).trim())) return 'date 格式需為 YYYY-MM-DD。';
  if (!item.day || !/^day\d+$/i.test(String(item.day).trim())) return 'day 欄請填像 day1、day2。';
  if (!item.order || !/^\d+$/.test(String(item.order).trim())) return 'order 必須是正整數。';
  if (!item.place_name || !String(item.place_name).trim()) return 'place_name 不可空白。';
  if (item.coords && !/^\s*-?\d+(\.\d+)?\s*,\s*-?\d+(\.\d+)?\s*$/.test(String(item.coords))) return 'coords 格式請填「緯度, 經度」。';
  if (item.google_maps_url && !/^https?:\/\//i.test(String(item.google_maps_url).trim())) return 'google_maps_url 必須是 http 或 https 開頭。';
  if (item.image_url && !/^https?:\/\//i.test(String(item.image_url).trim())) return 'image_url 必須是 http 或 https 開頭。';
  return '';
}

function buildRowObject_(item, preserveId) {
  const day = String(item.day || '').trim();
  const order = String(item.order || '').trim();
  const generatedId = preserveId || (item.id && String(item.id).trim()) || (day && order ? day + '-' + order + '-' + Date.now() : 'item-' + Date.now());

  return {
    id: generatedId,
    date: item.date || '',
    day: item.day || '',
    order: item.order || '',
    place_name: item.place_name || '',
    address: item.address || '',
    coords: item.coords || '',
    start_time: item.start_time || '',
    end_time: item.end_time || '',
    note: item.note || '',
    category: item.category || '',
    google_maps_url: item.google_maps_url || autoGoogleMapsUrl_(item),
    image_url: item.image_url || '',
    status: item.status || 'planned',
  };
}

function autoGoogleMapsUrl_(item) {
  if (item.google_maps_url) return String(item.google_maps_url).trim();
  if (item.coords && String(item.coords).includes(',')) {
    return 'https://www.google.com/maps/search/?api=1&query=' + encodeURIComponent(String(item.coords).trim());
  }
  const q = [item.place_name || '', item.address || ''].join(' ').trim();
  return q ? 'https://www.google.com/maps/search/?api=1&query=' + encodeURIComponent(q) : '';
}

function buildRowArray_(rowObject, headers) {
  return headers.map(header => rowObject[header] ?? '');
}

function findRowIndexById_(id) {
  const sheet = getTargetSheet_();
  const values = sheet.getDataRange().getDisplayValues();
  if (values.length < 2) return -1;

  const headers = values[0].map(h => String(h).trim());
  const idIndex = headers.indexOf('id');
  if (idIndex === -1) throw new Error('表頭缺少 id 欄位。');

  for (var i = 1; i < values.length; i += 1) {
    if (String(values[i][idIndex]).trim() === String(id).trim()) {
      return i + 1;
    }
  }
  return -1;
}

function addItemFromRequest_(params) {
  const auth = requireAdminKey_(params);
  if (!auth.ok) return auth;

  const item = params.item || params;
  const validationError = validateItem_(item, false);
  if (validationError) return { ok: false, message: validationError };

  const headers = getHeaders_();
  if (!headers.length) return { ok: false, message: '找不到表頭，無法寫入。' };

  const rowObject = buildRowObject_(item, '');
  const row = buildRowArray_(rowObject, headers);
  getTargetSheet_().appendRow(row);
  return { ok: true, message: '已新增到 Google Sheets。', id: rowObject.id };
}

function updateItemFromRequest_(params) {
  const auth = requireAdminKey_(params);
  if (!auth.ok) return auth;

  const item = params.item || params;
  const validationError = validateItem_(item, true);
  if (validationError) return { ok: false, message: validationError };

  const rowIndex = findRowIndexById_(item.id);
  if (rowIndex === -1) return { ok: false, message: '找不到要更新的資料。' };

  const headers = getHeaders_();
  const rowObject = buildRowObject_(item, item.id);
  const row = buildRowArray_(rowObject, headers);
  getTargetSheet_().getRange(rowIndex, 1, 1, row.length).setValues([row]);
  return { ok: true, message: '已更新 Google Sheets。', id: item.id };
}

function deleteItemFromRequest_(params) {
  const auth = requireAdminKey_(params);
  if (!auth.ok) return auth;

  if (!params.id || !String(params.id).trim()) {
    return { ok: false, message: '缺少 id。' };
  }

  const rowIndex = findRowIndexById_(params.id);
  if (rowIndex === -1) return { ok: false, message: '找不到要刪除的資料。' };

  getTargetSheet_().deleteRow(rowIndex);
  return { ok: true, message: '已刪除這筆資料。', id: params.id };
}
