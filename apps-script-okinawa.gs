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
        const parts = String(obj.coords).split(',').map(s => s.trim());
        obj.lat = parts[0] || '';
        obj.lng = parts[1] || '';
      }

      return obj;
    });
}

function doGet(e) {
  try {
    if (e && e.parameter && e.parameter.action === 'add_item') {
      const payload = addItemFromRequest_(e.parameter);
      return response_(payload, e.parameter.prefix);
    }

    const items = sheetToObjects_();
    const payload = {
      ok: true,
      updatedAt: new Date().toISOString(),
      count: items.length,
      items: items,
    };
    return response_(payload, e && e.parameter ? e.parameter.prefix : '');
  } catch (error) {
    return response_({ ok: false, message: error.message || String(error) }, e && e.parameter ? e.parameter.prefix : '');
  }
}

function doPost(e) {
  try {
    const body = parsePostBody_(e);
    if (!body || body.action !== 'add_item') {
      return jsonResponse_({ ok: false, message: '未知動作。' });
    }
    const payload = addItemFromRequest_(body);
    return jsonResponse_(payload);
  } catch (error) {
    return jsonResponse_({ ok: false, message: error.message || String(error) });
  }
}

function addItemFromRequest_(params) {
  const adminKey = params.admin_key || '';
  const expectedKey = PropertiesService.getScriptProperties().getProperty('ADMIN_KEY') || '';
  if (!expectedKey) {
    return { ok: false, message: '尚未設定 ADMIN_KEY。' };
  }
  if (adminKey !== expectedKey) {
    return { ok: false, message: '管理密碼錯誤。' };
  }

  const item = params.item || params;
  const validationError = validateItem_(item);
  if (validationError) {
    return { ok: false, message: validationError };
  }

  const sheet = getTargetSheet_();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0].map(h => String(h).trim());
  const rowObject = buildRowObject_(item);
  const row = headers.map(header => rowObject[header] ?? '');
  sheet.appendRow(row);

  return { ok: true, message: '已新增到 Google Sheets。' };
}

function parsePostBody_(e) {
  if (!e || !e.postData || !e.postData.contents) return null;
  const text = e.postData.contents;
  try {
    return JSON.parse(text);
  } catch (err) {
    return null;
  }
}

function validateItem_(item) {
  if (!item.date || !/^\d{4}-\d{2}-\d{2}$/.test(String(item.date).trim())) return 'date 格式需為 YYYY-MM-DD。';
  if (!item.day || !/^day\d+$/i.test(String(item.day).trim())) return 'day 欄請填像 day1、day2。';
  if (!item.order || !/^\d+$/.test(String(item.order).trim())) return 'order 必須是正整數。';
  if (!item.place_name || !String(item.place_name).trim()) return 'place_name 不可空白。';
  if (item.coords && !/^\s*-?\d+(\.\d+)?\s*,\s*-?\d+(\.\d+)?\s*$/.test(String(item.coords))) return 'coords 格式請填「緯度, 經度」。';
  if (item.google_maps_url && !/^https?:\/\//i.test(String(item.google_maps_url).trim())) return 'google_maps_url 必須是 http 或 https 開頭。';
  if (item.image_url && !/^https?:\/\//i.test(String(item.image_url).trim())) return 'image_url 必須是 http 或 https 開頭。';
  return '';
}

function buildRowObject_(item) {
  return {
    id: item.id || '',
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
  if (item.coords && String(item.coords).includes(',')) {
    return 'https://www.google.com/maps/search/?api=1&query=' + encodeURIComponent(String(item.coords).trim());
  }
  const q = [item.place_name || '', item.address || ''].join(' ').trim();
  return q ? 'https://www.google.com/maps/search/?api=1&query=' + encodeURIComponent(q) : '';
}

function response_(obj, prefix) {
  if (prefix) {
    return ContentService
      .createTextOutput(prefix + '(' + JSON.stringify(obj) + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return jsonResponse_(obj);
}

function jsonResponse_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
