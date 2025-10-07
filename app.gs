/**
 * 同人場行事曆＋攤位收藏 LINE Bot
 * Apps Script (V8) implementation.
 */
function testWebhook() {
  const url = 'https://script.google.com/macros/s/AKfycbw9EIV-GcePNZAyHuOZFWB__bGNVKibN8YkKWTCxYn3899FfidH_bAdMhpYcHgkaHpHtQ/exec'; // 你的 Web app URL
  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({ events: [] })
  });
  Logger.log('code=' + res.getResponseCode()); // 要 200
  Logger.log('body=' + res.getContentText());  // OK
}

function getDb() {
  const prop = PropertiesService.getScriptProperties();
  const saved = prop.getProperty('DB_ID');
  if (saved) return SpreadsheetApp.openById(saved);

  // 若已綁定某試算表，直接用它
  const active = SpreadsheetApp.getActive();
  if (active) return active;

  // 否則第一次執行就自動建立一份資料庫
  const ss = SpreadsheetApp.create('Doujin Bot DB');
  prop.setProperty('DB_ID', ss.getId());
  return ss;
}

const SHEET_USERS = 'Users';
const SHEET_EVENTS = 'Events';
const SHEET_BOOTHS = 'Booths';
const SHEET_FAVORITES = 'Favorites';
const SHEET_CONFIG = 'Config';
const TIMEZONE = 'Asia/Taipei';
const MAX_EVENT_LOOKAHEAD_DAYS = 60;
const DEFAULT_EVENT_LIST_COUNT = 5;
const MAX_EVENT_LIST_COUNT = 20;
const MAX_LIST_ITEMS = 10;
const DEFAULT_REMIND_MINS = 30;

/**
 * 初始化試算表工作表與標題列。
 */
function initSheets() {
  const ss = getDb();
  const definitions = [
    { name: SHEET_USERS, headers: ['id', 'lineUserId', 'name', 'lastEventId', 'lastEventName', 'createdAt'] },
    { name: SHEET_EVENTS, headers: ['id', 'name', 'start', 'end', 'location', 'url', 'openAt', 'note', 'createdAt'] },
    { name: SHEET_BOOTHS, headers: ['id', 'eventId', 'code', 'circle', 'works', 'zone', 'tableNo', 'start', 'end', 'note'] },
    { name: SHEET_FAVORITES, headers: ['id', 'userId', 'eventId', 'boothId', 'remindMins', 'createdAt'] },
    { name: SHEET_CONFIG, headers: ['KEY', 'VALUE'] }
  ];

  definitions.forEach(def => {
    let sheet = ss.getSheetByName(def.name);
    if (!sheet) {
      sheet = ss.insertSheet(def.name);
      Logger.log('Created sheet %s', def.name);
    }
    const range = sheet.getRange(1, 1, 1, def.headers.length);
    range.setValues([def.headers]);
  });
}

/**
 * 範例資料建立：1-2 場活動與攤位。
 */
function seedSample() {
  const ss = SpreadsheetApp.getActive();
  const nowIso = new Date().toISOString();
  const eventsSheet = ss.getSheetByName(SHEET_EVENTS);
  const boothsSheet = ss.getSheetByName(SHEET_BOOTHS);

  if (!eventsSheet || !boothsSheet) {
    throw new Error('請先執行 initSheets() 建立工作表');
  }

  const today = new Date();
  const event1Date = new Date(today.getTime() + 7 * 24 * 60 * 60 * 1000);
  const event2Date = new Date(today.getTime() + 14 * 24 * 60 * 60 * 1000);

  const eventRows = [
    ['EVT001', 'FF 同人祭', formatDate(event1Date), '', '花博爭艷館', 'https://example.com/ff', '10:00', ''],
    ['EVT002', 'CWT 漫畫博覽會', formatDate(event2Date), '', '台北世貿一館', 'https://example.com/cwt', '10:30', '']
  ];

  const eventData = eventRows.map(row => row.concat(nowIso));
  eventsSheet.getRange(eventsSheet.getLastRow() + 1, 1, eventData.length, eventData[0].length).setValues(eventData);

  const boothRows = [
    ['BO001', 'EVT001', 'A12', '星光社', '原創小說', 'A', '12', '', '', ''],
    ['BO002', 'EVT001', 'B08', '蔚藍工作室', '奇幻插畫', 'B', '08', '', '', ''],
    ['BO003', 'EVT001', 'C05', '山海誌', '神話設定集', 'C', '05', '', '', ''],
    ['BO004', 'EVT002', 'A01', '晨曦社', '輕小說', 'A', '01', '', '', ''],
    ['BO005', 'EVT002', 'B11', '青藍紙上', '漫畫周邊', 'B', '11', '', '', '']
  ];
  boothsSheet.getRange(boothsSheet.getLastRow() + 1, 1, boothRows.length, boothRows[0].length).setValues(boothRows);
}

// GET：健康檢查 / 後台
function doGet(e) {
  const a = (e && e.parameter && e.parameter.a || '').toLowerCase();
  if (a === 'admin') return handleAdmin(e); // 後台照舊
  return ContentService.createTextOutput('OK'); // 其他 GET 一律 200
}

// POST：LINE Webhook（不管發生什麼都回 200）
function doPost(e) {
  try {
    return handleCallback(e || {});
  } catch (err) {
    Logger.log('doPost error: ' + err.message);
    return ContentService.createTextOutput('OK');
  }
}

/**
 * 從 Config 工作表取得設定。
 */
function getConfigValue(key) {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(SHEET_CONFIG);
  if (!sheet) {
    return '';
  }
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === key) {
      return data[i][1];
    }
  }
  return '';
}

// 安全的 callback：容忍空 body/非 JSON
function handleCallback(e) {
  let bodyText = '';
  try { bodyText = (e && e.postData && e.postData.contents) || '{}'; } catch (_) {}
  let body; try { body = JSON.parse(bodyText); } catch (_) { body = {}; }
  const events = Array.isArray(body.events) ? body.events : [];
  if (events.length === 0) return ContentService.createTextOutput('OK'); // Verify 送空陣列也 200

  // 你的原本事件處理
  events.forEach(ev => {
    const uid = ev.source && ev.source.userId;
    if (!uid) return;
    const u = ensureUser(uid);
    if (ev.type === 'follow') {
      safe(() => linePush(uid, { type: 'text', text: '歡迎～輸入「場次」看近期活動，或輸入「指令」看用法。' }));
    } else if (ev.type === 'message' && ev.message && ev.message.type === 'text') {
      handleText(u, ev.message.text || '', ev.replyToken);
    }
  });
  return ContentService.createTextOutput('OK');
}
function onFollow(event) {
  const userId = event.source && event.source.userId;
  if (!userId) {
    return;
  }
  const profile = getUserProfile(userId);
  const displayName = profile && profile.displayName ? profile.displayName : '同好';
  const user = ensureUser(userId, displayName);
  const message = '歡迎～輸入「場次」看近期活動，或輸入「指令」看用法。';
  replyMessage(event.replyToken, [buildTextMessage(message)]);
  Logger.log('follow event handled for user %s (%s)', user.id, displayName);
}

function onTextMessage(event) {
  const userId = event.source && event.source.userId;
  const text = event.message && event.message.text ? event.message.text.trim() : '';
  if (!userId || !text) {
    return;
  }
  const profile = getUserProfile(userId);
  const displayName = profile && profile.displayName ? profile.displayName : '同好';
  const user = ensureUser(userId, displayName);
  const normalized = text.replace(/\s+/g, ' ').trim();

  if (normalized === '指令') {
    return replyMessage(event.replyToken, [buildTextMessage(buildHelpText())]);
  }

  const eventListMatch = normalized.match(/^場次(?:\s*(\d+))?$/);
  if (eventListMatch) {
    const countRaw = eventListMatch[1];
    const count = Math.min(Math.max(parseInt(countRaw, 10) || DEFAULT_EVENT_LIST_COUNT, 1), MAX_EVENT_LIST_COUNT);
    const events = listUpcomingEvents(count);
    const textOutput = events.length ? events.map(ev => `${ev.name}\n日期：${formatEventDate(ev)}\n地點：${ev.location}`).join('\n\n') : '找不到符合條件的場次，請稍後再試或請管理者執行 seedSample()。';
    replyMessage(event.replyToken, [buildTextMessage(textOutput)]);
    return;
  }

  const boothsMatch = normalized.match(/^攤位\s+(.+)/);
  if (boothsMatch) {
    const keyword = boothsMatch[1];
    const response = handleBoothsCommand(user, keyword);
    replyMessage(event.replyToken, [buildTextMessage(response.message)]);
    if (response.selectedEvent) {
      updateUserLastEvent(user.id, response.selectedEvent.id, response.selectedEvent.name);
    }
    return;
  }

  const searchMatch = normalized.match(/^搜攤\s+(.+)/);
  if (searchMatch) {
    const keyword = searchMatch[1];
    const textOutput = handleSearchBooth(keyword);
    replyMessage(event.replyToken, [buildTextMessage(textOutput)]);
    return;
  }

  const favMatch = normalized.match(/^收藏\s+(\S+)/);
  if (favMatch) {
    const code = favMatch[1];
    const message = handleFavorite(user, code, true);
    replyMessage(event.replyToken, [buildTextMessage(message)]);
    return;
  }

  const unfavMatch = normalized.match(/^取消收藏\s+(\S+)/);
  if (unfavMatch) {
    const code = unfavMatch[1];
    const message = handleFavorite(user, code, false);
    replyMessage(event.replyToken, [buildTextMessage(message)]);
    return;
  }

  if (normalized === '我的收藏') {
    const message = listUserFavorites(user);
    replyMessage(event.replyToken, [buildTextMessage(message)]);
    return;
  }

  const remindMatch = normalized.match(/^提醒\s+(\S+)\s+提前=(\d+)/);
  if (remindMatch) {
    const code = remindMatch[1];
    const mins = parseInt(remindMatch[2], 10);
    const message = handleReminder(user, code, mins);
    replyMessage(event.replyToken, [buildTextMessage(message)]);
    return;
  }

  replyMessage(event.replyToken, [buildTextMessage(buildHelpText())]);
}

function buildHelpText() {
  return '指令一覽：\n場次 [數字]\n攤位 <場次關鍵字>\n搜攤 <關鍵字>\n收藏 <攤位代碼>\n取消收藏 <攤位代碼>\n提醒 <攤位代碼> 提前=<分鐘>\n我的收藏\n提示：收藏/提醒會使用你最近查詢的場次作為上下文。';
}

function handleBoothsCommand(user, keyword) {
  const event = findEventByKeyword(keyword);
  if (!event) {
    return { message: '找不到符合的場次，請換個關鍵字或輸入「場次」查看列表。' };
  }
  const booths = listBoothsByEvent(event.id);
  if (!booths.length) {
    return { message: `場次「${event.name}」目前沒有攤位資料，請稍後再試。`, selectedEvent: event };
  }
  const topBooths = booths.slice(0, MAX_LIST_ITEMS);
  const message = `場次：${event.name}\n` + topBooths.map(b => `${b.code}｜${b.circle}｜${b.works}`).join('\n');
  return { message, selectedEvent: event };
}

function handleSearchBooth(keyword) {
  if (!keyword) {
    return '請提供搜尋關鍵字。';
  }
  const booths = getAllBooths();
  const eventsMap = getEventsMap();
  const lower = keyword.toLowerCase();
  const results = booths.filter(b => {
    const combined = `${b.circle} ${b.works}`.toLowerCase();
    return combined.indexOf(lower) !== -1;
  }).slice(0, MAX_LIST_ITEMS);
  if (!results.length) {
    return '找不到符合的攤位，請換關鍵字試試。';
  }
  return results.map(b => {
    const event = eventsMap[b.eventId];
    const eventName = event ? event.name : '未知場次';
    return `【${eventName}】 ${b.code}｜${b.circle}｜${b.works}`;
  }).join('\n');
}

function handleFavorite(user, code, isAdd) {
  if (!user.lastEventId) {
    return '請先輸入「攤位 場次關鍵字」設定要收藏的場次。';
  }
  const booth = findBoothByEventAndCode(user.lastEventId, code);
  if (!booth) {
    return '找不到該攤位，請確認代碼或重新輸入「攤位 場次關鍵字」。';
  }
  const favorite = getFavorite(user.id, booth.id);
  const eventsMap = getEventsMap();
  const eventName = eventsMap[user.lastEventId] ? eventsMap[user.lastEventId].name : (user.lastEventName || '未知場次');
  if (isAdd) {
    if (favorite) {
      return `你已收藏過 ${booth.code}｜${booth.circle}。`;
    }
    addFavorite(user.id, user.lastEventId, booth.id);
    return `已收藏：${booth.code}｜${booth.circle}（${eventName}）`;
  }
  if (!favorite) {
    return '尚未收藏該攤位。';
  }
  removeFavorite(favorite.id);
  return `已取消收藏：${booth.code}｜${booth.circle}`;
}

function handleReminder(user, code, mins) {
  if (!user.lastEventId) {
    return '請先輸入「攤位 場次關鍵字」。';
  }
  if (!mins || mins <= 0 || mins > 720) {
    return '提醒分鐘需介於 1 至 720。';
  }
  const booth = findBoothByEventAndCode(user.lastEventId, code);
  if (!booth) {
    return '找不到該攤位，請確認代碼或先輸入「攤位 場次關鍵字」。';
  }
  const favorite = getFavorite(user.id, booth.id);
  if (!favorite) {
    return '尚未收藏此攤位，請先輸入「收藏 ' + code + '」。';
  }
  updateFavoriteReminder(favorite.id, mins);
  return `已設定提醒：${booth.code} 提前 ${mins} 分鐘。`;
}

function listUserFavorites(user) {
  const favorites = getFavoritesByUser(user.id);
  if (!favorites.length) {
    return '目前沒有收藏的攤位，先輸入「攤位 關鍵字」探索吧！';
  }
  const eventsMap = getEventsMap();
  const boothsMap = getBoothsMap();
  const upcoming = favorites.filter(fav => {
    const event = eventsMap[fav.eventId];
    if (!event) {
      return false;
    }
    const startDate = parseDate(event.start);
    if (!startDate) {
      return true;
    }
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    return startDate >= today;
  });
  if (!upcoming.length) {
    return '收藏中沒有即將到來的攤位。';
  }
  const lines = upcoming.map(fav => {
    const event = eventsMap[fav.eventId];
    const booth = boothsMap[fav.boothId];
    const eventName = event ? event.name : '未知場次';
    const code = booth ? booth.code : '未知代碼';
    const circle = booth ? booth.circle : '未知社團';
    return `【${eventName}】 ${code}｜${circle}`;
  });
  return lines.join('\n');
}

/**
 * 每日 09:00 觸發的提醒任務。
 */
function dailyTomorrowFavorites() {
  try {
    const events = listEvents();
    const boothsMap = getBoothsMap();
    const users = listUsers();
    const targetDate = new Date();
    targetDate.setDate(targetDate.getDate() + 1);
    targetDate.setHours(0, 0, 0, 0);

    const eventsMap = {};
    events.forEach(ev => {
      const start = parseDate(ev.start);
      if (!start) {
        return;
      }
      if (start.getTime() === targetDate.getTime()) {
        eventsMap[ev.id] = ev;
      }
    });

    const favorites = listFavorites();
    const favoritesByUser = {};
    favorites.forEach(fav => {
      if (!eventsMap[fav.eventId]) {
        return;
      }
      if (!favoritesByUser[fav.userId]) {
        favoritesByUser[fav.userId] = [];
      }
      favoritesByUser[fav.userId].push(fav);
    });

    Object.keys(favoritesByUser).forEach(userId => {
      const favs = favoritesByUser[userId];
      if (!favs.length) {
        return;
      }
      const user = users.find(u => u.id === userId);
      if (!user) {
        return;
      }
      const messages = buildReminderMessages(favs, eventsMap, boothsMap);
      if (messages.length) {
        pushMessage(user.lineUserId, messages);
      }
    });
  } catch (error) {
    Logger.log('dailyTomorrowFavorites error: %s', error && error.stack ? error.stack : error);
  }
}

function buildReminderMessages(favorites, eventsMap, boothsMap) {
  const messages = [];
  favorites.forEach(fav => {
    const event = eventsMap[fav.eventId];
    const booth = boothsMap[fav.boothId];
    if (!event || !booth) {
      return;
    }
    const timeText = event.openAt ? event.openAt : '早上';
    const text = `明天是《${event.name}》！別忘了你收藏的：${booth.code}｜${booth.circle}｜${booth.works} 時間：${timeText} 地點：${event.location}`;
    messages.push(buildTextMessage(text));
  });
  return messages;
}

/**
 * 從 LINE 取得使用者資料。
 */
function getUserProfile(userId) {
  try {
    const token = getConfigValue('CHANNEL_TOKEN');
    if (!token) {
      return null;
    }
    const url = `https://api.line.me/v2/bot/profile/${encodeURIComponent(userId)}`;
    const response = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: { Authorization: 'Bearer ' + token },
      muteHttpExceptions: true
    });
    if (response.getResponseCode() === 200) {
      return JSON.parse(response.getContentText());
    }
  } catch (error) {
    Logger.log('getUserProfile error: %s', error && error.stack ? error.stack : error);
  }
  return null;
}

function ensureUser(lineUserId, name) {
  const users = listUsers();
  let user = users.find(u => u.lineUserId === lineUserId);
  if (user) {
    return user;
  }
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(SHEET_USERS);
  if (!sheet) {
    throw new Error('Users 工作表不存在，請先執行 initSheets()');
  }
  const id = generateId('USR');
  const nowIso = new Date().toISOString();
  sheet.appendRow([id, lineUserId, name, '', '', nowIso]);
  user = { id, lineUserId, name, lastEventId: '', lastEventName: '', createdAt: nowIso };
  return user;
}

function updateUserLastEvent(userId, eventId, eventName) {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(SHEET_USERS);
  if (!sheet) {
    return;
  }
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === userId) {
      sheet.getRange(i + 1, 4, 1, 2).setValues([[eventId, eventName]]);
      return;
    }
  }
}

function listUsers() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(SHEET_USERS);
  if (!sheet) {
    return [];
  }
  const data = sheet.getDataRange().getValues();
  const users = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    users.push({
      id: row[0],
      lineUserId: row[1],
      name: row[2],
      lastEventId: row[3],
      lastEventName: row[4],
      createdAt: row[5]
    });
  }
  return users;
}

function listEvents() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(SHEET_EVENTS);
  if (!sheet) {
    return [];
  }
  const data = sheet.getDataRange().getValues();
  const events = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    events.push({
      id: row[0],
      name: row[1],
      start: row[2],
      end: row[3],
      location: row[4],
      url: row[5],
      openAt: row[6],
      note: row[7],
      createdAt: row[8]
    });
  }
  return events;
}

function listBooths() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(SHEET_BOOTHS);
  if (!sheet) {
    return [];
  }
  const data = sheet.getDataRange().getValues();
  const booths = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    booths.push({
      id: row[0],
      eventId: row[1],
      code: row[2],
      circle: row[3],
      works: row[4],
      zone: row[5],
      tableNo: row[6],
      start: row[7],
      end: row[8],
      note: row[9]
    });
  }
  return booths;
}

function listFavorites() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(SHEET_FAVORITES);
  if (!sheet) {
    return [];
  }
  const data = sheet.getDataRange().getValues();
  const favorites = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    favorites.push({
      id: row[0],
      userId: row[1],
      eventId: row[2],
      boothId: row[3],
      remindMins: row[4],
      createdAt: row[5]
    });
  }
  return favorites;
}

function listUpcomingEvents(count) {
  const events = listEvents();
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const until = new Date(today.getTime() + MAX_EVENT_LOOKAHEAD_DAYS * 24 * 60 * 60 * 1000);
  return events.filter(ev => {
    const start = parseDate(ev.start);
    if (!start) {
      return false;
    }
    return start >= today && start <= until;
  }).sort((a, b) => {
    const aDate = parseDate(a.start);
    const bDate = parseDate(b.start);
    return aDate - bDate;
  }).slice(0, count);
}

function listBoothsByEvent(eventId) {
  return listBooths().filter(booth => booth.eventId === eventId);
}

function getAllBooths() {
  return listBooths();
}

function getEventsMap() {
  const events = listEvents();
  const map = {};
  events.forEach(ev => {
    map[ev.id] = ev;
  });
  return map;
}

function getBoothsMap() {
  const booths = listBooths();
  const map = {};
  booths.forEach(b => {
    map[b.id] = b;
  });
  return map;
}

function findEventByKeyword(keyword) {
  if (!keyword) {
    return null;
  }
  const lower = keyword.toLowerCase();
  const events = listEvents();
  return events.find(ev => ev.name.toLowerCase().indexOf(lower) !== -1) || null;
}

function findBoothByEventAndCode(eventId, code) {
  const normalizedCode = code.toUpperCase();
  const booths = listBoothsByEvent(eventId);
  return booths.find(b => b.code.toUpperCase() === normalizedCode) || null;
}

function getFavorite(userId, boothId) {
  const favorites = listFavorites();
  return favorites.find(fav => fav.userId === userId && fav.boothId === boothId) || null;
}

function addFavorite(userId, eventId, boothId) {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(SHEET_FAVORITES);
  if (!sheet) {
    throw new Error('Favorites 工作表不存在');
  }
  const id = generateId('FAV');
  const nowIso = new Date().toISOString();
  sheet.appendRow([id, userId, eventId, boothId, DEFAULT_REMIND_MINS, nowIso]);
}

function removeFavorite(favoriteId) {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(SHEET_FAVORITES);
  if (!sheet) {
    return;
  }
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === favoriteId) {
      sheet.deleteRow(i + 1);
      return;
    }
  }
}

function updateFavoriteReminder(favoriteId, mins) {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(SHEET_FAVORITES);
  if (!sheet) {
    return;
  }
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === favoriteId) {
      sheet.getRange(i + 1, 5).setValue(mins);
      return;
    }
  }
}

function getFavoritesByUser(userId) {
  return listFavorites().filter(fav => fav.userId === userId);
}

/**
 * LINE 推播與回覆。
 */
function replyMessage(replyToken, messages) {
  try {
    const token = getConfigValue('CHANNEL_TOKEN');
    if (!token) {
      Logger.log('CHANNEL_TOKEN 未設定');
      return;
    }
    UrlFetchApp.fetch('https://api.line.me/v2/bot/message/reply', {
      method: 'post',
      headers: {
        'Content-Type': 'application/json',
        Authorization: 'Bearer ' + token
      },
      payload: JSON.stringify({ replyToken, messages }),
      muteHttpExceptions: true
    });
  } catch (error) {
    Logger.log('replyMessage error: %s', error && error.stack ? error.stack : error);
  }
}

function pushMessage(to, messages) {
  try {
    const token = getConfigValue('CHANNEL_TOKEN');
    if (!token) {
      Logger.log('CHANNEL_TOKEN 未設定');
      return;
    }
    UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
      method: 'post',
      headers: {
        'Content-Type': 'application/json',
        Authorization: 'Bearer ' + token
      },
      payload: JSON.stringify({ to, messages }),
      muteHttpExceptions: true
    });
  } catch (error) {
    Logger.log('pushMessage error: %s', error && error.stack ? error.stack : error);
  }
}

function buildTextMessage(text) {
  return { type: 'text', text: truncateText(text, 2000) };
}

function truncateText(text, maxLength) {
  if (!text) {
    return '';
  }
  if (text.length <= maxLength) {
    return text;
  }
  return text.substring(0, maxLength - 1) + '…';
}

/**
 * 管理後台 Dashboard HTML。
 */
function renderAdminDashboard() {
  const users = listUsers();
  const events = listEvents();
  const booths = listBooths();
  const favorites = listFavorites();
  const spreadsheetUrl = SpreadsheetApp.getActive().getUrl();
  const template = HtmlService.createTemplateFromFile('admin');
  template.userCount = users.length;
  template.eventCount = events.length;
  template.boothCount = booths.length;
  template.favoriteCount = favorites.length;
  template.spreadsheetUrl = spreadsheetUrl;
  const output = template.evaluate();
  output.addMetaTag('viewport', 'width=device-width, initial-scale=1');
  output.setTitle('同人場行事曆 Dashboard');
  return output;
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * 工具函式。
 */
function createTextResponse(text) {
  return ContentService.createTextOutput(text);
}

function parseDate(value) {
  if (!value) {
    return null;
  }
  const date = new Date(value);
  if (isNaN(date.getTime())) {
    return null;
  }
  date.setHours(0, 0, 0, 0);
  return date;
}

function formatDate(date) {
  return Utilities.formatDate(date, TIMEZONE, 'yyyy-MM-dd');
}

function formatEventDate(event) {
  const start = event.start ? event.start : '';
  const end = event.end ? event.end : '';
  if (start && end) {
    return `${start} ~ ${end}`;
  }
  return start || '日期待定';
}

function generateId(prefix) {
  const random = Math.floor(Math.random() * 1e6).toString().padStart(6, '0');
  const timestamp = Utilities.formatDate(new Date(), TIMEZONE, 'yyyyMMddHHmmss');
  return `${prefix}${timestamp}${random}`;
}

/**
 * 版本資訊：主要函式索引。
 */
function listFunctionIndex() {
  return [
    'initSheets：建立必要工作表與欄位',
    'seedSample：建立範例場次與攤位',
    'doPost：LINE Webhook 入口',
    'doGet：後台與健康檢查入口',
    'dailyTomorrowFavorites：每日提醒任務',
    'handleLineEvent / onTextMessage：指令解析與回覆',
    'renderAdminDashboard：後台統計頁面'
  ];
}
