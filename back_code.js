/************************************
 * SVOY SHOP — Google Apps Script backend
 ************************************/

// ====== НАСТРОЙКИ ======
const BOT_TOKEN      = '8493140119:AAEEm0Ka5iqTsIDOpWlqydSrYeungZ7_AGk';
const SPREADSHEET_ID = '1kUhsycMz9fHYx_vwK_A70li6OVUdC4Ac34Cp6QM8tZQ';
const IMGBB_API_KEY = '3019595232f385628b1378a5d5d8f9ba';
// список админов по телефону (11 цифр, без "+")
const ADMIN_PHONES   = ['77782031551'];
const SPREADSHEET_REDEMPTIONS = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Redemptions');

// ====== ОБЩИЕ УТИЛИТЫ ======
function normPhone(raw) {
  let d = String(raw||'').replace(/[^\d]/g,'');
  if (d.startsWith('8') && d.length===11) d = '7'+d.slice(1);
  if (d.length===10) d = '7'+d;
  if (d.length>11) d = d.slice(-11);
  return d;
}
function prettyPhone(p11){
  if (!p11 || String(p11).length!==11) return String(p11||'');
  return `+7 ${p11.slice(1,4)} ${p11.slice(4,7)} ${p11.slice(7,9)} ${p11.slice(9,11)}`;
}
function toNum(x){ const n = Number(x); return isNaN(n)?0:n; }
function truthy(x){ const s=String(x||'').trim().toLowerCase(); return ['true','1','yes','y','да','on'].includes(s); }
function hash_(s){ return Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s)); }
function SS(){ return SpreadsheetApp.openById(SPREADSHEET_ID); }
function sh(name){ return SS().getSheetByName(name); }

function ensureSheet(name, headers){
  let s = sh(name);
  if (!s){ s = SS().insertSheet(name); s.appendRow(headers); }
  const first = s.getRange(1,1,1,s.getLastColumn()).getValues()[0].map(x=>String(x).trim());
  const miss = headers.filter(h => !first.includes(h));
  if (miss.length){
    s.insertColumnsAfter(s.getLastColumn(), miss.length);
    s.getRange(1,1,1, first.length+miss.length).setValues([first.concat(miss)]);
  }
  return s;
}

// ====== SETTINGS (курс балла и коэффициенты) ======
function getSettingsSheet_(){ return sh('Settings'); }

function getSetting_(key, defVal){
  const s = getSettingsSheet_(); if (!s || s.getLastRow()<2) return defVal;
  const vals = s.getRange(2,1,s.getLastRow()-1,2).getValues();
  for (let i=0;i<vals.length;i++){
    const k = String(vals[i][0]||'').trim();
    if (k === key) return vals[i][1];
  }
  return defVal;
}

function setSettingIfMissing_(key, value){
  const s = getSettingsSheet_(); if (!s) return;
  const last = s.getLastRow();
  if (last < 2){
    s.getRange(2,1,1,2).setValues([[key, value]]);
    return;
  }
  const vals = s.getRange(2,1,last-1,2).getValues();
  for (let i=0;i<vals.length;i++){
    if (String(vals[i][0]||'').trim() === key) return; // уже есть
  }
  s.appendRow([key, value]);
}

// Инициализация дефолтов (можете править потом прямо в листе Settings)
function initDefaultSettings_(){
  setSettingIfMissing_('BALL_RATE', 2);       // 1 балл = 2 тенге (пример)
  // Коэффициенты начисления в долях (0.07% = 0.0007; 0.01% = 0.0001)
  setSettingIfMissing_('COEFF_KV', 0.0007);   // КВ
  setSettingIfMissing_('COEFF_VP', 0.0007);   // ВП
  setSettingIfMissing_('COEFF_KP', 0.0001);   // КП
  setSettingIfMissing_('COEFF_PM', 0.0001);   // ПМ
}

// Текущий курс балла (тенге за 1 балл)
function getBallRate_() {
  const s = sh('Settings');
  const vals = s.getDataRange().getValues();
  const rateRow = vals.find(r => String(r[0]).trim() === 'BALL_RATE');
  return rateRow ? toNum(rateRow[1]) : 1;
}

function getCoeffByPremise_(type) {
  const s = sh('Settings');
  const vals = s.getDataRange().getValues();
  const map = {
    'КВ': 'COEFF_KV',
    'ВП': 'COEFF_VP',
    'ПМ': 'COEFF_PM',
    'КП': 'COEFF_KP'
  };
  const key = map[type];
  const row = vals.find(r => String(r[0]).trim() === key);
  return row ? toNum(row[1]) : 0;
}

function initSheets(){
  ensureSheet('Users', ['telegram_user_id','telegram_chat_id','tg_verify_code','tg_verify_expires','tg_verified_at','phone','full_name','dob','gender','created_at','password_hash','is_admin','reset_code','reset_expires','tg_link_code','tg_link_expires']);
  ensureSheet('Purchases',   ['contract_id','phone','permise_type','price','points','status','updated_at','comment']);
  ensureSheet('Catalog',     ['item_id','category','title','desc','points_price','stock','photo_url','is_active','description','price_tenge']);
  ensureSheet('Redemptions', ['redeem_id','phone','item_id','title','points_spent','status','created_at','manager_comment','pickup_code','delivered_at']);
  ensureSheet('Settings', ['KEY','VALUE']);
  initDefaultSettings_();
  return 'OK';
}

function requireAdmin_(token){
  const me = sessionGet_(token);
  if (!me || !me.is_admin) throw new Error('forbidden');
  return me;
}

function getTelegramIdByPhone_(phone11){
  const u = usersFindByPhone_(phone11);
  const chat = u && String(u.telegram_user_id||'').trim();
  return chat || '';
}

function sendTelegramMessage_(chat_id, text){
  if (!chat_id) return;
  const url = 'https://api.telegram.org/bot' + BOT_TOKEN + '/sendMessage';
  const payload = { chat_id, text };
  const params  = { method:'post', contentType:'application/json', payload: JSON.stringify(payload), muteHttpExceptions:true };
  UrlFetchApp.fetch(url, params);
}

function notifyBalanceChange_(phone11, kind, points, opts){
  try{
    const chat = getTelegramIdByPhone_(normPhone(phone11));
    if (!chat) return;
    const cid = opts && opts.contract_id ? String(opts.contract_id).trim() : '';
    const suffix = cid ? ` по договору ${cid}` : '';
    const pts = Number(points)||0;

    if (kind === 'pending'){
      sendTelegramMessage_(chat, `➕ Начисление${suffix}: ${pts} баллов.\nСтатус: ожидает подтверждения.`);
    } else if (kind === 'available'){
      sendTelegramMessage_(chat, `✅ Подтверждено начисление${suffix}: ${pts} баллов.\nБаланс доступен к использованию.`);
    } else if (kind === 'termination'){
      sendTelegramMessage_(chat, `⚠️ Расторжение${suffix}.\nОжидаемые баллы отменены.`);
    }
  } catch(e){
    Logger.log('notifyBalanceChange_ error: '+e);
  }
}

// ====== СЕССИИ ======
function sessStore_(){ return PropertiesService.getScriptProperties(); }
function sessionCreate_(user){
  const token = Utilities.getUuid();
  sessStore_().setProperty('sess_'+token, JSON.stringify({ phone: user.phone, full_name: user.full_name, is_admin: !!user.is_admin, t: Date.now() }));
  return token;
}
function sessionGet_(token){
  if (!token) return null;
  const raw = sessStore_().getProperty('sess_'+token);
  if (!raw) return null;
  try { return JSON.parse(raw); } catch(_){ return null; }
}
function sessionDelete_(token){ if (token) sessStore_().deleteProperty('sess_'+token); }

// ====== USERS ======
function usersAll_(){
  const s = sh('Users'); if (!s || s.getLastRow()<2) return [];
  const vals = s.getRange(1,1,s.getLastRow(), s.getLastColumn()).getValues();
  const hdr = vals[0].map(x=>String(x).trim());
  return vals.slice(1).filter(r=>r.join('')!=='').map(r => Object.fromEntries(hdr.map((h,i)=>[h,r[i]])));
}
function usersFindByPhone_(p11){
  const p = normPhone(p11);
  return usersAll_().find(u => normPhone(u.phone)===p) || null;
}
function usersAppend_(obj){
  const s = sh('Users');
  const hdr = s.getRange(1,1,1,s.getLastColumn()).getValues()[0].map(x=>String(x).trim());
  const row = hdr.map(h => obj[h] ?? '');
  s.appendRow(row);
}

function getDriveImageDataUrl(fileId){
  var file = DriveApp.getFileById(String(fileId));
  var blob = file.getBlob();
  var ct = blob.getContentType(); // например image/png
  var base64 = Utilities.base64Encode(blob.getBytes());
  return 'data:' + ct + ';base64,' + base64;
}

// ====== КАТАЛОГ ======
function getCatalogActive_(){
  const s = sh('Catalog'); if (!s || s.getLastRow()<2) return [];
  const v = s.getRange(1,1,s.getLastRow(), s.getLastColumn()).getValues();
  const h = v[0].map(x=>String(x).trim());
  const idI=h.indexOf('item_id'), catI=h.indexOf('category'), titleI=h.indexOf('title'),
        descI=h.indexOf('desc'), priceI=h.indexOf('points_price'), stockI=h.indexOf('stock'),
        photoI=h.indexOf('photo_url'), actI=h.indexOf('is_active');

  return v.slice(1).filter(r=>r.join('')!=='').map(r=>({
    item_id: String(r[idI]||''),
    category: r[catI]||'',
    title: String(r[titleI]||'').trim(),
    desc:  String(r[descI]||'').trim(),
    points_price: toNum(r[priceI]),
    stock: toNum(r[stockI]),
    photo_url: String(r[photoI]||'').trim(),
    is_active: truthy(r[actI]),
  })).filter(it => it.is_active && it.stock>0 && it.title && it.points_price>0);
}

// ====== БАЛАНС: РАСШИРЕННЫЕ МЕТРИКИ ======
function getBalanceStatsByPhone_(p11){
  const ss = SS();
  const phone = normPhone(p11);

  let earned_total = 0;   // все начисления, которые когда-либо стали доступными
  let pending      = 0;   // ещё не подтверждены (не входят в available)
  // Redemptions:
  let hold_now     = 0;   // waiting + approved (резервы сейчас)
  let spent_total  = 0;   // delivered (потрачено навсегда)
  // для формулы available:
  let waitingApprovedDelivered = 0;

  // Purchases
  const sP = ss.getSheetByName('Purchases');
  if (sP && sP.getLastRow()>=2){
    const v = sP.getRange(1,1,sP.getLastRow(), sP.getLastColumn()).getValues();
    const h = v[0].map(x=>String(x).trim());
    const m = Object.fromEntries(h.map((k,i)=>[k,i]));

    // сгруппируем все записи по contract_id
    const deals = {}; // contract_id -> массив записей этой сделки

    for (let i=1;i<v.length;i++){
      const r=v[i]; if (!r.join('')) continue;
      if (normPhone(r[m.phone])!==phone) continue;

      const contract_id = String(r[m.contract_id]||'').trim();
      if (!contract_id) continue;

      if (!deals[contract_id]) deals[contract_id] = [];

      deals[contract_id].push({
        status: String(r[m.status]||'').trim(),
        points: toNum(r[m.points]),
        ts:     parseTs_(r[m.updated_at])
      });
    }

    // теперь по каждой сделке смотрим её историю
    Object.keys(deals).forEach(cid=>{
      const recs = deals[cid];
      if (!recs.length) return;

      // отсортируем по времени от старых к новым
      recs.sort((a,b)=>a.ts - b.ts);

      // найдём последний статус
      const last = recs[recs.length-1];

      // найдём последний credited_available в истории, если был
      let lastAvailable = null;
      for (let i=0;i<recs.length;i++){
        if (recs[i].status === 'credited_available'){
          lastAvailable = recs[i]; // перезаписываем, чтобы в конце был реально последний available
        }
      }

      if (lastAvailable){
        // когда-либо сделка была подтверждена → эти баллы принадлежат человеку навсегда
        earned_total += lastAvailable.points;
      } else {
        // подтверждения не было ни разу
        // тогда если последний статус = credited_pending → это в ожидании
        if (last.status === 'credited_pending'){
          pending += last.points;
        }
        // если последний статус termination, и не было credited_available,
        // то ничего не добавляем (сделка умерла до начисления)
      }
    });
  }

  // Redemptions
  const sR = ss.getSheetByName('Redemptions');
  if (sR && sR.getLastRow()>=2){
    const v = sR.getRange(1,1,sR.getLastRow(), sR.getLastColumn()).getValues();
    const h = v[0].map(x=>String(x).trim());
    const pI=h.indexOf('phone'), ptsI=h.indexOf('points_spent'), stI=h.indexOf('status');
    for (let i=1;i<v.length;i++){
      const r=v[i]; if (!r.join('')) continue;
      if (normPhone(r[pI])!==phone) continue;
      const pts = toNum(r[ptsI]);
      const st  = String(r[stI]||'');
      if (st==='waiting' || st==='approved' || st==='ready' || st==='await_code' || st==='delivered') waitingApprovedDelivered += pts;
      if (st==='waiting' || st==='approved' || st==='ready' || st==='await_code') hold_now += pts;
      if (st==='delivered') spent_total += pts;                    // потрачено навсегда
    }
  }

  const available = Math.max(0, earned_total - waitingApprovedDelivered);
  return { available, pending, earned_total, spent_total, hold_now };
}

// ====== API: РЕГИСТРАЦИЯ / ВХОД / ВЫХОД ======
function api_register(payload){
  initSheets();
  const phone = normPhone(payload.phone);
  const full_name = String(payload.full_name||'').trim();
  const password  = String(payload.password||'').trim();
  const dob       = String(payload.dob||'').trim();
  const gender    = String(payload.gender||'').trim();

  if (!phone || phone.length!==11) throw new Error('Укажите корректный телефон');
  if (!full_name) throw new Error('Укажите ФИО');
  if (usersFindByPhone_(phone)) throw new Error('Этот телефон уже зарегистрирован');

  usersAppend_({
    telegram_user_id: '',
    phone, full_name, dob, gender,
    created_at: new Date(),
    password_hash: hash_(password),
    is_admin: ADMIN_PHONES.includes(phone)
  });
  return { ok:true };
}

function api_login(payload){
  const phone = normPhone(payload.phone);
  const password = String(payload.password||'');
  const u = usersFindByPhone_(phone);
  if (!u) throw new Error('Пользователь не найден');
  if (String(u.password_hash) !== hash_(password)) throw new Error('Неверный пароль');
  const token = sessionCreate_({ phone, full_name: u.full_name, is_admin: truthy(u.is_admin) || ADMIN_PHONES.includes(phone) });
  return { token, is_admin: truthy(u.is_admin) || ADMIN_PHONES.includes(phone) };
}

function api_logout(token){ sessionDelete_(token); return { ok:true }; }

// ====== API: ДАШБОРД КЛИЕНТА ======
function api_getDashboard(token){
  const me = sessionGet_(token); if (!me) throw new Error('auth required');
  const phone = normPhone(me.phone);

  const stats = getBalanceStatsByPhone_(phone); // {available, pending, earned_total, spent_total, hold_now}
  const catalog = getCatalogActive_().map(it=>{
    const can = stats.available >= it.points_price;
    return Object.assign({}, it, { can_afford: can, missing: can?0:(it.points_price-stats.available) });
  });

  // мои заявки
  const sR = sh('Redemptions');
  const red = [];
  if (sR && sR.getLastRow()>=2){
    const v = sR.getRange(1,1,sR.getLastRow(), sR.getLastColumn()).getValues();
    const h = v[0].map(x=>String(x).trim());
    const m = Object.fromEntries(h.map((k,i)=>[k,i]));
    for (let i=1;i<v.length;i++){
      const r=v[i]; if (!r.join('')) continue;
      if (normPhone(r[m.phone])===phone){
        const createdTs = parseTs_(r[m.created_at]);
        red.push({
          redeem_id: r[m.redeem_id],
          title: String(r[m.title]||''),
          points_spent: toNum(r[m.points_spent]),
          status: String(r[m.status]||''),
          created_at: r[m.created_at], // сырое значение из шита (Date/строка/число)
          created_at_ts: createdTs,    // миллисекунды — удобно сортировать
          created_at_display: Utilities.formatDate(
            new Date(createdTs),
            Session.getScriptTimeZone(),
            'dd.MM.yyyy HH:mm'
          )
        });
      }
    }
  }

  return {
    full_name: me.full_name || 'Покупатель',
    phone,
    phone_pretty: prettyPhone(phone),
    is_admin: !!me.is_admin,
    // баланс-метрики
    balance: stats.available,                 // для обратной совместимости
    balance_stats: stats,                     // новые поля
    catalog,
    redemptions: red.slice(-20).reverse()
  };
}

function usersUpdateByPhone_(phone11, patchObj){
  const p = normPhone(phone11);
  const s = sh('Users');
  if (!s || s.getLastRow() < 2) return false;

  const rng = s.getDataRange();
  const vals = rng.getValues();
  const hdr = vals[0].map(x=>String(x).trim());

  // карта имяКолонки -> индекс
  const colIndex = {};
  hdr.forEach((h,i)=>{ colIndex[h]=i; });

  for (let r = 1; r < vals.length; r++){
    const row = vals[r];
    if (!row.join('')) continue;
    if (normPhone(row[colIndex['phone']]) === p){
      // применяем patchObj
      Object.keys(patchObj).forEach(k=>{
        if (colIndex.hasOwnProperty(k)){
          row[colIndex[k]] = patchObj[k];
        }
      });
      vals[r] = row;
      rng.setValues(vals); // перезапись всего листа (для простоты)
      return true;
    }
  }
  return false;
}

// ====== TELEGRAM VERIFY ======
function api_generateTelegramCode(token){
  // 1) проверка сессии
  const me = sessionGet_(token);
  if (!me) throw new Error('auth required');

  const phone = normPhone(me.phone);
  const u = usersFindByPhone_(phone);
  if (!u) throw new Error('not found');

  // если уже привязан телеграм — не генерим новый код
  if (u.telegram_user_id && String(u.telegram_user_id).trim() !== ''){
    return { linked: true, code: '' };
  }

  // генерим короткий код, напр. TG-123456
  const rnd = Math.floor(100000 + Math.random()*900000); // шестизначный
  const linkCode = 'TG-' + rnd;

  // код живёт, скажем, 10 минут
  const expiresAt = Date.now() + 10*60*1000;

  usersUpdateByPhone_(phone, {
    tg_link_code: linkCode,
    tg_link_expires: expiresAt
  });

  return { linked:false, code: linkCode, valid_till: expiresAt };
}

function api_requestPasswordReset(payload) {
  const phone = normPhone(payload.phone);
  if (!phone) throw new Error('phone required');

  const u = usersFindByPhone_(phone);
  if (!u) return { status: 'not_found', message: 'Пользователь с таким телефоном не найден' };

  if (!u.telegram_user_id) {
    return { status: 'no_telegram', message: 'У аккаунта не привязан Telegram. Сначала свяжите Telegram в личном кабинете.' };
  }

  const code = 'RP-' + Math.floor(100000 + Math.random()*900000);
  const expiresAt = Date.now() + 10*60*1000; // 10 минут

  usersUpdateByPhone_(phone, { reset_code: code, reset_expires: expiresAt });

  // отправка в TG
  const text = '🔐 Восстановление пароля SVOY SHOP\n' +
               'Код подтверждения: ' + code + '\n\n' +
               'Срок действия ~10 минут.\n' +
               'Введите этот код на сайте, чтобы задать новый пароль.';
  sendTelegramMessage_(String(u.telegram_user_id), text);

  return { status: 'sent' };
}

function api_confirmPasswordReset(payload) {
  const phone = normPhone(payload.phone);
  const code  = String(payload.code || '').trim();
  const new_password = String(payload.new_password || '').trim();

  if (!phone || !code || !new_password) throw new Error('phone, code, new_password required');

  const u = usersFindByPhone_(phone);
  if (!u) return { status: 'not_found', message: 'Пользователь не найден' };

  const storedCode = String(u.reset_code || '').trim();
  const exp = Number(u.reset_expires || 0);
  if (!storedCode || !exp) return { status: 'bad_code', message: 'Код не запрашивался либо уже использован' };
  if (storedCode !== code) return { status: 'bad_code', message: 'Неверный код' };
  if (Date.now() > exp)   return { status: 'bad_code', message: 'Код просрочен. Запросите новый.' };

  usersUpdateByPhone_(phone, {
    password_hash: hash_(new_password),
    reset_code: '',
    reset_expires: ''
  });

  return { status: 'ok' };
}

function api_checkTelegramLink(token){
  const me = sessionGet_(token);
  if (!me) throw new Error('auth required');

  const phone = normPhone(me.phone);
  const u = usersFindByPhone_(phone);
  if (!u) throw new Error('not found');

  const linked = !!(u.telegram_user_id && String(u.telegram_user_id).trim() !== '');

  return { linked: linked };
}

function api_telegramConfirm(payload){
  // payload = { code: 'TG-123456', chat_id: '123456789' }

  const code = String(payload.code || '').trim();
  const chat_id = String(payload.chat_id || '').trim();
  if (!code || !chat_id) throw new Error('missing code/chat_id');

  const s = sh('Users');
  if (!s || s.getLastRow()<2) throw new Error('no users');

  const rng = s.getDataRange();
  const vals = rng.getValues();
  const hdr = vals[0].map(x=>String(x).trim());

  const col = {};
  hdr.forEach((h,i)=>{ col[h]=i; });

  let updated = false;
  const now = Date.now();

  for (let r=1; r<vals.length; r++){
    const row = vals[r];
    if (!row.join('')) continue;

    const rowCode   = String(row[col['tg_link_code']]||'').trim();
    const rowExpire = Number(row[col['tg_link_expires']]||0);

    if (rowCode === code){
      // проверим не истёк ли код
      if (rowExpire && now > rowExpire){
        throw new Error('code expired');
      }

      // ок, связываем
      row[col['telegram_user_id']]  = chat_id;
      row[col['tg_link_code']]      = '';
      row[col['tg_link_expires']]   = '';

      vals[r] = row;
      updated = true;
      break;
    }
  }

  if (!updated) throw new Error('code not found');

  rng.setValues(vals);
  return { ok:true };
}

// ====== API: СОЗДАНИЕ ЗАЯВКИ ======
function api_createRedemption(token, item_id){
  const me = sessionGet_(token); if (!me) throw new Error('auth required');
  const phone = normPhone(me.phone);
  // проверяем телеграм
  const u = usersFindByPhone_(phone);
  if (!u) throw new Error('user not found');
  const isLinkedToTelegram = !!(u.telegram_user_id && String(u.telegram_user_id).trim() !== '');
  if (!isLinkedToTelegram){
    throw new Error('Для оформления заявки сначала привяжите Telegram в кабинете');
  }

  const stats = getBalanceStatsByPhone_(phone);
  const available = stats.available;

  const cat = getCatalogActive_();
  const it = cat.find(x => String(x.item_id)===String(item_id));
  if (!it) throw new Error('Товар недоступен');
  if (toNum(it.points_price)>available) throw new Error('Недостаточно баллов');
  if (toNum(it.stock)<=0) throw new Error('Нет в наличии');

  // уменьшить stock
  const sC = sh('Catalog');
  const v = sC.getDataRange().getValues(); const h=v[0].map(x=>String(x).trim());
  const idI=h.indexOf('item_id'), stI=h.indexOf('stock');
  for (let r=1;r<v.length;r++){
    if (String(v[r][idI])===String(item_id)){ v[r][stI]=Math.max(0,toNum(v[r][stI])-1); break; }
  }
  sC.getDataRange().setValues(v);

  // записать заявку
  const sR = sh('Redemptions');
  const hdr = sR.getRange(1,1,1,sR.getLastColumn()).getValues()[0].map(x=>String(x).trim());
  const row = {
    redeem_id: 'R'+Math.floor(100000+Math.random()*900000),
    phone, item_id: String(it.item_id), title: String(it.title),
    points_spent: toNum(it.points_price),
    status: 'waiting', created_at: new Date(), manager_comment: ''
  };
  sR.appendRow(hdr.map(h=>row[h]??''));
  // --- уведомление в Telegram ---
  try {
    const chatId = u.telegram_user_id || getTelegramIdByPhone_(phone);
    if (chatId) {
      const title = String(it.title || 'товар');
      sendTelegramMessage_(chatId, `Заявка на «${title}» создана. Ожидайте подтверждения.`);
      Logger.log(`TG notify ok: redemption created for ${phone} (${title}) -> chatId=${chatId}`);
    } else {
      Logger.log(`TG notify skipped: no telegram_user_id for ${phone}`);
    }
  } catch (e) {
    Logger.log('Ошибка при уведомлении о создании заявки: ' + e);
  }
  return { ok:true };
}

function api_cancelRedemption(token, redeem_id, reason){
  const me = sessionGet_(token); 
  if (!me) throw new Error('auth required');
  const phone = normPhone(me.phone);

  // Лист Redemptions
  const sR = sh('Redemptions');
  if (!sR) throw new Error('Redemptions not found');
  const rng = sR.getDataRange();
  const v   = rng.getValues();
  const h   = v[0].map(x=>String(x).trim());
  const m   = Object.fromEntries(h.map((k,i)=>[k,i]));

  // Найдём строку заявки, которая принадлежит текущему пользователю
  let row = -1;
  for (let i=1;i<v.length;i++){
    if (!v[i].join('')) continue;
    if (String(v[i][m.redeem_id])===String(redeem_id) && normPhone(v[i][m.phone])===phone){
      row = i; break;
    }
  }
  if (row < 0) throw new Error('Заявка не найдена');

  const curStatus = String(v[row][m.status]||'').toLowerCase();
  if (!['waiting','approved','ready','await_code'].includes(curStatus)){
    throw new Error('Эту заявку уже нельзя отменить');
  }

  // Вернуть stock в Catalog по item_id
  try {
    const item_id = String(v[row][m.item_id]||'').trim();
    if (item_id){
      const sC = sh('Catalog');
      const c  = sC.getDataRange().getValues();
      const ch = c[0].map(x=>String(x).trim());
      const cIdI = ch.indexOf('item_id');
      const cStI = ch.indexOf('stock');
      for (let r=1;r<c.length;r++){
        if (String(c[r][cIdI])===item_id){
          c[r][cStI] = toNum(c[r][cStI]) + 1;
          break;
        }
      }
      sC.getDataRange().setValues(c);
    }
  } catch(e){
    Logger.log('stock rollback error: '+e);
  }

  // Обновить статус и комментарий
  v[row][m.status] = 'canceled';
  if (m.pickup_code !== undefined) v[row][m.pickup_code] = '';
  if (m.manager_comment !== undefined) v[row][m.manager_comment] = String(reason||'');

  rng.setValues(v);

  // TG-уведомление (необязательно)
  try{
    const chat = getTelegramIdByPhone_(phone);
    const title = String(v[row][m.title]||'').trim();
    if (chat){
      sendTelegramMessage_(chat, `❌ Заявка «${title}» отменена.\nПричина: ${reason||'не указана'}.`);
    }
  }catch(e){ Logger.log('TG notify cancel error: '+e); }

  return { ok:true };
}

// ====== API: АДМИН — заявки ======
function api_adminListRedemptions(token, statusFilter){
  const me = sessionGet_(token); if (!me || !me.is_admin) throw new Error('forbidden');
  const sR = sh('Redemptions');
  const out = [];
  if (sR && sR.getLastRow()>=2){
    const v = sR.getRange(1,1,sR.getLastRow(), sR.getLastColumn()).getValues();
    const h = v[0].map(x=>String(x).trim());
    const m = Object.fromEntries(h.map((k,i)=>[k,i]));
    for (let i=1;i<v.length;i++){
      const r=v[i]; if (!r.join('')) continue;
      const createdTs = parseTs_(r[m.created_at]);
      const rec = {
        redeem_id: r[m.redeem_id],
        phone: String(r[m.phone]),
        item_id: r[m.item_id],
        title: r[m.title],
        points_spent: toNum(r[m.points_spent]),
        status: String(r[m.status]||''),
        created_at: r[m.created_at],
        created_at_ts: createdTs,
        created_at_display: Utilities.formatDate(
          new Date(createdTs),
          Session.getScriptTimeZone(),
          'dd.MM.yyyy HH:mm'
        )
      };
      if (statusFilter && rec.status!==statusFilter) continue;
      out.push(rec);
    }
  }
  return { items: out.slice(-200).reverse() };
}
function api_adminUpdateRedemption(token, redeem_id, new_status){
  const me = sessionGet_(token); if (!me || !me.is_admin) throw new Error('forbidden');

  const sR = sh('Redemptions');
  const rng = sR.getDataRange();
  const v   = rng.getValues();
  const h   = v[0].map(x => String(x).trim());

  const idI = h.indexOf('redeem_id');
  const stI = h.indexOf('status');
  const pI  = h.indexOf('phone');
  const tI  = h.indexOf('title');

  for (let r = 1; r < v.length; r++){
    if (String(v[r][idI]) === String(redeem_id)){
      v[r][stI] = String(new_status);

      try{
        const phone = normPhone(v[r][pI]);
        const chat  = getTelegramIdByPhone_(phone);
        const title = String(v[r][tI] || '').trim();

        if (chat){
          if (String(new_status) === 'approved'){
            // Было: "готов к выдаче". Стало: подтверждено, ждите инфо по выдаче.
            sendTelegramMessage_(chat,
              `Заказ «${title}» подтверждён. Ожидайте информации по выдаче.`);
          }
          else if (String(new_status) === 'ready'){
            // Новый промежуточный статус
            sendTelegramMessage_(chat,
              `Заказ «${title}» готов к выдаче.\n` +
              `Адрес: Макатаева 168/1, офис SvoyDom.\n` +
              `Пн–Сб 09:00–21:00, Вс 09:00–18:00.`);
          }
          // Остальные статусы как были (await_code → код уходит в другом методе, delivered/canceled — свои уведомления)
        }
      } catch(_){}

      rng.setValues(v);
      return { ok: true };
    }
  }
  throw new Error('Заявка не найдена');
}

// ====== API: АДМИН — каталог ======
function api_adminListCatalog(token){
  const me = sessionGet_(token); if (!me || !me.is_admin) throw new Error('forbidden');
  return { items: getCatalogActive_().concat( getAllCatalogRaw_().filter(it=>!(it.is_active && it.stock>0 && it.points_price>0)) ) };
}
function getAllCatalogRaw_(){
  const s = sh('Catalog'); if (!s || s.getLastRow()<2) return [];
  const v = s.getRange(1,1,s.getLastRow(), s.getLastColumn()).getValues();
  const h = v[0].map(x=>String(x).trim());
  const idI=h.indexOf('item_id'), catI=h.indexOf('category'), titleI=h.indexOf('title'),
        descI=h.indexOf('desc'), priceI=h.indexOf('points_price'), stockI=h.indexOf('stock'),
        photoI=h.indexOf('photo_url'), actI=h.indexOf('is_active');
  return v.slice(1).filter(r=>r.join('')!=='').map(r=>({
    item_id:String(r[idI]||''), category:r[catI]||'', title:String(r[titleI]||'').trim(),
    desc:String(r[descI]||'').trim(), points_price:toNum(r[priceI]), stock:toNum(r[stockI]),
    photo_url:String(r[photoI]||'').trim(), is_active:truthy(r[actI])
  }));
}
function api_adminUpdateCatalogItem(token, payload){
  const me = sessionGet_(token);
  if (!me || !me.is_admin) throw new Error('forbidden');

  const s   = sh('Catalog');
  const rng = s.getDataRange();
  const v   = rng.getValues();
  const h   = v[0].map(x => String(x).trim());

  // индексы важных колонок
  const idI = h.indexOf('item_id');
  const stI = h.indexOf('stock');
  const aI  = h.indexOf('is_active');
  const pI  = h.indexOf('points_price'); // <--- НОВОЕ: колонка с ценой

  if (idI < 0 || stI < 0 || aI < 0 || pI < 0){
    throw new Error('Catalog sheet: missing columns');
  }

  for (let r = 1; r < v.length; r++){
    if (String(v[r][idI]) === String(payload.item_id)){
      v[r][stI] = toNum(payload.stock);
      v[r][aI]  = !!payload.is_active;
      v[r][pI]  = toNum(payload.points_price); // <--- НОВОЕ: сохранить новую цену

      rng.setValues(v);
      return { ok:true };
    }
  }

  throw new Error('Товар не найден');
}

function calcPointsByPremise_(premiseType, price){
  // price — в тенге. Сначала конвертируем в "балловую" базу через курс, затем умножаем на коэффициент.
  // Баллы начисляются по "новым настройкам" на момент добавления сделки.
  const rate   = getBallRate_();                     // тенге за 1 балл (например, 2)
  const coeff  = getCoeffByPremise_(premiseType);    // 0.0007/0.0001 и т.п.
  const p      = toNum(price);

  if (!(p > 0) || !(rate > 0) || !(coeff > 0)) return 0;

  // Пример: price=10 000 000 ₸, rate=2 ₸/балл → 5 000 000 балл-база; * 0.0007 = 3500 баллов
  return Math.round((p / rate) * coeff);
}

// ====== IMAGE: upload to imgbb ======
// dataUrl ("data:image/png;base64,....") -> POST в imgbb -> вернуть {url}
// Требуется IMGBB_API_KEY (см. константу выше)
function uploadImageToImgbb_(dataUrl) {
  if (!dataUrl || String(dataUrl).indexOf('data:') !== 0) {
    return { url: '' };
  }

  // dataUrl формата "data:image/png;base64,AAAA..."
  // нам нужно вытащить только base64 без заголовка
  var parts = String(dataUrl).split(',');
  var base64 = parts[1] || '';

  var payload = {
    key: IMGBB_API_KEY,
    image: base64
    // можно ещё добавить "name": "item_123", но не обязательно
  };

  var options = {
    method: 'post',
    payload: payload,
    muteHttpExceptions: true
  };

  var resp = UrlFetchApp.fetch('https://api.imgbb.com/1/upload', options);
  var code = resp.getResponseCode();
  if (code !== 200) {
    throw new Error('imgbb upload failed: ' + code + ' ' + resp.getContentText());
  }

  var json;
  try {
    json = JSON.parse(resp.getContentText());
  } catch (e) {
    throw new Error('imgbb bad JSON: ' + e);
  }

  // по контракту imgbb возвращает { data: { url: "...", display_url: "...", ... } }
  var url = (json && json.data && (json.data.url || json.data.display_url)) || '';
  if (!url) {
    throw new Error('imgbb: no URL returned');
  }

  return { url: url };
}

function api_adminAddPurchase(token, payload){
  const me = sessionGet_(token); if (!me || !me.is_admin) throw new Error('forbidden');

  const contract_id  = String(payload.contract_id||'').trim();
  const phone        = normPhone(payload.phone);
  const permise_type = String(payload.permise_type||'').trim(); // пишем, как просил
  const price        = toNum(payload.price);
  const status       = String(payload.status||'').trim(); // credited_pending | credited_available | termination

  if (!contract_id) throw new Error('Укажите номер договора');
  if (!phone || phone.length!==11) throw new Error('Неверный телефон');
  if (!permise_type) throw new Error('Выберите вид помещения');
  if (!(price > 0)) throw new Error('Цена должна быть > 0');
  if (!['credited_pending','credited_available','termination'].includes(status)) throw new Error('Некорректный статус');

  const points = calcPointsByPremise_(permise_type, price);

  const s = sh('Purchases'); if (!s) throw new Error('Лист Purchases не найден');
  const hdr = s.getRange(1,1,1,s.getLastColumn()).getValues()[0].map(x=>String(x).trim());
  const rateUsed  = getBallRate_();
  const coeffUsed = getCoeffByPremise_(permise_type);
  const row = {
    contract_id,
    phone,
    permise_type,
    price,
    points,              // положительным числом; "termination" учтём в расчётах
    status,
    updated_at: new Date(),
    comment: '',
    ball_rate_used: rateUsed,    // 👈 сохранится, если колонка есть
    coeff_used:    coeffUsed     // 👈 сохранится, если колонка есть
  };
  s.appendRow(hdr.map(h => row[h] !== undefined ? row[h] : ''));
  // === NEW: уведомления клиенту про движение баланса ===
  try{
    if (status === 'credited_pending'){
      notifyBalanceChange_(phone, 'pending', points, { contract_id });
    } else if (status === 'credited_available'){
      notifyBalanceChange_(phone, 'available', points, { contract_id });
    } else if (status === 'termination'){
      notifyBalanceChange_(phone, 'termination', points, { contract_id });
    }
  } catch(e){ Logger.log('addPurchase notify error: '+e); }
  return { ok:true };
}

function api_adminListPurchases(token, search){
  const me = sessionGet_(token); 
  if (!me || !me.is_admin) throw new Error('forbidden');

  const sP = sh('Purchases');
  const outByContract = {}; // contract_id -> latest snapshot

  if (sP && sP.getLastRow()>=2){
    const v = sP.getRange(1,1,sP.getLastRow(), sP.getLastColumn()).getValues();
    const h = v[0].map(x=>String(x).trim());
    const m = Object.fromEntries(h.map((k,i)=>[k,i]));

    const needle = String(search||'').trim().toLowerCase();

    for (let i=1;i<v.length;i++){
      const r = v[i];
      if (!r.join('')) continue;

      const contract_id  = String(r[m.contract_id] || '').trim();
      if (!contract_id) continue;

      const phoneRaw     = String(r[m.phone] || '').trim();
      const phoneNorm    = normPhone(phoneRaw);

      const ts_ms        = parseTs_(r[m.updated_at]);

      // проверка поиска (по договору и телефону)
      if (needle){
        const hay = (contract_id + ' ' + phoneNorm).toLowerCase();
        if (hay.indexOf(needle) === -1) {
          // не матчится, но возможно более свежая строка с тем же contract_id потом матчится?
          // поэтому не continue прямо сейчас. сначала просто сохраним,
          // а потом фильтранём после выбора latest.
        }
      }

      // если это самая свежая запись по данному contract_id — запоминаем
      const prev = outByContract[contract_id];
      if (!prev || ts_ms > prev.updated_at_ts){
        outByContract[contract_id] = {
          contract_id,
          phone:        phoneNorm,
          permise_type: String(r[m.permise_type] || '').trim(),
          price:        toNum(r[m.price]),
          points:       toNum(r[m.points]),
          status:       String(r[m.status] || '').trim(),
          updated_at_ts: ts_ms,
          updated_at_display: Utilities.formatDate(
            new Date(ts_ms),
            Session.getScriptTimeZone(),
            'dd.MM.yyyy HH:mm'
          )
        };
      }
    }
  }

  // Преобразуем карту -> массив
  let out = Object.keys(outByContract).map(k => outByContract[k]);

  // Теперь фильтруем по needle уже после того как выбрали latest
  const needle2 = String(search||'').trim().toLowerCase();
  if (needle2){
    out = out.filter(item => {
      const hay = (item.contract_id + ' ' + item.phone).toLowerCase();
      return hay.indexOf(needle2) !== -1;
    });
  }

  // сортируем по дате (новые сверху)
  out.sort((a,b)=> b.updated_at_ts - a.updated_at_ts);

  // ограничим например 200
  return { items: out.slice(0,200) };
}

function api_adminUpdatePurchase(token, contract_id, new_status, new_points){
  const me = sessionGet_(token); 
  if (!me || !me.is_admin) throw new Error('forbidden');

  if (!contract_id) throw new Error('contract_id required');

  if (!['credited_pending','credited_available','termination'].includes(new_status)){
    throw new Error('bad status');
  }

  const ptsNum = toNum(new_points);
  if (ptsNum < 0) throw new Error('points must be >= 0');

  const sP = sh('Purchases');
  if (!sP) throw new Error('Purchases sheet not found');

  // забираем все данные, чтобы найти самую свежую версию этой сделки
  const v = sP.getRange(1,1,sP.getLastRow(), sP.getLastColumn()).getValues();
  const h = v[0].map(x=>String(x).trim());
  const m = Object.fromEntries(h.map((k,i)=>[k,i]));

  let latest = null;
  for (let i=1;i<v.length;i++){
    const r = v[i]; if (!r.join('')) continue;
    if (String(r[m.contract_id]).trim() === String(contract_id).trim()){
      const ts = parseTs_(r[m.updated_at]);
      if (!latest || ts > latest.ts){
        latest = {
          phone:        normPhone(r[m.phone] || ''),
          permise_type: String(r[m.permise_type] || '').trim(),
          price:        toNum(r[m.price]),
          // points и status мы будем обновлять
        };
        latest.ts = ts;
      }
    }
  }

  if (!latest){
    throw new Error('contract not found');
  }

  // формируем новую запись (снимок состояния сделки)
  const rowObj = {
    contract_id:  contract_id,
    phone:        latest.phone,
    permise_type: latest.permise_type,
    price:        latest.price,
    points:       ptsNum,
    status:       new_status,
    updated_at:   new Date(),
    comment:      ''
  };

  // добавляем строку в конец
  const hdr = h;
  sP.appendRow(hdr.map(col => rowObj[col] !== undefined ? rowObj[col] : ''));
  // === NEW: уведомления клиенту про движение баланса ===
  try{
    if (new_status === 'credited_pending'){
      notifyBalanceChange_(latest.phone, 'pending', ptsNum, { contract_id });
    } else if (new_status === 'credited_available'){
      notifyBalanceChange_(latest.phone, 'available', ptsNum, { contract_id });
    } else if (new_status === 'termination'){
      notifyBalanceChange_(latest.phone, 'termination', ptsNum, { contract_id });
    }
  } catch(e){ Logger.log('updatePurchase notify error: '+e); }

  return { ok:true };
}

// === История операций с running balance ===
function parseTs_(v){
  // принимает Date | строку | число → миллисекунды
  if (v instanceof Date) return v.getTime();
  const s = String(v||'').trim();
  if (!s) return Date.now();
  // пробуем ISO / локальные форматы
  const d = new Date(s);
  if (!isNaN(d.getTime())) return d.getTime();
  // если в ячейке число (Excel/Sheets serial)
  const n = Number(s);
  if (!isNaN(n)) {
    // Google Sheets даёт "кол-во дней с 1899-12-30"
    const ms = (n - 25569) * 86400 * 1000;
    if (ms > 0) return ms;
  }
  return Date.now();
}
function api_getHistory(token){
  const me = sessionGet_(token); if (!me) throw new Error('auth required');
  const phone = normPhone(me.phone);
  const ss = SS();

  const events = [];

  // Purchases → кредиты
  const sP = ss.getSheetByName('Purchases');
  if (sP && sP.getLastRow()>=2){
    const v = sP.getRange(1,1,sP.getLastRow(), sP.getLastColumn()).getValues();
    const h = v[0].map(x=>String(x).trim());
    const m = Object.fromEntries(h.map((k,i)=>[k,i]));
    for (let i=1;i<v.length;i++){
      const r=v[i]; if (!r.join('')) continue;
      if (normPhone(r[m.phone])!==phone) continue;
      const st = String(r[m.status]||'');
      const pts = toNum(r[m.points]);
      if (!pts) continue;
      const ts = parseTs_(r[m.updated_at]);
      const base = {
        ts,
        ts_display: Utilities.formatDate(new Date(ts), Session.getScriptTimeZone(), 'dd.MM.yyyy HH:mm'),
        amount: pts,
        contract_id: String(r[m.contract_id]||''),
        comment: String(r[m.comment]||''),
        status: st,
      };
      if (st === 'credited_available') {
        events.push(Object.assign({}, base, {
          kind: 'credit',
          subtype: 'credited_available',
          sign: '+',
          title: `Начисление ${base.contract_id ? '('+base.contract_id+')' : ''}`,
          status_label: 'Подтверждено'
        }));
      } else if (st === 'credited_pending') {
        events.push(Object.assign({}, base, {
          kind: 'credit',
          subtype: 'credited_pending',
          sign: '+',
          title: `Ожидает подтверждения ${base.contract_id ? '('+base.contract_id+')' : ''}`,
          status_label: 'Ожидает подтверждения'
        }));
      } else if (st === 'termination') {
        events.push(Object.assign({}, base, {
          kind: 'credit',
          subtype: 'termination',
          sign: '−',
          title: `Расторжение ${base.contract_id ? '('+base.contract_id+')' : ''}`,
          status_label: 'Расторжение'
        }));
      }
    }
  }

  // Redemptions → дебеты
  const sR = ss.getSheetByName('Redemptions');
  if (sR && sR.getLastRow()>=2){
    const v = sR.getRange(1,1,sR.getLastRow(), sR.getLastColumn()).getValues();
    const h = v[0].map(x=>String(x).trim());
    const m = Object.fromEntries(h.map((k,i)=>[k,i]));
    for (let i=1;i<v.length;i++){
      const r=v[i]; if (!r.join('')) continue;
      if (normPhone(r[m.phone])!==phone) continue;
      const st  = String(r[m.status]||'');
      const pts = toNum(r[m.points_spent]);
      const ts  = parseTs_(r[m.created_at]);
      const base = {
        ts,
        ts_display: Utilities.formatDate(new Date(ts), Session.getScriptTimeZone(), 'dd.MM.yyyy HH:mm'),
        amount: pts,
        title: String(r[m.title]||''),
        status: st,
      };
      if (['waiting','approved','await_code','delivered'].includes(st)) {
        events.push(Object.assign({}, base, { kind:'debit', subtype:st, sign:'-' }));
      } else if (['canceled','rejected','failed'].includes(st)) {
        events.push(Object.assign({}, base, { kind:'neutral', subtype:st, sign:'' }));
      }
    }
  }

  // running balance с двумя треками:
  // confirmed = доступный баланс (зелёные + учёт резерва)
  // pendingSum = сумма синих "ожидает подтверждения"
  // 1. Сортируем по времени от старых к новым
  const asc = events.slice().sort((a,b)=>a.ts-b.ts);

  // Будем вести три "счётчика состояния" на момент каждой операции:
  let confirmedEarned = 0; // подтверждённые баллы (credited_available)
  let pendingEarned   = 0; // ожидание подтверждения (credited_pending минус расторжения до подтверждения)
  let reserved        = 0; // удержано за заявки (waiting/approved/delivered из Redemptions)

  asc.forEach(ev => {
    if (ev.kind === 'credit') {
      if (ev.subtype === 'credited_available') {
        // подтверждённые баллы
        confirmedEarned += ev.amount;

        // важно: тут мы pending не уменьшаем, потому что:
        // pending добавляется отдельным событием credited_pending ранее,
        // и этот шаг (credited_available) — отдельная новая запись,
        // а не "апдейт той же самой".
        // Это допустимо визуально. Если захочешь полностью идеально
        // вычитать pending на этот шаг — нужно будет тащить contract_id
        // и матчить сделки. Пока держим проще.

      } else if (ev.subtype === 'credited_pending') {
        // ещё не подтверждено, но обещано
        pendingEarned += ev.amount;

      } else if (ev.subtype === 'termination') {
        // расторжение.
        // Если мы расторгли сделку, которая была ещё на стадии pending,
        // надо убрать эти обещанные баллы из pending.
        // Мы используем "минус", который ты хотел видеть визуально.
        // => просто вычитаем
        pendingEarned = Math.max(0, pendingEarned - ev.amount);
      }

    } else if (ev.kind === 'debit') {
      // списания/брони баллов за награды
      if (['waiting','approved','await_code','delivered'].includes(ev.subtype)) {
        reserved += ev.amount;
      }
      // canceled/rejected мы не учитываем как удержание
    }

    // считаем доступный баланс после этой операции:
    ev.running_confirmed = Math.max(0, confirmedEarned - reserved);

    // считаем текущий неподтверждённый баланс после этой операции:
    ev.running_pending = Math.max(0, pendingEarned);
  });

  // 2. Теперь обратно сортируем по убыванию времени (последние сверху)
  const items = asc.sort((a,b)=>b.ts-a.ts);
  return { items };
}

// ====== ADMIN: загрузка картинки и добавление товара ======
function ensureFolderByName_(name){
  var iter = DriveApp.getFoldersByName(name);
  if (iter.hasNext()) return iter.next();
  return DriveApp.createFolder(name);
}

// dataURL ("data:image/png;base64,....") -> файл в Drive, вернуть {id, url}
function saveImageDataUrlToDrive_(dataUrl, baseName){
  if (!dataUrl || String(dataUrl).indexOf('data:')!==0) return { id:'', url:'' };
  var parts = String(dataUrl).split(',');
  var meta  = parts[0];                 // "data:image/png;base64"
  var b64   = parts[1] || '';
  var mime  = meta.substring(5, meta.indexOf(';')) || 'application/octet-stream';
  var bytes = Utilities.base64Decode(b64);

  var folder = ensureFolderByName_('SVOYSHOP_CATALOG_IMAGES');
  var file   = folder.createFile(Utilities.newBlob(bytes, mime, (baseName||'image')+'.png'));

  // сделать доступным по ссылке
  try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch(_){}

  var id  = file.getId();
  var url = 'https://drive.google.com/uc?export=view&id=' + id;
  return { id:id, url:url };
}

function api_adminAddCatalogItem(token, payload){
  var me = sessionGet_(token); 
  if (!me || !me.is_admin) throw new Error('forbidden');

  // === Нормализация входа ===
  var title       = String(payload.title||'').trim();
  var category    = String(payload.category||'').trim();
  var desc        = String(payload.desc||'').trim();
  var priceTenge  = toNum(payload.price_tenge);           // НОВОЕ: цена в тенге
  var pointsIn    = toNum(payload.points_price);          // Может прийти (если UI уже посчитал), но не обязателен
  var stock       = toNum(payload.stock);
  var active      = !!payload.is_active;

  if (!title) throw new Error('Укажите название товара');
  if (!(priceTenge > 0)) throw new Error('Укажите цену в тенге > 0');

  // === Определяем points_price ===
  // Если пришёл валидный points_price — используем его.
  // Иначе считаем по текущему курсу ball_rate и фиксируем.
  var rate = 1;
  try {
    rate = Math.max(0.000001, toNum(getBallRate_()));    // защита от нуля
  } catch (_){ rate = 1; }

  var points_price = (pointsIn > 0)
    ? Math.round(pointsIn)
    : Math.max(1, Math.round(priceTenge / rate));

  // === Картинка (imgbb) ===
  var photo_url = '';
  if (payload.image_data_url){
    var saved = uploadImageToImgbb_(String(payload.image_data_url));
    photo_url = saved.url || '';
  }

  // === Генерируем item_id ===
  var item_id = 'I' + Math.floor(100000 + Math.random()*900000);

  // === Лист Catalog ===
  var s = sh('Catalog'); 
  if (!s) throw new Error('Лист Catalog не найден');

  // Соберём заголовки (ensureSheet ранее должен был добавить недостающие)
  var hdr = s.getRange(1,1,1,s.getLastColumn()).getValues()[0].map(function(x){
    return String(x).trim();
  });

  // Подготовим строку к записи
  var row = {
    item_id:         item_id,
    category:        category,
    title:           title,
    desc:            desc,
    price_tenge:     priceTenge,       // НОВОЕ: сохраняем исходную цену в тенге
    points_price:    points_price,     // Фиксированная цена в баллах на момент добавления
    ball_rate_used:  rate,             // НОВОЕ: курс, по которому считали (для аудита/прозрачности)
    stock:           stock,
    photo_url:       photo_url,
    is_active:       active
  };

  // Записываем в порядке колонок заголовка
  s.appendRow(hdr.map(function(h){ 
    return row[h] !== undefined ? row[h] : ''; 
  }));

  return { ok:true, item_id:item_id, photo_url:photo_url, points_price: points_price };
}

// ===========================
// 🔸 1. Генерация кода выдачи
// ===========================
function api_adminGeneratePickupCode(token, redeem_id) {
  requireAdmin_(token);

  const s = sh('Redemptions');
  if (!s) throw new Error('Redemptions not found');

  const v = s.getDataRange().getValues();
  const h = v[0].map(x=>String(x).trim());
  const m = Object.fromEntries(h.map((k,i)=>[k,i])); // индексы

  const row = v.findIndex((r,i)=> i>0 && String(r[m.redeem_id])===String(redeem_id));
  if (row < 1) throw new Error('Заявка не найдена');

  // 4-значный код
  const code = String(Math.floor(1000 + Math.random()*9000));

  v[row][m.status] = 'await_code';
  if (m.pickup_code === undefined) throw new Error('Добавьте колонку "pickup_code" в Redemptions');
  v[row][m.pickup_code] = code;

  s.getDataRange().setValues(v);

  // Телеграм
  const phone = normPhone(v[row][m.phone]);
  const chat  = getTelegramIdByPhone_(phone);
  const title = String(v[row][m.title]||'').trim();
  if (chat){
    sendTelegramMessage_(chat, `Код для получения по заявке «${title}»: ${code}\nПокажите этот код менеджеру при выдаче.`);
  }

  return { ok:true, code };
}

// ===========================
// 🔸 2. Подтверждение выдачи
// ===========================
function api_adminConfirmPickupCode(token, redeem_id, inputCode) {
  requireAdmin_(token);

  const s = sh('Redemptions');
  const v = s.getDataRange().getValues();
  const h = v[0].map(x=>String(x).trim());
  const m = Object.fromEntries(h.map((k,i)=>[k,i]));

  const row = v.findIndex((r,i)=> i>0 && String(r[m.redeem_id])===String(redeem_id));
  if (row < 1) throw new Error('Заявка не найдена');

  if (m.pickup_code === undefined) throw new Error('Нет колонки pickup_code');
  const real = String(v[row][m.pickup_code]||'').trim();
  if (!real || String(inputCode).trim() !== real) throw new Error('Неверный код выдачи');

  v[row][m.status] = 'delivered';
  v[row][m.pickup_code] = '';
  if (m.delivered_at === undefined) throw new Error('Добавьте колонку delivered_at');
  v[row][m.delivered_at] = new Date();

  s.getDataRange().setValues(v);

  const phone = normPhone(v[row][m.phone]);
  const chat  = getTelegramIdByPhone_(phone);
  const title = String(v[row][m.title]||'').trim();
  if (chat){
    sendTelegramMessage_(chat, `Заказ «${title}» выдан.\nСпасибо, что участвуете в программе SVOY SHOP.`);
  }

  return { ok:true };
}

// ===========================
// 🔸 3. Отмена заявки
// ===========================
function api_adminCancelRedemption(token, redeem_id, comment) {
  requireAdmin_(token);

  const s = sh('Redemptions');
  const v = s.getDataRange().getValues();
  const h = v[0].map(x=>String(x).trim());
  const m = Object.fromEntries(h.map((k,i)=>[k,i]));

  const row = v.findIndex((r,i)=> i>0 && String(r[m.redeem_id])===String(redeem_id));
  if (row < 1) throw new Error('Заявка не найдена');

  v[row][m.status] = 'canceled';
  if (m.pickup_code !== undefined) v[row][m.pickup_code] = '';
  if (m.manager_comment !== undefined) v[row][m.manager_comment] = String(comment||'');
  // ⚙️ вернуть stock в Catalog по item_id (как в user cancel)
  try {
    const item_id = String(v[row][m.item_id]||'').trim();
    if (item_id){
      const sC = sh('Catalog');
      const c  = sC.getDataRange().getValues();
      const ch = c[0].map(x=>String(x).trim());
      const cIdI = ch.indexOf('item_id');
      const cStI = ch.indexOf('stock');
      for (let r=1;r<c.length;r++){
        if (String(c[r][cIdI])===item_id){
          c[r][cStI] = toNum(c[r][cStI]) + 1;
          break;
        }
      }
      sC.getDataRange().setValues(c);
    }
  } catch(e){
    Logger.log('stock rollback error (admin cancel): '+e);
  }

  s.getDataRange().setValues(v);

  const phone = normPhone(v[row][m.phone]);
  const chat  = getTelegramIdByPhone_(phone);
  const title = String(v[row][m.title]||'').trim();
  if (chat){
    sendTelegramMessage_(chat, `Ваша заявка на «${title}» отменена.\nПричина: ${comment||'не указана'}.`);
  }

  return { ok:true };
}

// ====== WEB APP ======
function doGet() {
  return HtmlService.createHtmlOutputFromFile('app')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Универсальный helper, чтобы навесить CORS-заголовки на ответ
function withCors_(output) {
  return output
    .setMimeType(ContentService.MimeType.JSON)
    .setHeader('Access-Control-Allow-Origin', '*')
    .setHeader('Access-Control-Allow-Methods', 'POST, GET, OPTIONS')
    .setHeader('Access-Control-Allow-Headers', 'Content-Type');
}

// Ответ на preflight-запросы (OPTIONS)
function doOptions(e) {
  // Просто отвечаем 200 OK с нужными CORS заголовками, без тела
  return withCors_( ContentService.createTextOutput('') );
}

// Обновлённый doPost с CORS
function doPost(e) {
  try {
    var data = {};

    // === 1. Парсим тело запроса ===
    if (e && e.postData && e.postData.contents) {
      var ct = (e.postData.type || "").toLowerCase();

      if (ct.indexOf("application/json") !== -1) {
        data = JSON.parse(e.postData.contents || "{}");
      } else {
        // form-urlencoded
        var params = {};
        var rawBody = String(e.postData.contents || "");
        rawBody.split("&").forEach(function (pair) {
          var kv = pair.split("=");
          var k = decodeURIComponent((kv[0] || "").replace(/\+/g, " "));
          var v = decodeURIComponent((kv[1] || "").replace(/\+/g, " "));
          params[k] = v;
        });
        data = params;
      }
    }

    var action = data.action;
    var token  = data.token || "";
    var result;

    // === 2. USER AUTH ===
    if (action === 'register') {
      result = api_register({
        phone:      data.phone,
        full_name:  data.full_name,
        password:   data.password,
        dob:        data.dob,
        gender:     data.gender,
      });
    }

    else if (action === 'login') {
      result = api_login({
        phone:    data.phone,
        password: data.password,
      });
    }

    else if (action === 'logout') {
      result = api_logout(token);
    }

    // === 3. CLIENT ZONE ===
    else if (action === 'getDashboard') {
      result = api_getDashboard(token);
    }

    else if (action === 'getHistory') {
      result = api_getHistory(token);
    }

    else if (action === 'redeem') {
      result = api_createRedemption(token, data.item_id);
    }

    else if (action === 'createRedemption') {
      result = api_createRedemption(token, data.item_id);
    }

    else if (action === 'cancelRedemption') {
      result = api_cancelRedemption(token, data.redeem_id, data.reason || '');
    }

    // === 4. TELEGRAM ===
    else if (action === 'generateTelegramCode') {
      result = api_generateTelegramCode(token);
    }

    else if (action === 'checkTelegramLink') {
      // теперь возвращает true/false по наличию telegram_user_id
      result = api_checkTelegramLink(token);
    }

    else if (action === 'telegramConfirm') {
      // вызывется ботом (без токена)
      result = api_telegramConfirm({
        code: data.code,
        chat_id: data.chat_id
      });
    }

    // === 5. ADMIN ZONE ===
    else if (action === 'adminListRedemptions') {
      result = api_adminListRedemptions(token, data.statusFilter || '');
    }

    else if (action === 'adminGeneratePickupCode') {
      result = api_adminGeneratePickupCode(token, data.redeem_id);
    }
    else if (action === 'adminConfirmPickupCode') {
      result = api_adminConfirmPickupCode(token, data.redeem_id, data.code);
    }
    else if (action === 'adminCancelRedemption') {
      result = api_adminCancelRedemption(token, data.redeem_id, data.comment);
    }

    else if (action === 'adminUpdateRedemption') {
      result = api_adminUpdateRedemption(token, data.redeem_id, data.new_status);
    }

    else if (action === 'adminListCatalog') {
      result = api_adminListCatalog(token);
    }

    else if (action === 'adminUpdateCatalogItem') {
      result = api_adminUpdateCatalogItem(token, JSON.parse(data.payload));
    }

    else if (action === 'adminAddPurchase') {
      result = api_adminAddPurchase(token, JSON.parse(data.payload));
    }

    else if (action === 'adminListPurchases') {
      result = api_adminListPurchases(token, data.search || '');
    }

    else if (action === 'adminUpdatePurchase') {
      result = api_adminUpdatePurchase(
        token,
        data.contract_id,
        data.new_status,
        data.new_points
      );
    }

    else if (action === 'adminAddCatalogItem') {
      var incomingPayload = data.payload || data;
      if (typeof incomingPayload === 'string') {
        try { incomingPayload = JSON.parse(incomingPayload); } catch (_e) {
          throw new Error('Bad payload JSON for adminAddCatalogItem');
        }
      }
      result = api_adminAddCatalogItem(token, incomingPayload);
    }

    else if (action === 'adminApproveRedemption') {
      result = api_adminApproveRedemption(token, data.redeem_id);
    }

    else if (action === 'requestPasswordReset') {
      result = api_requestPasswordReset({
        phone: data.phone
      });
    }

    else if (action === 'confirmPasswordReset') {
      result = api_confirmPasswordReset({
        phone: data.phone,
        code: data.code,
        new_password: data.new_password
      });
    }

    // === 6. UNKNOWN ===
    else {
      throw new Error("Unknown action: " + action);
    }

    // === 7. SUCCESS ===
    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, data: result }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    // === 8. ERROR ===
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.message || String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}