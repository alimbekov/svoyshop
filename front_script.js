<script>
// === Форматирование чисел ===
function formatNum(n) {
  const num = Number(n || 0);
  const withSpaces = num.toLocaleString('ru-RU');
  const noWrap = withSpaces.replace(/ /g, '\u00A0'); // неразрывные пробелы
  return noWrap;
}

function formatPoints(n) {
  return formatNum(n) + '\u00A0баллов';
}
// Дата/время: берём display, иначе *_ts (секунды)
function fmtDateTime(display, ts){
  if (display && String(display).trim()) return display;
  if (typeof ts === 'number' && ts > 0){
    try { return new Date(ts * 1000).toLocaleString('ru-RU'); } catch(e){}
  }
  return '';
}

function esc(str){
  return String(str || '')
    .replace(/&/g,'&amp;')
    .replace(/</g,'&lt;')
    .replace(/>/g,'&gt;')
    .replace(/"/g,'&quot;')
    .replace(/'/g,'&#39;');
}
// если хотите сохранять привычную замену плюса:
function safeText(str){
  return esc(str).replace(/\+/g, ' ');
}

// форматируем created_at_* в "ДД.ММ.ГГГГ, ЧЧ:ММ"
function fmtDateTimeRu(val){
  if (val === undefined || val === null || val === '') return '';
  const opts = { day:'2-digit', month:'2-digit', year:'numeric', hour:'2-digit', minute:'2-digit' };
  if (typeof val === 'number') return new Date(val*1000).toLocaleString('ru-RU', opts);
  const t = Date.parse(String(val));
  return isNaN(t) ? String(val) : new Date(t).toLocaleString('ru-RU', opts);
}
function getCreatedAtForRedemption(r){
  return r.created_at_display || (r.created_at_ts!=null ? fmtDateTimeRu(r.created_at_ts)
         : (r.created_at ? fmtDateTimeRu(r.created_at) : ''));
}

/* =========================
   CONFIG / HELPERS
========================= */
const API_URL = 'https://svoyshop-proxy.alimbekov98.workers.dev/';
const HERO_URL = 'https://i.ibb.co/zWMg8cCS/logo.png';
let tgLinked = false; // кэш состояния привязки Telegram

function apiCall(payload) {
  return fetch(API_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(payload || {})
  })
  .then(r => r.json())
  .then(json => {
    if (!json.ok) {
      throw new Error(json.error || 'Ошибка запроса');
    }
    return json.data;
  });
}

async function adminRedemptionsReload(){
  if (!isAdmin){ nav('cabinet'); return; }

  const stSel = document.getElementById('adm_status_filter_ru');
  const statusFilter = stSel ? (stSel.value || '') : '';

  const wrap  = document.getElementById('adm_orders');
  const empty = document.getElementById('adm_orders_empty');
  if (wrap)  wrap.innerHTML = '⏳ Загружаем…';
  if (empty) empty.style.display = 'none';

  setBusy(true);
  try{
    const res = await apiCall({
      action:'adminListRedemptions',
      token,
      statusFilter // серверный фильтр по статусу
    });
    _adminRedemptionsCache = Array.isArray(res.items) ? res.items : [];
    renderAdminRedemptions();
  }catch(e){
    if (wrap) wrap.innerHTML = `<div class="alert">Ошибка: ${e.message}</div>`;
  }
  setBusy(false);
}

// Универсальный POST (подставьте свою переменную URL и токен, как в других ваших вызовах)
async function apiPost(action, data = {}) {
  const body = { action, ...data };
  const res = await fetch(API_URL, {
    method: 'POST',
    headers: {'Content-Type':'application/json'},
    body: JSON.stringify(body)
  });
  const json = await res.json();
  if (!json.ok) throw new Error(json.error || 'API error');
  return json.data;
}

// ===== Админ: действия по заявкам =====
async function admGeneratePickupCode(redeem_id) {
  try{
    setBusy?.(true);
    await apiPost('adminGeneratePickupCode', { token, redeem_id });
    // никаких уведомлений и показа кода администратору
    await adminLoadRedemptions(); // статус обновится на await_code и появится кнопка "Подтвердить выдачу"
  } finally { setBusy?.(false); }
}

async function admConfirmPickupCode(redeem_id) {
  const code = prompt('Введите 4-значный код выдачи:');
  if (!code) return;
  try{
    setBusy?.(true);
    await apiPost('adminConfirmPickupCode', { token, redeem_id, code });
    alert('Выдача подтверждена. Клиенту отправлено уведомление.');
    await adminLoadRedemptions();
  } finally { setBusy?.(false); }
}

async function admCancelRedemption(redeem_id) {
  const comment = prompt('Укажите причину отмены:');
  if (comment === null) return; // нажали Cancel
  try{
    setBusy?.(true);
    await apiPost('adminCancelRedemption', { token, redeem_id, comment });
    alert('Заявка отменена. Клиенту отправлено уведомление в Telegram.');
    await adminLoadRedemptions();
  } finally { setBusy?.(false); }
}

// ===== Пользователь: отмена своей заявки =====
async function userCancelRedemption(redeem_id){
  const reason = prompt('Укажите причину отмены (необязательно):') || '';
  try{
    setBusy?.(true);
    await apiPost('cancelRedemption', { token, redeem_id, reason });
    alert('Заявка отменена. Резерв баллов разблокирован.');
    await loadCabinet(); // здесь у Вас уже есть логика обновления дашборда
  } finally { setBusy?.(false); }
}

// utils как у тебя
function digitsAfterCountry_(raw){ let d=String(raw||'').replace(/\D/g,''); if(d.startsWith('7')) d=d.slice(1); return d; }
function normalizeDigitsWithForced7_(raw){ let d=digitsAfterCountry_(raw); d=d.replace(/^7?/, '7'); d=d.slice(0,10); return d.split(''); }
function getPlainPhoneForApi(rawPhone) {
  // берём только цифры, насильно ставим первую цифру 7, ограничиваемся 11 цифрами
  // например "+7 778 203 15 51" -> "77782031551"
  const arr = normalizeDigitsWithForced7_(rawPhone); // это массив цифр без лишнего
  return arr.join('');
}
function buildFormattedFromDigits_(arr){ const groups=[3,3,2,2]; let out='+7 ', caret=out.length, i=0,g=0; while(i<arr.length&&g<groups.length){ const need=groups[g]; const take=Math.min(need,arr.length-i); out+=arr.slice(i,i+take).join(''); i+=take; caret=out.length; if(i<arr.length){ out+=' '; caret=out.length; } g++; } return {text:out, caret}; }
function setCaret_(el,pos){ requestAnimationFrame(()=>{ try{ el.setSelectionRange(pos,pos);}catch(_){}}); }
function applyPhoneMask(sel){
  const el = document.querySelector(sel); if(!el) return;
  const MIN_POS = 5;

  const render = ()=>{
    const digs = normalizeDigitsWithForced7_(el.value||'');
    const { text, caret } = buildFormattedFromDigits_(digs);
    el.value = text;
    setCaret_(el, Math.max(caret, MIN_POS));
  };

  if (!el.value.trim()){
    const {text}=buildFormattedFromDigits_(['7']);
    el.value=text;
  } else {
    render();
  }

  el.addEventListener('keydown', e=>{
    const pos=el.selectionStart||0;
    if ((e.key==='Backspace' && pos<=4) || (e.key==='Delete' && pos<4)){
      e.preventDefault();
      setCaret_(el,5);
      return;
    }
    if (e.key==='Backspace' && el.value.charAt(pos-1)===' '){
      e.preventDefault();
      setCaret_(el,pos-1);
    }
  });

  el.addEventListener('input', render);

  el.addEventListener('paste', e=>{
    e.preventDefault();
    const t=(e.clipboardData||window.clipboardData).getData('text')||'';
    const d=normalizeDigitsWithForced7_(t);
    const {text,caret}=buildFormattedFromDigits_(d.length?d:['7']);
    el.value=text;
    setCaret_(el,caret);
  });

  el.addEventListener('focus', ()=>{
    if(!/\+7 7/.test(el.value)){
      const {text}=buildFormattedFromDigits_(['7']);
      el.value=text;
    }
    render();
  });
}

function applyDobMask(selector){
  const el=document.querySelector(selector); if(!el) return;
  const fmt=(raw)=>{
    let d=(raw||'').replace(/\D/g,'').slice(0,8);
    const dd=d.slice(0,2), mm=d.slice(2,4), yy=d.slice(4,8);
    let out='';
    if(dd) out+=dd;
    if(mm) out+=(out?'.':'')+mm;
    if(yy) out+=(out?'.':'')+yy;
    return out;
  };
  const set=v=>el.value=v;
  el.addEventListener('input', ()=>set(fmt(el.value)));
  el.addEventListener('paste', e=>{
    e.preventDefault();
    const t=(e.clipboardData||window.clipboardData).getData('text')||'';
    set(fmt(t));
  });
  el.addEventListener('keydown', e=>{
    const ok=['Backspace','Delete','ArrowLeft','ArrowRight','Tab','.'];
    if(ok.includes(e.key)) return;
    if(!/^\d$/.test(e.key)) e.preventDefault();
  });
}
function dobMaskedToISO(m){
  m=String(m||'').trim();
  if(!m) return '';
  const p=m.split('.');
  if(p.length!==3) return '';
  const [dd,mm,yy]=p;
  if(yy.length!==4||dd.length!==2||mm.length!==2) return '';
  return `${yy}-${mm}-${dd}`;
}

function togglePasswordSimple(inputId, btn){
  const inp = document.getElementById(inputId);
  if(!inp || !btn) return;
  const isHidden = inp.type === 'password';
  inp.type = isHidden ? 'text' : 'password';
  btn.textContent = isHidden ? 'Скрыть' : 'Показать';
  try { btn.blur(); } catch(e){}
}

/* =========================
   GLOBAL STATE
========================= */
let token = localStorage.getItem('svoyshop_token')||'';
let isAdmin = false;
let adminCatalogItems = [];
let currentView = 'landing';
let lastDashboard = null;
let _busyCount = 0;
let _newImageDataUrl = '';
let _adminDealsCache = [];
let _adminRedemptionsCache = [];

/* =========================
   BUSY/LOCK UI
========================= */
function setBusy(on){
  _busyCount += on ? 1 : -1;
  if (_busyCount < 0) _busyCount = 0;
  const active = _busyCount > 0;

  document.body.classList.toggle('is-busy', active);

  const lock = document.getElementById('ui_lock');
  if (lock) lock.style.display = active ? 'block' : 'none';

  const overlay = document.getElementById('loadingOverlay');
  if (overlay) overlay.style.display = active ? 'block' : 'none';

  document.body.setAttribute('aria-busy', String(active));
}

function flashMsg(elId, html, ms=5000){
  const el = document.getElementById(elId);
  if (!el) return;
  el.innerHTML = html;
  if (el._flashTimer) clearTimeout(el._flashTimer);
  el._flashTimer = setTimeout(()=>{ el.innerHTML=''; }, ms);
}

/* =========================
   VIEWS / NAV
========================= */
function showLanding(){
  currentView='landing';
  document.getElementById('view-landing').style.display   = 'block';
  document.getElementById('view-register').style.display  = 'none';
  document.getElementById('view-login').style.display     = 'none';
  ;['cabinet','catalog','orders','history','admin'].forEach(v=>{
    const el = document.getElementById('view-'+v);
    if (el) el.style.display='none';
  });
  document.getElementById('authNav').style.display = 'none';
  document.getElementById('tabsBar').style.display = 'none';
}

function show(view){
  currentView=view;
  document.getElementById('view-landing').style.display  = 'none';
  document.getElementById('view-register').style.display = 'none';
  document.getElementById('view-login').style.display    = 'none';
  ;['cabinet','catalog','orders','history','admin'].forEach(v=>{
    const el = document.getElementById('view-'+v);
    if (el) el.style.display='none';
  });

  if (view==='register' || view==='login'){
    document.getElementById('tabsBar').style.display = 'flex';
    document.getElementById('authNav').style.display = 'none';
    document.getElementById('view-'+view).style.display = 'block';
    document.getElementById('tab-register').setAttribute('aria-selected', String(view==='register'));
    document.getElementById('tab-login').setAttribute('aria-selected', String(view==='login'));
  } else {
    showLanding();
  }
}

function nav(view){
  currentView = view;
  if (_busyCount > 0) return;
  document.getElementById('view-landing').style.display  = 'none';
  document.getElementById('view-register').style.display = 'none';
  document.getElementById('view-login').style.display    = 'none';
  ;['cabinet','catalog','orders','history','admin'].forEach(v=>{
    const el = document.getElementById('view-'+v);
    if (el) el.style.display = (v===view)?'block':'none';
  });
  document.getElementById('tabsBar').style.display = 'none';
  document.getElementById('authNav').style.display = 'flex';

  if (view==='cabinet') loadCabinet();
  if (view==='catalog') loadCatalogFresh();
  if (view==='orders')  { loadOrdersFresh(); }
  if (view==='history') loadHistory();
  if (view==='admin')   { adminShow('orders'); }
}

/* =========================
   AUTH
========================= */
function register(){
  const phone = document.getElementById('reg_phone').value.trim();
  const full_name = document.getElementById('reg_name').value.trim();
  const password = document.getElementById('reg_password').value.trim();
  const dobISO   = dobMaskedToISO(document.getElementById('reg_dob').value.trim());
  const gender   = document.getElementById('reg_gender').value;
  const msg = document.getElementById('reg_msg');
  const btn = document.getElementById('btn_register');
  const agreeRules = document.getElementById('agree_rules')?.checked;
  const agreeOffer = document.getElementById('agree_offer')?.checked;
  if (!agreeRules || !agreeOffer){
      msg.innerHTML = '<div class="alert">Для регистрации отметьте согласие с Правилами и Публичной офертой.</div>';
      return;
  }

  // сбрасываем предыдущее сообщение
  msg.textContent = '';
  msg.innerHTML = '';

  // ===== ВАЛИДАЦИЯ ПАРОЛЯ =====
  // 1) длина >= 6
  const passOkLength = password.length >= 6;

  // 2) только латиница / цифры / знаки
  //    разрешаем: A-Z a-z 0-9 и эти символы !@#$%^&*()_+-=[]{};:'",.<>/?\|`~
  //    запрещаем: кириллица, пробелы, эмодзи и т.п.
  const passOkCharset = /^[A-Za-z0-9!@#$%^&*()_\+\-\=\[\]{};:'",.<>\/?\\|`~]+$/.test(password);

  if (!passOkLength || !passOkCharset){
    msg.innerHTML = `<div class="alert">
      Пароль должен быть не короче 6 символов и содержать только латинские буквы, цифры и символы.
    </div>`;
    if (btn) btn.disabled = false;
    return;
  }

  // прошли проверку → продолжаем стандартный флоу
  msg.textContent='⏳ Отправка...';
  if (btn) btn.disabled = true;

  setBusy(true);
  apiCall({
      action: 'register',
      phone,
      full_name,
      password,
      dob: dobISO,
      gender,
      agree_rules: true,
      agree_offer: true,
      agreements_meta: {
        rules_version: '2025-10-31',
        offer_version: '2025-10-31',
        accepted_at: new Date().toISOString()
      }
    })
  .then(() => {
    setBusy(false);
    if (btn) btn.disabled = false;
    msg.innerHTML='✅ Регистрация успешна. Теперь войдите.';
    document.getElementById('log_phone').value = phone;
    show('login');
  })
  .catch(err => {
    setBusy(false);
    if (btn) btn.disabled = false;
    msg.innerHTML=`<div class="alert">Ошибка: ${err.message}</div>`;
  });
}

function login(){
  const phone = document.getElementById('log_phone').value.trim();
  const password = document.getElementById('log_password').value.trim();
  const msg = document.getElementById('log_msg');
  const btn = document.getElementById('btn_login');

  msg.textContent='⏳ Проверяем...';
  if (btn) btn.disabled = true;

  setBusy(true);
  apiCall({
    action: 'login',
    phone,
    password
  })
  .then(res => {
    setBusy(false);
    if (btn) btn.disabled = false;
    token = res.token;
    isAdmin = !!res.is_admin;
    localStorage.setItem('svoyshop_token', token);
    msg.textContent='';
    document.getElementById('authNav').style.display = 'flex';
    document.getElementById('nav_admin').style.display = isAdmin ? 'inline-block' : 'none';
    nav('cabinet');
  })
  .catch(err => {
    setBusy(false);
    if (btn) btn.disabled = false;
    msg.innerHTML=`<div class="alert">Ошибка: ${err.message}</div>`;
  });
}

// === Модалки Правила/Оферта ===
function openModal(id){ const m=document.getElementById(id); if(m) m.classList.add('open'); }
function closeModal(id){ const m=document.getElementById(id); if(m) m.classList.remove('open'); }
function closeModalOutside(e,id){ if(e.target && e.target.id===id) closeModal(id); }

// Кнопка регистрации активна только при 2 галочках
function updateRegisterBtnState(){
  const ok = !!document.getElementById('agree_rules')?.checked && !!document.getElementById('agree_offer')?.checked;
  const btn = document.getElementById('btn_register');
  if (btn) btn.disabled = !ok || (_busyCount>0);
}

// следим за изменениями чекбоксов
document.addEventListener('change', (e)=>{
  if (e.target && (e.target.id==='agree_rules' || e.target.id==='agree_offer')) {
    updateRegisterBtnState();
  }
});

function openResetFlow(stayOnStep1) {
  // показать экран восстановления
  currentView = 'reset';

  // скрыть все остальные
  document.getElementById('view-landing').style.display  = 'none';
  document.getElementById('view-register').style.display = 'none';
  document.getElementById('view-login').style.display    = 'none';
  ['cabinet','catalog','orders','history','admin'].forEach(v=>{
    const el = document.getElementById('view-'+v);
    if (el) el.style.display='none';
  });

  // показать сам reset
  document.getElementById('view-reset').style.display = 'block';

  // показать нужный шаг
  const s1 = document.getElementById('reset_step1');
  const s2 = document.getElementById('reset_step2');
  if (stayOnStep1) {
    s1.style.display = 'block';
    s2.style.display = 'none';
  } else {
    // по умолчанию открываем шаг 1
    s1.style.display = 'block';
    s2.style.display = 'none';
  }

  // очистим сообщения
  const msg1 = document.getElementById('reset_msg1');
  const msg2 = document.getElementById('reset_msg2');
  if (msg1) msg1.innerHTML = '';
  if (msg2) msg2.innerHTML = '';

  // телефон автозаполним тем, что в логине, чтобы не вводить заново
  const logPhone = document.getElementById('log_phone')?.value || '';
  const resetPhoneEl = document.getElementById('reset_phone');
  if (resetPhoneEl && logPhone && !resetPhoneEl.value) {
    resetPhoneEl.value = logPhone;
  }

  // и маску телефона навешиваем (если не висит)
  applyPhoneMask('#reset_phone');
}

async function resetRequestCode() {
  const phone = document.getElementById('reset_phone').value.trim(); // <-- без нормализации
  const msg1  = document.getElementById('reset_msg1');
  const btn   = document.getElementById('btn_reset_req');

  if (!phone) {
    msg1.innerHTML = `<div class="alert">Укажите телефон</div>`;
    return;
  }

  msg1.textContent = '⏳ Отправляем код в Telegram...';
  if (btn) btn.disabled = true;

  setBusy(true);
  try {
    const resp = await apiCall({
      action: 'requestPasswordReset',
      phone: phone // <-- отправляем как есть, формат "+7 778 203 15 51"
    });

    setBusy(false);
    if (btn) btn.disabled = false;

    if (resp.status === 'sent') {
      document.getElementById('reset_step1').style.display = 'none';
      document.getElementById('reset_step2').style.display = 'block';

      msg1.innerHTML = '';
      const msg2 = document.getElementById('reset_msg2');
      msg2.innerHTML = `
        <div style="font-size:14px;line-height:1.4;color:#065f46;font-weight:600;">
          ✅ Код отправлен в ваш Telegram.
          <br/>Проверьте диалог с ботом.
        </div>
      `;

      // dataset.norm больше не нужен
    }
    else if (resp.status === 'not_found') {
      msg1.innerHTML = `<div class="alert">Пользователь с таким номером не найден</div>`;
    }
    else if (resp.status === 'no_telegram') {
      msg1.innerHTML = `<div class="alert">У этого аккаунта не привязан Telegram. Восстановление пароля возможно только через привязанный Telegram.</div>`;
    }
    else {
      msg1.innerHTML = `<div class="alert">Неизвестный ответ сервера</div>`;
    }

  } catch (err) {
    setBusy(false);
    if (btn) btn.disabled = false;
    msg1.innerHTML = `<div class="alert">Ошибка: ${err.message}</div>`;
  }
}

async function resetConfirm() {
  const phone    = document.getElementById('reset_phone').value.trim(); // <- берём прямо из поля
  const code     = document.getElementById('reset_code').value.trim();
  const newPass  = document.getElementById('reset_newpass').value.trim();
  const msg2     = document.getElementById('reset_msg2');
  const btn      = document.getElementById('btn_reset_confirm');

  // валидация пароля остаётся как была:
  const passOkLength  = newPass.length >= 6;
  const passOkCharset = /^[A-Za-z0-9!@#$%^&*()_\+\-\=\[\]{};:'",.<>\/?\\|`~]+$/.test(newPass);

  if (!phone || !code || !newPass) {
    msg2.innerHTML = `<div class="alert">Заполните все поля</div>`;
    return;
  }
  if (!passOkLength || !passOkCharset) {
    msg2.innerHTML = `<div class="alert">Пароль должен быть не короче 6 символов и содержать только латинские буквы, цифры и символы.</div>`;
    return;
  }

  msg2.textContent = '⏳ Сохраняем новый пароль...';
  if (btn) btn.disabled = true;

  setBusy(true);
  try {
    const resp = await apiCall({
      action: 'confirmPasswordReset',
      phone: phone,       // <-- отправляем в том же виде
      code: code,
      new_password: newPass
    });

    setBusy(false);
    if (btn) btn.disabled = false;

    if (resp.status === 'ok') {
      // переключаемся на логин
      show('login');
    
      // подставляем телефон в форму входа
      const logPhoneEl = document.getElementById('log_phone');
      if (logPhoneEl) {
        logPhoneEl.value = phone;
      }
    
      // ставим зелёное уведомление в блоке входа
      const logMsg = document.getElementById('log_msg');
      if (logMsg) {
        logMsg.innerHTML = `
          <div style="
            display:flex;
            align-items:flex-start;
            gap:8px;
            font-size:14px;
            line-height:1.4;
            color:#065f46;
            font-weight:600;
            background:#ecfdf5;
            border:1px solid #6ee7b7;
            border-radius:8px;
            padding:8px 10px;
            max-width:360px;
          ">
            <span style="
              display:inline-block;
              width:16px;
              height:16px;
              border-radius:4px;
              border:2px solid #065f46;
              font-size:12px;
              font-weight:700;
              line-height:1;
              text-align:center;
            ">✓</span>
            <span>Новый пароль сохранён. Войдите.</span>
          </div>
        `;
    
        // автоочистка через 5 секунд
        setTimeout(() => {
          logMsg.innerHTML = '';
        }, 5000);
      }
    
      // подчистить поля восстановления (на всякий случай, чтобы если человек вернётся —
      // всё было пусто и без старого кода)
      document.getElementById('reset_code').value = '';
      document.getElementById('reset_newpass').value = '';
      const msg1 = document.getElementById('reset_msg1'); if (msg1) msg1.innerHTML = '';
      const msg2 = document.getElementById('reset_msg2'); if (msg2) msg2.innerHTML = '';
    
      // скрыть сам экран восстановления
      const resetView = document.getElementById('view-reset');
      if (resetView) {
        resetView.style.display = 'none';
      }
    }
    else if (resp.status === 'bad_code') {
      msg2.innerHTML = `<div class="alert">${resp.message || 'Неверный или просроченный код'}</div>`;
    }
    else if (resp.status === 'not_found') {
      msg2.innerHTML = `<div class="alert">Пользователь не найден</div>`;
    }
    else {
      msg2.innerHTML = `<div class="alert">Неизвестный ответ сервера</div>`;
    }

  } catch (err) {
    setBusy(false);
    if (btn) btn.disabled = false;
    msg2.innerHTML = `<div class="alert">Ошибка: ${err.message}</div>`;
  }
}


function logout(){
  localStorage.removeItem('svoyshop_token');
  // вызов logout на бэке, но без ожидания критичного результата
  apiCall({ action:'logout', token }).catch(()=>{});
  token=''; isAdmin=false;
  document.getElementById('authNav').style.display = 'none';
  showLanding();
}

// ====== ВОССТАНОВЛЕНИЕ ПАРОЛЯ ======

function cancelReset(){
  // просто вернуться на экран логина
  show('login');
}

// та же логика проверки пароля, что при регистрации
function passwordIsValid_(pw){
  const passOkLength  = pw.length >= 6;
  const passOkCharset = /^[A-Za-z0-9!@#$%^&*()_\+\-\=\[\]{};:'",.<>\/?\\|`~]+$/.test(pw);
  return passOkLength && passOkCharset;
}

/* =========================
   CABINET / DASHBOARD
========================= */
function colorizeMetrics(){
  const setCls = (id, cls)=>{
    const el = document.getElementById(id);
    if (el){ const card = el.closest('.metric'); if (card) card.classList.add(cls); }
  };
  setCls('m_available','metric-available');
  setCls('m_pending','metric-pending');
  setCls('m_earned','metric-earned');
  setCls('m_spent','metric-spent');
  setCls('m_hold','metric-hold');
}

async function loadCabinet(){
  const cb = document.getElementById('cab_msg');
  cb.textContent = '⏳ Загружаем...';

  setBusy(true); // покажем глобальный лоадер и заблокируем интерфейс

  try {
    // шаг 1. получить дашборд
    const d = await apiCall({
      action: 'getDashboard',
      token
    });

    // шаг 2. узнать телеграм-связку
    const linked = await fetchTelegramLink();

    // сохраним глобально
    lastDashboard = d;
    isAdmin = !!d.is_admin;
    cb.textContent = '';

    // рендерим ФИО и телефон
    document.getElementById('cab_name').textContent  = d.full_name || 'Покупатель';
    document.getElementById('cab_phone').textContent = d.phone_pretty;

    // баланс
    const s = d.balance_stats || {available:0,pending:0,earned_total:0,spent_total:0,hold_now:0};
    document.getElementById('m_available').textContent = formatPoints(s.available);
    document.getElementById('m_pending').textContent   = formatPoints(s.pending);
    document.getElementById('m_earned').textContent    = formatPoints(s.earned_total);
    document.getElementById('m_spent').textContent     = formatPoints(s.spent_total);
    document.getElementById('m_hold').textContent      = formatPoints(s.hold_now);

    // показать / скрыть "Админ-панель"
    document.getElementById('nav_admin').style.display = isAdmin ? 'inline-block' : 'none';
    document.getElementById('authNav').style.display = 'flex';

    // окрасить карточки метрик (один раз)
    colorizeMetrics();

    // применяем телеграм-статус к UI (прячем блок, блокируем покупки и т.д.)
    applyTelegramLockUI(linked);

  } catch(err){
    cb.innerHTML = `<div class="alert">Ошибка: ${err.message}</div>`;
    // выкидываем в логин, если токен невалидный
    show('login');
  }

  setBusy(false); // убираем глобальный лоадер ТОЛЬКО СЕЙЧАС, т.е. после всего
}

async function startTelegramVerify(){
  const msgEl = document.getElementById('tg_verify_msg');
  if (msgEl) {
    msgEl.textContent = '⏳ Генерируем код...';
  }

  setBusy(true);
  try {
    const res = await apiCall({
      action: 'generateTelegramCode',
      token
    });

    setBusy(false);

    const code = res.code || '';

    if (msgEl) {
      msgEl.innerHTML = `
        <div style="font-size:14px; line-height:1.5; text-align:center;">
          <div style="margin-bottom:8px; color:#111827;">
            ✅ Код создан:
          </div>

          <div style="display:flex; flex-wrap:wrap; justify-content:center; gap:8px; align-items:center; margin-bottom:12px;">
            <code id="tg_code_value"
                  style="font-size:16px; font-weight:700; background:#f3f4f6; padding:6px 10px; border-radius:8px; border:1px solid #d1d5db; min-width:100px; text-align:center;">
              ${code}
            </code>

            <button onclick="copyTelegramCode()"
                    style="padding:8px 12px; border-radius:8px; border:none; background:#0a7c59; color:#fff; font-weight:600; cursor:pointer;">
              Скопировать
            </button>

            <a href="https://t.me/svoy_shop_bot"
               target="_blank"
               style="padding:8px 12px; border-radius:8px; background:#1d4ed8; color:#fff; font-weight:600; text-decoration:none; display:inline-block;">
              Открыть бота
            </a>
          </div>

          <div style="font-size:13px; color:#4b5563; line-height:1.4; text-align:left; max-width:400px; margin:0 auto;">
            <div style="margin-bottom:4px;">1️⃣ Откройте Telegram и перейдите к нашему боту.</div>
            <div style="margin-bottom:4px;">2️⃣ Отправьте боту этот код.</div>
            <div>3️⃣ Вернитесь сюда и обновите кабинет.</div>
          </div>
        </div>
      `;
    }
  } catch(e){
    setBusy(false);
    if (msgEl) {
      msgEl.innerHTML = `<div class="alert">Ошибка: ${e.message}</div>`;
    }
  }
}

function copyTelegramCode(){
  const el = document.getElementById('tg_code_value');
  if (!el) return;
  const text = el.textContent.trim();
  navigator.clipboard.writeText(text).then(()=>{
    // можно сделать маленький визуальный отклик:
    alert('Код скопирован: ' + text);
  }).catch(()=>{
    alert('Не удалось скопировать, скопируй вручную.');
  });
}

// проверяем, привязан ли уже телеграм
async function fetchTelegramLink() {
  try {
    const res = await apiCall({ action: 'checkTelegramLink', token });
    tgLinked = !!res.linked;
    return tgLinked;
  } catch(e) {
    console.warn('fetchTelegramLink error', e);
    tgLinked = false;
    return false;
  }
}

function applyTelegramLockUI(isLinked){
  // 1. показать/спрятать блок "Подтвердите аккаунт через Telegram"
  const block = document.getElementById('tg_verify_block');
  if (block){
    block.style.display = isLinked ? 'none' : 'block';
  }

  // 2. заблокировать или разблокировать кнопки "Купить"
  //    (эти кнопки появляются только внутри каталога)
  document.querySelectorAll('.btn-buy').forEach(btn=>{
    if (isLinked){
      btn.disabled = false;
      btn.style.opacity = '';
      btn.style.cursor = '';
      btn.title = '';
    } else {
      btn.disabled = true;
      btn.style.opacity = '0.6';
      btn.style.cursor = 'not-allowed';
      btn.title = 'Недоступно. Подтвердите аккаунт через Telegram';
    }
  });
}

/* =========================
   HISTORY
========================= */
function loadHistory(){
  const list = document.getElementById('history_list');
  const empty = document.getElementById('history_empty');
  list.innerHTML = '';
  empty.style.display = 'none';

  setBusy(true);
  apiCall({
    action: 'getHistory',
    token
  })
  .then(res => {
    setBusy(false);
    const items = res.items || [];
    if (!items.length){ empty.style.display='block'; return; }

    items.forEach(ev => {
      let color, sign, mainTitle;
      if (ev.kind === 'credit') {
        if (ev.subtype === 'credited_available') {
          color = '#059669'; sign = '+'; mainTitle = ev.title || 'Начисление';
        } else if (ev.subtype === 'credited_pending') {
          color = '#2563eb'; sign = '+'; mainTitle = ev.title || 'Ожидает подтверждения';
        } else if (ev.subtype === 'termination') {
          color = '#6b7280'; sign = '−'; mainTitle = ev.title || 'Расторжение';
        } else {
          color = '#6b7280'; sign = '';  mainTitle = ev.title || '';
        }
      } else if (ev.kind === 'debit') {
        color = '#ef4444'; sign = '−'; mainTitle = ev.title || '';
      } else {
        color = '#6b7280'; sign = ''; mainTitle = ev.title || '';
      }

      const rightHTML = `
          <div style="text-align:right">
            <div style="font-weight:800">${formatPoints(ev.running_confirmed ?? 0)}</div>
            <div class="muted" style="font-size:12px">доступно после операции</div>
        
            <div style="font-weight:800; margin-top:8px; color:#2563eb">
              ${formatPoints(ev.running_pending ?? 0)}
            </div>
            <div class="muted" style="font-size:12px">неподтверждённый баланс</div>
          </div>`;

      const el = document.createElement('div');
      el.className = 'item';
      el.innerHTML = `
        <div class="flex">
          <div>
            <div style="display:flex;align-items:center;gap:8px">
              <div style="width:8px;height:8px;border-radius:50%;background:${color}"></div>
              <div style="font-weight:700">${mainTitle}</div>
            </div>
            <div class="muted" style="margin-top:4px">${ev.ts_display || ''}</div>
          </div>
          ${rightHTML}
        </div>

        <div class="flex" style="margin-top:8px">
          <div class="amount" style="color:${color}; font-weight:700;">
              ${sign} ${formatNum(ev.amount)}
            </div>
          <div class="muted">${ev.status || ev.status_label || ''}</div>
        </div>
      `;
      list.appendChild(el);
    });
  })
  .catch(err => {
    setBusy(false);
    list.innerHTML = `<div class="alert">Ошибка: ${err.message}</div>`;
  });
}

/* =========================
   CATALOG
========================= */
function updateCatalogBalance(){
  const el = document.getElementById('catalog_balance_num');
  if (!el) return;
  const av = Number(lastDashboard?.balance_stats?.available ?? 0);
  el.textContent = formatPoints(av);
}

function populateCategoryFilter(){
  const sel = document.getElementById('cat_filter');
  if (!sel || !lastDashboard) return;
  const cats = Array.from(
    new Set((lastDashboard.catalog || [])
      .map(it => (it.category || '').trim())
      .filter(Boolean))
  ).sort((a,b)=> a.localeCompare(b, 'ru'));

  const prev = sel.value;
  sel.innerHTML = '<option value="">Все категории</option>' +
    cats.map(c => `<option value="${c}">${c}</option>`).join('');

  if (cats.includes(prev)) sel.value = prev;
}

function renderCatalogSections(isLinked){
  if (!lastDashboard) return;
  if (typeof isLinked === 'undefined') {
    isLinked = tgLinked;
  }
  
  // --- Показать или скрыть предупреждение под "Доступно сейчас"
    const existingNotice = document.getElementById('tg_catalog_notice');
    if (existingNotice) existingNotice.remove();
    
    if (!isLinked) {
      const notice = document.createElement('div');
      notice.id = 'tg_catalog_notice';
      notice.innerHTML = `
        <div style="
          background:#fff7ed;
          border:1px solid #fed7aa;
          border-radius:12px;
          padding:10px 14px;
          margin:-8px 0 14px 0;
          color:#9a3412;
          font-weight:600;
          display:flex;
          align-items:center;
          gap:8px;
        ">
          <span>Активируйте аккаунт через Telegram, чтобы совершать покупки. Для этого перейдите в Мой кабинет</span>
        </div>
      `;
      const h3 = document.querySelector('#view-catalog h3:nth-of-type(1)');
      if (h3) h3.insertAdjacentElement('afterend', notice);
    }

  function safeText(str){
    return String(str || '').replace(/\+/g, ' ');
  }

  updateCatalogBalance();

  const onlyAvail = document.getElementById('only_available_toggle').checked;
  const cat = (document.getElementById('cat_filter').value || '').trim();
  const dir = (document.getElementById('sort_dir').value || 'asc');

  let items = (lastDashboard.catalog || []).slice();
  if (cat) items = items.filter(it => String(it.category || '').trim() === cat);

  let avail = items.filter(it => it.can_afford);
  let unavail = items.filter(it => !it.can_afford);

  const byPrice = (a,b) => {
    const d = (a.points_price||0) - (b.points_price||0);
    return dir === 'asc' ? d : -d;
  };
  avail.sort(byPrice);
  unavail.sort(byPrice);

  const secA=document.getElementById('catalog_available'); secA.innerHTML='';
  const secU=document.getElementById('catalog_unavailable'); secU.innerHTML='';

  document.getElementById('catalog_available_empty').style.display = avail.length?'none':'block';

  // ===== ДОСТУПНЫЕ
  avail.forEach(it => {
    const disabledAttr = isLinked ? '' : 'disabled';
    const disabledStyle = isLinked ? '' : 'style="opacity:0.6;cursor:not-allowed"';
    const disabledTitle = isLinked ? '' : 'Недоступно. Подтвердите аккаунт через Telegram';

    const el = document.createElement('div');
    el.className = 'item';

    el.innerHTML = `
      ${it.photo_url ? `<div class="thumb" onclick="openImg('${it.photo_url}')">
        <img src="${it.photo_url}" alt="${safeText(it.title)}">
      </div>` : ''}

      <div style="font-weight:700; word-break:break-word;">${safeText(it.title)}</div>
      <div class="muted" style="margin:4px 0; word-break:break-word;">${safeText(it.desc || '')}</div>

      <div class="flex">
        <div style="white-space:nowrap;">
          <b>${formatNum(it.points_price)}</b>\u00A0баллов
        </div>
        <button
          class="btn-buy"
          data-item="${it.item_id}"
          onclick="redeem('${it.item_id}')"
          ${disabledAttr}
          ${disabledStyle}
          title="${disabledTitle}"
        >Купить</button>
      </div>

      <div class="muted" style="margin-top:4px">
        Остаток: ${formatNum(it.stock)}
      </div>
    `;
    secA.appendChild(el);
  });

  // ===== НЕДОСТУПНЫЕ
  document.getElementById('catalog_unavailable_empty').style.display =
    (onlyAvail || unavail.length) ? 'none' : 'block';

  const sectionHeaderUnavailable = document.querySelector('#view-catalog h3:nth-of-type(2)');
  if (sectionHeaderUnavailable) {
    sectionHeaderUnavailable.style.display = onlyAvail ? 'none' : 'block';
  }

  document.getElementById('catalog_unavailable').style.display =
    onlyAvail ? 'none' : 'grid';

  if (!onlyAvail){
    unavail.forEach(it => {
      const el = document.createElement('div');
      el.className = 'item';

      el.innerHTML = `
        ${it.photo_url ? `<div class="thumb" onclick="openImg('${it.photo_url}')">
          <img src="${it.photo_url}" alt="${safeText(it.title)}">
        </div>` : ''}

        <div style="font-weight:700; word-break:break-word;">${safeText(it.title)}</div>
        <div class="muted" style="margin:4px 0; word-break:break-word;">${safeText(it.desc || '')}</div>

        <div class="flex">
          <div style="white-space:nowrap;">
            <b>${formatNum(it.points_price)}</b>\u00A0баллов
          </div>
          <button disabled>Купить</button>
        </div>

        <div class="muted" style="margin-top:4px">
          Не хватает: <b>${formatNum(it.missing)}</b> • Остаток: ${formatNum(it.stock)}
        </div>
      `;
      secU.appendChild(el);
    });
  }
}

async function loadCatalogFresh() {
  setBusy(true);

  try {
    // свежие данные дашборда (чтобы были актуальные балансы и каталог)
    const d = await apiCall({
      action: 'getDashboard',
      token
    });
    lastDashboard = d;

    // узнаём телеграм-состояние
    const linked = await fetchTelegramLink();

    // теперь рендер
    populateCategoryFilter();
    updateCatalogBalance();
    renderCatalogSections(linked); // <-- передаём в отрисовку

    // применяем телеграм-статус к кнопкам "Купить" и блоку tg_verify_block
    applyTelegramLockUI(linked);

  } catch(err){
    flashMsg(
      'cat_msg',
      `<div class="alert">Ошибка обновления каталога: ${err.message}</div>`,
      7000
    );
  }

  setBusy(false);
}

function redeem(item_id){
  const item = (lastDashboard?.catalog || []).find(x => String(x.item_id) === String(item_id));
  const title = item?.title || 'товар';
  const price = Number(item?.points_price || 0).toLocaleString('ru-RU');

  const ok = confirm(`Купить «${title}» за ${price} баллов?\n\nПодтвердить списание и создать заявку?`);
  if (!ok) return;

  const msg = document.getElementById('cat_msg');
  if (msg) msg.textContent='⏳ Создаём заявку...';

  setBusy(true);
  apiCall({
    action:'redeem',
    token,
    item_id
  })
  .then(() => {
    flashMsg('cat_msg', '✅ Заявка создана');

    // подтянуть свежий дашборд и перейти в "Мои заявки"
    return apiCall({
      action:'getDashboard',
      token
    }).then(d=>{
      setBusy(false);
      lastDashboard = d;
      nav('orders');
      renderOrders(lastDashboard?.redemptions || []);
    });
  })
  .catch(err => {
    setBusy(false);
    flashMsg('cat_msg', `<div class="alert">Ошибка: ${err.message}</div>`, 7000);
  });
}

/* =========================
   ORDERS (моё)
========================= */
function dtLabel_(r){
  // отдаём красивую строку, если бэк уже прислал
  if (r.created_at_display) return r.created_at_display;
  // иначе соберём из Unix-ts (секунды)
  if (r.created_at_ts){
    try{
      const d = new Date(r.created_at_ts * 1000);
      return d.toLocaleString('ru-RU', {
        day:'2-digit', month:'2-digit', year:'numeric',
        hour:'2-digit', minute:'2-digit'
      });
    }catch(_){}
  }
  return '';
}

async function loadOrdersFresh(){
  const wrap  = document.getElementById('orders');
  const empty = document.getElementById('orders_empty');
  if (wrap)  wrap.innerHTML = '⏳ Загружаем данные...';
  if (empty) empty.style.display = 'none';

  setBusy(true);
  try{
    const d = await apiCall({ action:'getDashboard', token });
    lastDashboard = d;
    renderOrders(d.redemptions || []);
  } catch(e){
    if (wrap) wrap.innerHTML = `<div class="alert">Ошибка: ${e.message}</div>`;
  }
  setBusy(false);
}

function renderOrders(list){
  const wrap = document.getElementById('orders');
  const empty = document.getElementById('orders_empty');
  wrap.innerHTML = '';
  if (!list || !list.length){ empty.style.display='block'; return; }
  empty.style.display='none';

  const statusRu = {
    waiting:   'в ожидании',
    approved:  'одобрено',
    ready:     'готов к выдаче',
    delivered: 'выдано',
    canceled:  'отменено',
    rejected:  'отклонено',
    failed:    'ошибка'
  };

  const order = ['approved','ready','waiting','delivered','canceled'];
  const titles = {
      approved:  'Подтверждено (ожидает подготовки к выдаче)',
      ready:     'Готов к выдаче',
      waiting:   'В обработке',
      delivered: 'Выдано',
      canceled:  'Отменено'
  };

  const groups = {};
  list.forEach(r=>{
    const st = String(r.status||'').toLowerCase();
    const key = order.includes(st) ? st : '_other';
    (groups[key] ||= []).push(r);
  });

  order.concat('_other').forEach(key=>{
    const items = groups[key]; if (!items || !items.length) return;

    const title = key==='_other' ? 'Прочее' : titles[key];
    const h = document.createElement('div');
    h.className = 'status-group-title';
    h.textContent = title;
    wrap.appendChild(h);

    items.forEach(r=>{
      const st = String(r.status||'').toLowerCase();
      const stRu = statusRu[st] || r.status;

      const el = document.createElement('div');
      el.className = 'item ' + (key==='_other' ? '' : ('status-' + key));
      const dt = fmtDateTime(r.created_at_display, r.created_at_ts);
      el.innerHTML = `
          <div class="flex">
            <div><b>${r.title}</b></div>
            <div class="muted" style="text-align:right">#${r.redeem_id}</div>
          </div>
        
          <div class="flex" style="margin-top:6px">
            <div>${r.points_spent} баллов</div>
            <div><span class="muted">статус:</span> <b class="status-pill">${stRu}</b></div>
          </div>
        
          <div class="muted" style="margin-top:6px; font-size:12px;">
            ${getCreatedAtForRedemption(r)}
          </div>
        `;
      wrap.appendChild(el);
    });
  });
}

function adminOrdersInit(){
  // очистим список и поля
  const wrap  = document.getElementById('adm_orders');
  const empty = document.getElementById('adm_orders_empty');
  if (wrap)  wrap.innerHTML = '';
  if (empty) empty.style.display = 'none';

  // навесим обработчики: локальная фильтрация по кэшу
  const fStatus = document.getElementById('adm_status_filter_ru');
  const fPhone  = document.getElementById('adm_phone_search');
  const fFrom   = document.getElementById('adm_date_from');
  const fTo     = document.getElementById('adm_date_to');

  [fStatus, fPhone, fFrom, fTo].forEach(el=>{
    if (!el || el._wired) return;
    el._wired = true;
    el.addEventListener(el.tagName === 'INPUT' ? 'input' : 'change', renderAdminRedemptions);
  });

  // показываем пусто до нажатия "Обновить"
  renderAdminRedemptions();
}

/* =========================
   ADMIN: COMMON
========================= */
function adminShow(tab){
  const ordersEl  = document.getElementById('adm_tab_orders');
  const catalogEl = document.getElementById('adm_tab_catalog');
  const dealsEl   = document.getElementById('adm_tab_deals');

  ordersEl.style.display  = (tab==='orders')  ? 'block' : 'none';
  catalogEl.style.display = (tab==='catalog') ? 'block' : 'none';
  dealsEl.style.display   = (tab==='deals')   ? 'block' : 'none';

  document.getElementById('adm_tab_btn_orders').setAttribute('aria-selected', String(tab==='orders'));
  document.getElementById('adm_tab_btn_catalog').setAttribute('aria-selected', String(tab==='catalog'));
  document.getElementById('adm_tab_btn_deals').setAttribute('aria-selected', String(tab==='deals'));

  if (tab==='orders')  adminOrdersInit();
  if (tab==='catalog') adminLoadCatalog();
  if (tab==='deals')   adminDealsShow('add');
}

/* =========================
   ADMIN: REDEMPTIONS
========================= */
function adminLoadRedemptions(){
  if (!isAdmin){ nav('cabinet'); return; }
  const filter = document.getElementById('adm_status_filter').value || '';

  const statusRu = {
    waiting:   'в ожидании',
    approved:  'одобрено',
    ready:     'готов к выдаче',
    delivered: 'выдано',
    canceled:  'отменено',
    rejected:  'отклонено',
    failed:    'ошибка'
  };

  setBusy(true);
  apiCall({
    action:'adminListRedemptions',
    token,
    statusFilter: filter
  })
  .then(res => {
    setBusy(false);
    const wrap = document.getElementById('adm_orders'); wrap.innerHTML = '';
    const empty = document.getElementById('adm_orders_empty');
    empty.style.display = res.items.length ? 'none' : 'block';

    res.items.forEach(r=>{
      const st = String(r.status || '').toLowerCase();
      const stRu = statusRu[st] || r.status;

      const el = document.createElement('div');
      const canApprove = (r.status === 'waiting');       // Подтвердить
      const canMarkReady = (r.status === 'approved');    // Готов к выдаче
      const canGenCode = (r.status === 'ready');         // Выдать код
      const canConfirm = (r.status === 'await_code');    // Подтвердить выдачу
      const canCancel  = ['waiting','approved','ready','await_code'].includes(r.status);
      el.className = 'item ' + (st ? ('status-' + st) : '');
      el.innerHTML = `
          <div class="flex">
            <div><b>${r.title}</b></div>
            <div class="muted">#${r.redeem_id}</div>
          </div>
          <div class="muted" style="margin:4px 0">${r.phone}</div>
        
          <div class="flex" style="margin:6px 0">
            <div>${r.points_spent} баллов</div>
            <div><b>${stRu}</b></div>
          </div>
        
          <!-- Оставляем только эту нижнюю дату -->
          <div class="muted" style="margin-top:6px; font-size:12px;">
            ${getCreatedAtForRedemption(r)}
          </div>
        
          <div class="btn-row">
            ${canApprove  ? `<button onclick="admSet('${r.redeem_id}','approved')">Подтвердить</button>` : ''}
            ${canMarkReady? `<button onclick="admSet('${r.redeem_id}','ready')">Готов к выдаче</button>` : ''}
            ${canGenCode  ? `<button onclick="admGeneratePickupCode('${r.redeem_id}')">Выдать код</button>` : ''}
            ${canConfirm  ? `<button onclick="admConfirmPickupCode('${r.redeem_id}')">Подтвердить выдачу</button>` : ''}
            ${canCancel   ? `<button onclick="admCancelRedemption('${r.redeem_id}')">Отменить</button>` : ''}
          </div>
        `;
      wrap.appendChild(el);
    });
  })
  .catch(err => {
    setBusy(false);
    alert('Ошибка: '+err.message);
  });
}

function _digits(s){ return String(s||'').replace(/\D/g,''); }

function _tsFromAny(r){
  // 1) явный числовой unix (секунды/миллисекунды)
  if (r.created_at_ts != null) {
    const n = Number(r.created_at_ts);
    if (!Number.isNaN(n) && n > 0) {
      return n > 2e10 ? Math.floor(n/1000) : Math.floor(n); // 13 знаков -> мс
    }
  }

  const raw0 = (r.created_at_display || r.created_at || '').trim();
  if (!raw0) return null;

  // 2) если строка — число
  if (/^\d{10,13}$/.test(raw0)) {
    const n = Number(raw0);
    return n > 2e10 ? Math.floor(n/1000) : n;
  }

  // 3) формат ДД.ММ.ГГГГ[ HH:MM[:SS]]
  const m = raw0.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?$/);
  if (m) {
    const [, dd, mm, yyyy, hh='0', mi='0', ss='0'] = m;
    const d = new Date(
      Number(yyyy),
      Number(mm) - 1,
      Number(dd),
      Number(hh),
      Number(mi),
      Number(ss)
    );
    return Math.floor(d.getTime()/1000);
  }

  // 4) ISO / “YYYY-MM-DD HH:MM” и т.п.
  const iso = raw0.replace(' ', 'T');
  const t = Date.parse(iso);
  return Number.isNaN(t) ? null : Math.floor(t/1000);
}

function _dateToStartTs(dateStr){ // 'YYYY-MM-DD' -> ts начала дня (лок)
  if (!dateStr) return null;
  const d = new Date(dateStr + 'T00:00:00');
  return Math.floor(d.getTime()/1000);
}
function _dateToEndTs(dateStr){ // конец дня
  if (!dateStr) return null;
  const d = new Date(dateStr + 'T23:59:59');
  return Math.floor(d.getTime()/1000);
}

function renderAdminRedemptions(){
  const wrap  = document.getElementById('adm_orders');
  const empty = document.getElementById('adm_orders_empty');
  if (!wrap) return;
  wrap.innerHTML = '';

  const st = document.getElementById('adm_status_filter_ru')?.value || '';
  const phoneNeedle = _digits(document.getElementById('adm_phone_search')?.value || '');
  const fromTs = _dateToStartTs(document.getElementById('adm_date_from')?.value || '');
  const toTs   = _dateToEndTs  (document.getElementById('adm_date_to')?.value   || '');

  // русские подписи
  const statusRu = {
    waiting:   'в ожидании',
    approved:  'одобрено',
    ready:     'готов к выдаче',
    await_code:'ожидает код',
    delivered: 'выдано',
    canceled:  'отменено',
    rejected:  'отклонено',
    failed:    'ошибка'
  };

  // фильтрация по кэшу
  let items = _adminRedemptionsCache.slice();

  if (st) items = items.filter(r => String(r.status||'') === st);

  if (phoneNeedle){
    items = items.filter(r => _digits(r.phone).includes(phoneNeedle));
  }

  if (fromTs || toTs){
    items = items.filter(r=>{
      const t = _tsFromAny(r);
      if (t == null) return false;
      if (fromTs && t < fromTs) return false;
      if (toTs   && t > toTs)   return false;
      return true;
    });
  }

  if (!items.length){
    if (empty) empty.style.display='block';
    return;
  }
  if (empty) empty.style.display='none';

  // Рендер (то же, что было, но без верхних дат — вы их уже убрали ранее)
  items.forEach(r=>{
    const stRu = statusRu[String(r.status||'').toLowerCase()] || r.status;
    const canApprove  = (r.status === 'waiting');
    const canMarkReady= (r.status === 'approved');
    const canGenCode  = (r.status === 'ready');
    const canConfirm  = (r.status === 'await_code');
    const canCancel   = ['waiting','approved','ready','await_code'].includes(r.status);

    const el = document.createElement('div');
    el.className = 'item ' + (r.status ? ('status-' + r.status) : '');
    el.innerHTML = `
      <div class="flex">
        <div><b>${r.title}</b></div>
        <div class="muted">#${r.redeem_id}</div>
      </div>
      <div class="muted" style="margin:4px 0">${r.phone}</div>

      <div class="flex" style="margin:6px 0">
        <div>${r.points_spent} баллов</div>
        <div><b>${stRu}</b></div>
      </div>

      <div class="muted" style="margin-top:6px; font-size:12px;">
        ${getCreatedAtForRedemption(r)}
      </div>

      <div class="btn-row">
        ${canApprove  ? `<button onclick="admSet('${r.redeem_id}','approved')">Подтвердить</button>` : ''}
        ${canMarkReady? `<button onclick="admSet('${r.redeem_id}','ready')">Готов к выдаче</button>` : ''}
        ${canGenCode  ? `<button onclick="admGeneratePickupCode('${r.redeem_id}')">Выдать код</button>` : ''}
        ${canConfirm  ? `<button onclick="admConfirmPickupCode('${r.redeem_id}')">Подтвердить выдачу</button>` : ''}
        ${canCancel   ? `<button onclick="admCancelRedemption('${r.redeem_id}')">Отменить</button>` : ''}
      </div>
    `;
    wrap.appendChild(el);
  });
}

function admSet(id, st){
  if (!isAdmin) return;
  setBusy(true);
  apiCall({
    action:'adminUpdateRedemption',
    token,
    redeem_id: id,
    new_status: st
  })
  .then(() => {
    setBusy(false);
    adminLoadRedemptions();
  })
  .catch(err => {
    setBusy(false);
    alert('Ошибка: '+err.message);
  });
}

/* =========================
   ADMIN: CATALOG
========================= */
function _wireCategorySelect(){
  const sel = document.getElementById('new_category_select');
  const custom = document.getElementById('new_category_custom');
  if (!sel || !custom) return;
  sel.addEventListener('change', ()=>{
    if (sel.value === '__add__'){
      custom.style.display = 'block';
      custom.focus();
    } else {
      custom.style.display = 'none';
      custom.value = '';
    }
  });
}

function populateAdminCategoryOptionsFromList(items){
  const sel = document.getElementById('new_category_select');
  if (!sel) return;

  const cats = Array.from(new Set(
    (items||[])
      .map(it => String(it.category || '').trim())
      .filter(Boolean)
  )).sort((a,b)=> a.localeCompare(b, 'ru'));

  const prev = sel.value;
  const addVal = '__add__';
  sel.innerHTML = `<option value="">— выбрать —</option>` +
    cats.map(c=>`<option value="${c}">${c}</option>`).join('') +
    `<option value="${addVal}">+ Добавить категорию…</option>`;

  if (cats.includes(prev)) sel.value = prev;
  else if (prev === addVal) sel.value = addVal;

  _wireCategorySelect();
}

function renderAdminCatalog(){
  const wrap = document.getElementById('adm_catalog');
  if (!wrap) return;
  wrap.innerHTML = '';

  let items = adminCatalogItems.slice();

  const catSel   = document.getElementById('adm_cat_filter');
  const searchEl = document.getElementById('adm_search_filter');
  const catVal   = (catSel?.value || '').trim().toLowerCase();
  const needle   = (searchEl?.value || '').trim().toLowerCase();

  if (catVal){
    items = items.filter(it => String(it.category||'').trim().toLowerCase() === catVal);
  }
  if (needle){
    items = items.filter(it => ( (it.title||'') + ' ' + (it.desc||'') ).toLowerCase().includes(needle));
  }

  if (!items.length){
    wrap.innerHTML = '<div class="muted">Ничего не найдено.</div>';
    return;
  }

  items.forEach(it=>{
    const el = document.createElement('div');
    el.className = 'item';

    const imgBlock = it.photo_url
      ? `<div class="thumb small" onclick="openImg('${it.photo_url}')">
          <img src="${it.photo_url}" alt="${(it.title||'').replace(/"/g,'&quot;')}">
        </div>`
      : '';

    el.innerHTML = `
      ${imgBlock}
      <div style="font-weight:700">${it.title || ''}</div>
      <div class="muted">${it.category || ''}</div>

      <div class="flex" style="margin-top:6px; flex-wrap:wrap; gap:12px; align-items:flex-end">
        <div>
          <label class="muted" style="font-size:12px;display:block;margin-bottom:4px">
            Цена (баллы)
          </label>
          <input
            id="price_${it.item_id}"
            type="number"
            min="1"
            step="1"
            value="${Number(it.points_price||0)}"
            style="width:110px;padding:8px;border:1px solid #d1d5db;border-radius:10px;font-size:14px"
          />
        </div>

        <div>
          <label class="muted" style="font-size:12px;display:block;margin-bottom:4px">
            stock
          </label>
          <input
            id="stk_${it.item_id}"
            type="number"
            min="0"
            step="1"
            value="${Number(it.stock||0)}"
            style="width:86px;padding:8px;border:1px solid #d1d5db;border-radius:10px;font-size:14px"
          />
        </div>
      </div>

      <div class="flex" style="margin-top:6px;align-items:center;gap:8px">
        <div class="muted">Активен:</div>
        <input id="act_${it.item_id}" type="checkbox" ${it.is_active ? 'checked' : ''}>
      </div>

      <div style="margin-top:8px">
        <button onclick="admSaveItem('${it.item_id}')">Сохранить</button>
      </div>
    `;
    wrap.appendChild(el);
  });
}

function adminLoadCatalog(){
  if (!isAdmin) return;

  const wrap  = document.getElementById('adm_catalog');
  if (wrap) wrap.innerHTML = '⏳ Загружаем каталог...';

  setBusy(true);
  apiCall({
    action:'adminListCatalog',
    token
  })
  .then(res => {
    setBusy(false);
    const items = Array.isArray(res?.items) ? res.items : [];
    adminCatalogItems = items;

    // заполнить категории для добавления товара
    try { populateAdminCategoryOptionsFromList(items); } catch(e){ console.warn(e); }

    // заполнить фильтр категорий в админке
    const sel = document.getElementById('adm_cat_filter');
    if (sel){
      const cats = [...new Set(
        items.map(it => String(it.category || '').trim()).filter(Boolean)
      )].sort((a,b)=>a.localeCompare(b,'ru'));

      const prev = sel.value;
      sel.innerHTML = `<option value="">Все категории</option>` +
        cats.map(c=>`<option value="${c}">${c}</option>`).join('');
      if (cats.includes(prev)) sel.value = prev;
    }

    renderAdminCatalog();
  })
  .catch(err => {
    setBusy(false);
    if (wrap) wrap.innerHTML = `<div class="alert">Ошибка: ${err.message}</div>`;
  });
}

function admSaveItem(item_id){
  if (!isAdmin) return;

  const stockEl  = document.getElementById('stk_'   + item_id);
  const activeEl = document.getElementById('act_'   + item_id);
  const priceEl  = document.getElementById('price_' + item_id);

  const stock         = Number(stockEl?.value  || 0);
  const is_active     = !!activeEl?.checked;
  const points_price  = Number(priceEl?.value  || 0);

  setBusy(true);
  apiCall({
    action:'adminUpdateCatalogItem',
    token,
    payload: JSON.stringify({ item_id, stock, is_active, points_price })
  })
  .then(() => {
    setBusy(false);
    adminLoadCatalog();
  })
  .catch(err => {
    setBusy(false);
    alert('Ошибка: ' + err.message);
  });
}

/* Добавление товара (с картинкой) */
document.addEventListener('change', e=>{
  if (e.target && e.target.id==='new_image') {
    const file = e.target.files && e.target.files[0];
    const prev = document.getElementById('new_image_preview');
    _newImageDataUrl = '';
    if (!file){ if (prev) prev.style.display='none'; return; }
    const reader = new FileReader();
    reader.onload = function(evt){
      _newImageDataUrl = String(evt.target.result || '');
      if (prev){
        prev.src = _newImageDataUrl;
        prev.style.display = 'block';
      }
    };
    reader.readAsDataURL(file);
  }
});

function adminAddItem(){
  if (!isAdmin){ nav('cabinet'); return; }

  const btn = document.getElementById('btn_add_item');
  const msg = document.getElementById('adm_add_msg');

  const titleEl   = document.getElementById('new_title');
  const sel       = document.getElementById('new_category_select');
  const custom    = document.getElementById('new_category_custom');
  const descEl    = document.getElementById('new_desc');
  const priceKztEl= document.getElementById('new_price_tenge'); // ← НОВОЕ имя поля
  const stockEl   = document.getElementById('new_stock');
  const activeEl  = document.getElementById('new_active');

  const title = (titleEl?.value || '').trim();
  let category = '';
  if (sel){
    category = sel.value === '__add__' ? (custom?.value.trim() || '') : (sel.value || '');
  }
  const desc        = (descEl?.value || '').trim();
  const priceTenge  = Number(priceKztEl?.value || 0);   // ← читаем цену в тенге
  const stock       = Number(stockEl?.value || 0);
  const is_active   = !!(activeEl?.checked);

  if (!title){ msg.innerHTML = '<div class="alert">Укажите название</div>'; return; }
  if (!(priceTenge > 0)){ msg.innerHTML = '<div class="alert">Укажите корректную цену в тенге</div>'; return; }

  msg.textContent = '⏳ Сохраняем...';
  if (btn) btn.disabled = true;
  setBusy(true);

  apiCall({
    action:'adminAddCatalogItem',
    token,
    // Бэку теперь отправляем price_tenge; points_price бэк сам посчитает по текущему ball_rate
    payload: JSON.stringify({
      title,
      category,
      desc,
      price_tenge: priceTenge,          // ← ВАЖНО: поле в тенге
      // points_price: можно опционально передать, если хотите вручную зафиксировать баллы
      stock,
      is_active,
      image_data_url: _newImageDataUrl || ''
    })
  })
  .then(() => {
    setBusy(false);
    if (btn) btn.disabled = false;
    flashMsg('adm_add_msg', '✅ Товар добавлен');

    // Безопасная очистка полей (с null-гардами)
    if (titleEl) titleEl.value = '';
    if (sel) sel.value = '';
    if (custom){ custom.value=''; custom.style.display='none'; }
    if (descEl) descEl.value = '';
    if (priceKztEl) priceKztEl.value = '';   // ← чистим правильное поле
    if (stockEl) stockEl.value = '';
    if (activeEl) activeEl.checked = true;

    _newImageDataUrl = '';
    const prev = document.getElementById('new_image_preview');
    if (prev){ prev.src=''; prev.style.display='none'; }
    const inp = document.getElementById('new_image'); if (inp) inp.value = '';

    adminLoadCatalog();
  })
  .catch(err => {
    setBusy(false);
    if (btn) btn.disabled = false;
    flashMsg('adm_add_msg', `<div class="alert">Ошибка: ${err.message}</div>`, 7000);
  });

  // защитный таймер, чтобы не осталась навечно задизейблена
  setTimeout(()=>{ if (btn) btn.disabled = false; }, 2000);
}

/* =========================
   ADMIN: DEALS
========================= */
function adminDealsShow(which){
  const isAdd = (which === 'add');

  document.getElementById('adm_deal_add').style.display  = isAdd ? 'block' : 'none';
  document.getElementById('adm_deal_list').style.display = isAdd ? 'none'  : 'block';

  document.getElementById('deal_tab_add').setAttribute('aria-selected',  String(isAdd));
  document.getElementById('deal_tab_list').setAttribute('aria-selected', String(!isAdd));

  if (!isAdd){
    resetDealsListView();
  }
}

let _currentBallRate = 1;

function initBallRatePreview(){
  // 1. Получаем курс с сервера
  google.script.run
    .withSuccessHandler(rate => {
      _currentBallRate = rate;
      const input = document.getElementById('new_price_tenge');
      const preview = document.getElementById('price_preview');
      if (!input || !preview) return;

      input.addEventListener('input', () => {
        const val = Number(input.value || 0);
        if (val > 0 && _currentBallRate > 0){
          const pts = Math.round(val / _currentBallRate);
          preview.innerText = `≈ ${pts.toLocaleString('ru-RU')} баллов (по курсу ${_currentBallRate} ₸/балл)`;
        } else {
          preview.innerText = '';
        }
      });
    })
    .getBallRate_();
}

// Запускаем при загрузке админ-панели
document.addEventListener('DOMContentLoaded', initBallRatePreview);

function calcPointsPreview_(){
  const t = (document.getElementById('deal_type').value||'').toUpperCase();
  const price = Number(document.getElementById('deal_price').value||0);
  let pts = 0;
  if (['КВ','ВП'].includes(t)) pts = Math.round(price*0.0007);
  else if (['КП','ПМ'].includes(t)) pts = Math.round(price*0.001);
  document.getElementById('deal_points').value = pts ? pts.toLocaleString('ru-RU') : '';
}
document.addEventListener('input', e=>{
  if (e.target && (e.target.id==='deal_type' || e.target.id==='deal_price')){
    calcPointsPreview_();
  }
});

function adminAddDeal(){
  if (!isAdmin){ nav('cabinet'); return; }

  const contract_id  = document.getElementById('deal_contract').value.trim();
  const phone        = document.getElementById('deal_phone').value.trim();
  const permise_type = document.getElementById('deal_type').value.trim();
  const price        = Number(document.getElementById('deal_price').value||0);
  const status       = document.getElementById('deal_status').value;
  const msg = document.getElementById('adm_deal_msg');

  if (!contract_id){ msg.innerHTML='<div class="alert">Укажите номер договора</div>'; return; }
  if (!phone){ msg.innerHTML='<div class="alert">Укажите телефон</div>'; return; }
  if (!permise_type){ msg.innerHTML='<div class="alert">Выберите вид помещения</div>'; return; }
  if (!(price>0)){ msg.innerHTML='<div class="alert">Цена должна быть > 0</div>'; return; }

  setBusy(true);
  apiCall({
    action:'adminAddPurchase',
    token,
    payload: JSON.stringify({
      contract_id,
      phone,
      permise_type,
      price,
      status
    })
  })
  .then(() => {
    setBusy(false);
    flashMsg(
      'adm_deal_msg',
      '<span style="display:flex;align-items:center;gap:8px;font-weight:600;color:#065f46;">' +
        '<span style="width:18px;height:18px;border:2px solid #065f46;border-radius:4px;display:flex;align-items:center;justify-content:center;font-size:14px;line-height:1;">✓</span>' +
        '<span>Сделка добавлена</span>' +
      '</span>',
      5000
    );

    document.getElementById('deal_contract').value='';
    document.getElementById('deal_phone').value='';
    document.getElementById('deal_type').value='';
    document.getElementById('deal_price').value='';
    document.getElementById('deal_status').value='credited_pending';
    document.getElementById('deal_points').value='';

    // обновим Dashboard (баланс юзера может поменяться)
    return apiCall({
      action:'getDashboard',
      token
    }).then(d=>{
      lastDashboard = d;
      if (currentView==='catalog') {
        updateCatalogBalance();
        renderCatalogSections();
      }
    });
  })
  .catch(err => {
    setBusy(false);
    msg.innerHTML = `<div class="alert">Ошибка: ${err.message}</div>`;
  });
}

function resetDealsListView(){
  _adminDealsCache = [];
  const wrap  = document.getElementById('deal_list_wrap');
  const empty = document.getElementById('deal_list_empty');
  const msg   = document.getElementById('deal_list_msg');
  if (wrap)  wrap.innerHTML = '';
  if (msg)   msg.innerHTML  = '';
  if (empty) empty.style.display = 'none';
}

function adminLoadDeals(){
  if (!isAdmin){ nav('cabinet'); return; }

  const q = document.getElementById('deal_search').value.trim() || '';

  setBusy(true);
  apiCall({
    action:'adminListPurchases',
    token,
    search:q
  })
  .then(res => {
    setBusy(false);
    _adminDealsCache = Array.isArray(res.items) ? res.items : [];
    renderDealsList();
  })
  .catch(err => {
    setBusy(false);
    flashMsg('deal_list_msg', `<div class="alert">Ошибка: ${err.message}</div>`, 7000);
  });
}

function renderDealsList(){
  const wrap  = document.getElementById('deal_list_wrap');
  const empty = document.getElementById('deal_list_empty');
  const sortSel = document.getElementById('deal_sort');

  if (!wrap) return;
  wrap.innerHTML = '';

  let items = _adminDealsCache.slice();
  const dir = (sortSel?.value || 'desc');
  items.sort((a,b)=>{
    return dir === 'asc'
      ? (a.updated_at_ts - b.updated_at_ts)
      : (b.updated_at_ts - a.updated_at_ts);
  });

  if (!items.length){
    if (empty) empty.style.display='block';
    return;
  }
  if (empty) empty.style.display='none';

  items.forEach(deal=>{
    const el = document.createElement('div');
    el.className = 'item';
    el.style.background = '#fff';
    el.style.border = '1px solid #e5e7eb';
    el.style.borderRadius = '12px';
    el.style.padding = '12px';

    const statusSelectId = `deal_status_${deal.contract_id}`;
    const pointsInputId  = `deal_points_${deal.contract_id}`;

    el.innerHTML = `
      <div class="flex">
        <div style="font-weight:700">${deal.contract_id}</div>
        <div class="muted">${deal.updated_at_display}</div>
      </div>

      <div class="grid" style="grid-template-columns:repeat(auto-fit,minmax(140px,1fr));gap:8px;margin-top:8px">
        <div>
          <div class="muted" style="font-size:12px">Телефон</div>
          <div style="font-weight:600">${deal.phone}</div>
        </div>
        <div>
          <div class="muted" style="font-size:12px">Тип</div>
          <div>${deal.permise_type || '-'}</div>
        </div>
        <div>
          <div class="muted" style="font-size:12px">Цена</div>
          <div>${Number(deal.price||0).toLocaleString('ru-RU')}</div>
        </div>
        <div>
          <div class="muted" style="font-size:12px">Баллы</div>
          <input
            id="${pointsInputId}"
            type="number"
            min="0"
            step="1"
            value="${Number(deal.points||0)}"
            style="width:100%;padding:8px;border:1px solid #d1d5db;border-radius:10px;font-size:14px"
          />
        </div>
      </div>

      <div class="flex" style="margin-top:10px; flex-wrap:wrap; gap:12px; align-items:flex-end">
        <div>
          <div class="muted" style="font-size:12px">Статус</div>
          <select id="${statusSelectId}" style="min-width:170px;padding:8px;border:1px solid #d1d5db;border-radius:10px;font-size:14px">
            <option value="credited_pending"   ${deal.status==='credited_pending'?'selected':''}>credited_pending</option>
            <option value="credited_available" ${deal.status==='credited_available'?'selected':''}>credited_available</option>
            <option value="termination"        ${deal.status==='termination'?'selected':''}>termination</option>
          </select>
        </div>

        <div>
          <button onclick="adminSaveDealStatus('${deal.contract_id}', '${statusSelectId}', '${pointsInputId}')">
            Сохранить
          </button>
        </div>
      </div>
    `;

    wrap.appendChild(el);
  });
}

function adminSaveDealStatus(contract_id, selectId, pointsId){
  if (!isAdmin) return;

  const sel = document.getElementById(selectId);
  const pt  = document.getElementById(pointsId);
  if (!sel || !pt) return;

  const newStatus = sel.value;
  const newPoints = Number(pt.value || 0);

  setBusy(true);
  apiCall({
    action:'adminUpdatePurchase',
    token,
    contract_id,
    new_status: newStatus,
    new_points: newPoints
  })
  .then(() => {
    flashMsg('deal_list_msg', '✅ Изменения сохранены', 5000);

    return apiCall({
      action:'adminListPurchases',
      token,
      search: document.getElementById('deal_search').value.trim() || ''
    }).then(res=>{
      _adminDealsCache = Array.isArray(res.items) ? res.items : [];
      renderDealsList();

      // баланс тоже мог измениться
      return apiCall({
        action:'getDashboard',
        token
      }).then(d=>{
        lastDashboard = d;
        if (currentView==='catalog') {
          updateCatalogBalance();
          renderCatalogSections();
        }
        setBusy(false);
      });
    });
  })
  .catch(err => {
    setBusy(false);
    flashMsg('deal_list_msg', `<div class="alert">Ошибка: ${err.message}</div>`, 7000);
  });
}

/* =========================
   IMAGE MODAL
========================= */
function openImg(url){
  const m = document.getElementById('img_modal');
  const pic = document.getElementById('img_modal_pic');
  if (!m || !pic) return;
  pic.src = url || '';
  m.classList.add('open');
  document.body.style.overflow = 'hidden';
}
function closeImg(){
  const m = document.getElementById('img_modal');
  const pic = document.getElementById('img_modal_pic');
  if (!m || !pic) return;
  m.classList.remove('open');
  pic.src = '';
  document.body.style.overflow = '';
}
document.addEventListener('keydown', e=>{ if (e.key==='Escape') closeImg(); });

/* =========================
   INIT
========================= */
window.addEventListener('load', ()=> {
  applyPhoneMask('#deal_phone');
  applyPhoneMask('#reg_phone');
  applyPhoneMask('#log_phone');
  applyDobMask('#reg_dob');
  applyPhoneMask('#reset_phone');

  // hero-картинка
  const heroImg = document.getElementById('hero_img');
  if (heroImg) {
    heroImg.src = HERO_URL;
    heroImg.style.display = 'block';
  }

  const availTgl = document.getElementById('only_available_toggle');
  if (availTgl)  availTgl.addEventListener('change', () => renderCatalogSections(tgLinked));
  const catFilter = document.getElementById('cat_filter');
  if (catFilter) catFilter.addEventListener('change', () => renderCatalogSections(tgLinked));
  const sortDir = document.getElementById('sort_dir');
  if (sortDir)   sortDir.addEventListener('change', () => renderCatalogSections(tgLinked));

  const admCatFilter = document.getElementById('adm_cat_filter');
  if (admCatFilter) admCatFilter.addEventListener('change', renderAdminCatalog);
  const admSearchFilter = document.getElementById('adm_search_filter');
  if (admSearchFilter) admSearchFilter.addEventListener('input', renderAdminCatalog);

  // глобальный Enter: регистрация/вход
  document.addEventListener('keydown', e=>{
    if (e.key!=='Enter') return;
    const tag=(e.target.tagName||'').toLowerCase();
    if(tag==='button'||tag==='a') return;
    if (_busyCount>0) return;
    if (currentView==='register') { e.preventDefault(); register(); }
    else if (currentView==='login') { e.preventDefault(); login(); }
  });

  if (token) {
    document.getElementById('authNav').style.display='flex';
    nav('cabinet');
  } else {
    showLanding();
  }
  updateRegisterBtnState();
});
</script>