/******************************************************
 * 無知ノ知 撮影管理 - 完全安定統合版（Part1）
 * 2025-10-16（全量・整合済）
 ******************************************************/

/* ================= 基本設定 ================= */
const CONFIG = {
  TZ: 'Asia/Tokyo',
  SS_ID: '',
  SHEETS: {
    MAIN: '顧客管理',
    PRICE: '価格',
    SETTINGS: '設定',
    HEARING: 'ムービーヒアリング',
    LOCS: '撮影地リスト'
  },
  COLS: {
    LINK: '顧客用ページ',
    INTERNAL_LINK: '社内用ページ',
BRIDE: '新婦様お名前',
GROOM: '新郎様お名前',
    PHOTO: '撮影日',
    PHOTO_DONE: '写真納品',
    VIDEO_DONE: '動画納品',
    PLAN_AUTO: 'プラン（自動）',
    PLAN_MAN: 'プラン（手動）',
    STATUS_R: 'スケジュール',
    LOC: '撮影地',
    LOC_FIX: '撮影地（確定）',
    CAMERA: 'カメラマン',
    DONE: '最終完了'
  },
  PRICE_INCLUDE_KEYS: {
    HAIR: '含有_ヘアメイク',
    SALON: '含有_サロン',
    DRESS: '含有_ドレス',
    BOUQUET: '含有_ブーケ',
    TUX: '含有_タキシード',
    PROFILE: '含有_プロフィール'
  },
  FEATURE_HEADERS: ['ヘアメイク','サロン','スケジュール','ドレス','ブーケ','タキシード','プロフィール'],
  DEADLINE: {
     CALENDAR_ID_SHOOT: 'c_db085b08ac1ca83a0bb99674620e263339e81999c7f4ffb4de0d190e8369858f@group.calendar.google.com',   // 📸 撮影用
  CALENDAR_ID_DEADLINE: 'c_9c2aa6354b2d7955a57aedb7c7490339c25237ddbc6ce182cfbd6c56ffa5c42b@group.calendar.google.com', // ⏰ 締切用
    CHAT_WEBHOOK: 'https://chat.googleapis.com/v1/spaces/AAQARnwfhmQ/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=1SibxR4B0U6F50iyPlV3qBolb7tEoBNFmQ6MIGPzp6A',
    VALUE_UNDECIDED: '未決定',
    VALUE_NONE: 'なし',
    PROP_NS: 'deadlineMgr_v9',
    ITEMS: {
      'ヘアメイク':   { col: 'P', offsetDays: -30, type: 'undecided', title: 'ヘアメイク締切' },
      'サロン':       { col: 'Q', offsetDays: -25, type: 'undecided', title: 'サロン締切' },
      'スケジュール': { col: 'R', offsetDays: -15, type: 'undecided', title: 'スケジュール締切' },
      'ブーケ':       { col: 'T', offsetDays: -20, type: 'undecided', title: 'ブーケ締切' },
      'プロフィール': { col: 'V', offsetDays: 10,  type: 'undecided', title: 'プロフィール締切' },
      '写真納品':     { chkCol: 'H', offsetDays: 13, type: 'checkbox', title: '写真納品締切' },
      '動画納品':     { chkCol: 'J', offsetDays: 30, type: 'checkbox', title: '動画納品締切' }
    },
    REMIND: {
      buildOffsets(offset){
        if (offset < 0) {
          // 撮影日前の締切（例：ヘアメイク30日前）
          // → 締切まで残り○日前リマインド
          return [offset + 1, offset + 2, offset + 3, offset + 5];
        } else {
          // 撮影後の締切（例：写真納品13日後、動画納品30日後）
          // → 締切の○日前リマインド（前倒し通知）
          return [offset - 1, offset - 2, offset - 3, offset - 5];
        }
      },
      OVERDUE_MAX_DAYS: 30
    }
  }
};


/* ================= Utility ================= */
const U = {
  ss(){ try{const a=SpreadsheetApp.getActiveSpreadsheet();if(a)return a;}catch(e){} if(!CONFIG.SS_ID) throw new Error('SS_ID未設定'); return SpreadsheetApp.openById(CONFIG.SS_ID); },
  sh(name){ const s=U.ss().getSheetByName(name); if(!s) throw `シートが見つかりません: ${name}`; return s; },
  fmt(d,p='yyyy/MM/dd'){ return Utilities.formatDate(d, CONFIG.TZ, p); },
  todayYmd(){ const n=new Date(); return new Date(n.getFullYear(),n.getMonth(),n.getDate()); },
  getHeaders(sh){ return sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(v=>String(v||'').trim()); },
  colOf(hs,n){ const i=hs.indexOf(n); if(i===-1) throw `ヘッダーが見つかりません: ${n}`; return i+1; },
  getVal(sh,c,r){ return sh.getRange(`${c}${r}`).getValue(); },
  setVal(sh,c,r,v){ sh.getRange(`${c}${r}`).setValue(v); },
  rich(t,u){ return SpreadsheetApp.newRichTextValue().setText(t).setLinkUrl(u).build(); },
  safeDate(v){
    // すでに Date 型ならそのまま
    if (v instanceof Date) return v;
    if (v == null || v === '') return null;

    let s = String(v).trim();

    // パターン1: 2026-04-30（ハイフン区切り）
    if (/^\d{4}-\d{2}-\d{2}$/.test(s)) {
      const [y,m,d] = s.split('-').map(Number);
      const dt = new Date(y, m - 1, d);
      return isNaN(dt) ? null : dt;
    }

    // パターン2: 2026/04/30 や 2026年4月30日(土)
    s = s
      .replace(/[年月]/g, '/') // 年・月 → /
      .replace('日', '')
      .replace(/[^\d/]/g, ''); // 数字と / 以外を削除（曜日など）

    const t = new Date(s);
    return isNaN(t) ? null : t;
  },  clean(s){ return String(s??'').replace(/（.*?）/g,'').replace(/\(.*?\)/g,'').replace(/[　\s]/g,'').trim(); },
  daysBetween(a,b){return Math.floor((b - a) / (1000 * 60 * 60 * 24));},
  a1(colLetter, row){ return `${colLetter}${row}`; }
};

// 数値化
function num_(v){ const n = Number(String(v).replace(/[^0-9.-]/g, '')); return isNaN(n) ? 0 : n; }
// 括弧削除ユーティリティ（税別表記も確実に削除）
function removeParenJP_(s) {
  let t = String(s || '');
  return t
    .replace(/[（(][^）)]*[）)]/g, '')  // 全角/半角括弧削除
    .replace(/税別.*?円/g, '')          // 「税別19万」などを削除
    .replace(/[　\s]/g, '')              // 空白削除
    .trim();
}



/* ================= 設定/価格読込 ================= */
const Settings = {
  read(){
    const sh=U.sh(CONFIG.SHEETS.SETTINGS);
    const parentFolderId=String(sh.getRange('A2').getValue()||'').trim();
    if(!parentFolderId) throw '設定!A2 親フォルダIDが空';
    const data=sh.getDataRange().getValues();
    const templateIds=data.filter((r,i)=>i>0&&r[1]&&String(r[2])!=='社内用テンプレ').map(r=>String(r[1]).trim());
    const internalRow=data.find(r=>r[2]==='社内用テンプレ');
    const internalDocId=internalRow?String(internalRow[1]||''):'';
    return {parentFolderId,templateIds,internalDocId};
  }
};

const Price = {
  cache: null,

  load(){
    if (Price.cache) return Price.cache;
    const sh = U.sh(CONFIG.SHEETS.PRICE);
    const vals = sh.getDataRange().getValues();
    const headers = vals[0].map(v => String(v||'').trim());
    const idxName  = headers.indexOf('表示名');
    const idxPrice = headers.indexOf('税別価格');
    if (idxName === -1 || idxPrice === -1) {
      throw new Error('価格シートに「表示名」または「税別価格」列がありません。');
    }

    const map = {};             // 行オブジェクト
    const priceByName = {};     // 表示名→税別価格（数値）
    for (let i = 1; i < vals.length; i++){
      const row = vals[i];
      const nameRaw = String(row[idxName] || '').trim();
      if (!nameRaw) continue;
      const nameKey = U.clean(nameRaw);
      const rec = {};
      headers.forEach((h, idx) => rec[h] = row[idx]);
      map[nameKey] = rec;

      const priceNum = Number(String(row[idxPrice]||'').replace(/[^0-9.-]/g,''));
      if (!isNaN(priceNum)) priceByName[nameKey] = priceNum;
    }

    Price.cache = { headers, map, priceByName };
    return Price.cache;
  },

  // プラン含有チェック
  includes(planName, key){
    if (!planName) return false;
    const { map } = Price.load();
    const p = U.clean(planName);
    const hit = Object.keys(map).find(k => U.clean(k) === p) ||
                Object.keys(map).find(k => p.includes(U.clean(k)));
    if (!hit) return false;
    return String(map[hit][key] || '') === '○';
  },

  // 表示名→税別価格（数値）
  priceOf(name){
    if (!name) return 0;
    const { priceByName } = Price.load();
    const key = U.clean(removeParenJP_(name));
    if (key in priceByName) return priceByName[key];
    const hit = Object.keys(priceByName).find(k => key.includes(k) || k.includes(key));
    return hit ? priceByName[hit] : 0;
  }
};
/* ================= プラン（自動）名 正規化 ================= */
/** 例: 「ムービープラン（税別19万）」→価格表を見て「ムービープラン（19万円（税別））」に整形 */
function normalizePlanAuto_(planAutoRaw) {
  if (!planAutoRaw) return '';

  // --- 括弧と税別表記を削除してベース名抽出 ---
  const baseName = String(planAutoRaw)
    .replace(/[（(][^）)]*[）)]/g, '')  // 全角/半角括弧削除
    .replace(/税別.*?円/g, '')          // 「税別19万」などを削除
    .trim();

  // --- キャッシュを安全にロード ---
  if (!Price.cache) Price.load();
  const cache = Price.cache || {};
  const priceByName = cache.priceByName || {};

  // --- 価格取得 ---
  let price = 0;
  if (priceByName && baseName in priceByName) {
    price = priceByName[baseName];
  } else {
    // 部分一致対応
    const hit = Object.keys(priceByName).find(k => baseName.includes(k) || k.includes(baseName));
    if (hit) price = priceByName[hit];
  }

  if (!price || isNaN(price)) {
    console.warn(`価格未取得: ${baseName}`);
    return baseName; // 該当なしの場合はそのまま返す
  }

  // --- 整形出力 ---
  return `${baseName}（${price / 10000}万円（税別））`;
}


/* ================= 顧客情報読込（列番号版） ================= */
function readRowInfo(row, opts = { includePrice: true }){
  const sh = U.sh(CONFIG.SHEETS.MAIN);
  const hs = U.getHeaders(sh);   // 他の処理用にヘッダーだけ保持（列検索は使わない）

  // ★ 列マッピング（今のシート構成前提）
  // A:1 B:2 C:3 D:4 E:5 F:6 G:7 H:8 I:9 J:10 K:11 L:12 M:13 N:14 O:15 ...
  const COL = {
    LINK:         1,   // 顧客用ページ
    INTERNAL:     2,   // 社内用ページ
    BRIDE:        4,   // 新婦様お名前
    GROOM:        5,   // 新郎様お名前
    PHOTO:        6,   // 撮影日
    PLAN_AUTO:   12,   // プラン（自動） L
    PLAN_MAN:    13,   // プラン（手動） M
    LOC:         14,   // 撮影地 N
    LOC_FIX:     15,   // 撮影地（確定） O
    CAMERA:      11    // カメラマン K
  };

  const groom   = sh.getRange(row, COL.GROOM).getDisplayValue().trim();
  const bride   = sh.getRange(row, COL.BRIDE).getDisplayValue().trim();
  const planAuto= sh.getRange(row, COL.PLAN_AUTO).getDisplayValue().trim();
  const planMan = sh.getRange(row, COL.PLAN_MAN).getDisplayValue().trim();
  const camera  = sh.getRange(row, COL.CAMERA).getDisplayValue().trim();

  // 撮影地：O列優先（確定）、空ならN列
  const loc    = sh.getRange(row, COL.LOC).getDisplayValue().trim();
  const locFx  = sh.getRange(row, COL.LOC_FIX).getDisplayValue().trim();

  // 撮影日（F列）→ 文字列でも safeDate で Date に変換
  const photoRaw  = sh.getRange(row, COL.PHOTO).getValue();
  const photoDate = U.safeDate(photoRaw);
  const photoDisp = photoDate
    ? Utilities.formatDate(photoDate, CONFIG.TZ, 'yyyy年MM月dd日')
    : '';

  // 顧客フォルダURL（A列）
  const folderCell = sh.getRange(row, COL.LINK);
  let folderUrl = '';
  try {
    folderUrl = folderCell.getRichTextValue()?.getLinkUrl() || '';
  } catch (_) {}

  // 社内用URL（B列）
  const internalCell = sh.getRange(row, COL.INTERNAL);
  let internalUrl = '';
  try {
    internalUrl = internalCell.getRichTextValue()?.getLinkUrl() || '';
  } catch (_) {}

  const planAutoNorm = opts.includePrice ? normalizePlanAuto_(planAuto) : planAuto;

  return {
    row,
    hs,
    groom,
    bride,
    planAuto,
    planAutoNorm,
    planMan,
    camera,
    // 元コードと同じプロパティ名を維持
    location: locFx || loc,
    photoDate,
    photoDisp,
    folderUrl,
    internalUrl
  };
}


/* ================= テンプレ種別推定 ================= */
function detectBaseTitle(srcName){
  const n = String(srcName || '');
  if (n.includes('請求') || n.toLowerCase().includes('invoice')) return '請求書';
  if (n.includes('案内状')) return '案内状';
  if (n.includes('よくあるご質問')) return 'よくあるご質問';
  if (n.includes('撮影準備') || n.includes('準備編')) return 'ウェディング撮影準備編';
  return n;
}

