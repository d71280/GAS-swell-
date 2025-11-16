/******************************************************
* 無知ノ知 撮影管理 - 完全安定統合版（Part3／カレンダー手動限定版）
* 2025-10-18
******************************************************/

/* ================ 補助：列番号→列記号 ================ */
function __colLetter(n){
  let s = '';
  while(n>0){ const m=(n-1)%26; s=String.fromCharCode(65+m)+s; n=Math.floor((n-1)/26); }
  return s;
}

/* ================ P〜V（機能列）自動反映 ================ */
/**
* プラン（自動）/（手動）から、ヘアメイク/サロン/スケジュール/ドレス/ブーケ/タキシード/プロフィール を自動反映
* カレンダー同期は削除し、セルの値のみを更新する
*/
function updateFeaturesRow(row){
  const sh = U.sh(CONFIG.SHEETS.MAIN);
  const hs = U.getHeaders(sh);

  const idx = (name)=> hs.indexOf(name)+1;
  const get = (c)=> String(sh.getRange(row, c).getDisplayValue() || '').trim();

  const info = readRowInfo(row);

  // 手動（複数）を配列化
  const manualItems = (info.planMan || '')
    .split(/[,、\s]+/).map(s=>s.trim()).filter(Boolean);
  const manualLower = manualItems.map(s=>s.toLowerCase());

  // 判定ヘルパー
  const hasWord = (word)=> manualLower.some(t => t.includes(word.toLowerCase()));
  const denyWord = (word)=> hasWord(word+'なし') || hasWord(word+'不要') || hasWord('no '+word);

  // 自動プラン含有チェック
  const inc = {
    HAIR:     Price.includes(info.planAuto, CONFIG.PRICE_INCLUDE_KEYS.HAIR),
    SALON:    Price.includes(info.planAuto, CONFIG.PRICE_INCLUDE_KEYS.SALON),
    DRESS:    Price.includes(info.planAuto, CONFIG.PRICE_INCLUDE_KEYS.DRESS),
    BOUQUET:  Price.includes(info.planAuto, CONFIG.PRICE_INCLUDE_KEYS.BOUQUET),
    TUX:      Price.includes(info.planAuto, CONFIG.PRICE_INCLUDE_KEYS.TUX),
    PROFILE:  Price.includes(info.planAuto, CONFIG.PRICE_INCLUDE_KEYS.PROFILE)
  };

  function decide(featureName, includeAuto){
    const f = featureName;
    if (denyWord(f)) return CONFIG.DEADLINE.VALUE_NONE;
    if (includeAuto) return CONFIG.DEADLINE.VALUE_UNDECIDED;
    if (hasWord(f))  return CONFIG.DEADLINE.VALUE_UNDECIDED;
    return CONFIG.DEADLINE.VALUE_NONE;
  }

  const col = {
    HAIR: idx('ヘアメイク'),
    SALON: idx('サロン'),
    SCHEDULE: idx('スケジュール'),
    DRESS: idx('ドレス'),
    BOUQUET: idx('ブーケ'),
    TUX: idx('タキシード'),
    PROFILE: idx('プロフィール')
  };

  const cur = {
    HAIR: col.HAIR ? sh.getRange(row,col.HAIR).getDisplayValue() : '',
    SALON: col.SALON ? sh.getRange(row,col.SALON).getDisplayValue() : '',
    SCHEDULE: col.SCHEDULE ? sh.getRange(row,col.SCHEDULE).getDisplayValue() : '',
    DRESS: col.DRESS ? sh.getRange(row,col.DRESS).getDisplayValue() : '',
    BOUQUET: col.BOUQUET ? sh.getRange(row,col.BOUQUET).getDisplayValue() : '',
    TUX: col.TUX ? sh.getRange(row,col.TUX).getDisplayValue() : '',
    PROFILE: col.PROFILE ? sh.getRange(row,col.PROFILE).getDisplayValue() : ''
  };

  const next = {
    HAIR: decide('ヘアメイク', inc.HAIR),
    SALON: decide('サロン', inc.SALON),
    SCHEDULE: CONFIG.DEADLINE.VALUE_UNDECIDED,
    DRESS: decide('ドレス', inc.DRESS),
    BOUQUET: decide('ブーケ', inc.BOUQUET),
    TUX: decide('タキシード', inc.TUX),
    PROFILE: decide('プロフィール', inc.PROFILE)
  };

Object.entries(next).forEach(([key, val]) => {
  const c = col[key];
  if (!c) return;

  const currentValue = String(cur[key] || '').trim();

  // --- 「決定」は上書き禁止 ---
  if (currentValue === '決定') {
    console.log(`🛑 ${key} は「決定」のためスキップ`);
    return;
  }

  // --- それ以外のみ自動反映 ---
  if (currentValue !== val) {
    sh.getRange(row, c).setValue(val);
  }
});

}

/* ================ onEdit：列番号固定版 ================ */
function onEdit(e){
  try{
    if (!e || !e.range) return;

    const sh = e.range.getSheet();
    if (!sh || sh.getName() !== CONFIG.SHEETS.MAIN) return;

    const row = e.range.getRow();
    const col = e.range.getColumn();
    if (row <= 1) return; // 見出し行は無視

    // 列番号（現行のシート構成前提）
    const COL = {
      CAMERA:    11, // K列 カメラマン
      PLAN_AUTO: 12, // L列 プラン（自動）
      PLAN_MAN:  13, // M列 プラン（手動）
      LOC_FIX:   15, // O列 撮影地（確定）
      LINK:       1, // A列 顧客用ページ
      INTERNAL:   2  // B列 社内用ページ
    };

    // --- L / M列の変更時は P〜V列を自動反映 ---
    if (col === COL.PLAN_AUTO || col === COL.PLAN_MAN) {
      updateFeaturesRow(row);
    }

// --- M列（プラン手動）変更時：Docs/フォルダ & カレンダー同期 ---
if (col === COL.PLAN_MAN) {
  // ロック取得（最大30秒待機）
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    
    const info = readRowInfo(row);
    const parentFolderId = Settings.read().parentFolderId;
    const parent = DriveApp.getFolderById(parentFolderId);
    const folderName = `${info.groom} × ${info.bride}　様`;
    
    // フォルダの実在チェック
    const existingFolder = parent.getFoldersByName(folderName);
    const folderExists = existingFolder.hasNext();
    
    const hasA = !!sh.getRange(row, COL.LINK).getDisplayValue();
    const hasB = !!sh.getRange(row, COL.INTERNAL).getDisplayValue();

    if (folderExists || hasA || hasB) {
      refreshExistingForRow_(row);
    } else {
      createOrUpdateClientFiles(row, { refreshOnly: false });
    }

    calendarSyncForRow_(row);
    
  } catch (err) {
    console.error('M列処理エラー:', err);
    SpreadsheetApp.getActive().toast('⚠️ 処理中にエラーが発生しました');
  } finally {
    lock.releaseLock();
  }
  return;
}

    // --- K列 / O列 変更時： Docs反映 + カレンダー同期 ---
    if (col === COL.CAMERA || col === COL.LOC_FIX) {
      refreshExistingForRow_(row);   // 社内用ページなど更新
      calendarSyncForRow_(row);      // カレンダーも更新
      return;
    }

  } catch(err){
    console.log('onEdit error', err);
  }
}


/* ================ デイリー（Chat通知のみ／軽量化版） ================ */
function dailyReminderJob(){
  const sh = U.sh(CONFIG.SHEETS.MAIN);
  const vals = sh.getDataRange().getValues();
  const headers = vals[0];
  const idx = (h) => headers.indexOf(h);
  const now = U.todayYmd();
  const cal = CalendarApp.getCalendarById(CONFIG.DEADLINE.CALENDAR_ID_DEADLINE);
  const notices = [];

  // === 行ループ（ヘッダー除く） ===
  for (let i = 1; i < vals.length; i++) {
    const row = vals[i];
    const rowNum = i + 1;
    const photoDate = row[idx('撮影日')];
    const done = row[idx('最終完了')];

    // ① 撮影日あり & 最終完了未チェックのみ
    if (!(photoDate instanceof Date) || done === true) continue;

    // 顧客情報（価格スキップ）
    const info = readRowInfo(rowNum, { includePrice: false });
    if (!info.photoDate) continue;

    // === ② P〜V列をチェック：未決定のみ対象 ===
    Object.entries(CONFIG.DEADLINE.ITEMS).forEach(([label, def]) => {
   if (def.type !== 'undecided' && def.type !== 'checkbox') return;
   // ✅ チェックボックス完了除外（H/J）
if (def.chkCol) {
  const colIndexChk = headers.indexOf(def.chkCol);
  if (colIndexChk !== -1) {
    const chkVal = row[colIndexChk];
    if (chkVal === true || String(chkVal).toLowerCase() === "true") return;
  }
}

      const colIndex = idx(label);
      if (colIndex === -1) return;
      const val = row[colIndex];
      if (val !== CONFIG.DEADLINE.VALUE_UNDECIDED) return; // 「未決定」以外除外

      // === ③ カレンダーに該当締切イベントがある場合のみ ===
      const title = DL.buildTitle(info, label, def);
      const events = cal.getEvents(
        new Date(photoDate.getTime() - 90 * 86400000),
        new Date(photoDate.getTime() + 90 * 86400000)
      );
      const exist = events.some(e => e.getTitle().includes(title));
      if (!exist) return;

      // --- 日数差分を計算 ---
      const date = new Date(photoDate.getTime() + def.offsetDays * 86400000);
      const diff = U.daysBetween(now, date);

      // --- Chat通知判定 ---
      if (diff === 0) {
        // 本日締切
        notices.push(
          '📅【' + label + '】本日が締切です\n' +
          info.groom + ' × ' + info.bride +
          (info.folderUrl ? '\n📂 ' + info.folderUrl : '')
        );

      } else if (CONFIG.DEADLINE.REMIND.buildOffsets(def.offsetDays).indexOf(diff) !== -1) {
        // 締切◯日前リマインド
        var remain = Math.abs(diff);
        notices.push(
          '⏰【リマインド】' + label + 'まで残り' + remain + '日\n' +
          info.groom + ' × ' + info.bride +
          (info.folderUrl ? '\n📂 ' + info.folderUrl : '')
        );

      } else if (diff < 0 && Math.abs(diff) <= CONFIG.DEADLINE.REMIND.OVERDUE_MAX_DAYS) {
        // 期限超過
        notices.push(
          '⚠️【期限超過】' + label + '（' + Math.abs(diff) + '日経過）\n' +
          info.groom + ' × ' + info.bride +
          (info.folderUrl ? '\n📂 ' + info.folderUrl : '')
        );
      }
      // ここまで通知判定
    }); // ← forEach閉じ
  } // ← forループ閉じ（ここが無かった！）

  // === Chat通知をまとめて1回送信 ===
  if (notices.length > 0) {
    UrlFetchApp.fetch(CONFIG.DEADLINE.CHAT_WEBHOOK, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ text: notices.join('\n\n') })
    });
    console.log('✅ ' + notices.length + '件 通知送信完了');
  } else {
    console.log('✅ 対象なし：通知なし');
  }
} // ← dailyReminderJob の閉じ



/* ================ メニュー ================ */
function onOpen(){
  SpreadsheetApp.getUi().createMenu('📂 顧客管理メニュー')
    .addItem('①新規予約の一括処理（選択行）','runNewBookingForSelectedRow_')
    .addItem('②既存データ更新（選択行）','runRefreshExistingForSelectedRow_')
    .addItem('③カレンダー同期（選択行）','runCalendarSyncForSelectedRow_')
    .addSeparator()
    .addItem('④スケジュール生成＋案内状/社内ページ反映','runScheduleApplyForSelectedRow_')
    .addSeparator()
    // 🆕 以下を追加
    .addItem('⑤請求書PDF化（選択行）','runExportInvoiceForSelectedRow_')
    .addToUi();
}


/* ================ メニュー操作関数群 ================ */

// 新規予約
function runNewBookingForSelectedRow_(){
  const sh = U.sh(CONFIG.SHEETS.MAIN);
  const ranges = sh.getActiveRangeList().getRanges();
  ranges.forEach(r=>{
    const row = r.getRow();
    if(row<=1) return;
    createOrUpdateClientFiles(row, { refreshOnly:false });
    updateFeaturesRow(row);
  });
  SpreadsheetApp.getActive().toast('✅ 新規予約テンプレ生成');
}

// ===============================
// ②既存データ更新（1行）
//  - 顧客フォルダのファイルは全削除して作り直し
//  - 社内Docセクションも作り直しつつ「その他」は保持
// ===============================
function refreshExistingForRow_(row) {
  const sh  = U.sh(CONFIG.SHEETS.MAIN);
  if (row <= 1) return;

  const info = readRowInfo(row);
  const set  = Settings.read();

  // === 顧客フォルダ取得 or 作成 ===
  const parent = DriveApp.getFolderById(set.parentFolderId);
  const folderName = `${info.groom} × ${info.bride}　様`;
  const folder = DriveX.getOrCreateChild(parent, folderName);

  // === 既存ファイルを全部削除 ===
  const files = folder.getFiles();
  while (files.hasNext()) {
    const f = files.next();
    console.log(`🗑️ 旧ファイル削除: ${f.getName()}`);
    f.setTrashed(true);
  }

  // === 最新テンプレートから再生成 ===
  set.templateIds.forEach(tid => {
    const src  = DriveApp.getFileById(tid);
    const base = detectBaseTitle(src.getName());
    const newName = `${base}_${info.groom}${info.bride}`;
    const f = src.makeCopy(newName, folder);
    applyPairsByMime(f, buildCommonPairs(info));
    console.log(`📄 再生成: ${f.getName()}`);
  });

  // === 社内ページも再生成（セクション削除＋作り直し／その他だけ保持） ===
  const docId = set.internalDocId;
  if (docId) {
    const doc  = DocumentApp.openById(docId);
    const body = doc.getBody();
    const titleText = `📸 ${info.groom} × ${info.bride}　様`;

    // 既存見出し検索
    let titlePara = body.getParagraphs().find(p => p.getText().trim() === titleText);
    let otherMemo = "";

    if (!titlePara) {
      // 初回：見出し新規
      body.appendPageBreak();
      titlePara = body.appendParagraph(titleText)
        .setHeading(DocumentApp.ParagraphHeading.HEADING1);
      body.appendParagraph('');
    } else {
      // 既存セクションから「その他」だけ救出
      otherMemo = getOtherMemoFromSection_(body, titlePara);
      // 既存セクションをまるごとクリア
      clearSectionAfterHeading_(body, titlePara);
    }

    const insertAt = body.getChildIndex(titlePara) + 1;

    // 顧客管理テーブル用のデータ
    const mainData  = sh.getDataRange().getValues();
    const headers   = mainData[0];
    const headerMap = {};
    headers.forEach((h, i) => headerMap[h] = i);
    const rowVals   = mainData[row - 1];

    // 「その他」列がある場合はDoc側の値を優先
    const idxOther = headers.indexOf('その他');
    if (idxOther !== -1 && otherMemo) {
      rowVals[idxOther] = otherMemo;
    }

    const tableData = [["項目", "内容"]];
    headers.forEach(h => tableData.push([h, rowVals[headerMap[h]] ?? ""]));

    // 「その他」列がシートに無い場合は、テーブル末尾に追加
    if (idxOther === -1) {
      tableData.push(["その他", otherMemo || ""]);
    }

    // 社内スケジュール行（テンプレ）
    tableData.push(["社内スケジュール", "{{社内スケジュール}}"]);

    // 顧客管理テーブル挿入
    insertTableAt_(body, insertAt, tableData);

    // 🗓 社内スケジュール見出し
    body.insertParagraph(insertAt + 1, "🗓 社内スケジュール")
      .setHeading(DocumentApp.ParagraphHeading.HEADING2);
    body.insertParagraph(insertAt + 2, "{{社内スケジュール}}");

    // 🎥 ムービーヒアリング情報
    const hSheet   = U.sh(CONFIG.SHEETS.HEARING);
    const hData    = hSheet.getDataRange().getValues();
    const hHeaders = hData[0];
    const hKey     = `${info.groom}_${info.bride}`;
    const hearingRow = hData.find(
      r => `${r[hHeaders.indexOf("新郎名")]}_${r[hHeaders.indexOf("新婦名")]}` === hKey
    );

    if (hearingRow) {
      body.insertParagraph(insertAt + 3, "🎥 ムービーヒアリング情報")
        .setHeading(DocumentApp.ParagraphHeading.HEADING2);
      const hTable = [["項目", "内容"]];
      hHeaders.forEach(h =>
        hTable.push([h, hearingRow[hHeaders.indexOf(h)] ?? ""])
      );
      insertTableAt_(body, insertAt + 4, hTable);
    }

    // ブックマーク再利用＋重複掃除
    let bm = doc.getBookmarks().find(b => {
      const el = b.getPosition().getElement();
      return el &&
        el.getType() === DocumentApp.ElementType.PARAGRAPH &&
        el.asParagraph().getText().trim() === titleText;
    });

    if (!bm) {
      bm = doc.addBookmark(doc.newPosition(titlePara, 0));
    }

    // 同じ見出しにぶら下がっている古いブックマークを削除
    doc.getBookmarks().forEach(b => {
      if (b.getId() === bm.getId()) return;
      const el = b.getPosition().getElement();
      if (!el || el.getType() !== DocumentApp.ElementType.PARAGRAPH) return;
      if (el.asParagraph().getText().trim() === titleText) {
        doc.removeBookmark(b);
        console.log(`🧹 重複ブックマーク削除: ${titleText}`);
      }
    });

    const colInternal = U.colOf(info.hs, CONFIG.COLS.INTERNAL_LINK);
    const linkUrl =
      `https://docs.google.com/document/d/${docId}/edit#bookmark=${bm.getId()}`;
    sh.getRange(row, colInternal)
      .setRichTextValue(U.rich('🗒 社内ページ', linkUrl));



    doc.saveAndClose();
  }

  console.log(`✅ 再生成完了: ${info.groom} × ${info.bride}`);
}

// ===============================
// ②既存データ再生成（選択行）メニュー
// ===============================
function runRefreshExistingForSelectedRow_() {
  const sh = U.sh(CONFIG.SHEETS.MAIN);
  const ranges = sh.getActiveRangeList().getRanges();

  ranges.forEach(r => {
    const row = r.getRow();
    if (row <= 1) return;
    try {
      refreshExistingForRow_(row);
    } catch (err) {
      console.log('再生成エラー:', err);
    }
  });

  SpreadsheetApp.getActive().toast(
    '🆕 既存データをテンプレートから再生成しました（顧客用＋社内用／その他は保持）'
  );
}

function calendarSyncForRow_(row) {
  const sh = U.sh(CONFIG.SHEETS.MAIN);
  const info = readRowInfo(row);
  if (!info) return;

  if (!info.photoDate || !(info.photoDate instanceof Date)) {
    console.log('calendarSync skip: 撮影日なし', row);
    return;
  }

  // 撮影地は O列「撮影地（確定）」があればそれを最優先
  info.location =
    (info.locFix && String(info.locFix).trim()) || // ※ locFix/loc は無くても既存 location がそのまま使われる
    (info.location && String(info.location).trim()) ||
    (info.loc && String(info.loc).trim()) ||
    '';

  console.log('📅 calendarSync start', {
    row,
    groom: info.groom,
    bride: info.bride,
    photoDate: info.photoDate,
    location: info.location
  });

  // ① このお客さんの撮影／締切イベントを全部削除（失敗したら即エラー）
  try {
    const summary = DL.clearAllEventsFor(info);
    console.log('🧹 clearAllEventsFor summary:', JSON.stringify(summary));
  } catch (err) {
    // メニュー実行や単発関数からわかりやすいように行番号＋新郎新婦を付けて投げる
    throw new Error(
      `calendarSyncForRow_: カレンダー削除に失敗しました。` +
      `行: ${row}, 新郎: ${info.groom}, 新婦: ${info.bride} / 理由: ${err.message}`
    );
  }

  // ② 撮影イベント（撮影カレンダー）作成＋説明更新
  DL.ensureShootEvent(info);
  DL.refreshShootEventDescription(info);

  // ③ 締切イベント：未決定のものだけ作成
  Object.entries(CONFIG.DEADLINE.ITEMS).forEach(([label, def]) => {
    if (def.type !== 'undecided') return;

    const val = String(U.getVal(sh, def.col, row) || '');
    if (val !== CONFIG.DEADLINE.VALUE_UNDECIDED) return;

    DL.createDeadlineIfNeeded(info, label, def);
  });

  // ④ 顧客フォルダ／社内ページのリンクを説明欄に追記
  DL.appendFolderUrlToEvents(info);

  console.log('📅 calendarSync 完了', info.groom, '×', info.bride);
}





// スケジュール反映
function runScheduleApplyForSelectedRow_(){
  const sh = U.sh(CONFIG.SHEETS.MAIN);
  const ranges = sh.getActiveRangeList().getRanges();
  ranges.forEach(r=>{
    const row = r.getRow();
    if(row<=1) return;
    const info = readRowInfo(row);
    DL.refreshShootEventDescription(info);
  });
  SpreadsheetApp.getActive().toast('📋 スケジュール・案内状反映を更新しました');
}

