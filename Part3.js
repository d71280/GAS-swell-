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

/* ================ ムービーヒアリング自動同期 ================ */

/**
 * 社内ページのムービーヒアリング情報のみを更新（軽量版）
 * 顧客フォルダのファイルは触らない
 */
function updateInternalPageOnly_(row, hearingData) {
  const info = readRowInfo(row, { includePrice: false });
  const set = Settings.read();

  if (!set.internalDocId) {
    console.warn('⚠️ 社内用ドキュメントIDが設定されていません');
    return;
  }

  const doc = DocumentApp.openById(set.internalDocId);
  const body = doc.getBody();
  const titleText = `📸 ${info.groom} × ${info.bride}　様`;

  // 見出しを検索
  const titlePara = body.getParagraphs().find(p => p.getText().trim() === titleText);
  if (!titlePara) {
    console.warn(`⚠️ 社内ページに見出しが見つかりません: ${titleText}`);
    return;
  }

  // ムービーヒアリング情報のセクションを削除して再作成
  const startIdx = body.getChildIndex(titlePara);
  let deleteEnd = body.getNumChildren();

  // 次の見出し（📸）を探す
  for (let i = startIdx + 1; i < body.getNumChildren(); i++) {
    const child = body.getChild(i);
    if (child.getType() === DocumentApp.ElementType.PARAGRAPH &&
        child.asParagraph().getText().trim().startsWith('📸 ')) {
      deleteEnd = i;
      break;
    }
  }

  // 🗓 社内スケジュールの位置を特定
  let scheduleHeadingIdx = -1;
  for (let i = startIdx + 1; i < deleteEnd; i++) {
    const child = body.getChild(i);
    if (child.getType() === DocumentApp.ElementType.PARAGRAPH &&
        child.asParagraph().getText().trim() === '🗓 社内スケジュール') {
      scheduleHeadingIdx = i;
      break;
    }
  }

  // 既存のムービーヒアリング見出しを削除
  for (let i = deleteEnd - 1; i > startIdx; i--) {
    const child = body.getChild(i);
    if (child.getType() === DocumentApp.ElementType.PARAGRAPH) {
      const text = child.asParagraph().getText().trim();
      if (text === '🎥 ムービーヒアリング情報') {
        // この見出しから次のセクションまでを削除
        let endDelete = i + 1;
        for (let j = i + 1; j < deleteEnd; j++) {
          const c = body.getChild(j);
          if (c.getType() === DocumentApp.ElementType.PARAGRAPH &&
              (c.asParagraph().getText().trim().startsWith('📋') ||
               c.asParagraph().getText().trim().startsWith('🗓'))) {
            endDelete = j;
            break;
          }
        }

        // 逆順で削除
        for (let k = endDelete - 1; k >= i; k--) {
          try {
            body.removeChild(body.getChild(k));
            // 削除後、社内スケジュールの位置を調整
            if (scheduleHeadingIdx > k) {
              scheduleHeadingIdx--;
            }
          } catch (e) {
            console.log('削除スキップ:', e);
          }
        }
        break;
      }
    }
  }

  // ムービーヒアリング情報があれば追加（社内スケジュールの直後に挿入）
  if (hearingData && hearingData.length > 0) {
    const hHeaders = hearingData[0];
    const hRow = hearingData[1]; // データは2行目（1行目はヘッダー）

    // 社内スケジュールの直後に挿入（見出し＋内容の後）
    let insertIdx = deleteEnd; // デフォルトはセクションの最後

    if (scheduleHeadingIdx !== -1) {
      // 社内スケジュールの見出しとその内容（{{社内スケジュール}}）の後
      insertIdx = scheduleHeadingIdx + 2; // 見出し + 内容段落の後
    }

    body.insertParagraph(insertIdx, '🎥 ムービーヒアリング情報')
      .setHeading(DocumentApp.ParagraphHeading.HEADING2);

    const hTable = [['項目', '内容']];
    hHeaders.forEach((h, idx) => hTable.push([h, hRow[idx] ?? '']));
    insertTableAt_(body, insertIdx + 1, hTable);
  }

  doc.saveAndClose();
  console.log(`📝 社内ページ更新完了（軽量）: ${info.groom} × ${info.bride}`);
}

/**
 * 回答者名から名前部分を抽出（日付・場所より前）
 * 例: "櫻井真優 11/8 城ヶ島" → "櫻井真優"
 */
function extractNameFromRespondent_(respondentName) {
  if (!respondentName) return '';

  const str = String(respondentName).trim();

  // スペース、数字、日付パターンより前の部分を抽出
  const match = str.match(/^([^\s\d]+)/);
  if (match) {
    return match[1].trim();
  }

  // マッチしない場合は最初のスペースまで
  const spaceIdx = str.indexOf(' ');
  if (spaceIdx > 0) {
    return str.substring(0, spaceIdx).trim();
  }

  return str;
}

/**
 * ムービーヒアリングシート全体をチェックして、社内ページを更新
 * 時間ベーストリガーで定期実行される（変更検出＋軽量更新）
 */
function syncAllMovieHearings() {
  console.log('🎥 ムービーヒアリング全件同期開始');

  const hearingSheet = U.sh(CONFIG.SHEETS.HEARING);
  const hearingData = hearingSheet.getDataRange().getValues();
  const hearingHeaders = hearingData[0];

  // 回答者名 列を探す
  const respondentIdx = hearingHeaders.indexOf('回答者名');

  if (respondentIdx === -1) {
    console.error('❌ ムービーヒアリングシートに「回答者名」列がありません');
    return;
  }

  const mainSheet = U.sh(CONFIG.SHEETS.MAIN);
  const mainData = mainSheet.getDataRange().getValues();
  const mainHeaders = mainData[0];

  const mainGroomIdx = mainHeaders.indexOf('新郎様お名前');
  const mainBrideIdx = mainHeaders.indexOf('新婦様お名前');

  if (mainGroomIdx === -1 || mainBrideIdx === -1) {
    console.error('❌ 顧客管理シートに「新郎様お名前」または「新婦様お名前」列がありません');
    return;
  }

  const normalize = (str) => String(str || '').replace(/[\s　]/g, '');
  const props = PropertiesService.getScriptProperties();
  let updateCount = 0;
  let skipCount = 0;

  // ムービーヒアリングの各行をチェック
  for (let i = 1; i < hearingData.length; i++) {
    const respondentName = String(hearingData[i][respondentIdx] || '').trim();

    if (!respondentName) continue;

    // 回答者名から名前部分を抽出（例: "櫻井真優 11/8 城ヶ島" → "櫻井真優"）
    const extractedName = extractNameFromRespondent_(respondentName);

    if (!extractedName) {
      console.warn(`⚠️ 名前抽出失敗: ${respondentName}`);
      continue;
    }

    const targetNameNorm = normalize(extractedName);

    // データのハッシュ値を計算（変更検出用）
    const rowDataStr = JSON.stringify(hearingData[i]);
    const currentHash = Utilities.computeDigest(
      Utilities.DigestAlgorithm.MD5,
      rowDataStr,
      Utilities.Charset.UTF_8
    ).map(b => (b < 0 ? b + 256 : b).toString(16).padStart(2, '0')).join('');

    const hashKey = `hearing_hash_${targetNameNorm}`;
    const lastHash = props.getProperty(hashKey);

    // 変更がない場合はスキップ
    if (lastHash === currentHash) {
      skipCount++;
      continue;
    }

    // 一致する顧客を検索（新郎名 OR 新婦名で照合）
    for (let j = 1; j < mainData.length; j++) {
      const mainGroom = normalize(mainData[j][mainGroomIdx]);
      const mainBride = normalize(mainData[j][mainBrideIdx]);

      // 新郎名または新婦名のどちらかに一致
      if (mainGroom === targetNameNorm || mainBride === targetNameNorm) {
        const matchedRow = j + 1;
        const groomDisplay = mainData[j][mainGroomIdx];
        const brideDisplay = mainData[j][mainBrideIdx];

        try {
          // 軽量更新：社内ページのみ
          updateInternalPageOnly_(matchedRow, [hearingHeaders, hearingData[i]]);

          // ハッシュ値を保存
          props.setProperty(hashKey, currentHash);

          updateCount++;
          console.log(`✅ 更新: 行${matchedRow} - ${groomDisplay} × ${brideDisplay} (照合: ${extractedName} 元: ${respondentName})`);
        } catch (err) {
          console.error(`❌ 更新エラー (行${matchedRow}):`, err);
        }
        break;
      }
    }
  }

  console.log(`🎥 同期完了: ${updateCount}件更新, ${skipCount}件スキップ（変更なし）`);
  return { updated: updateCount, skipped: skipCount };
}

/**
 * ムービーヒアリング自動同期の時間ベーストリガーをセットアップ
 */
function setupMovieHearingAutoSync() {
  const ss = SpreadsheetApp.getActive();

  // 既存の同期トリガーを削除
  const triggers = ScriptApp.getUserTriggers(ss);
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'syncAllMovieHearings') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // 新しいトリガーを作成（1分ごと）
  ScriptApp.newTrigger('syncAllMovieHearings')
    .timeBased()
    .everyMinutes(1)
    .create();

  SpreadsheetApp.getUi().alert(
    '✅ ムービーヒアリング自動同期をセットアップしました！\n\n' +
    '1分ごとに自動的にチェックして、社内ページを更新します。\n' +
    '変更検出により、変更があった行だけ更新されます。\n\n' +
    '※このセットアップは初回のみ実行すればOKです。'
  );

  console.log('✅ ムービーヒアリング自動同期トリガー（1分間隔）をセットアップしました');
}

/* ================ ムービーヒアリング編集時の処理 ================ */
/**
 * ムービーヒアリングシートが編集されたときに、対応する顧客の社内ページを自動更新
 */
function handleHearingEdit_(hearingSheet, editedRow) {
  console.log(`🎥 ムービーヒアリング編集検知: 行${editedRow}`);

  // 編集された行の新郎・新婦名を取得
  const hearingData = hearingSheet.getDataRange().getValues();
  const hearingHeaders = hearingData[0];

  const groomIdx = hearingHeaders.indexOf('新郎名');
  const brideIdx = hearingHeaders.indexOf('新婦名');

  if (groomIdx === -1 || brideIdx === -1) {
    console.error('❌ ムービーヒアリングシートに「新郎名」または「新婦名」列がありません');
    return;
  }

  const editedRowData = hearingData[editedRow - 1];
  const hearingGroom = String(editedRowData[groomIdx] || '').trim();
  const hearingBride = String(editedRowData[brideIdx] || '').trim();

  if (!hearingGroom || !hearingBride) {
    console.log('⏭️ 新郎・新婦名が空のためスキップ');
    return;
  }

  console.log(`👰 検索: ${hearingGroom} × ${hearingBride}`);

  // 顧客管理シートで一致する行を検索
  const mainSheet = U.sh(CONFIG.SHEETS.MAIN);
  const mainData = mainSheet.getDataRange().getValues();
  const mainHeaders = mainData[0];

  const mainGroomIdx = mainHeaders.indexOf('新郎様お名前');
  const mainBrideIdx = mainHeaders.indexOf('新婦様お名前');

  if (mainGroomIdx === -1 || mainBrideIdx === -1) {
    console.error('❌ 顧客管理シートに「新郎様お名前」または「新婦様お名前」列がありません');
    return;
  }

  // 名前の正規化（空白を削除して比較）
  const normalize = (str) => String(str || '').replace(/[\s　]/g, '');
  const targetGroomNorm = normalize(hearingGroom);
  const targetBrideNorm = normalize(hearingBride);

  // 一致する顧客を検索
  for (let i = 1; i < mainData.length; i++) {
    const mainGroom = normalize(mainData[i][mainGroomIdx]);
    const mainBride = normalize(mainData[i][mainBrideIdx]);

    if (mainGroom === targetGroomNorm && mainBride === targetBrideNorm) {
      const matchedRow = i + 1;
      console.log(`✅ 一致: 行${matchedRow} - ${hearingGroom} × ${hearingBride}`);

      try {
        // 社内ページを自動更新
        refreshExistingForRow_(matchedRow);
        console.log(`🔄 社内ページ更新完了: 行${matchedRow}`);
        SpreadsheetApp.getActive().toast(
          `🎥 ムービーヒアリング情報を社内ページに反映しました\n${hearingGroom} × ${hearingBride}`,
          '自動更新完了',
          5
        );
      } catch (err) {
        console.error(`❌ 更新エラー (行${matchedRow}):`, err);
        SpreadsheetApp.getActive().toast(
          `⚠️ 社内ページの更新に失敗しました: ${err.message}`,
          'エラー',
          5
        );
      }
      return;
    }
  }

  console.warn(`⚠️ 一致する顧客が見つかりません: ${hearingGroom} × ${hearingBride}`);
  SpreadsheetApp.getActive().toast(
    `⚠️ 顧客管理シートに一致する顧客が見つかりませんでした\n${hearingGroom} × ${hearingBride}`,
    'ムービーヒアリング',
    5
  );
}

/* ================ onEdit：列番号固定版 ================ */
/**
 * インストール可能トリガーで実行される onEdit ハンドラー
 * シンプルトリガーとの重複を避けるため、関数名を変更
 */
function onEditHandler(e){
  try{
    if (!e || !e.range) return;

    const sh = e.range.getSheet();
    if (!sh) return;

    const sheetName = sh.getName();
    const row = e.range.getRow();
    const col = e.range.getColumn();
    if (row <= 1) return; // 見出し行は無視

    // ========== ムービーヒアリングシート編集時の処理 ==========
    if (sheetName === CONFIG.SHEETS.HEARING) {
      handleHearingEdit_(sh, row);
      return;
    }

    // ========== 顧客管理シート編集時の処理 ==========
    if (sheetName !== CONFIG.SHEETS.MAIN) return;

    // 列番号（現行のシート構成前提）
    const COL = {
      CAMERA:    11, // K列 カメラマン
      PLAN_AUTO: 12, // L列 プラン（自動）
      PLAN_MAN:  13, // M列 プラン（手動）
      LOC_FIX:   15, // O列 撮影地（確定）
      LINK:       1, // A列 顧客用ページ
      INTERNAL:   2  // B列 社内用ページ
    };

    // === 連続実行防止：同じセルを1秒以内に編集した場合はスキップ ===
    const cellKey = `${sh.getName()}_${row}_${col}`;
    const props = PropertiesService.getScriptProperties();
    const lastEditKey = `lastEdit_${cellKey}`;
    const lastEditTime = props.getProperty(lastEditKey);
    const now = new Date().getTime();

    if (lastEditTime && (now - Number(lastEditTime)) < 1000) {
      console.warn(`⚠️ 連続実行防止: ${cellKey} は1秒以内に編集されたためスキップ`);
      return;
    }

    props.setProperty(lastEditKey, String(now));

    // --- 顧客名チェック：新郎・新婦名が空の場合は処理をスキップ ---
    const groomCell = sh.getRange(row, 5).getDisplayValue().trim(); // E列：新郎
    const brideCell = sh.getRange(row, 4).getDisplayValue().trim(); // D列：新婦

    if (!groomCell || !brideCell) {
      console.log(`⏭️ 顧客名が空のためスキップ (行${row})`);
      return;
    }

    // --- L / M列の変更時は P〜V列を自動反映 ---
    if (col === COL.PLAN_AUTO || col === COL.PLAN_MAN) {
      updateFeaturesRow(row);
    }

// --- M列（プラン手動）変更時：既存データ更新 & カレンダー同期のみ ---
if (col === COL.PLAN_MAN) {
  // ロック取得（最大30秒待機）
  const lock = LockService.getScriptLock();
  try {
    // ロック取得を試みる（30秒待機）
    const hasLock = lock.tryLock(30000);
    if (!hasLock) {
      console.warn('⚠️ 既に処理中のため、この編集はスキップされました');
      return;
    }

    const info = readRowInfo(row);
    const hasA = !!sh.getRange(row, COL.LINK).getDisplayValue();
    const hasB = !!sh.getRange(row, COL.INTERNAL).getDisplayValue();

    // A列またはB列にリンクがある場合のみ、既存データを更新
    if (hasA || hasB) {
      console.log(`📝 M列変更: 既存データ更新 (行${row})`);
      refreshExistingForRow_(row);
      calendarSyncForRow_(row);
    } else {
      // リンクがない場合は、カレンダー同期のみ実行
      // 新規作成はメニュー「①新規予約の一括処理」から実行してください
      console.warn(`⚠️ M列変更: リンク未設定のためカレンダー同期のみ実行 (行${row})`);
      console.warn('💡 新規作成はメニュー「①新規予約の一括処理」から実行してください');
      calendarSyncForRow_(row);
    }

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
      const lock = LockService.getScriptLock();
      try {
        const hasLock = lock.tryLock(30000);
        if (!hasLock) {
          console.warn('⚠️ 既に処理中のため、この編集はスキップされました');
          return;
        }

        const colName = col === COL.CAMERA ? 'K列（カメラマン）' : 'O列（撮影地確定）';
        console.log(`📝 ${colName}変更: 既存データ更新 + カレンダー同期 (行${row})`);

        refreshExistingForRow_(row);   // 社内用ページなど更新
        calendarSyncForRow_(row);      // カレンダーも更新

      } catch (err) {
        console.error(`K/O列処理エラー (行${row}):`, err);
        SpreadsheetApp.getActive().toast('⚠️ 処理中にエラーが発生しました');
      } finally {
        lock.releaseLock();
      }
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



/* ================ インストール可能トリガーのセットアップ ================ */
/**
 * onEdit の自動更新を有効にするためのセットアップ関数
 * 初回のみ実行してください（メニューから実行）
 */
function setupAutoUpdateTrigger() {
  const ss = SpreadsheetApp.getActive();

  // 既存の onEdit / onEditHandler トリガーを全て削除
  const triggers = ScriptApp.getUserTriggers(ss);
  triggers.forEach(trigger => {
    const funcName = trigger.getHandlerFunction();
    if (funcName === 'onEdit' || funcName === 'onEditHandler') {
      ScriptApp.deleteTrigger(trigger);
      console.log(`🗑️ 古いトリガー削除: ${funcName}`);
    }
  });

  // 新しいインストール可能トリガーを作成
  ScriptApp.newTrigger('onEditHandler')
    .forSpreadsheet(ss)
    .onEdit()
    .create();

  SpreadsheetApp.getUi().alert(
    '✅ 自動更新トリガーをセットアップしました！\n\n' +
    '古い onEdit トリガーを削除し、新しい onEditHandler トリガーを作成しました。\n' +
    'これで M列・K列・O列の編集時に自動更新が動作します。'
  );

  console.log('✅ インストール可能トリガーをセットアップしました');
}

/* ================ メニュー ================ */
function onOpen(){
  SpreadsheetApp.getUi().createMenu('📂 顧客管理メニュー')
    .addItem('⚙️ 自動更新トリガーをセットアップ','setupAutoUpdateTrigger')
    .addItem('🎥 ムービーヒアリング自動同期を有効化','setupMovieHearingAutoSync')
    .addSeparator()
    .addItem('①新規予約の一括処理（選択行）','runNewBookingForSelectedRow_')
    .addItem('②既存データ更新（選択行）','runRefreshExistingForSelectedRow_')
    .addItem('③カレンダー同期（選択行）','runCalendarSyncForSelectedRow_')
    .addSeparator()
    .addItem('④スケジュール生成＋案内状/社内ページ反映','runScheduleApplyForSelectedRow_')
    .addSeparator()
    .addItem('⑤請求書PDF化（選択行）','runExportInvoiceForSelectedRow_')
    .addSeparator()
    .addItem('🔍 カレンダー削除テスト（選択行）','testCalendarDelete')
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
  if (!info) {
    console.warn(`⚠️ calendarSync skip: 顧客情報の取得に失敗 (行${row})`);
    return;
  }

  // 新郎・新婦名のチェック
  if (!info.groom || !info.bride) {
    console.warn(`⚠️ calendarSync skip: 新郎・新婦名が空 (行${row})`);
    return;
  }

  if (!info.photoDate || !(info.photoDate instanceof Date)) {
    console.warn(`⚠️ calendarSync skip: 撮影日なし (行${row}, 新郎: ${info.groom}, 新婦: ${info.bride})`);
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
    console.log(`🧹 clearAllEventsFor summary: ${JSON.stringify(summary)}`);

    // 削除件数をログ出力
    const totalDeleted = summary.reduce((sum, s) => sum + s.deleted, 0);
    console.log(`✅ カレンダーイベント削除完了: ${totalDeleted}件 (行${row}, ${info.groom} × ${info.bride})`);
  } catch (err) {
    // メニュー実行や単発関数からわかりやすいように行番号＋新郎新婦を付けて投げる
    const errMsg = `カレンダー削除に失敗しました。行: ${row}, 新郎: ${info.groom}, 新婦: ${info.bride} / 理由: ${err.message}`;
    console.error(`❌ ${errMsg}`);
    throw new Error(errMsg);
  }

  // ② 撮影イベント（撮影カレンダー）作成＋説明更新
  DL.ensureShootEvent(info);
  DL.refreshShootEventDescription(info);

  // ③ 締切イベント作成
  Object.entries(CONFIG.DEADLINE.ITEMS).forEach(([label, def]) => {
    // type: 'undecided' の場合
    if (def.type === 'undecided') {
      const val = String(U.getVal(sh, def.col, row) || '');
      console.log(`📋 ${label}: 列${def.col} = "${val}"`);
      if (val === CONFIG.DEADLINE.VALUE_UNDECIDED) {
        DL.createDeadlineIfNeeded(info, label, def);
        console.log(`  ✅ 締切イベント作成: ${label}`);
      } else {
        console.log(`  ⏭️ スキップ（値が"未決定"ではない）`);
      }
    }
    // type: 'checkbox' の場合（写真納品・動画納品）
    else if (def.type === 'checkbox' && def.chkCol) {
      // 列記号を使って直接値を取得
      const chkVal = U.getVal(sh, def.chkCol, row);
      console.log(`📋 ${label}: 列${def.chkCol} = ${chkVal}`);
      // チェックされていない場合のみ締切イベントを作成
      if (chkVal !== true && String(chkVal).toLowerCase() !== 'true') {
        DL.createDeadlineIfNeeded(info, label, def);
        console.log(`  ✅ 締切イベント作成: ${label}`);
      } else {
        console.log(`  ⏭️ スキップ（チェック済み）`);
      }
    }
  });

  // ④ 顧客フォルダ／社内ページのリンクを説明欄に追記
  DL.appendFolderUrlToEvents(info);

  console.log('📅 calendarSync 完了', info.groom, '×', info.bride);
}





// カレンダー同期（選択行）
function runCalendarSyncForSelectedRow_(){
  const sh = U.sh(CONFIG.SHEETS.MAIN);
  const ranges = sh.getActiveRangeList().getRanges();
  ranges.forEach(r=>{
    const row = r.getRow();
    if(row<=1) return;
    try {
      calendarSyncForRow_(row);
      console.log(`✅ カレンダー同期完了: 行${row}`);
    } catch (err) {
      console.error(`❌ カレンダー同期エラー (行${row}):`, err);
      SpreadsheetApp.getActive().toast(`⚠️ 行${row}のカレンダー同期に失敗しました: ${err.message}`);
    }
  });
  SpreadsheetApp.getActive().toast('📅 カレンダー同期を実行しました');
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

// ===== テスト用：カレンダー削除のデバッグ関数 =====
function testCalendarDelete() {
  const sh = U.sh(CONFIG.SHEETS.MAIN);
  const row = sh.getActiveRange().getRow();

  Logger.log('=== カレンダー削除テスト開始 ===');

  if (row <= 1) {
    Logger.log('❌ データ行を選択してください');
    console.error('データ行を選択してください');
    return;
  }

  const info = readRowInfo(row);

  Logger.log(`行: ${row}`);
  Logger.log(`新郎: ${info.groom}`);
  Logger.log(`新婦: ${info.bride}`);
  Logger.log(`撮影日: ${info.photoDate}`);
  Logger.log(`撮影地: ${info.location}`);

  if (!info.groom || !info.bride) {
    Logger.log('❌ 新郎・新婦名が空です');
    console.error('新郎・新婦名が空です');
    return;
  }

  if (!info.photoDate) {
    Logger.log('❌ 撮影日が空です');
    console.error('撮影日が空です');
    return;
  }

  try {
    const summary = DL.clearAllEventsFor(info);
    Logger.log('✅ 削除結果: ' + JSON.stringify(summary));

    const totalDeleted = summary.reduce((sum, s) => sum + s.deleted, 0);
    const msg = `✅ 削除完了: ${totalDeleted}件のイベントを削除しました`;
    Logger.log(msg);
    console.log(msg);

    // UIが利用可能な場合のみアラート表示
    try {
      SpreadsheetApp.getUi().alert(msg + '\n詳細はログを確認してください');
    } catch (e) {
      // UIが利用できない場合は無視
    }
  } catch (err) {
    const errMsg = '❌ エラー: ' + err.message;
    Logger.log(errMsg);
    console.error(errMsg);

    // UIが利用可能な場合のみアラート表示
    try {
      SpreadsheetApp.getUi().alert(errMsg);
    } catch (e) {
      // UIが利用できない場合は無視
    }
  }
}

