/******************************************************
* 無知ノ知 撮影管理 - 完全安定統合版（Part2）
* 修正版 - エラー解消済み
******************************************************/

/** Doc本文の index 位置に 2次元配列 tableData をテーブルとして安全に挿入する */
function insertTableAt_(body, index, tableData){
  // 一旦末尾にテーブルを作ってからコピー→挿入→元を削除（最も安定）
  const tmp = body.appendTable(
    tableData.map(row => row.map(v => v == null ? '' : String(v)))
  );
  const copy = tmp.copy();              // Detached な Table
  body.removeChild(tmp);                // 一度消して
  body.insertTable(index, copy);        // 目的位置へ挿入
}

/** セクション（見出し行の次～次の見出し直前）を安全にクリアする */
function clearSectionAfterHeading_(body, headingParagraph){
  const start = body.getChildIndex(headingParagraph);
  let end = body.getNumChildren();

  // 次の「📸 」見出しを探してそこまでを削除対象に
  for(let i = start + 1; i < body.getNumChildren(); i++){
    const child = body.getChild(i);
    if (child.getType() === DocumentApp.ElementType.PARAGRAPH &&
        child.asParagraph().getText().trim().startsWith("📸 ")) {
      end = i;
      break;
    }
  }

  // 最終段落削除エラー回避のため番兵の空行を末尾に追加
  body.appendParagraph("");

  // 逆順で削除（ドキュメントの最終段落は削除しない）
  const lastDeletableIndex = body.getNumChildren() - 2; // 末尾1つは保持
  for (let i = Math.min(end - 1, lastDeletableIndex); i > start; i--) {
    try {
      body.removeChild(body.getChild(i));
    } catch (e) {
      console.log("削除スキップ:", e);
    }
  }
}

/** セクション内の「その他」セルの内容を取得（なければ空文字） */
function getOtherMemoFromSection_(body, headingParagraph){
  const start = body.getChildIndex(headingParagraph);
  const ET = DocumentApp.ElementType;
  let other = "";

  for (let i = start + 1; i < body.getNumChildren(); i++) {
    const child = body.getChild(i);

    // 次の「📸 」見出しが来たら終了
    if (child.getType() === ET.PARAGRAPH &&
        child.asParagraph().getText().trim().startsWith("📸 ")) {
      break;
    }

    if (child.getType() === ET.TABLE) {
      const t = child.asTable();
      for (let r = 0; r < t.getNumRows(); r++) {
        const row = t.getRow(r);
        if (row.getNumCells() < 2) continue;
        const label = row.getCell(0).getText().trim();
        if (label === "その他") {
          other = row.getCell(1).getText();
        }
      }
    }
  }
  return other;
}



/* ================ Drive / Docs Utility ================ */
const DriveX = {
  getOrCreateChild(parent,name){
    const it=parent.getFoldersByName(name);
    return it.hasNext()?it.next():parent.createFolder(name);
  },
  // ファイル名を正規化（連続する空白を1つに統一）
  normalizeName(name){
    return String(name).replace(/\s+/g, ' ').trim();
  },
  copyIfMissing(folder,templateId,newName){
    // ファイル名を正規化
    const normalizedNewName = DriveX.normalizeName(newName);

    // 既存ファイルを全て取得して、正規化した名前で比較
    const files = folder.getFiles();
    while(files.hasNext()){
      const existingFile = files.next();
      const existingName = DriveX.normalizeName(existingFile.getName());
      if(existingName === normalizedNewName){
        console.log(`既存ファイルを使用: ${existingFile.getName()}`);
        return existingFile;
      }
    }

    // 既存ファイルがなければ新規作成
    const src=DriveApp.getFileById(templateId);
    console.log(`新規ファイル作成: ${normalizedNewName}`);
    return src.makeCopy(normalizedNewName,folder);
  }
};

const Docs = {
  replaceInDoc(id,pairs){
    try{
      const d=DocumentApp.openById(id);
      const b=d.getBody();
      Object.entries(pairs).forEach(([k,v])=>b.replaceText(k,String(v??'')));
      d.saveAndClose();
    }catch(e){
      console.log('DocErr',e)
    }
  },
  replaceInSlides(id,pairs){
    try{
      const s=SlidesApp.openById(id);
      Object.entries(pairs).forEach(([k,v])=>s.replaceAllText(k,String(v??'')));
    }catch(e){
      console.log('SlideErr',e)
    }
  },
  replaceInSheets(id,pairs){
    try{
      const ss=SpreadsheetApp.openById(id);
      ss.getSheets().forEach(sh=>{
        Object.entries(pairs).forEach(([k,v])=>sh.createTextFinder(k).replaceAllWith(String(v??'')));
      });
    }catch(e){
      console.log('SheetErr',e)
    }
  }
};

/* ================ 請求書対応関数群 ================ */
// ===== 請求書：行別に値を計算（①＝自動、②〜⑤＝手動最大4件） =====
function buildInvoiceRows_(planAutoName, manualItems){
  const rows = [{},{},{},{},{}]; // 最大5行
  const used = new Set();

  // 正規化関数（比較用に括弧と空白を削除）
  function normalize(s){
    return String(s || '')
      .replace(/[（(][^）)]*[）)]/g, '')  // 括弧とその中身を削除
      .replace(/[　\s]/g, '')              // 空白削除
      .trim();
  }

  // ===== ① 自動プラン =====
  if (planAutoName){
    const normalizedForDupe = normalize(planAutoName);
    
    // ★修正：Price.priceOf()を使用
    const unit1 = Price.priceOf(planAutoName);
    
    rows[0] = {
      desc: planAutoName || '',
      qty: planAutoName ? 1 : '',
      unit: unit1 || '',
      amount: unit1 || ''
    };
    
    if (normalizedForDupe) used.add(normalizedForDupe);
  }

  // ===== ②〜⑤ 手動プラン =====
  let pos = 1;
  for (const raw of manualItems){
    const label = String(raw || '').trim();
    if (!label) continue;
    
    const normalizedForDupe = normalize(label);
    if (!normalizedForDupe || used.has(normalizedForDupe)) continue;
    used.add(normalizedForDupe);

    // ★修正：Price.priceOf()を使用
    const unit = Price.priceOf(label);
    
    rows[pos] = {
      desc: label,
      qty: 1,
      unit: unit || '',
      amount: unit || ''
    };

    pos++;
  if (pos > 5) break;
}

  // ===== 合計計算 =====
  const subtotal = rows.reduce((s,r)=> s + (num_(r.amount)||0), 0);
  const tax = Math.round(subtotal * 0.10);
  const total = subtotal + tax;

  return { rows, subtotal, tax, total };
}

// === プラン詳細生成関数（《...》形式） ===
function getPlanDetail_(planName) {
  const details = {
    '挙式準備完璧プラン': '《オープニングムービー(〜90秒)・フォト(レタッチ込み200枚〜)\n新郎新婦アテンド付きヘアメイク(当日アテンド・新郎新婦ヘアメイク・フィッティング・新婦ヘアチェンジ・\nスマホオフショット)・テンプレプロフィールムービー・提携衣装店ドレス&タキシード・レンタルブーケ&ベール・\nウェルカムボード制作》',
    '衣装プラン': '《オープニングムービー(〜90秒)・フォト(レタッチ込み100枚〜)\n新郎新婦アテンド付きヘアメイク(当日アテンド・新婦ヘアメイク・フィッティング・新婦ヘアチェンジ\nスマホオフショット)・提携衣装店ドレス&タキシード》',
    'ヘアメイクプラン': '《オープニングムービー(〜90秒)・フォト(レタッチ込み100枚〜)\n新郎新婦アテンド付きヘアメイク(当日アテンド・新婦ヘアメイク・フィッティング・新婦ヘアチェンジ\nスマホオフショット)》',
    'ムービープラン': '《オープニングムービー(〜90秒)・新郎新婦アテンド付きヘアメイク(当日アテンド\n新婦ヘアメイク・フィッティング・新婦ヘアチェンジ・スマホオフショット)》',
    'フォトプラン': '《フォト(レタッチ込み100枚〜)・新郎新婦アテンド付きヘアメイク(当日アテンド\n新婦ヘアメイク・フィッティング・新婦ヘアチェンジ・スマホオフショット)》'
  };

  if (!planName) return '';
  const clean = removeParenJP_(planName);
  const match = Object.keys(details).find(k => clean.includes(k));
  return match ? details[match] : '';
}

function buildCommonPairs(info){
  const issue = new Date();
  const due = new Date(issue.getTime());
  due.setMonth(due.getMonth() + 1);

  // ---- プラン手動：純粋な「なし」だけは最初から除外 ----
  const manualItemsRaw = String(info.planMan || '')
    .split(/[,、\s]+/)
    .map(s => s.trim())
    .filter(Boolean);

  const manualItems = manualItemsRaw.filter(s => {
    const t = s.replace(/[　\s]/g, '');
    return !/^(なし|無し|ナシ)$/.test(t);
  });

  const inv = buildInvoiceRows_(info.planAuto, manualItems);

  // プラン手動が「なし」だけの場合は空文字列にする
  const planManDisplay = manualItems.length > 0 ? info.planMan : '';

  const pairs = {
    '{{新郎名}}': info.groom,
    '{{新婦名}}': info.bride,
    '{{新郎}}': info.groom,
    '{{新婦}}': info.bride,
    '{{撮影日}}': info.photoDisp,
    '{{撮影地}}': info.location,
    '{{カメラマン}}': info.camera,
    '{{プラン自動}}': info.planAuto,
    '{{プラン（自動）}}': info.planAuto,
    '{{プラン手動}}': planManDisplay,
    '{{プラン（手動）}}': planManDisplay,
    '{{今日}}': Utilities.formatDate(issue, CONFIG.TZ, 'yyyy年MM月dd日'),
    '{{発行日}}': U.fmt(issue),
    '{{お支払い期限}}': U.fmt(due),
    '{{宛名}}': `${info.groom}　様 / ${info.bride}　様`,
    '{{件名}}': `${info.groom}　様 × ${info.bride}　様 ウェディング前撮り`,
    '{{小計}}': inv.subtotal.toLocaleString(),
    '{{消費税}}': inv.tax.toLocaleString(),
    '{{合計}}': inv.total.toLocaleString(),
    '{{合計金額}}': `¥${inv.total.toLocaleString()}`,
    '{{数量}}': 1
  };

  // --- 自動プラン（①） ---
  const auto = inv.rows[0] || {};
  pairs['{{プラン自動}}'] = auto.desc || '';
  pairs['{{数量①}}'] = auto.desc ? 1 : '';
  pairs['{{金額①}}'] = auto.unit ? `¥${Number(auto.unit).toLocaleString()}` : '';
  pairs['{{合計①}}'] = auto.amount ? `¥${Number(auto.amount).toLocaleString()}` : '';

  // --- 手動プラン（②〜⑤） ---
  const manualNums = ['①','②','③','④'];
  const moneyNums  = ['②','③','④','⑤'];
  const manualList = [];

  for (let i = 0; i < 4; i++) {
    const r  = inv.rows[i + 1] || {}; // rows[1]〜rows[4]
    const n1 = manualNums[i];
    const n2 = moneyNums[i];

    let desc = String(r.desc || '').trim();
    const t = desc.replace(/[　\s]/g, '');
    const isNone = /^(なし|無し|ナシ)$/.test(t);

    // 純粋な「なし」は Doc 上では完全に空扱い
    if (isNone) desc = '';

    pairs[`{{プラン手動${n1}}}`] = desc;
    pairs[`{{数量${n2}}}`]       = desc ? 1 : '';
    pairs[`{{金額${n2}}}`]       = !desc || !r.unit
                                   ? ''
                                   : `¥${Number(r.unit).toLocaleString()}`;
    pairs[`{{合計${n2}}}`]       = !desc || !r.amount
                                   ? ''
                                   : `¥${Number(r.amount).toLocaleString()}`;

    if (!desc) continue; // ここから下は案内状用まとめ

    manualList.push(desc);
  }

  // --- プラン詳細（案内状用） ---
  pairs['{{プラン詳細}}'] = getPlanDetail_(info.planAuto);

  // --- 案内状用まとめ（◯付き） ---
  pairs['{{プラン手動まとめ}}'] = manualList.length
    ? `⚪︎${manualList.join('、')}`
    : '';

  return pairs;
}



function applyPairsByMime(file, pairs) {
  const mt = file.getMimeType();
  const name = file.getName();

  // 🆕 案内状テンプレ（ファイル名に「案内」または「案内状」を含む場合）
  // 改行（\n）をスペースに変換して1行化
  if (name.includes('案内') || name.includes('案内状')) {
    Object.keys(pairs).forEach(k => {
      if (typeof pairs[k] === 'string') {
        pairs[k] = pairs[k].replace(/\n+/g, ' '); // 改行→半角スペース
      }
    });
  }

  if (mt === MimeType.GOOGLE_DOCS) Docs.replaceInDoc(file.getId(), pairs);
  else if (mt === MimeType.GOOGLE_SLIDES) Docs.replaceInSlides(file.getId(), pairs);
  else if (mt === MimeType.GOOGLE_SHEETS) Docs.replaceInSheets(file.getId(), pairs);
}


function createOrUpdateClientFiles(row, opts = { refreshOnly:false }) {
  const set = Settings.read();              // 設定シート（親フォルダID、テンプレ群、internalDocId取得）
  const info = readRowInfo(row);            // 顧客情報読み込み
  const sh = U.sh(CONFIG.SHEETS.MAIN);
  const groom = info.groom, bride = info.bride;
  if(!groom || !bride) throw '氏名が空です';
  
  /* ===============================
   * ① 顧客フォルダ＋顧客Docs生成
   * =============================== */
  const parent = DriveApp.getFolderById(set.parentFolderId);
  const folder = DriveX.getOrCreateChild(parent, `${groom} × ${bride}　様`);
  const folderUrl = folder.getUrl();
  const colLink = U.colOf(info.hs, CONFIG.COLS.LINK);
  if(!sh.getRange(row, colLink).getDisplayValue()) {
    sh.getRange(row, colLink).setRichTextValue(U.rich('📂 フォルダ', folderUrl));
  }

  if(!opts.refreshOnly){
    set.templateIds.forEach(tid=>{
      const src=DriveApp.getFileById(tid);
      const base=detectBaseTitle(src.getName());
      const newName=`${base}_${groom} × ${bride}　様`;
      const f=DriveX.copyIfMissing(folder,tid,newName);
      applyPairsByMime(f,buildCommonPairs(info));
    });
  }

/* ===============================
 * ② 社内用ページ（単一Doc）更新
 * =============================== */
if (set.internalDocId){
  const doc  = DocumentApp.openById(set.internalDocId);
  const body = doc.getBody();

  // 顧客管理＆ヘッダ
  const mainData = sh.getDataRange().getValues();
  const headers  = mainData[0];
  const headerMap = {};
  headers.forEach((h,i)=> headerMap[h] = i);

  // ムービーヒアリング
  const hSheet   = U.sh(CONFIG.SHEETS.HEARING);
  const hData    = hSheet.getDataRange().getValues();
  const hHeaders = hData[0];
  const hearingMap = {};
  hData.slice(1).forEach(r=>{
    const key = `${r[hHeaders.indexOf("新郎名")]}_${r[hHeaders.indexOf("新婦名")]}`;
    const o = {};
    hHeaders.forEach((h,i)=> o[h] = r[i]);
    hearingMap[key] = o;
  });

  const titleText = `📸 ${groom} × ${bride}　様`;

  // 既存見出し検索
  let titlePara = body.getParagraphs().find(p => p.getText().trim() === titleText);

  // なければ見出し新規
  if (!titlePara){
    body.appendPageBreak();
    titlePara = body.appendParagraph(titleText)
      .setHeading(DocumentApp.ParagraphHeading.HEADING1);
    body.appendParagraph(""); // 余白
  } else {
    // 既存セクションを安全にクリア
    clearSectionAfterHeading_(body, titlePara);
  }

  /**
 * 顧客ブロック内のテーブルから「その他」行の内容だけ拾う
 */
function getOtherMemoForSection_(body, titleText) {
  const ET = DocumentApp.ElementType;
  let start = -1;
  let end = body.getNumChildren();

  // そのお客さんのセクション範囲を特定
  for (let i = 0; i < body.getNumChildren(); i++) {
    const el = body.getChild(i);
    if (el.getType() === ET.PARAGRAPH) {
      const txt = el.asParagraph().getText().trim();
      if (txt === titleText) {
        start = i;
      } else if (start >= 0 && txt.startsWith("📸 ")) {
        end = i;
        break;
      }
    }
  }
  if (start < 0) return '';

  // セクション内のテーブルを舐めて「その他」行を探す
  let memo = '';
  for (let i = start + 1; i < end; i++) {
    const el = body.getChild(i);
    if (el.getType() !== ET.TABLE) continue;

    const table = el.asTable();
    for (let r = 0; r < table.getNumRows(); r++) {
      const row = table.getRow(r);
      if (row.getNumCells() < 2) continue;

      const label = row.getCell(0).getText().trim();
      if (label === 'その他') {
        memo = row.getCell(1).getText(); // 右側のセルそのまま
      }
    }
  }
  return memo;
}


  // この時点の挿入位置（見出し直後の位置）を固定
  const insertAt = body.getChildIndex(titlePara) + 1;

  // 見出し直後に「顧客管理情報」見出し
  body.insertParagraph(insertAt, "📋 顧客管理情報")
      .setHeading(DocumentApp.ParagraphHeading.HEADING2);

  // --- 顧客管理テーブル ---
  const rowVals   = mainData[row-1];
  const tableData = [["項目","内容"]];
  headers.forEach(h => tableData.push([h, rowVals[headerMap[h]] ?? ""]));

  // 「その他」列がシートに無ければ、空行として追加（Doc側で自由に書く用）
  if (headers.indexOf('その他') === -1) {
    tableData.push(["その他", ""]);
  }

  // 社内スケジュールもテーブルに残す（プレースホルダ）
  tableData.push(["社内スケジュール", "{{社内スケジュール}}"]);
  insertTableAt_(body, insertAt + 1, tableData);

  // --- 🗓 社内スケジュールテンプレ（独立ブロック） ---
  body.insertParagraph(insertAt + 2, "🗓 社内スケジュール")
    .setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.insertParagraph(insertAt + 3, "{{社内スケジュール}}");



  // ムービーヒアリング（存在時のみ）
  const hKey = `${groom}_${bride}`;
  if (hearingMap[hKey]){
    body.insertParagraph(insertAt + 2, "🎥 ムービーヒアリング情報")
        .setHeading(DocumentApp.ParagraphHeading.HEADING2);

    const hRow   = hearingMap[hKey];
    const hTable = [["項目","内容"]];
    hHeaders.forEach(h => hTable.push([h, hRow[h] ?? ""]));
    insertTableAt_(body, insertAt + 3, hTable);
  }

  // ブックマーク（見出しに付与 or 既存を利用＋重複掃除）
  let bm = doc.getBookmarks().find(b => {
    const el = b.getPosition().getElement();
    return el &&
      el.getType() === DocumentApp.ElementType.PARAGRAPH &&
      el.asParagraph().getText().trim() === titleText;
  });

  if (!bm) {
    bm = doc.addBookmark(doc.newPosition(titlePara, 0));
  }

  // 🔄 同じ見出しにぶら下がる古いブックマークを削除
  doc.getBookmarks().forEach(b => {
    if (b.getId() === bm.getId()) return;
    const el = b.getPosition().getElement();
    if (!el || el.getType() !== DocumentApp.ElementType.PARAGRAPH) return;
    if (el.asParagraph().getText().trim() === titleText) {
      doc.removeBookmark(b);
      console.log(`🧹 重複ブックマーク削除: ${titleText}`);
    }
  });

  // シートB列（社内用ページ）にブックマークリンク
  const colInternal = U.colOf(info.hs, CONFIG.COLS.INTERNAL_LINK);
  const linkUrl = `https://docs.google.com/document/d/${doc.getId()}/edit#bookmark=${bm.getId()}`;
  sh.getRange(row, colInternal).setRichTextValue(U.rich('🗒 社内ページ', linkUrl));

  doc.saveAndClose();
}


  /* ===============================
   * ③ refreshOnly時はDocs再差込
   * =============================== */
  if(opts.refreshOnly){
    const pairs = buildCommonPairs(info);
    const it = folder.getFiles();
    while(it.hasNext()){
      const f = it.next();
      applyPairsByMime(f, pairs);
    }
  }
}

/******************************************************
 * 📄 請求書PDF化＆スプレッドシート削除＋一覧追記
 ******************************************************/

/**
 * 顧客フォルダ内の「請求書」スプレッドシートをPDF化して削除し、
 * 「請求書一覧」シートに記録を追加する
 * @param {string} folderUrl 顧客フォルダのURL
 */
function exportInvoiceToPdfAndDelete_(folderUrl) {
  if (!folderUrl) throw new Error("顧客フォルダURLが未設定です。");

  const folderId = folderUrl.match(/[-\w]{25,}/)[0];
  const folder = DriveApp.getFolderById(folderId);

  // 請求書スプレッドシートを検索
  const files = folder.getFiles();
  let target = null;
  while (files.hasNext()) {
    const f = files.next();
    if (f.getName().includes("請求書") && f.getMimeType() === MimeType.GOOGLE_SHEETS) {
      target = f;
      break;
    }
  }
  if (!target) {
    SpreadsheetApp.getUi().alert("請求書スプレッドシートが見つかりません。");
    return;
  }

  const ssId = target.getId();
  const pdfName = `${target.getName()}.pdf`;

  // PDF出力URLを生成
  const url = `https://docs.google.com/spreadsheets/d/${ssId}/export?exportFormat=pdf&format=pdf&size=A4&portrait=true&fitw=true&sheetnames=false&printtitle=false&gridlines=false`;

  const token = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(url, { headers: { Authorization: `Bearer ${token}` } });

  // フォルダにPDF保存
  const blob = response.getBlob().setName(pdfName);
  const pdfFile = folder.createFile(blob);

  // === 請求書一覧シートに追記 ===
  try {
    const ss = U.ss();
    const shList = ss.getSheetByName('請求書一覧');
    if (shList) {
      // 対応する顧客情報を取得
      const shMain = U.sh(CONFIG.SHEETS.MAIN);
      const activeRow = shMain.getActiveRange().getRow();
      const info = readRowInfo(activeRow);
      const manualItems = String(info.planMan || '').split(/[,、\s]+/).filter(Boolean);
      const inv = buildInvoiceRows_(info.planAuto, manualItems);

      const issueDate = new Date();
      const due = new Date(issueDate);
      due.setMonth(due.getMonth() + 1);

      shList.appendRow([
        info.groom,                      // 新郎様お名前
        info.bride,                      // 新婦様お名前
        info.photoDisp,                  // 撮影日
        inv.total,                       // 合計金額
        U.fmt(issueDate),                // 発行日
        U.fmt(due),                      // 振込期日
        ''                               // 入金済み（空欄）
      ]);
      console.log(`🧾 請求書一覧に追記: ${info.groom} × ${info.bride}`);
    } else {
      console.warn('⚠️ シート「請求書一覧」が見つかりません。');
    }
  } catch (err) {
    console.error('請求書一覧追記エラー:', err);
  }

  SpreadsheetApp.getActive().toast(`📄 PDF出力＆一覧追記完了：${pdfFile.getName()}`);
  console.log(`✅ PDF保存: ${pdfFile.getUrl()}`);
}


/**
 * 選択行の顧客フォルダから請求書PDF出力
 */
function runExportInvoiceForSelectedRow_() {
  const sh = U.sh(CONFIG.SHEETS.MAIN);
  const row = sh.getActiveRange().getRow();
  if (row <= 1) return;

  const info = readRowInfo(row);
  if (!info.folderUrl) {
    SpreadsheetApp.getUi().alert("顧客フォルダURLが未設定です。");
    return;
  }

  exportInvoiceToPdfAndDelete_(info.folderUrl);
}

/******************************************************
* 📅 DL: Deadline Manager（締切・リマインド・Chat通知・カレンダー制御）
******************************************************/
const DL = {
  // ===== タイトル生成（締切表記なし／統一形式） =====
  // 例）「ヘアメイク - 山田太郎 × 山田花子」
  buildTitle(info, label, def) {
    return `${label} - ${info.groom || ''} × ${info.bride || ''}`;
  },

  // ===== イベント検出（完全一致・90日前後範囲） =====
  findEvent(info, label) {
    const cal = CalendarApp.getCalendarById(CONFIG.DEADLINE.CALENDAR_ID_DEADLINE);
    if (!cal || !info.photoDate) return null;

    const from = new Date(info.photoDate.getTime() - 90 * 86400000);
    const to   = new Date(info.photoDate.getTime() + 90 * 86400000);
    const all  = cal.getEvents(from, to);

    const title = this.buildTitle(info, label, {});
    return all.find(e => e.getTitle() === title) || null;
  },

  // ===== 締切イベント作成 =====
  createDeadlineIfNeeded(info, label, def) {
    if (!info.photoDate) return;
    const cal = CalendarApp.getCalendarById(CONFIG.DEADLINE.CALENDAR_ID_DEADLINE);
    if (!cal) return;

    const title = this.buildTitle(info, label, def);
    const exist = this.findEvent(info, label);
    if (exist) return; // 既存あればスキップ

    const date = new Date(info.photoDate.getTime() + (def.offsetDays || 0) * 86400000);
    cal.createAllDayEvent(title, date, {
      description: `${label} 締切\n${info.groom} × ${info.bride}`
    });
    console.log(`📅 追加: ${title} (${U.fmt(date)})`);
  },

  // ===== 締切イベント削除 =====
  deleteDeadlineIfExists(info, label) {
    const cal = CalendarApp.getCalendarById(CONFIG.DEADLINE.CALENDAR_ID_DEADLINE);
    const ev = this.findEvent(info, label);
    if (ev) {
      ev.deleteEvent();
      console.log(`🗑 削除: ${ev.getTitle()}`);
    }
  },

  // ===== 撮影日イベント（撮影地を前に追加） =====
  ensureShootEvent(info) {
    if (!info.photoDate) return;
    const cal = CalendarApp.getCalendarById(CONFIG.DEADLINE.CALENDAR_ID_SHOOT);
    if (!cal) return;

    // info.location は「O列があればO、なければN」を使う想定
    const locPart = info.location ? `${info.location} - ` : '';
    const title = `${locPart}${info.groom || ''} × ${info.bride || ''}`;

    const events = cal.getEventsForDay(info.photoDate);
    const ex = events.find(e => e.getTitle() === title);
    if (!ex) {
      cal.createAllDayEvent(title, info.photoDate, {
        description: `撮影地: ${info.location}\nカメラマン: ${info.camera}\nプラン: ${info.planAuto}`
      });
      console.log(`📸 撮影日イベント作成: ${title}`);
    }
  },

  // ===== 撮影イベント更新 =====
  refreshShootEventDescription(info) {
    if (!info.photoDate) return;
    const cal = CalendarApp.getCalendarById(CONFIG.DEADLINE.CALENDAR_ID_SHOOT);
    if (!cal) return;

    const locPart = info.location ? `${info.location} - ` : '';
    const title = `${locPart}${info.groom || ''} × ${info.bride || ''}`;
    const evs = cal.getEventsForDay(info.photoDate);
    const ev = evs.find(e => e.getTitle() === title);
    if (ev) {
      ev.setDescription(
        `撮影地: ${info.location}\nカメラマン: ${info.camera}\nプラン: ${info.planAuto}`
      );
      console.log(`📝 撮影イベント説明更新: ${title}`);
    }
  },

  // ===== Chat通知送信 =====
  notifyToChat(text) {
    const url = CONFIG.DEADLINE.CHAT_WEBHOOK;
    if (!url) return;
    const payload = { text };
    UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload)
    });
  },

  // ===== 締切リマインダー処理 =====
  processRemindersRow(info) {
    if (!info.photoDate) return;
    const now = U.todayYmd();
    const cal = CalendarApp.getCalendarById(CONFIG.DEADLINE.CALENDAR_ID_DEADLINE);
    const sh = U.sh(CONFIG.SHEETS.MAIN);
    const hs = U.getHeaders(sh);
    const row = info.row;

    Object.entries(CONFIG.DEADLINE.ITEMS).forEach(([label, def]) => {
      if (def.type !== 'undecided') return;

      const title = this.buildTitle(info, label, def);
      const events = cal.getEvents(
        new Date(info.photoDate.getTime() - 90 * 86400000),
        new Date(info.photoDate.getTime() + 90 * 86400000)
      );
      const exist = events.some(e => e.getTitle() === title);
      if (!exist) return;

      const val = String(U.getVal(sh, def.col, row) || '');
      if (val !== CONFIG.DEADLINE.VALUE_UNDECIDED) return;

      const doneIdx = hs.indexOf('最終完了') + 1;
      if (doneIdx > 0) {
        const chkVal = sh.getRange(row, doneIdx).getValue();
        if (chkVal === true || String(chkVal).trim().toUpperCase() === 'TRUE') return;
      }

      const date = new Date(info.photoDate.getTime() + def.offsetDays * 86400000);
      const diff = U.daysBetween(now, date);

      if (def.offsetDays < 0 &&
          diff <= 0 &&
          diff > -CONFIG.DEADLINE.REMIND.OVERDUE_MAX_DAYS) {
        this.notifyToChat(
          `⚠️【期限超過】${label} (${Math.abs(diff)}日経過)\n${info.groom} × ${info.bride}`
        );
      } else if (CONFIG.DEADLINE.REMIND.buildOffsets(def.offsetDays).includes(diff)) {
        this.notifyToChat(
          `⏰【リマインド】${label}まで残り${diff}日\n${info.groom} × ${info.bride}`
        );
      }
    });
  },
/**
 * 今日ではなく「撮影日」を基準に ±365日のイベントを削除する
 * 削除に失敗した場合は Error を投げて理由を返す
 */
clearAllEventsFor(info) {
  const groom = String(info.groom || '').trim();
  const bride = String(info.bride || '').trim();

  if (!groom || !bride) {
    throw new Error('clearAllEventsFor: 新郎/新婦名が空です');
  }
  if (!info.photoDate || !(info.photoDate instanceof Date)) {
    throw new Error('clearAllEventsFor: 撮影日が不正（未設定 or Date 型でない）');
  }

  const key = `${groom} × ${bride}`;

  // 撮影日を中心に ±365日（1年）を削除対象にする
  const base = info.photoDate;
  const from = new Date(base.getTime() - 365 * 86400000);
  const to   = new Date(base.getTime() + 365 * 86400000);

  const calIds = [
    CONFIG.DEADLINE.CALENDAR_ID_SHOOT,
    CONFIG.DEADLINE.CALENDAR_ID_DEADLINE
  ];

  // 名前を正規化（空白・記号を削除）
  const normalizeForMatch = (str) => {
    return String(str || '')
      .replace(/[　\s]/g, '')  // 全角・半角スペース削除
      .replace(/[様さん]/g, '')  // 敬称削除
      .toLowerCase();
  };

  const groomNorm = normalizeForMatch(groom);
  const brideNorm = normalizeForMatch(bride);

  const summary = [];

  calIds.forEach(id => {
    try {
      if (!id) {
        throw new Error('カレンダーIDが設定されていません');
      }

      const cal = CalendarApp.getCalendarById(id);
      if (!cal) {
        throw new Error(`カレンダーが見つかりません: ${id}`);
      }

      const events = cal.getEvents(from, to);
      let count = 0;

      events.forEach(ev => {
        const title = ev.getTitle() || '';
        const titleNorm = normalizeForMatch(title);

        // タイトルに新郎・新婦両方の名前が含まれているものだけ削除（正規化して比較）
        if (titleNorm.includes(groomNorm) && titleNorm.includes(brideNorm)) {
          try {
            ev.deleteEvent();
            count++;
            console.log(`🗑 削除: [${id}] ${title}`);
          } catch (e) {
            // 個々のイベント削除失敗は即エラーにする
            throw new Error(`イベント削除失敗: ${title} / ${e.message}`);
          }
        }
      });

      summary.push({ calendarId: id, deleted: count });

    } catch (e) {
      // どのカレンダーで失敗したのか分かるように包んで投げる
      throw new Error(`clearAllEventsFor: カレンダー [${id}] 処理中にエラー: ${e.message}`);
    }
  });

  console.log(`✅ clearAllEventsFor: ${key} の削除結果: ${JSON.stringify(summary)}`);
  return summary;
},





  // ===== 顧客フォルダURLをカレンダーイベントに追加 =====
  appendFolderUrlToEvents(info) {
    const folderUrl   = info.folderUrl   || '';
    const internalUrl = info.internalUrl || '';
    if (!folderUrl && !internalUrl) return;

    const htmlLines = [];
    if (folderUrl)   htmlLines.push(`<a href="${folderUrl}">📂 顧客フォルダを開く</a>`);
    if (internalUrl) htmlLines.push(`<a href="${internalUrl}">🗒 社内ページを開く</a>`);
    const htmlBlock = htmlLines.join('<br>');

    // --- 撮影カレンダー ---
    const shootCal = CalendarApp.getCalendarById(CONFIG.DEADLINE.CALENDAR_ID_SHOOT);
    if (shootCal && info.photoDate) {
      const locPart = info.location ? `${info.location} - ` : '';
      const shootTitle = `${locPart}${info.groom || ''} × ${info.bride || ''}`;
      shootCal.getEventsForDay(info.photoDate).forEach(e => {
        if (e.getTitle() === shootTitle) {
          const desc = e.getDescription() || '';
          if (!desc.includes(folderUrl) && !desc.includes(internalUrl)) {
            e.setDescription(desc + '\n\n' + htmlBlock);
            console.log(`📎 撮影イベントにリンク追加: ${shootTitle}`);
          }
        }
      });
    }

    // --- 締切カレンダー ---
    const deadlineCal = CalendarApp.getCalendarById(CONFIG.DEADLINE.CALENDAR_ID_DEADLINE);
    if (!deadlineCal || !info.photoDate) return;

    Object.entries(CONFIG.DEADLINE.ITEMS).forEach(([label, def]) => {
      const titlePart = this.buildTitle(info, label, def);
      const events = deadlineCal.getEvents(
        new Date(info.photoDate.getTime() - 90 * 86400000),
        new Date(info.photoDate.getTime() + 90 * 86400000)
      );
      events.forEach(e => {
        if (e.getTitle() === titlePart) {
          const desc = e.getDescription() || '';
          if (!desc.includes(folderUrl) && !desc.includes(internalUrl)) {
            e.setDescription(desc + '\n\n' + htmlBlock);
            console.log(`📎 締切イベントにリンク追加: ${titlePart}`);
          }
        }
      });
    });
  }
};  // ← DL 終わり

