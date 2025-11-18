/******************************************************
 * 無知ノ知 撮影管理 - 工程表生成（Part4／Slides対応・安全版）
 * 2025-10-24
 ******************************************************/

/**
 * 📋 スケジュール生成＆Docs/Slides反映
 * - 顧客用Slides: {{当日スケジュール}}
 * - 社内用Docs: {{社内スケジュール}}
 */
function runScheduleApplyForSelectedRow_() {
  const sh = U.sh(CONFIG.SHEETS.MAIN);
  const ranges = sh.getActiveRangeList().getRanges();

  ranges.forEach(r => {
    const row = r.getRow();
    if (row <= 1) return;

    const info = readRowInfo(row);
    if (!info.photoDate || !info.location) {
      console.log(`⚠️ 撮影日または撮影地が未設定: row ${row}`);
      return;
    }

    // ===== 撮影地リストから緯度経度を取得 =====
    const latLng = getLatLngFromSheet(info.location);
    if (!latLng) {
      console.log(`⚠️ 撮影地「${info.location}」の緯度経度が未登録`);
      return;
    }

    // ===== 日没APIからスケジュール生成 =====
    const sunset = fetchSunsetTime(latLng, info.photoDate);
    const clientText = generateScheduleTextForClient(sunset, info.location);
    const internalText = generateScheduleTextForInternal(sunset, info.location);

    // ===== 顧客フォルダ内の案内状を検索 =====
    if (info.folderUrl) {
      const folderId = info.folderUrl.match(/[-\w]{25,}/)[0];
      const folder = DriveApp.getFolderById(folderId);
      const files = folder.getFiles();

      while (files.hasNext()) {
        const f = files.next();
        const name = f.getName();
        const normalizedName = name.replace(/[　\s]/g, "");
        const groomKey = info.groom.replace(/[　\s様さん]/g, "");
        const brideKey = info.bride.replace(/[　\s様さん]/g, "");

        if (normalizedName.includes("案内状") && normalizedName.includes(groomKey) && normalizedName.includes(brideKey)) {
          const mime = f.getMimeType();
          if (mime === MimeType.GOOGLE_SLIDES) {
            const slide = SlidesApp.openById(f.getId());
            slide.replaceAllText("{{当日スケジュール}}", clientText);
            console.log(`🎞️ 案内状（Slides）更新: ${name}`);
          } else if (mime === MimeType.GOOGLE_DOCS) {
            const doc = DocumentApp.openById(f.getId());
            doc.getBody().replaceText("{{当日スケジュール}}", clientText);
            doc.saveAndClose();
            console.log(`📘 案内状（Docs）更新: ${name}`);
          }
        }
      }
    }

    // ===== 社内ページにも反映 =====
    const set = Settings.read();
    if (set.internalDocId) {
      const doc = DocumentApp.openById(set.internalDocId);
      const body = doc.getBody();
      const title = `📸 ${info.groom} × ${info.bride}　様`; // ← 全角スペース統一

      // === セクション探索（顧客ブロック単位） ===
      let startIdx = -1, endIdx = body.getNumChildren();
      for (let i = 0; i < body.getNumChildren(); i++) {
        const el = body.getChild(i);
        if (el.getType() === DocumentApp.ElementType.PARAGRAPH) {
          const text = el.asParagraph().getText().trim();
          if (text === title) startIdx = i;
          else if (startIdx >= 0 && text.startsWith("📸 ")) {
            endIdx = i;
            break;
          }
        }
      }

      // === プレースホルダー置換 ===
      let found = false;
      if (startIdx >= 0) {
        for (let i = startIdx; i < endIdx; i++) {
          if (_replaceInElement(body.getChild(i), "{{社内スケジュール}}", internalText)) {
            found = true;
          }
        }
      }

      // === 全文にも保険で置換 ===
      if (!found) {
        body.replaceText("{{社内スケジュール}}", internalText);
      }

      // === 見つからなければ末尾追記 ===
      if (!found && !body.getText().includes(internalText)) {
        body.appendParagraph("📋 社内スケジュール").setHeading(DocumentApp.ParagraphHeading.HEADING2);
        body.appendParagraph(internalText);
      }

      doc.saveAndClose();
      console.log(`🗒 社内ページ更新: ${title}`);
    }

  });

  SpreadsheetApp.getActive().toast("📋 当日スケジュール（Slides案内状・社内用Docs）を反映しました");
}

/* === 再帰置換 === */
function _replaceInElement(el, placeholder, value) {
  let hit = false;
  const ET = DocumentApp.ElementType;
  switch (el.getType()) {
    case ET.PARAGRAPH:
    case ET.LIST_ITEM:
      if (el.asText().getText().includes(placeholder)) {
        el.asText().replaceText(placeholder, value);
        hit = true;
      }
      break;
    case ET.TABLE:
      const t = el.asTable();
      for (let r = 0; r < t.getNumRows(); r++) {
        const row = t.getRow(r);
        for (let c = 0; c < row.getNumCells(); c++) {
          if (_replaceInElement(row.getCell(c), placeholder, value)) hit = true;
        }
      }
      break;
    case ET.TABLE_ROW:
      const row = el.asTableRow();
      for (let c = 0; c < row.getNumCells(); c++) {
        if (_replaceInElement(row.getCell(c), placeholder, value)) hit = true;
      }
      break;
    case ET.TABLE_CELL:
      const cell = el.asTableCell();
      for (let i = 0; i < cell.getNumChildren(); i++) {
        if (_replaceInElement(cell.getChild(i), placeholder, value)) hit = true;
      }
      break;
    default:
      if (el.getNumChildren) {
        for (let i = 0; i < el.getNumChildren(); i++) {
          if (_replaceInElement(el.getChild(i), placeholder, value)) hit = true;
        }
      }
  }
  return hit;
}

/* === 日没・座標・テキスト生成 === */
function getLatLngFromSheet(location) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.LOCS);
  const vals = sheet.getDataRange().getValues();
  for (let i = 1; i < vals.length; i++) {
    const name = String(vals[i][0]).trim();
    if (name === location) return { lat: Number(vals[i][2]), lng: Number(vals[i][3]) };
  }
  return null;
}

function fetchSunsetTime(latLng, date) {
  // 自動リトライ機能付き（最大3回試行）
  const maxRetries = 3;
  let lastError = null;

  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      console.log(`🌅 日没API呼び出し (試行 ${attempt}/${maxRetries})`);

      // tzidパラメータを削除（APIでサポートされていない）
      const api = `https://api.sunrise-sunset.org/json?lat=${latLng.lat}&lng=${latLng.lng}&date=${Utilities.formatDate(date, CONFIG.TZ, "yyyy-MM-dd")}&formatted=0`;

      const res = UrlFetchApp.fetch(api, {
        muteHttpExceptions: true,
        validateHttpsCertificates: true
      });

      const statusCode = res.getResponseCode();
      if (statusCode !== 200) {
        throw new Error(`HTTPステータス ${statusCode}`);
      }

      const json = JSON.parse(res.getContentText());

      if (json.status !== 'OK') {
        throw new Error(`APIステータス: ${json.status}`);
      }

      if (!json.results || !json.results.sunset) {
        throw new Error("日没データなし");
      }

      // UTC時刻をJSTに変換
      let sunset = new Date(json.results.sunset);
      if (sunset.getHours() < 9) {
        sunset = new Date(sunset.getTime() + 9 * 3600000);
      }

      console.log(`✅ 日没取得成功: ${Utilities.formatDate(sunset, CONFIG.TZ, 'HH:mm')}`);
      return sunset;

    } catch (err) {
      lastError = err;
      console.warn(`⚠️ 試行 ${attempt} 失敗: ${err.message}`);

      // 最終試行以外は1秒待機してリトライ
      if (attempt < maxRetries) {
        Utilities.sleep(1000);
      }
    }
  }

  // 全ての試行が失敗
  throw new Error(`日没時刻の取得に失敗しました（${maxRetries}回試行）: ${lastError.message}`);
}

/**
 * 📘 顧客用スケジュール（ヘアメイク2時間20分）
 */
function generateScheduleTextForClient(sunset, location) {
  const shootEnd = roundDown30(new Date(sunset));
  const shootStart = new Date(shootEnd.getTime() - 3.5 * 3600000);
  const moveStart = new Date(shootStart.getTime() - 3600000);
  const hairStart = new Date(moveStart.getTime() - 140 * 60000);
  const t = d => Utilities.formatDate(d, "Asia/Tokyo", "HH:mm");
  return [
    `${t(hairStart)}　サロン集合`,
    `${t(hairStart)}〜${t(moveStart)}　ヘアメイク`,
    `${t(moveStart)}〜${t(shootStart)}　移動・準備`,
    `${t(shootStart)}〜${t(shootEnd)}　撮影（ロケ地：${location}）`,
    `${t(shootEnd)}　撮影終了`
  ].join("\n");
}

/**
 * 🗒 社内用スケジュール（ヘアメイク2時間30分）
 */
function generateScheduleTextForInternal(sunset, location) {
  const shootEnd = roundDown30(new Date(sunset));
  const shootStart = new Date(shootEnd.getTime() - 3.5 * 3600000);
  const moveStart = new Date(shootStart.getTime() - 3600000);
  const hairStart = new Date(moveStart.getTime() - 150 * 60000);
  const t = d => Utilities.formatDate(d, "Asia/Tokyo", "HH:mm");
  return [
    `${t(hairStart)}　サロン集合`,
    `${t(hairStart)}〜${t(moveStart)}　ヘアメイク`,
    `${t(moveStart)}〜${t(shootStart)}　移動・準備`,
    `${t(shootStart)}〜${t(shootEnd)}　撮影（ロケ地：${location}）`,
    `${t(shootEnd)}　撮影終了`
  ].join("\n");
}

function roundDown30(date) {
  const d = new Date(date);
  const minutes = d.getMinutes();

  if (minutes <= 14) {
    d.setMinutes(0, 0, 0);
  } else if (minutes >= 15 && minutes < 45) {
    d.setMinutes(30, 0, 0);
  } else {
    d.setHours(d.getHours() + 1);
    d.setMinutes(0, 0, 0);
  }

  return d;
}


function replacePlaceholder(fileId, placeholder, value) {
  const file = DriveApp.getFileById(fileId);
  const mt = file.getMimeType();
  if (mt === MimeType.GOOGLE_DOCS) {
    const doc = DocumentApp.openById(fileId);
    doc.getBody().replaceText(placeholder, value);
    doc.saveAndClose();
  } else if (mt === MimeType.GOOGLE_SLIDES) {
    const slide = SlidesApp.openById(fileId);
    slide.replaceAllText(placeholder, value);
  }
}

