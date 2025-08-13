// === 設定定数 ===
const SHEET_NAME_UI = "UI";
const SHEET_NAME_RARITY = "レアリティと条件";
const SHEET_NAME_DATA = "data";
const SHEET_NAME_LOG = "log";

const RARITY_MAP = {
  0: "☆",
  1: "★",
  2: "★★",
  3: "★★★",
  4: "★★★★",
  5: "★★★★★"
};

const RARITY_INDEXES = [0, 1, 2, 3, 4, 5];

function logMessage(logSheet, timestamp, msg) {
  logSheet.appendRow([`[${timestamp}] ${msg}`]);
}

function runGacha(name, mode, rarity, count) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const recordSheet = ss.getSheetByName("record");
  const raritySheet = ss.getSheetByName(SHEET_NAME_RARITY);
  const dataSheet = ss.getSheetByName(SHEET_NAME_DATA);
  const logSheet = ss.getSheetByName(SHEET_NAME_LOG);

  const now = new Date();
  const timestamp = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");

  if (mode === "complete") {
    logMessage(logSheet, timestamp, `${name} さんが挑戦した COMPLETE ガチャの結果発表！`);
  } else if (mode === "choice") {
    logMessage(logSheet, timestamp, `${name} さんが挑戦した CHOICE ガチャ ${RARITY_MAP[rarity] || rarity} ${count}回の結果発表！`);
  } else {
    logMessage(logSheet, timestamp, `${name} さんが挑戦した ${mode} ガチャ ${count}回の結果発表！`);
  }

  const dataValues = dataSheet.getDataRange().getValues();
  const itemList = dataValues
    .slice(1)
    .filter(r => r.length >= 3 && r[1] !== "" && r[2] !== "")
    .map(r => ({ name: r[1].toString().trim(), rarity: parseInt(r[2], 10) }));

  let results = [];

  if (mode === "complete") {
    results = itemList;
  } else if (mode === "choice") {
    const parsedRarity = parseInt(rarity, 10);
    const filtered = itemList.filter(r => r.rarity === parsedRarity);
    for (let i = 0; i < count; i++) {
      const pick = filtered[Math.floor(Math.random() * filtered.length)];
      results.push(pick);
    }
  } else {
    const headerRow = raritySheet.getRange(1, 1, 1, raritySheet.getLastColumn()).getValues()[0];
    const modeColIndex = headerRow.indexOf(mode) + 1;
    if (modeColIndex <= 0) {
      logMessage(logSheet, timestamp, `エラー: ポイント種別「${mode}」が見つかりません`);
      return;
    }
    const rawRates = raritySheet.getRange(2, modeColIndex, 6, 1).getValues().flat();
    const rarityRates = rawRates.map(v => typeof v === 'number' ? v : parseFloat(v) || 0);
    const total = rarityRates.reduce((sum, v) => sum + v, 0);
    if (total === 0) {
      logMessage(logSheet, timestamp, `エラー: ポイント種別「${mode}」の排出確率がすべて0です`);
      return;
    }
    const weights = rarityRates.map(v => v / total);

    for (let i = 0; i < count; i++) {
      const rarityIndex = pickByWeight(RARITY_INDEXES, weights);
      const pool = itemList.filter(r => r.rarity === rarityIndex);

      if (pool.length === 0) {
        results.push({ rarity: rarityIndex, name: null });
        continue;
      }

      const pick = pool[Math.floor(Math.random() * pool.length)];
      results.push(pick);
    }
  }

  results.forEach((r, idx) => {
    if (r && r.name) {
      const itemNames = itemList.map(i => i.name);
      const targetRow = itemNames.indexOf(r.name) + 2; // header offset
      if (targetRow > 1) {
        const lastCol = recordSheet.getLastColumn();
        const headerRange = recordSheet.getRange(1, 2, 1, Math.max(1, lastCol - 1));
        const header = headerRange.getValues()[0];
        let userCol = header.indexOf(name) + 2;
        if (userCol < 2) {
          userCol = lastCol + 1;
          recordSheet.getRange(1, userCol).setValue(name);
        }
        const cell = recordSheet.getRange(targetRow, userCol);
        if (!cell.getValue()) {
          cell.setValue("◯");
        }
      }
    }

    if (!r || r.rarity === undefined || r.name === null) {
      logMessage(logSheet, timestamp, `[${String(idx + 1).padStart(3, '0')}] エラー: ${RARITY_MAP[r?.rarity] || r?.rarity} に該当するアイテムがありません`);
    } else {
      logMessage(logSheet, timestamp, `[${String(idx + 1).padStart(3, '0')}] ${RARITY_MAP[r.rarity]}: ${r.name}`);
    }
  });
}

function pickByWeight(items, weights) {
  const rand = Math.random();
  let sum = 0;
  for (let i = 0; i < weights.length; i++) {
    sum += weights[i];
    if (rand <= sum) return items[i];
  }
  return items[0]; // fallback to first if something goes wrong
}

function runGachaFromUI() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const uiSheet = ss.getSheetByName(SHEET_NAME_UI);
  if (!uiSheet) {
    SpreadsheetApp.getUi().alert("UI シートが見つかりません！");
    return;
  }

  const name = uiSheet.getRange("B2").getValue().toString().trim();
  const mode = uiSheet.getRange("B3").getValue().toString().trim();
  const rarity = uiSheet.getRange("B4").getValue().toString().trim();
  const countRaw = uiSheet.getRange("B5").getValue();
  const count = parseInt(countRaw, 10);

  if (!name) {
    SpreadsheetApp.getUi().alert("名前が入力されていません。");
    return;
  }
  if (!mode) {
    SpreadsheetApp.getUi().alert("モードが入力されていません。");
    return;
  }
  if (mode === "choice") {
    if (rarity === "") {
      SpreadsheetApp.getUi().alert("choiceモードではrarityの入力が必要です。0〜5で入力してください。");
      return;
    }
    const rarityNum = parseInt(rarity, 10);
    if (isNaN(rarityNum) || rarityNum < 0 || rarityNum > 5) {
      SpreadsheetApp.getUi().alert("choiceモードではrarityは0〜5の数値で指定してください。");
      return;
    }
  } else {
    if (rarity !== "") {
      SpreadsheetApp.getUi().alert("レアリティはchoiceモードでのみ指定可能です。modeをchoiceにするか、レアリティ欄を空にしてください。");
      return;
    }
  }
  if (mode !== "complete" && (isNaN(count) || count < 1)) {
    SpreadsheetApp.getUi().alert("回数が正しく入力されていません。");
    return;
  }

  const actualRarity = mode === "choice" ? parseInt(rarity, 10) : "";
  const actualCount = mode === "complete" ? 0 : count;

  runGacha(name, mode, actualRarity, actualCount);
  SpreadsheetApp.getUi().alert("ガチャ実行が完了しました！");
}
