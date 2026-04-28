/**
 * ウェブアプリの表示
 */
function doGet() {
  const title = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName('Config')
      ?.getRange('D2').getValue() || 'ワードウルフ オンライン';

  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle(title)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * 現在のゲーム状態を取得
 * ConfigシートのStatusが空ならWAITING、Limitが空なら300を自動セットします
 */
function getGameState() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const configSheet = ss.getSheetByName('Config');
    const playersSheet = ss.getSheetByName('Players');
    
    if (!configSheet || !playersSheet) {
      throw new Error("シート「Config」または「Players」が見つかりません。");
    }

    // Config設定の取得と自動補完
    const configRange = configSheet.getRange('A2:F2');
    const configValues = configRange.getValues()[0];
    let status = configValues[0];
    let startTime = configValues[1];
    let limit = configValues[2];
    const title = configValues[3] || 'ワードウルフ オンライン';
    const wolfName = configValues[4] || '人狼';
    const citizenName = configValues[5] || '村人';
    let needsUpdate = false;

    // Statusが空ならWAITINGをセット
    if (!status) {
      status = 'WAITING';
      needsUpdate = true;
    }
    // Limitが空なら300(5分)をセット
    if (limit === "" || limit === null || isNaN(limit)) {
      limit = 300;
      needsUpdate = true;
    }

    // もし空欄を埋めた場合は、A2:C2 のみ書き込んで同期する（D〜Fはそのまま）
    if (needsUpdate) {
      configSheet.getRange('A2:C2').setValues([[status, startTime, limit]]);
    }

    const email = Session.getActiveUser().getEmail();
    const playersData = playersSheet.getDataRange().getValues();
    
    // ユーザー特定
    const userData = playersData.find(row => row[0] === email);
    if (!userData) return { status: 'UNAUTHORIZED', email: email };

    let playerCount = 0;
    let checkedCount = 0;
    let allResults = [];
    let playerList = [];
    let voteTallies = {};

    // プレイヤーデータの集計（2行目以降）
    for (let i = 1; i < playersData.length; i++) {
      const row = playersData[i];
      const pNickname = row[1]; // B列
      const pRole = row[2];     // C列
      const pWord = row[3];     // D列
      const pChecked = row[4];  // E列
      const pVote = row[5];     // F列

      if (!row[0]) continue; // メールアドレスが空の行はスキップ

      playerCount++;
      if (pChecked === true) checkedCount++;
      if (pNickname) playerList.push(pNickname);
      if (pVote) voteTallies[pVote] = (voteTallies[pVote] || 0) + 1;

      // 結果発表フェーズなら全員の正解データを送る
      if (status === 'REVEALED') {
        allResults.push({
          name: pNickname || "名無し",
          role: pRole,
          word: pWord,
          votedFor: pVote
        });
      }
    }

    // 終了時刻の計算
    let endTime = null;
    if (startTime instanceof Date && !isNaN(limit)) {
      endTime = startTime.getTime() + (Number(limit) * 1000);
    }

    return {
      status: status,
      endTime: endTime,
      title: title,
      wolfName: wolfName,
      citizenName: citizenName,
      myWord: userData[3] || "",
      isChecked: userData[4] === true,
      isFacilitator: userData[6] === true, // G列
      myVote: userData[5] || null,        // F列
      countStr: `${checkedCount} / ${playerCount}`,
      allResults: allResults,
      playerList: playerList,
      voteTallies: voteTallies
    };
  } catch (e) {
    console.error("getGameState Error: " + e.message);
    return { status: 'ERROR', message: e.message };
  }
}

/**
 * お題を配布する
 */
function setupWords(wolfCount) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const configSheet = ss.getSheetByName('Config');
    
    const configValues = configSheet.getRange('A2:F2').getValues()[0];
    if (configValues[0] !== 'WAITING') {
      return getGameState();
    }
    const wolfName = configValues[4] || '人狼';
    const citizenName = configValues[5] || '村人';

    const wordsSheet = ss.getSheetByName('Words');
    if (!wordsSheet) throw new Error("Wordsシートが見つかりません。");
    const words = wordsSheet.getDataRange().getValues();
    if (words.length < 2) throw new Error("Wordsシートにお題が登録されていません。");

    const pair = words[Math.floor(Math.random() * (words.length - 1)) + 1];
    const isSwapped = Math.random() < 0.5;
    const citizenWord = isSwapped ? pair[1] : pair[0];
    const wolfWord = isSwapped ? pair[0] : pair[1];

    const playerSheet = ss.getSheetByName('Players');
    const lastRow = playerSheet.getLastRow();
    if (lastRow < 2) throw new Error("プレイヤーが1人も参加していません。");

    const players = playerSheet.getRange(2, 1, lastRow - 1, 1).getValues();
    
    let playerIndices = players.map((_, i) => i);
    let wolfIndices = [];
    const count = parseInt(wolfCount) || 1;
    for (let i = 0; i < count; i++) {
      if (playerIndices.length === 0) break;
      const r = Math.floor(Math.random() * playerIndices.length);
      wolfIndices.push(playerIndices.splice(r, 1)[0]);
    }

    const updates = players.map((row, i) => [
      (wolfIndices.includes(i)) ? wolfName : citizenName,
      (wolfIndices.includes(i)) ? wolfWord : citizenWord,
      false
    ]);
    playerSheet.getRange(2, 3, updates.length, 3).setValues(updates);
    configSheet.getRange('A2').setValue('DISTRIBUTING');
  } catch(e) {
    return { status: 'ERROR', message: e.message };
  } finally {
    lock.releaseLock();
  }

  return getGameState();
}

/**
 * 投票を実行
 */
function castVote(targetNickname) {
  const email = Session.getActiveUser().getEmail();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Players');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === email) {
      sheet.getRange(i + 1, 6).setValue(targetNickname);
      break;
    }
  }
  return getGameState();
}

function startGame() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.getSheetByName('Config').getRange('A2:B2').setValues([['STARTED', new Date()]]);
  return getGameState();
}

function startVoting() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.getSheetByName('Config').getRange('A2').setValue('VOTING');
  return getGameState();
}

function revealAnswer() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.getSheetByName('Config').getRange('A2').setValue('REVEALED');
  return getGameState();
}

function resetGame() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.getSheetByName('Config').getRange('A2:B2').setValues([['WAITING', null]]);
  const playerSheet = ss.getSheetByName('Players');
  const lastRow = playerSheet.getLastRow();
  if (lastRow > 1) {
    playerSheet.getRange(2, 3, lastRow - 1, 4).clearContent();
  }
  return getGameState();
}

function markAsChecked() {
  const email = Session.getActiveUser().getEmail();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Players');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === email) {
      sheet.getRange(i + 1, 5).setValue(true);
      break;
    }
  }
  return getGameState();
}
