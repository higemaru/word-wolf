const STATE_CACHE_KEY = 'gameStateCache';

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

// ============================================================
// バリデーション・ヘルパー
// ============================================================

/**
 * 現在のキャッシュを取得する（なければ再構築）
 */
function getCache() {
  const raw = PropertiesService.getScriptProperties().getProperty(STATE_CACHE_KEY);
  return raw ? JSON.parse(raw) : rebuildCache();
}

/**
 * 呼び出し元が進行役かどうかチェックする
 * 進行役でなければ例外をスローする
 */
function assertFacilitator(email, cache) {
  const player = cache.players.find(p => p.email === email);
  if (!player || !player.isFacilitator) {
    throw new Error("進行役のみ実行できます。");
  }
}

/**
 * ゲームが期待するステータスかチェックする
 * 一致しなければ例外をスローする
 */
function assertState(cache, expectedStatus) {
  if (cache.status !== expectedStatus) {
    throw new Error(`この操作は ${expectedStatus} フェーズでのみ実行できます。（現在: ${cache.status}）`);
  }
}

/**
 * 進行役専用のステータス遷移をまとめたヘルパー (#4)
 * @param {string|null} requiredStatus - 遷移前の必須ステータス（null なら不問）
 * @param {function} writeFn - スプレッドシートへの書き込み処理（ss を受け取る）
 */
function facilitatorTransition(requiredStatus, writeFn) {
  const email = Session.getActiveUser().getEmail();
  const cache = getCache();
  assertFacilitator(email, cache);
  if (requiredStatus) assertState(cache, requiredStatus);
  writeFn(SpreadsheetApp.getActiveSpreadsheet());
  return buildUserState(rebuildCache(), email);
}

// ============================================================
// キャッシュ管理
// ============================================================

/**
 * スプレッドシートから全データを読み込み、キャッシュを再構築して返す
 * 状態変更を行う関数の末尾で呼ぶ
 */
function rebuildCache() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName('Config');
  const playersSheet = ss.getSheetByName('Players');

  if (!configSheet || !playersSheet) {
    throw new Error("シート「Config」または「Players」が見つかりません。");
  }

  const configValues = configSheet.getRange('A2:F2').getValues()[0];
  const status    = configValues[0] || 'WAITING';
  const startTime = configValues[1];
  const title       = configValues[3] || 'ワードウルフ オンライン';
  const wolfName    = configValues[4] || '人狼';
  const citizenName = configValues[5] || '村人';

  // Limitの評価は1回だけ行う (#3)
  const rawLimit = configValues[2];
  const limit = (rawLimit === "" || rawLimit === null || isNaN(rawLimit))
                ? 300 : Number(rawLimit);

  // Statusが空ならWAITING、Limitが空なら300を自動セット (#3)
  const needsStatus = !configValues[0];
  const needsLimit  = (rawLimit === "" || rawLimit === null || isNaN(rawLimit));
  if (needsStatus || needsLimit) {
    configSheet.getRange('A2:C2').setValues([[status, startTime, limit]]);
  }

  let endTime = null;
  if (startTime instanceof Date && !isNaN(limit)) {
    endTime = startTime.getTime() + limit * 1000;
  }

  const lastRow = playersSheet.getLastRow();
  const players = [];
  if (lastRow >= 2) {
    const rows = playersSheet.getRange(2, 1, lastRow - 1, 7).getValues();
    rows.forEach((row, i) => {
      if (!row[0] || !row[1]) return; // A・B列が空の行はスキップ
      players.push({
        email:         String(row[0]).trim(),
        nickname:      String(row[1]).trim(),
        role:          row[2] || "",
        word:          row[3] || "",
        checked:       row[4] === true,
        vote:          row[5] || "",
        isFacilitator: row[6] === true,
        sheetRow:      i + 2
      });
    });
  }

  const cache = { status, endTime, title, wolfName, citizenName, players };
  PropertiesService.getScriptProperties()
    .setProperty(STATE_CACHE_KEY, JSON.stringify(cache));

  return cache;
}

/**
 * キャッシュからユーザー向けのゲーム状態を生成して返す
 */
function buildUserState(cache, email) {
  const { status, endTime, title, wolfName, citizenName, players } = cache;

  const userData = players.find(p => p.email === email);
  if (!userData) return { status: 'UNAUTHORIZED', email };

  const playerCount  = players.length;
  const checkedCount = players.filter(p => p.checked).length;
  const playerList   = players.map(p => p.nickname);

  const voteTallies = {};
  players.forEach(p => {
    if (p.vote) voteTallies[p.vote] = (voteTallies[p.vote] || 0) + 1;
  });

  const allResults = status === 'REVEALED'
    ? players.map(p => ({
        name:     p.nickname || "名無し",
        role:     p.role,
        word:     p.word,
        votedFor: p.vote
      }))
    : [];

  return {
    status,
    endTime,
    title,
    wolfName,
    citizenName,
    myWord:        userData.word,
    isChecked:     userData.checked,
    isFacilitator: userData.isFacilitator,
    myVote:        userData.vote || null,
    countStr:      `${checkedCount} / ${playerCount}`,
    allResults,
    playerList,
    voteTallies
  };
}

// ============================================================
// 公開関数
// ============================================================

/**
 * 現在のゲーム状態を取得
 * キャッシュがあればスプレッドシートに触らない
 */
function getGameState() {
  try {
    const email = Session.getActiveUser().getEmail();
    return buildUserState(getCache(), email);
  } catch(e) {
    console.error("getGameState Error: " + e.message);
    return { status: 'ERROR', message: e.message };
  }
}

/**
 * お題を配布する
 */
function setupWords(wolfCount) {
  // email を try の外で取得しておく (#1)
  const email = Session.getActiveUser().getEmail();
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    const cache = getCache();

    // バリデーション
    assertFacilitator(email, cache);
    assertState(cache, 'WAITING');

    const count = parseInt(wolfCount);
    if (isNaN(count) || count < 1) {
      throw new Error("人狼の数は1以上の整数で指定してください。");
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const configSheet = ss.getSheetByName('Config');
    const configValues = configSheet.getRange('A2:F2').getValues()[0];
    const wolfName    = configValues[4] || '人狼';
    const citizenName = configValues[5] || '村人';

    const wordsSheet = ss.getSheetByName('Words');
    if (!wordsSheet) throw new Error("Wordsシートが見つかりません。");

    const words = wordsSheet.getDataRange().getValues();
    if (words.length < 2) throw new Error("Wordsシートにお題が登録されていません。");

    const pair = words[Math.floor(Math.random() * (words.length - 1)) + 1];
    const isSwapped   = Math.random() < 0.5;
    const citizenWord = isSwapped ? pair[1] : pair[0];
    const wolfWord    = isSwapped ? pair[0] : pair[1];

    const playerSheet = ss.getSheetByName('Players');
    const lastRow = playerSheet.getLastRow();
    if (lastRow < 2) throw new Error("プレイヤーが1人も参加していません。");

    // 行番号を保持しつつ、A列・B列が両方埋まっている行のみ対象にする
    const allRows = playerSheet.getRange(2, 1, lastRow - 1, 2).getValues();
    const players = allRows
      .map((row, i) => ({ email: row[0], nickname: row[1], sheetRow: i + 2 }))
      .filter(p => p.email !== "" && p.nickname !== "");

    if (players.length === 0) throw new Error("有効なプレイヤーが1人も参加していません。");

    // 人狼の数がプレイヤー数以上なら弾く
    if (count >= players.length) {
      throw new Error(`人狼の数はプレイヤー数（${players.length}名）より少なくしてください。`);
    }

    // ウルフをランダムに決定
    let playerIndices = players.map((_, i) => i);
    const wolfIndices = [];
    for (let i = 0; i < count; i++) {
      if (playerIndices.length === 0) break;
      const r = Math.floor(Math.random() * playerIndices.length);
      wolfIndices.push(playerIndices.splice(r, 1)[0]);
    }

    // 各プレイヤーの実際の行に書き込む（空白行はスキップ）
    players.forEach((player, i) => {
      const isWolf = wolfIndices.includes(i);
      playerSheet.getRange(player.sheetRow, 3, 1, 3).setValues([[
        isWolf ? wolfName : citizenName,
        isWolf ? wolfWord : citizenWord,
        false
      ]]);
    });

    configSheet.getRange('A2').setValue('DISTRIBUTING');

  } catch(e) {
    return { status: 'ERROR', message: e.message };
  } finally {
    lock.releaseLock();
  }

  // try の外なので email 変数をそのまま使い回せる (#1)
  return buildUserState(rebuildCache(), email);
}

/**
 * ワードを確認済みにする
 */
function markAsChecked() {
  try {
    const email = Session.getActiveUser().getEmail();
    const cache = getCache();
    assertState(cache, 'DISTRIBUTING');

    // sheetRow を使って1セルだけ書き込む（全行読み取り不要）(#2)
    const player = cache.players.find(p => p.email === email);
    if (!player) throw new Error("プレイヤーが見つかりません。");

    SpreadsheetApp.getActiveSpreadsheet()
      .getSheetByName('Players')
      .getRange(player.sheetRow, 5)
      .setValue(true);

    return buildUserState(rebuildCache(), email);
  } catch(e) {
    return { status: 'ERROR', message: e.message };
  }
}

/**
 * 投票を実行
 */
function castVote(targetNickname) {
  try {
    const email = Session.getActiveUser().getEmail();
    const cache = getCache();
    assertState(cache, 'VOTING');

    if (typeof targetNickname !== 'string' || targetNickname.trim() === '') {
      throw new Error("投票先が不正です。");
    }
    const trimmed = targetNickname.trim();

    // 存在するニックネームか確認
    const validNames = cache.players.map(p => p.nickname);
    if (!validNames.includes(trimmed)) {
      throw new Error("存在しないプレイヤーへの投票です。");
    }

    // sheetRow を使って最小限のアクセスで処理する (#2)
    const player = cache.players.find(p => p.email === email);
    if (!player) throw new Error("プレイヤーが見つかりません。");

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Players');

    // 二重投票チェックはシートから直接読む（キャッシュが古い場合の保険）(#2)
    const currentVote = sheet.getRange(player.sheetRow, 6).getValue();
    if (currentVote) throw new Error("すでに投票済みです。");

    sheet.getRange(player.sheetRow, 6).setValue(trimmed);

    return buildUserState(rebuildCache(), email);
  } catch(e) {
    return { status: 'ERROR', message: e.message };
  }
}

/**
 * ゲームを開始する
 */
function startGame() {
  try {
    return facilitatorTransition('DISTRIBUTING', ss => {
      ss.getSheetByName('Config')
        .getRange('A2:B2').setValues([['STARTED', new Date()]]);
    });
  } catch(e) {
    return { status: 'ERROR', message: e.message };
  }
}

/**
 * 投票フェーズへ移行する
 */
function startVoting() {
  try {
    return facilitatorTransition('STARTED', ss => {
      ss.getSheetByName('Config').getRange('A2').setValue('VOTING');
    });
  } catch(e) {
    return { status: 'ERROR', message: e.message };
  }
}

/**
 * 正解を表示する
 */
function revealAnswer() {
  try {
    return facilitatorTransition('VOTING', ss => {
      ss.getSheetByName('Config').getRange('A2').setValue('REVEALED');
    });
  } catch(e) {
    return { status: 'ERROR', message: e.message };
  }
}

/**
 * ゲームをリセットする
 */
function resetGame() {
  try {
    return facilitatorTransition(null, ss => {
      ss.getSheetByName('Config').getRange('A2:B2').setValues([['WAITING', null]]);
      const playerSheet = ss.getSheetByName('Players');
      const lastRow = playerSheet.getLastRow();
      if (lastRow > 1) {
        playerSheet.getRange(2, 3, lastRow - 1, 4).clearContent();
      }
    });
  } catch(e) {
    return { status: 'ERROR', message: e.message };
  }
}
