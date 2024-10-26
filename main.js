const IS_DEBUG = false; // デバッグモードを有効にするかどうか
const IS_MOCK_METRICS = false; // テストモードを有効にするかどうか

// userRetrieval.gs
/**
 * シートから指定された範囲のデータを取得する関数
 * @param {Sheet} sheet - Google Sheetsのシートオブジェクト
 * @param {Array} ranges - データを取得する範囲の配列
 * @returns {Array} 取得したデータの配列
 */
const getValuesFromRanges = (sheet, ranges) => 
  ranges.map(range => sheet.getRange(range).getValues());

/**
 * シート情報をログに出力する関数
 * @param {Object} data - シート情報のオブジェクト
 */
const logJson = data => Logger.log(JSON.stringify(data, null, 2));

/**
 * 担当者苗字をキーにしてアカウント名を管理する関数
 * @param {Object} sheetInfo - シート情報のオブジェクト
 * @returns {Object} チーム名、担当者名、アカウント名を含むオブジェクト
 */
const getManagerAndUsername = sheetInfo => {
  logJson(sheetInfo);

  const sheet = SpreadsheetApp.openById(sheetInfo.sheetId).getSheetByName(sheetInfo.sheetName);

  const teamData = getValuesFromRanges(sheet, sheetInfo.ranges.team);
  const managerData = getValuesFromRanges(sheet, sheetInfo.ranges.manager);
  const userNameData = getValuesFromRanges(sheet, sheetInfo.ranges.userName);

  const result = {};

  teamData.forEach((teamRange, rangeIndex) => {
    const managerRange = managerData[rangeIndex];
    const userNameRange = userNameData[rangeIndex];

    userNameRange.forEach((row, rowIndex) => {
      const userName = sanitizeUserName(row[0]);
      const managerName = managerRange[rowIndex][0];
      const teamName = teamRange[rowIndex][0];

      addUserNameToResult(result, teamName, managerName, userName);
    });
  });

  return result;
};

/**
 * ユーザー名をサニタイズする関数
 * @param {string} userName - サニタイズするユーザー名
 * @returns {string} サニタイズされたユーザー名
 */
const sanitizeUserName = userName => String(userName).replace(/[@＠]/g, "").trim();

/**
 * 結果オブジェクトにユーザー名を追加する関数
 * @param {Object} result - 結果オブジェクト
 * @param {string} teamName - チーム名
 * @param {string} managerName - 担当者名
 * @param {string} userName - ユーザー名
 */
const addUserNameToResult = (result, teamName, managerName, userName) => {
  const teamKey = String(teamName);
  const managerKey = String(managerName);
  const userKey = String(userName);

  if (!result[teamKey]) {
    result[teamKey] = {};
  }
  if (!result[teamKey][managerKey]) {
    result[teamKey][managerKey] = [];
  }
  result[teamKey][managerKey].push(userKey);
};

/**
 * 指定したユーザーのメトリクスを取得する
 * @param {string} userName - ユーザー名
 * @param {string} apiKey - APIキー
 * @param {boolean} [isMockMetrics=false] - テストモードかどうか
 * @returns {Object} フォロワー数とフォロー数を含むオブジェクト
 */
const getUserMetrics = (userName, apiKey, isMockMetrics = false) => {
  if (isMockMetrics) {
    Logger.log("テストモードが有効です。ランダムなメトリクスを返します。");
    return {
      followers: Math.floor(Math.random() * 1000),
      following: Math.floor(Math.random() * 1000),
    };
  }

  const url = `https://api.twitter.com/2/users/by/username/${userName}?user.fields=public_metrics`;
  const headers = {
    Authorization: `Bearer ${apiKey}`,
  };

  try {
    const response = UrlFetchApp.fetch(url, { headers });
    const data = JSON.parse(response.getContentText());

    if (data.errors) {
      throw new Error(data.errors);
    }

    return {
      followers: data.data.public_metrics.followers_count,
      following: data.data.public_metrics.following_count,
    };
  } catch (error) {
    Logger.log(`Error fetching metrics for ${userName}: ${error.message}`);
    return {
      followers: error.message,
      following: error.message,
    };
  }
};

// writeSheet.gs
/**
 * GASでMATCH関数のような動作をする検索関数
 * @param {Sheet} sheet - シートオブジェクト
 * @param {*} searchValue - 検索する値
 * @param {string|Range} range - 検索対象の範囲（例: "A1:A10"）
 * @param {boolean} [exactMatch=true] - 完全一致を求めるかどうか
 * @returns {string} マッチしたセルのアドレス（例: "A17"）
 * @throws {Error} 該当する値が見つからない場合にエラーをスロー
 */
const match = (sheet, searchValue, range, exactMatch = true) => {
  // 指定されたシートで範囲を取得する
  if (typeof range === "string") {
    range = sheet.getRange(range);
  }

  const values = range.getValues(); // 2次元配列で取得
  const numRows = values.length;
  const numCols = values[0].length;

  for (let row = 0; row < numRows; row++) {
    for (let col = 0; col < numCols; col++) {
      const cellValue = String(values[row][col]);
      const searchValueStr = String(searchValue);

      // 完全一致または部分一致をチェック
      if (
        (exactMatch && cellValue === searchValueStr) ||
        (!exactMatch && cellValue.includes(searchValueStr))
      ) {
        return range.getCell(row + 1, col + 1).getA1Notation(); // セルのアドレスを返す
      }
    }
  }
  throw new Error("該当する値が見つかりませんでした。");
};

/**
 * 列名を列番号に変換
 * @param {string} columnName - 列名（例: "A", "B", "AA", "AB"）
 * @returns {number} 列番号
 */
const columnNameToNumber = columnName => {
	// ...演算子で文字列を分割し、reduce関数で列番号に変換、accは累積値
  return [...columnName].reduce((acc, char) => {
    return acc * 26 + char.charCodeAt(0) - "A".charCodeAt(0) + 1;
  }, 0);
};

/**
 * シートにデータを書き込む
 * @param {Object} sheetInfo - シート情報
 * @param {string} manager - 担当者名
 * @param {Object} accountData - アカウントデータ
 * @param {boolean} [isDebug=false] - デバッグモードかどうか
 */
const writeToSheet = (sheetInfo, manager, accountData, isDebug = false) => {
  const sheet = SpreadsheetApp.openById(sheetInfo.sheetId).getSheetByName(sheetInfo.sheetName);

  logSheetInfo(sheet, sheetInfo, isDebug);

  const todayCell = getTodayCell(sheet, sheetInfo.dateRange);
  if (!todayCell) {
    throw new Error("今日の日付がシート内に見つかりませんでした。");
  }

  let rowStart = todayCell.getRow();
  if (sheetInfo.surnameRange) {
    rowStart = findManagerRow(sheet, sheetInfo.surnameRange, manager, rowStart);
  }

  writeMetricsToSheet(sheet, sheetInfo, accountData, rowStart, todayCell.getColumn(), isDebug);
};

/**
 * シート情報をログに出力する
 * @param {Sheet} sheet - シートオブジェクト
 * @param {Object} sheetInfo - シート情報
 * @param {boolean} isDebug - デバッグモードかどうか
 */
const logSheetInfo = (sheet, sheetInfo, isDebug) => {
  Logger.log(`
    [${isDebug ? "DEBUG" : "WRITE"}]
    ファイル名 ${sheet.getName()} 
    シート名 ${sheetInfo.sheetName}
  `);
};

/**
 * 今日の日付が記載されているセルを取得する
 * @param {Sheet} sheet - シートオブジェクト
 * @param {string} dateRange - 日付範囲
 * @returns {Range} 今日の日付が記載されているセル
 */
const getTodayCell = (sheet, dateRange) => {
  const range = sheet.getRange(dateRange);
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const todayCellAddress = match(sheet, today, range, true);
  const todayCell = sheet.getRange(todayCellAddress);
  Logger.log(`今日の日付が記載されているセル: ${todayCell.getA1Notation()}`);
  return todayCell;
};

/**
 * 担当者の行を見つける
 * @param {Sheet} sheet - シートオブジェクト
 * @param {string} surnameRange - 担当者名の範囲
 * @param {string} manager - 担当者名
 * @param {number} rowStart - 開始行
 * @returns {number} 担当者の行
 */
const findManagerRow = (sheet, surnameRange, manager, rowStart) => {
  const range = sheet.getRange(surnameRange);
  const values = range.getValues();

  for (let i = rowStart - range.getRow(); i < values.length; i++) {
    for (let j = 0; j < values[i].length; j++) {
      if (values[i][j] === manager) {
        return range.getRow() + i;
      }
    }
  }
  return rowStart;
};

/**
 * メトリクスをシートに書き込む
 * @param {Sheet} sheet - シートオブジェクト
 * @param {Object} sheetInfo - シート情報
 * @param {Object} accountData - アカウントデータ
 * @param {number} rowStart - 開始行
 * @param {number} colStart - 開始列
 * @param {boolean} isDebug - デバッグモードかどうか
 */
const writeMetricsToSheet = (sheet, sheetInfo, accountData, rowStart, colStart, isDebug) => {
  Object.keys(sheetInfo.cell).forEach((metrics, index) => {
    const rowOffsets = sheetInfo.cellOffsets.row;
    const colOffsets = sheetInfo.cellOffsets.col;

    const currentCol = getCurrentColumn(sheetInfo, metrics, colStart, colOffsets, index);

    rowOffsets.forEach((rowOffset, rowIndex) => {
      const username = Object.keys(accountData)[rowIndex];
      if (!accountData[username].hasOwnProperty(metrics)) {
        return;
      }

      const currentRow = rowStart + rowOffset;
      const valueToWrite = accountData[username][metrics];
      const currentCell = sheet.getRange(currentRow, currentCol).getA1Notation();

      Logger.log(`セル (${currentCell}) に "${valueToWrite}" を書き込みます。`);
      if (!isDebug) {
        sheet.getRange(currentRow, currentCol).setValue(valueToWrite);
      }
    });
  });
};

/**
 * 現在の列を取得する
 * @param {Object} sheetInfo - シート情報
 * @param {string} metrics - メトリクス
 * @param {number} colStart - 開始列
 * @param {Array} colOffsets - 列オフセット
 * @param {number} index - インデックス
 * @returns {number} 現在の列
 */
const getCurrentColumn = (sheetInfo, metrics, colStart, colOffsets, index) => {
  if (colOffsets && colOffsets[index] !== undefined) {
    return colStart + colOffsets[index];
  }
  return columnNameToNumber(sheetInfo.cell[metrics][0]);
};
// main.gs
const main = () => {
	const accountInfo =
		sheetInfo["SNSマーケインターン生シフト"]["アカウント一覧｜個人"];
	// チーム名、担当者名、アカウント名を取得
	const userData = getManagerAndUsername(accountInfo);
	logJson(userData);
	//チームごとに処理
	

	// // 担当者ごとに処理
	// Object.keys(userData).forEach((manager) => {
	// 	const accounts = userData[manager]; // アカウント名の配列を取得
	// 	userData[manager] = {}; // 担当者名でオブジェクトを初期化

	// 	// アカウントごとに処理
	// 	accounts.forEach((account, index) => {
	// 		try {
	// 			const apiKey = API_KEYS[index % API_KEYS.length]; // APIキーを循環させる
	// 			const metrics = getUserMetrics(account, apiKey, IS_MOCK_METRICS); // メトリクスを取得　テストモード
	// 			userData[manager][account] = metrics; // ユーザー名をキーにしてメトリクスを格納
	// 		} catch (error) {
	// 			Logger.log(`Error fetching metrics for ${account}: ${error.message}`);
	// 		}
	// 		// 1分待機
	// 		if (!IS_MOCK_METRICS) Utilities.sleep(1000 * 60);
	// 	});
	// 	Logger.log(JSON.stringify(userData, null, 2));

		/*
        writeToSheet(
            sheetInfo.personalManagementReport,
            manager,
            userData[manager],
            true
        ); // メトリクスを書き込む
        //*/

		sheetInfo.shiftTableBySurname.sheetName = manager; // 担当者名をシート名に設定
		writeToSheet(
			sheetInfo.shiftTableBySurname,
			manager,
			userData[manager],
			IS_DEBUG
		); // メトリクスを書き込む
	});
	/*
	// sheetInfoを繰り返して、sheetInfo.nameがsurnameと一致するものを探す
	// accountMetricsのキーがsheetInfo.nameに含まれていたら、sheetInfo.sheetNameにsheetNameを設定
	// その後、writeToSheetを実行
	sheetInfo.forEach((sheet) => {
		if (sheet.name.includes(teamName)) {
			writeToSheet(sheet, manager, userData[manager], IS_DEBUG);
			}
			});
			// */
};

// トリガー設定
function setupTrigger() {
	ScriptApp.newTrigger("main").timeBased().everyMinutes(15).create();
}
