// userRetrieval.gs
/**
 * 担当者苗字をキーにしてアカウント名を管理する関数
 * @param {*} sheetInfo
 * @returns {Object} 担当者名をキーにしたアカウント名のオブジェクト
 */
function getManagerAndUsername(sheetInfo) {
	Logger.log(JSON.stringify(sheetInfo, null, 2)); // sheetInfoの内容を確認

	const sheet = SpreadsheetApp.openById(sheetInfo.sheetId).getSheetByName(
		sheetInfo.sheetName
	);

	const result = {};

	// managerRange と usernameRange が配列であることを前提に処理
	sheetInfo.ranges.manager.forEach((managerRange, index) => {
		const usernameRange = sheetInfo.ranges.userName[index];

		const managerNames = sheet.getRange(managerRange).getValues();
		const usernames = sheet.getRange(usernameRange).getValues();

		usernames.forEach((row, rowIndex) => {
			let userName = row[0];
			userName = userName.replace(/[@＠]/g, "").trim(); // 「@」および「＠」記号の除去とトリミング

			const managerName = managerNames[rowIndex][0];

			// 担当者名をキーにして、対応するアカウント名を配列に追加
			if (!result[managerName]) {
				result[managerName] = [];
			}

			result[managerName].push(userName);
		});
	});

	return result;
}

// apiService.gs

/**
 * 指定したユーザーのメトリクスを取得する
 * @param {*} userName
 * @param {*} apiKey
 * @param {boolean} isMockMetrics - テストモードかどうか
 * @returns {Object} フォロワー数とフォロー数を含むオブジェクト
 */
function getUserMetrics(userName, apiKey, isMockMetrics = false) {
	if (isMockMetrics) {
		Logger.log("テストモードが有効です。ランダムなメトリクスを返します。");
		// テストモードの場合、ランダムなフォロワー数とフォロー数を返す
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
}

// writeSheet.gs

/**
 * GASでMATCH関数のような動作をする検索関数
 * @param {*} sheet	- シートオブジェクト
 * @param {*} searchValue - 検索する値
 * @param {*} range - 検索対象の範囲（例: "A1:A10"）
 * @param {*} exactMatch - 完全一致を求めるかどうか
 * @returns {string} マッチしたセルのアドレス（例: "A17"）
 */
function match(sheet, searchValue, range, exactMatch = true) {
	// 指定されたシートで範囲を取得する
	if (typeof range === "string") {
		range = sheet.getRange(range);
	}

	const values = range.getValues(); // 2次元配列で取得
	const numRows = values.length;
	const numCols = values[0].length;

	for (let row = 0; row < numRows; row++) {
		for (let col = 0; col < numCols; col++) {
			const cellValue = values[row][col];

			// セルの値を文字列に変換
			const cellValueStr = String(cellValue);
			const searchValueStr = String(searchValue);

			// 完全一致または部分一致をチェック
			if (exactMatch && cellValueStr === searchValueStr) {
				return range.getCell(row + 1, col + 1).getA1Notation(); // セルのアドレスを返す
			} else if (!exactMatch && cellValueStr.includes(searchValueStr)) {
				return range.getCell(row + 1, col + 1).getA1Notation();
			}
		}
	}
	throw new Error("該当する値が見つかりませんでした。");
}

/**
 * 列名を列番号に変換
 * @param {string} columnName - 列名（例: "A", "B", "AA", "AB"）
 * @returns {number} 列番号
 */
function columnNameToNumber(columnName) {
	let columnNumber = 0;
	const length = columnName.length;

	for (let i = 0; i < length; i++) {
		columnNumber *= 26;
		columnNumber += columnName.charCodeAt(i) - "A".charCodeAt(0) + 1;
	}

	return columnNumber;
}

/**
 * シートにデータを書き込む
 * @param {*} sheetInfo - シート情報
 * @param {*} manager - 担当者名
 * @param {*} accountData - アカウントデータ
 * @param {boolean} isDebug - デバッグモードかどうか
 */
function writeToSheet(sheetInfo, manager, accountData, isDebug = false) {
	const sheet = SpreadsheetApp.openById(sheetInfo.sheetId).getSheetByName(
		sheetInfo.sheetName
	);

	Logger.log(
		`
		[${isDebug ? "DEBUG" : "WRITE"}]
		ファイル名 ${sheet.getName()} 
		シート名 ${sheetInfo.sheetName}
		`
	); // null?

	// 日付列の範囲を取得
	const dateRange = sheet.getRange(sheetInfo.dateRange);

	// 今日の日付を取得
	const today = new Date();
	today.setHours(0, 0, 0, 0); // 時間をリセット

	// 今日の日付があるセルを探す
	let todayCell = match(sheet, today, dateRange, true);
	// sheetオブジェクトに変換
	todayCell = sheet.getRange(todayCell);
	Logger.log(`今日の日付が記載されているセル: ${todayCell.getA1Notation()}`);

	if (!todayCell) {
		throw new Error("今日の日付がシート内に見つかりませんでした。");
	}

	// 書き込み開始位置を取得
	let rowStart = todayCell.getRow();
	let colStart = todayCell.getColumn();

	// 日報シートの場合、日付行の下に担当者名が記載されている場合がある
	if (sheetInfo.surnameRange) {
		// 担当者名が記載されているセルを探す
		const surnameRange = sheet.getRange(sheetInfo.surnameRange);
		const surnameValues = surnameRange.getValues();

		// 日付行以下の範囲で担当者名を検索
		for (
			let i = rowStart - surnameRange.getRow();
			i < surnameValues.length;
			i++
		) {
			for (let j = 0; j < surnameValues[i].length; j++) {
				if (surnameValues[i][j] === manager) {
					rowStart = surnameRange.getRow() + i;
					break;
				}
			}
		}
	}

	// 行と列のオフセットを用いて、データを書き込む
	// sheetInfo.cell は、keyがメトリクス名、valueが列名のオブジェクト
	Object.keys(sheetInfo.cell).forEach((metrics, index) => {
		const rowOffsets = sheetInfo.cellOffsets.row;
		const colOffsets = sheetInfo.cellOffsets.col;

		let currentCol;
		// 列のオフセットが指定されている場合は、オフセットを適用
		if (colOffsets && colOffsets[index] !== undefined) {
			currentCol = todayCell.getColumn() + colOffsets[index]; // 日付セルからの列オフセット
		} else {
			// オフセットが指定されていない場合は、列名を基準にする
			currentCol = columnNameToNumber(sheetInfo.cell[metrics][0]);
		}

		// セルのオフセットを適用して、書き込む行を決定
		// rowOffsetsが [0, 1, 2, 3, 4] なら、担当者名の行から+0行、+1行、+2行、+3行、+4行のセルに書き込む
		rowOffsets.forEach((rowOffset, rowIndex) => {
			const username = Object.keys(accountData)[rowIndex]; // ユーザー名を取得
			// メトリクス名が accountData に存在するかをチェック
			if (!accountData[username].hasOwnProperty(metrics)) {
				return; // 存在しない場合は次のメトリクスに進む
			}

			const currentRow = rowStart + rowOffset; // 担当者名に基づく行
			const valueToWrite = accountData[username][metrics]; // ユーザーデータを取得

			const currentCell = sheet
				.getRange(currentRow, currentCol)
				.getA1Notation();

			Logger.log(`セル (${currentCell}) に "${valueToWrite}" を書き込みます。`);
			if (!isDebug) {
				sheet.getRange(currentRow, currentCol).setValue(valueToWrite);
			}
		});
	});
}

// main.gs
const main = () => {
	// 担当者名とアカウント名を取得
	const userData = getManagerAndUsername(accountInfo);
	Logger.log(JSON.stringify(userData, null, 2));

	// 担当者ごとに処理
	Object.keys(userData).forEach((manager) => {
		const accounts = userData[manager]; // アカウント名の配列を取得
		userData[manager] = {}; // 担当者名でオブジェクトを初期化

		// アカウントごとに処理
		accounts.forEach((account, index) => {
			try {
				const apiKey = API_KEYS[index % API_KEYS.length]; // APIキーを循環させる
				const metrics = getUserMetrics(account, apiKey, IS_MOCK_METRICS); // メトリクスを取得　テストモード
				userData[manager][account] = metrics; // ユーザー名をキーにしてメトリクスを格納
			} catch (error) {
				Logger.log(`Error fetching metrics for ${account}: ${error.message}`);
			}
			// 1分待機
			if (!IS_MOCK_METRICS) Utilities.sleep(1000 * 60);
		});
		Logger.log(JSON.stringify(userData, null, 2));

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
//*/
