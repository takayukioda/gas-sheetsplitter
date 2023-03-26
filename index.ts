const onOpen = (event: GoogleAppsScript.Events.SheetsOnOpen) => {
    const menu = SpreadsheetApp.getUi().createAddonMenu();
    menu.addItem("分割する", "run");
    menu.addToUi();
}

// appendAndApply 短期間で複数回 `appendRow` を呼び出す際に上書きされないよう対策している補助関数
//
// `appendRow` は短い間隔で複数回呼ばれると（for文の中で呼ぶなど）前の `appendRow` が保存される前にシート
// に対して行を追加する事によって行の上書きが発生する時がある。
// `appendRow` を呼び出した直後に変更を保存する関数を呼び出す事でこの問題を回避している。
const appendAndApply = (sheet: GoogleAppsScript.Spreadsheet.Sheet, row: any[]) => {
    sheet.appendRow(row);
    // NOTE: flush という名前が若干勘違いしそうだけど、「変更を保存する」という関数
    SpreadsheetApp.flush();
}

// uniq は配列を受け取り、重複する値を排除した配列を返す
const uniq = <T = any>(arr: T[]): T[] => {
    const list: T[] = [];
    for (const value of arr) {
        if (list.indexOf(value) < 0) {
            list.push(value);
        }
    }
    return list;
}

const run = () => {
    const ui = SpreadsheetApp.getUi();

    const keyColResp = ui.prompt("確認", "何列目の値をキーにしますか？列番号は1始まりの数値です。", ui.ButtonSet.OK_CANCEL);
    if (keyColResp.getSelectedButton() === ui.Button.CANCEL) {
        ui.alert("作業停止", "キャンセルされました", ui.ButtonSet.OK);
        return;
    }

    const keyColInput = Number(keyColResp.getResponseText());
    if (isNaN(keyColInput)) {
        ui.alert("エラー", "数値を指定してください", ui.ButtonSet.OK);
        return;
    }

    console.log("アクティブなシートを取得します。");
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getActiveSheet();
    const lastRowNum = sheet.getLastRow();
    const lastColNum = sheet.getLastColumn();
    console.log(`${sheet.getSheetName()} というシートを取得しました。サイズは${lastRowNum}行${lastColNum}列です`);

    if (lastColNum < keyColInput) {
        ui.alert("エラー", "指定した値がカラム数より大きいです", ui.ButtonSet.OK);
        return;
    }

    const HEADER_ROW_NUMBER = 1
    const KEY_COLUMN_NUMBER = keyColInput
    const EMPTY_KEY = "empty"

    console.log("ヘッダーを取得します。");
    const header = sheet.getRange(1, 1, HEADER_ROW_NUMBER, lastColNum).getValues().flat();

    const readyResp = ui.alert("確認", `「${header[KEY_COLUMN_NUMBER - 1]}」の列をキーに分割を始めます`, ui.ButtonSet.OK_CANCEL);
    if (readyResp === ui.Button.CANCEL) {
        ui.alert("作業停止", "キャンセルされました", ui.ButtonSet.OK);
        return;
    }

    // NOTE: ヘッダーを除いた、取得したいデータのサイズ
    const rowSize = sheet.getLastRow() - HEADER_ROW_NUMBER;
    const colSize = lastColNum;

    const keyColumn = sheet.getRange(HEADER_ROW_NUMBER + 1, KEY_COLUMN_NUMBER, rowSize).getValues().flat();
    const keys = uniq<string>(keyColumn).map(key => key.length ? key : EMPTY_KEY);
    console.log(`分割キーは${keys.length}件あります。`);

    console.log("データを取得します。");
    const values: string[][] = sheet.getRange(HEADER_ROW_NUMBER + 1, 1, rowSize, colSize).getValues();
    console.log(`データは ${values.length} 件あります。`);

    const separated = keys.reduce((map, key) => {
        return map.set(key, values.filter(row => row[KEY_COLUMN_NUMBER - 1] === key));
    }, new Map<string, string[][]>());

    for (const [key, rows] of Array.from(separated.entries())) {
        if (key === EMPTY_KEY && !rows.length) {
            continue;
        }
        const existingSheet = spreadsheet.getSheetByName(key);
        // NOTE: 二重呼び出しされた場合は洗い替えする。
        if (existingSheet) {
            console.log(`シート ${key} は洗い替えします。`);
            existingSheet.clearContents();
            SpreadsheetApp.flush();
        }
        console.log(`シート ${key} に ${rows.length} 行のデータを追加します。`);
        const newSheet = existingSheet || spreadsheet.insertSheet(key);
        appendAndApply(newSheet, header);
        rows.forEach(row => appendAndApply(newSheet, row));
    }
}
