'use strict';
/*
 * GDriveフォルダのURLまたはIDからフォルダIDを抽出する
 */
function extractFolderId(input) {
  // /drive/folders/{id} 形式（/u/0/ などのユーザー番号を含む場合も対応）
  const match = input.match(/\/folders\/([a-zA-Z0-9_-]+)/);
  if (match) return match[1];
  // URLではなくIDが直接入力された場合はそのまま返す
  return input;
}

/*
 * ダイアログに指定された、GDriveフォルダ内にあるファイル/フォルダ情報を取得する
 */
function getFileLists() {
  const sh = SpreadsheetApp.getActiveSheet();

  let rawInput = Browser.inputBox(CONF.MSG.ID_INPUT, Browser.Buttons.OK_CANCEL);

  if(rawInput === 'cancel'){
    return;
  }

  const folderId = extractFolderId(rawInput.trim());

  try{
    if(folderId === ''){
      throw new Error(CONF.MSG.ERROR_01);
    }

    initTable();

    // getFolderById を1回だけ呼び出す
    const folder = DriveApp.getFolderById(folderId);
    const files = folder.getFiles();
    const folders = folder.getFolders();

    const rows = [];

    while(files.hasNext()) {
      const file = files.next();
      rows.push(['', file.getId(), file.getName(), '', '']);
    }

    while(folders.hasNext()) {
      const f = folders.next();
      rows.push([CONF.FLAG.FOLDER, f.getId(), f.getName(), '', '']);
    }

    if(rows.length > 0){
      // まとめて1回のAPI呼び出しで書き込む
      sh.getRange(CONF.ROW.START_DATA, CONF.COL.DIR, rows.length, 5).setValues(rows);
    }
  }
  catch(error){
    console.error(error);
    Browser.msgBox('エラー:\n' + error);
  }
}

/*
 * ファイル／フォルダ名を一括変更する
 */
function renameFiles(){
  const sh = SpreadsheetApp.getActiveSheet();
  const lastRow = sh.getLastRow();

  if(lastRow < CONF.ROW.START_DATA) return;

  const numDataRows = lastRow - CONF.ROW.HEADER;

  // 処理結果をクリア
  sh.getRange(CONF.ROW.START_DATA, CONF.COL.RESULT, numDataRows, 1).clearContent();

  // 対象データを一括取得
  const dataRange = sh.getRange(CONF.ROW.START_DATA, CONF.COL.DIR, numDataRows, CONF.COL.RESULT);
  const data = dataRange.getValues();
  const results = data.map(row => [row[CONF.COL.RESULT - 1]]);

  for(let i = 0; i < data.length; i++){
    const dirFlg    = data[i][CONF.COL.DIR - 1];
    const fileID    = data[i][CONF.COL.FILE_ID - 1];
    const fileRename = data[i][CONF.COL.RENAME - 1];

    if(fileRename === '' || fileID === '') continue;

    try{
      if(dirFlg === ''){
        DriveApp.getFileById(fileID).setName(fileRename);
      }else{
        DriveApp.getFolderById(fileID).setName(fileRename);
      }
      results[i][0] = CONF.FLAG.DONE;
    }catch(error){
      console.error(error);
      results[i][0] = CONF.FLAG.ERROR;
    }
  }

  // 結果を一括書き込み
  sh.getRange(CONF.ROW.START_DATA, CONF.COL.RESULT, results.length, 1).setValues(results);
}

/*
 * テーブルを初期化（データをクリアしてヘッダを追加）
 */
function initTable(){
  const sh = SpreadsheetApp.getActiveSheet();

  sh.clearContents();

  // ヘッダを一括セット
  sh.getRange(CONF.ROW.HEADER, CONF.COL.DIR, 1, 5).setValues([[
    CONF.TITLE.DIR,
    CONF.TITLE.ID,
    CONF.TITLE.NAME,
    CONF.TITLE.RENAME,
    CONF.TITLE.STATUS
  ]]);

  sh.setColumnWidth(CONF.COL.DIR, 50);
  sh.setColumnWidth(CONF.COL.FILE_ID, 300);
  sh.setColumnWidth(CONF.COL.FILE_NAME, 400);
  sh.setColumnWidth(CONF.COL.RENAME, 400);
  sh.setColumnWidth(CONF.COL.RESULT, 50);

  sh.getRange(1, 1, 1, CONF.COL.RESULT).setBackground(CONF.TITLE.COLOR);
}
