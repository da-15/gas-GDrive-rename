'use strict';
const CONF = {
  ROW: {
    HEADER: 1, //ヘッダ行
    START_DATA: 2 // データ開始行
  },
  COL:{
    DIR:1, // DIRカラム番号
    FILE_ID:2, // ファイルIDカラム番号
    FILE_NAME:3, // ファイル名 カラム番号
    RENAME:4, // リネーム名指定 カラム番号
    RESULT:5 // 結果出力用 カラム番号
  },
  TITLE:{
    DIR:'Dir',
    ID:'ID',
    NAME:'ファイル / フォルダ名（元）',
    RENAME:'※変更したい名前を入力。空はスキップされます。\nファイル / フォルダ名（変更後）',
    STATUS:'処理',
    COLOR:'#cbdcf6'
  },
  FLAG:{
    DONE:'済',
    FOLDER:'d'
  },
  MSG:{
    ID_INPUT: 'GDriveのフォルダIDまたはURLを入力してください。',
    ERROR_01: 'GDriveのフォルダIDが指定されていません。'
  }
};

/*
 * シートを開いた時の処理
 * メニューの追加
 */
function onOpen(){
  //メニュー配列
  SpreadsheetApp.getUi()
    .createMenu('GDrive名前一括変換')
    .addItem('フォルダのパスを指定する', 'getFileLists')
    .addItem('名前を一括変換する', 'renameFiles')
    .addSeparator()
    .addItem('一覧をクリアする', 'initTable')
    .addItem('TIPSを表示', 'dispTips')
    .addToUi();
  
  //初期説明ダイアログの表示
  dispTips();
}

/*
 * TIPSを表示する
 */
function dispTips(){
  const msg = '' +
    '【 TIPS 】\\n' +
    'メニュー「GDrive名前一括変換」から処理をはじめます。\\n\\n' + 
    '1. メニュー「フォルダのパスを指定する」にて対象のURLを指定する\\n' + 
    '2. 一覧にて変更したい名前を指定する\\n' + 
    '3. メニュー「名前を一括変換する」にて変換を開始する';
   Browser.msgBox(msg);
}

/*
 * ダイアログに指定された、GDriveフォルダ内にあるファイル/フォルダ情報を取得する
 */
function getFileLists() {
  const sh = SpreadsheetApp.getActiveSheet();
  
  //フォルダIDを取得（GDriveのURLが入力された場合はID前後のパスを削除）
  let folderId = Browser.inputBox(CONF.MSG.ID_INPUT, Browser.Buttons.OK_CANCEL);
  folderId = folderId.replace('https://drive.google.com/drive/folders/', '');
  folderId = folderId.replace(/\?.*/, '');
  
  try{
    if(folderId === ''){
      // ダイアログに何も入力されなかった場合
      throw new Error(CONF.MSG.ERROR_01);
    }else if(folderId === 'cancel'){
      // ダイアログがキャンセルされた場合→終了
      return;
    }

    // テーブルを初期化
    initTable();

    // ファイルリストを取得したい親フォルダをセット
    const files = DriveApp.getFolderById(folderId).getFiles(); 
    // フォルダリストを取得したい親フォルダセット
    const folders = DriveApp.getFolderById(folderId).getFolders(); 
    
    // 取得したファイル情報を書き出し
    let i;
    for(i = CONF.ROW.START_DATA; files.hasNext(); i++) {
        const file = files.next();
        sh.getRange(i, CONF.COL.FILE_ID).setValue(file.getId());
        sh.getRange(i, CONF.COL.FILE_NAME).setValue(file.getName());
    }
    
    // 取得したフォルダ情報を書き出し
    for(; folders.hasNext(); i++){
        const folder = folders.next();
        sh.getRange(i, CONF.COL.DIR).setValue(CONF.FLAG.FOLDER);
        sh.getRange(i, CONF.COL.FILE_ID).setValue(folder.getId());
        sh.getRange(i, CONF.COL.FILE_NAME).setValue(folder.getName());
    }
  }
  catch(error){
    console.error(error);
    Browser.msgBox('エラー:\\n' + error);
  }
}

/*
 * ファイル／フォルダ名を一括変更する
 */
function renameFiles(){
  const sh = SpreadsheetApp.getActiveSheet();
  
  // 処理結果をクリア
  if(sh.getLastRow() - CONF.ROW.START_DATA >= 0){
    sh.getRange(CONF.ROW.START_DATA, CONF.COL.RESULT, 
      sh.getLastRow() - CONF.ROW.HEADER, 1).clearContent();
  }

  for(let i = CONF.ROW.START_DATA; i<=sh.getLastRow(); i++){
    const dirFlg = sh.getRange(i, CONF.COL.DIR).getValue();
    const fileID = sh.getRange(i, CONF.COL.FILE_ID).getValue();
    const fileRename = sh.getRange(i, CONF.COL.RENAME).getValue();

    if(fileRename !== ''){
      if(dirFlg === ''){
        // ファイル名の変更
        DriveApp.getFileById(fileID).setName(fileRename);
      }else{
        //フォルダ名の変更
        DriveApp.getFolderById(fileID).setName(fileRename);
      }
      // 処理カラムにチェック
      sh.getRange(i, CONF.COL.RESULT).setValue(CONF.FLAG.DONE);
    }
  }
}

/* 
 * テーブルを初期化（データをクリアしてヘッダを追加）
 */
function initTable(){
  const sh = SpreadsheetApp.getActiveSheet();

  // シートのデータをクリア
  sh.clearContents();

  // ヘッダ情報
  sh.getRange(CONF.ROW.HEADER, CONF.COL.DIR).setValue(CONF.TITLE.DIR);
  sh.getRange(CONF.ROW.HEADER, CONF.COL.FILE_ID).setValue(CONF.TITLE.ID);
  sh.getRange(CONF.ROW.HEADER, CONF.COL.FILE_NAME).setValue(CONF.TITLE.NAME);
  sh.getRange(CONF.ROW.HEADER, CONF.COL.RENAME).setValue(CONF.TITLE.RENAME);
  sh.getRange(CONF.ROW.HEADER, CONF.COL.RESULT).setValue(CONF.TITLE.STATUS);

  // ヘッダの幅調整
  sh.setColumnWidth(CONF.COL.DIR, 50);
  sh.setColumnWidth(CONF.COL.FILE_ID, 300);
  sh.setColumnWidth(CONF.COL.FILE_NAME, 400);
  sh.setColumnWidth(CONF.COL.RENAME, 400);
  sh.setColumnWidth(CONF.COL.RESULT, 50);
  
  //ヘッダの色
  sh.getRange(1,1,1,CONF.COL.RESULT).setBackground(CONF.TITLE.COLOR);
}

