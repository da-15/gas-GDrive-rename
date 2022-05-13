'use strict';
// mod：2022/05/13
// グローバル定数 --------------------------------------------------
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
  }
};

/*
 * シートを開いた時の処理
 * メニューの追加
 */
function onOpen(){
  //メニュー配列
  SpreadsheetApp.getUi()
    .createMenu('マクロ実行')
    .addItem('ファイル一覧の取得', 'getFileLists')
    .addItem('名前の一括変換', 'renameFiles')
    .addSeparator()
    .addItem('データクリア', 'initTable')
    .addToUi();
  
  //初期説明ダイアログの表示
  Browser.msgBox('TIPS:\\nメニュー："マクロ実行" から処理を開始してください。');
}

/* 
 * テーブルを初期化（データをクリアしてヘッダを追加）
 */
function initTable(){
  let sh = SpreadsheetApp.getActiveSheet();

  // シートのデータをクリア
  sh.clearContents();
  
  // ヘッダ情報
  sh.getRange(CONF.ROW.HEADER, CONF.COL.DIR).setValue('Dir');
  sh.getRange(CONF.ROW.HEADER, CONF.COL.FILE_ID).setValue('ID');
  sh.getRange(CONF.ROW.HEADER, CONF.COL.FILE_NAME).setValue('Name');
  sh.getRange(CONF.ROW.HEADER, CONF.COL.RENAME).setValue('Name（リネームしたいもののみ入力）');
  sh.getRange(CONF.ROW.HEADER, CONF.COL.RESULT).setValue('処理');

  sh.getRange(1,1,1,5).setBackground('#c5d7dc');
}

/*
 * ダイアログに指定された、GDriveフォルダ内にあるファイル/フォルダ情報を取得する
 */
function getFileLists() {
  let files;
  let file;
  let folders;
  let folder;
  let i;
  let sh = SpreadsheetApp.getActiveSheet();
  let folderId = Browser.inputBox('GDriveのフォルダIDまたはURLを入力してください。', 
    Browser.Buttons.OK_CANCEL);
  
  //GDriveのURLが入力されたときにID前後のパスを削除
  folderId = folderId.replace('https://drive.google.com/drive/folders/', '');
  folderId = folderId.replace(/\?.*/, '');
  
  try{
    if(folderId === ''){
      // ダイアログに何も入力されなかった場合→終了
      throw new Error('A Folder ID is not defined.');
    }else if(folderId === 'cancel'){
      // ダイアログがキャンセルされた場合→終了
      throw new Error('Dialog canceled');
    }

    // テーブルを初期化
    initTable();

    // ファイルリストを取得したい親フォルダをセット
    files = DriveApp.getFolderById(folderId).getFiles(); 
    // フォルダリストを取得したい親フォルダセット
    folders = DriveApp.getFolderById(folderId).getFolders(); 
    
    // 取得したファイル情報を書き出し
    for(i = CONF.ROW.START_DATA; files.hasNext(); i++) {
        file = files.next();
        sh.getRange(i, CONF.COL.FILE_ID).setValue(file.getId());
        sh.getRange(i, CONF.COL.FILE_NAME).setValue(file.getName());
    }
    
    // 取得したフォルダ情報を書き出し
    for(; folders.hasNext(); i++){
        folder = folders.next();
        sh.getRange(i, CONF.COL.DIR).setValue('d');
        sh.getRange(i, CONF.COL.FILE_ID).setValue(folder.getId());
        sh.getRange(i, CONF.COL.FILE_NAME).setValue(folder.getName());
    }
  }
  catch(error){
    console.error(error);
  }
}

/*
 * ファイル／フォルダ名を一括変更する
 */
function renameFiles(){
  let sh = SpreadsheetApp.getActiveSheet();
  let dirFlg = '';
  let fileID = '';
  let fileRename = '';
  let i;

  // 処理結果をクリア
  if(sh.getLastRow() - CONF.ROW.START_DATA >= 0){
    sh.getRange(CONF.ROW.START_DATA, CONF.COL.RESULT, 
      sh.getLastRow() - CONF.ROW.HEADER, 1).clearContent();
  }

  for(i = CONF.ROW.START_DATA; i<=sh.getLastRow(); i++){
    dirFlg = sh.getRange(i, CONF.COL.DIR).getValue();
    fileID = sh.getRange(i, CONF.COL.FILE_ID).getValue();
    fileRename = sh.getRange(i, CONF.COL.RENAME).getValue();

    if(fileRename !== ''){
      if(dirFlg === ''){
        // ファイル名の変更
        DriveApp.getFileById(fileID).setName(fileRename);
      }else{
        //フォルダ名の変更
        DriveApp.getFolderById(fileID).setName(fileRename);
      }
      // 処理カラムにチェック
      sh.getRange(i, CONF.COL.RESULT).setValue('o');
    }
  }
}


