// グローバル定数
const ROW_HEADER = 1;
const ROW_START_DATA = 2;

const COL_DIR = 1;
const COL_FILE_ID = 2;
const COL_FILE_NAME = 3;
const COL_RENAME = 4;
const COL_RESULT = 5;


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
 * ダイアログに指定された、GDriveフォルダ内にあるファイル/フォルダ情報を取得する
 */
function getFileLists() {
  let files;
  let file;
  let folders;
  let folder;
  let i;
  let sh = SpreadsheetApp.getActiveSheet();
  let folderId = Browser.inputBox('GDriveのフォルダIDまたはURLを入力してください。', Browser.Buttons.OK_CANCEL);
  
  //GDriveのURLが入力されたときにID前のパスを削除
  folderId = folderId.replace('https://drive.google.com/drive/folders/', '');

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
    for(i = ROW_START_DATA; files.hasNext(); i++) {
        file = files.next();
        sh.getRange(i, COL_FILE_ID).setValue(file.getId());
        sh.getRange(i, COL_FILE_NAME).setValue(file.getName());
    }
    
    // 取得したフォルダ情報を書き出し
    for(; folders.hasNext(); i++){
        folder = folders.next();
        sh.getRange(i, COL_DIR).setValue('d');
        sh.getRange(i, COL_FILE_ID).setValue(folder.getId());
        sh.getRange(i, COL_FILE_NAME).setValue(folder.getName());
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
  if(sh.getLastRow() - ROW_START_DATA >= 0){
    sh.getRange(ROW_START_DATA, COL_RESULT, sh.getLastRow()-ROW_HEADER, 1).clearContent();
  }

  for(i=ROW_START_DATA; i<=sh.getLastRow(); i++){
    dirFlg = sh.getRange(i, COL_DIR).getValue();
    fileID = sh.getRange(i, COL_FILE_ID).getValue();
    fileRename = sh.getRange(i, COL_RENAME).getValue();

    if(fileRename !== ''){
      if(dirFlg === ''){
        // ファイル名の変更
        DriveApp.getFileById(fileID).setName(fileRename);
      }else{
        //フォルダ名の変更
        DriveApp.getFolderById(fileID).setName(fileRename);
      }
      // 処理カラムにチェック
      sh.getRange(i, COL_RESULT).setValue('o');
    }
  }
}

/* 
 * テーブルを初期化（データをクリアしてヘッダを追加）
 */
function initTable(){
  let sh = SpreadsheetApp.getActiveSheet();

  // シートのデータをクリア
  sh.clearContents();
  
  // ヘッダ情報
  sh.getRange(ROW_HEADER, COL_DIR).setValue('Dir');
  sh.getRange(ROW_HEADER, COL_FILE_ID).setValue('ID');
  sh.getRange(ROW_HEADER, COL_FILE_NAME).setValue('Name');
  sh.getRange(ROW_HEADER, COL_RENAME).setValue('ReName（一括変換したい名前を入力）');
  sh.getRange(ROW_HEADER, COL_RESULT).setValue('処理');
}
