const ROW_HEADER = 1;
const ROW_START_DATA = 2;

const COL_DIR = 1;
const COL_FILE_ID = 2;
const COL_FILE_NAME = 3;
const COL_RENAME = 4;
const COL_RESULT = 5;

function onOpen(){
  //メニュー配列
  var myMenu=[
    {name: "ファイル一覧の取得", functionName: "getFileInfo"},
    {name: "名前の一括変換", functionName: "renameFile"}
  ];
  
  //メニューを追加
  SpreadsheetApp.getActiveSpreadsheet().addMenu("スクリプト",myMenu);
}

/*
 * ダイアログに指定された、GDriveフォルダ内にあるファイル/フォルダ情報を取得する
 */
function getFileInfo() {
  let files;
  let file;
  let folders;
  let folder;
  let i;
  let sh = SpreadsheetApp.getActiveSheet();
  let folderId = Browser.inputBox('取得したいGDriveのフォルダIDまたはURLを入力してください。');
  
  //GDriveのURLが入力されたときにID前のパスを削除
  folderId = folderId.replace('https://drive.google.com/drive/folders/', '');

  
  try{
    // ダイアログに何も入力されなかった場合→終了
    if(folderId == ''){
      throw new Error('A Folder ID is not defined.');
    }

    // シートの中身を削除
    sh.clearContents();

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

  // ヘッダ情報
  sh.getRange(ROW_HEADER, COL_DIR).setValue('Dir');
  sh.getRange(ROW_HEADER, COL_FILE_ID).setValue('ID');
  sh.getRange(ROW_HEADER, COL_FILE_NAME).setValue('Name');
  sh.getRange(ROW_HEADER, COL_RENAME).setValue('ReName（一括変換したい名前を入力）');
  sh.getRange(ROW_HEADER, COL_RESULT).setValue('処理');
}


function renameFile(){
  let sh = SpreadsheetApp.getActiveSheet();
  let dirFlg = '';
  let fileID = '';
  let fileRename = '';
  let i;

  // 処理結果をクリア
  sh.getRange(ROW_START_DATA, COL_RESULT, sh.getLastRow()-ROW_HEADER, 1).clearContent();

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






