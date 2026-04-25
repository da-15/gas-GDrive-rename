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
    ERROR:'エラー',
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
  SpreadsheetApp.getUi()
    .createMenu('GDrive名前一括変換')
    .addItem('フォルダのパスを指定する', 'getFileLists')
    .addItem('名前を一括変換する', 'renameFiles')
    .addSeparator()
    .addItem('一覧をクリアする', 'initTable')
    .addItem('TIPSを表示', 'dispTips')
    .addToUi();
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
