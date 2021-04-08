//mainクラス
function main() {
  //シートを指定
  var sheet = SpreadsheetApp.getActiveSheet();
  //Slideを指定
  var url = '-------------------------------';
  var presentation = SlidesApp.openByUrl(url);

  //行数取得(最後の行まで取得)
  //行が増えても自動更新
  var rows = sheet.getLastRow() - 1;

  //スプレッドシートの指定範囲のデータを二次元配列で取得する
  var matchDatas = getList(sheet, rows);

  //オブジェクトIDログ出力
  findObject(presentation);

  //Slideのテンプレートを紐付け、全オブジェクトを返す
  findById(presentation, matchDatas);

  //既読フラグを付ける
  addFlag(sheet, rows);
}


//----------------------------------------------------------------------------------------------------------------//


//SS(DB)内の指定したデータを二次元配列で返す
//既読フラグなしの行のみ取得する
function getList(sheet, rows) {
  //データを取得する範囲
  var range = sheet.getRange(2, 1, rows, 9);

  //rangeの範囲だけセルを取得
  var datas = range.getValues();

  var match = [];
  //二次元配列から取得する
  for (i = 0; i < datas.length; i++) {
    var cup = datas[i][0];
    var home = datas[i][1];
    var away = datas[i][2];
    var stadium = datas[i][3];
    var month = datas[i][4];
    var day = datas[i][5];
    var hTS = datas[i][6];
    var aTS = datas[i][7];
    var flag = datas[i][8];

    if (flag !== 1) {
      match.push(cup, home, away, stadium, month, day, hTS, aTS);
      return match;
    }
  }
}


//オブジェクトIDログ出力
function findObject(presentation) {
  //スライドを選択
  var slide = presentation.getSlides()[0];
  //Shape型で取得
  var shapes = slide.getShapes();

  //Shapeごとに出力
  for (var i = 0; i < shapes.length; i++) {
    var shape = slide.getShapes()[i];
    //オブジェクトIDとテキストを取得
    Logger.log(shape.getObjectId() + "：" + shape.getText().asString());
  }
}


//Slideの各オブジェクトを紐付け、値を出力する
function findById(presentation, matchDatas) {
  //オブジェクトIDを指定して文字を置換
  //大会名オブジェクト
  var cup_id = 'i1';
  //日付オブジェクト
  var date_id = 'gcedf0568fa_0_10';
  //会場オブジェクト
  var stadium_id = 'gcedf0568fa_0_11';
  //HOMEチームオブジェクト
  var home_team_id = 'gcedf0568fa_0_7';
  //HOMEチームスコアオブジェクト
  var home_team_score_id = 'gcedf0568fa_0_5';
  //AWAYチームオブジェクト
  var away_team_id = 'gcedf0568fa_0_6';
  //AWAYチームスコアオブジェクト
  var away_team_score_id = 'gcedf0568fa_0_8';

  //idに該当するShapeを紐付け
  var ShapeCup = presentation.getPageElementById(cup_id).asShape();
  var ShapeDate = presentation.getPageElementById(date_id).asShape();
  var ShapeStadium = presentation.getPageElementById(stadium_id).asShape();
  var ShapeHomeTeam = presentation.getPageElementById(home_team_id).asShape();
  var ShapeHomeScore = presentation.getPageElementById(home_team_score_id).asShape();
  var ShapeAwayTeam = presentation.getPageElementById(away_team_id).asShape();
  var ShapeAwayScore = presentation.getPageElementById(away_team_score_id).asShape();

  var strC = matchDatas[0];
  var strD = matchDatas[4] + "/" + matchDatas[5];
  var strSta = matchDatas[3];
  var strHT = matchDatas[1];
  var strHS = matchDatas[6];
  var strAT = matchDatas[2];
  var strAS = matchDatas[7];

  ShapeCup.getText().setText(strC);
  ShapeDate.getText().setText(strD);
  ShapeStadium.getText().setText(strSta);
  ShapeHomeTeam.getText().setText(strHT);
  ShapeHomeScore.getText().setText(strHS);
  ShapeAwayTeam.getText().setText(strAT);
  ShapeAwayScore.getText().setText(strAS);
}


//既読フラグを付ける
function addFlag(sheet, rows) {
  var flag = 1;
  sheet.getRange(rows + 1, 9).setValue(flag);
}