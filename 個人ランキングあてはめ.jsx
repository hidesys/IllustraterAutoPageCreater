#include "json2.js";

//********
//環境変数の設定
//********

//ディレクトリの構造は以下を考えている
// C:\hoge~\作業量ディレクトリ\
//                            |- テンプレート.ai
//                            |- users.json  #=> http://mantropy.net/users.json を使う
//                            |- 入力自画像\
//                            |- 入力総評\
//                            |- 出力AI\
//                            |- 出力PDF\
//                            |- 出力PNG\
var pwdPath = "C:\\Users\\washi\\Desktop\\個人ランキング\\作業用ディレクトリ\\";

//各Pathを用意
var inputTemplatePath = pwdPath + "テンプレート.ai";
var inputRankingsPath = pwdPath + "users.json";
var inputSelfieDirectory = pwdPath + "入力自画像\\";
var inputCommentDirectory = pwdPath + "入力総評\\";
var outputAiDirectory = pwdPath + "出力AI\\";
var outputPdfDirectory = pwdPath + "出力PDF\\";
var outputPngDirectory = pwdPath + "出力PNG\\";

//テンプレートで画像を埋め込む位置と画像の大きさ
// 正方形を仮定してる
// XとYは自画像領域の重心。当たり前だけどこれらはテンプレート.aiから算出する
var selfieWidthHeight = 115.0;
var selfieX = 88.812;
var selfieY = -93.329;

//置換用の辞書を用意する
// 総評.txtの何行目がどの文字列に変換されるか、の辞書
var replaceDictionary = {
  0: "メインコメント置換",
  1: "糞コメント置換",
  2: "判決置換",
  3: "性別",
  4: "生年月日",
  5: "ジャンル置換",
  6: "回生",
  7: "入トロ歴",
  8: "雑誌置換",
  9: "漫画家置換",
  10: "好きな漫画置換",
  11: "他の所属",
  12: "自己紹介"
};

//********
//ここからスクリプト本体
//******** 

//users.jsonを読み込む
var rankingsFile = new File(inputRankingsPath);
rankingsFile.open("r");
var rankings = JSON.parse(rankingsFile.read())["users"];
rankingsFile.close();

//総評ディレクトリの総評一覧を取得
var commentFileList = (new Folder(inputCommentDirectory)).getFiles("*.txt");

//以下、総評のエントリ順に処理していく
for (var i = 0; i < commentFileList.length; i++){
  var handleName = commentFileList[i].name.split(".")[0];
  var entry = getUserEntry(rankings, handleName);

  //テンプレートファイルを開く
  var docObj = open(new File(inputTemplatePath));

  //ハンドルネームの置換
  chikan(docObj, "ハンドルネーム", handleName);

  //ランキングの置換
  if (entry !== null){
    for (var j = 5; j > 0; j--){
      var rank = getRank(entry["ranks"], "糞" + j);
      chikan(docObj, "作家名糞" + j + "位", rank["author"]);
      chikan(docObj, "糞ラン" + j + "位", rank["name"]);
    }
    for (var j = 30; j > 0; j--){
      var rank = getRank(entry["ranks"], "" + j);
      chikan(docObj, "" + j + "位作家名", rank["author"]);
      chikan(docObj, "ランキング" + j + "位", rank["name"]);
    }
  }

  //総評ファイルを読み込む
  var commentFile = new File(inputCommentDirectory + handleName + ".txt");
  if (commentFile.exists){
    commentFile.open("r");
    var comments = commentFile.read().split(/\r\n|\r|\n/);
    commentFile.close();
    //総評の置換
    for (var key in replaceDictionary){
      chikan(docObj, replaceDictionary[key], comments[key]);
    }
  }

  //自画像の挿入
  var selfieFile = new File(inputSelfieDirectory + handleName + ".png");
  if (!selfieFile.exists){
    selfieFile = new File(inputSelfieDirectory + handleName + ".jpg");
  }
  if (selfieFile.exists){
    var selfie = docObj.placedItems.add();
    selfie.file = selfieFile;
    var width = selfie.width;
    var height = selfie.height;
    if (width > height){
      selfie.width = selfieWidthHeight;
      selfie.height = selfieWidthHeight * height / width;
    }else{
      selfie.width = selfieWidthHeight * width / height;
      selfie.height = selfieWidthHeight;
    }
    selfie.position = [selfieX - selfie.width / 2, selfieY + selfie.height / 2];
  }

  //保存
  // ページをJPEGで出力
  var pngOption = new ExportOptionsPNG8();
  pngOption.artBoardClipping = true;
  pngOption.transparency = false;
  pngOption.horizontalScale = 300;
  pngOption.verticalScale = 300;
  docObj.exportFile(new File(outputPngDirectory + handleName + ".png"), ExportType.PNG8, pngOption);
  
  // ページをPNGで出力
  var pdfOption = new PDFSaveOptions();
  pdfOption.compatibility = PDFCompatibility.ACROBAT5;
  pdfOption.generateThumbnails = true;
  pdfOption.preserveEditability = true;
  pdfOption.optimization = false;
  docObj.saveAs(new File(outputPdfDirectory + handleName + ".pdf"), pdfOption);

  // ページをAI別名で保存
  var aiOption = new IllustratorSaveOptions;
  aiOption.pdfCompatible = true;
  aiOption.embedLinkedFiles = true;
  aiOption.embedICCProfile = true;
  aiOption.compressed = true;
  docObj.saveAs(new File(outputAiDirectory + handleName + ".ai"), aiOption);

  //閉じる
  docObj.close(SaveOptions.DONOTSAVECHANGES);
}

function getUserEntry(rankings, handleName){
  for (var i = 0; i < rankings.length; i++){
    if (rankings[i]["name"] == handleName){
      return rankings[i];
    }
  }
  return null;
}

function chikan(docObj, moto, saki){
  for (var i = 0; i < docObj.textFrames.length; i++) {
    var text = app.activeDocument.textFrames[i].contents;
    if (text !== null){
      app.activeDocument.textFrames[i].contents = text.replace(moto, saki);
    }
  }
}

function getRank(ranks_, rank){
  for (var i = 0; i < ranks_.length; i++){
    if (ranks_[i]["rank"] == rank){
      return ranks_[i];
    }
  }
}
