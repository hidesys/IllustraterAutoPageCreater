//********
//環境変数の設定
//********

//ディレクトリの構造は以下を考えている
// C:\hoge~\作業量ディレクトリ\
//                            |- 入力AI\
//                            |- 出力AI\
//                            |- 出力PNG\
//                            |- 出力PDF\
var pwdPath = "C:\\Users\\washi\\Desktop\\個人ランキング\\作業用ディレクトリ\\";

//各Pathを用意
var inputAiDirectory = pwdPath + "入力AI\\";
var outputAiDirectory = pwdPath + "出力AI\\";
var outputPdfDirectory = pwdPath + "出力PDF\\";
var outputPngDirectory = pwdPath + "出力PNG\\";

//********
//ここからスクリプト本体
//******** 

//一覧を取得
var fileList = (new Folder(inputAiDirectory)).getFiles("*.ai");

//すでに出力されているものは除く
for (var i = 0; i < fileList.lenght; i++){
  var handleName = fileList[i].name.split(".")[0];
  if ((new File(outputPngDirectory + handleName  + ".png")).exists){
    fileList.splice(i--, 1);
  }
}

for (var i = 0; i < fileList.length; i++){
  //ファイルを開く
  var docObj = open(new File(fileList[i].fsName));

  //保存用に名前を取得
  var handleName = fileList[i].name.split(".")[0];

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
