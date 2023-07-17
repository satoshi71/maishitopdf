const folderID = '画像の入ったフォルダID';
const pdfFoloderID = 'PDFを出力するフォルダID';
const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

function createPDF(){
  let doc_Folder = DriveApp.getFolderById(folderID);
  let pdf_Folder = DriveApp.getFolderById(pdfFoloderID);
 
  files = doc_Folder.getFiles();
  while(files.hasNext()) {
    var buff = files.next();
    try {
      var doc  = DocumentApp.openById(buff.getId());
      let pdf_name = buff.getName() + '.pdf';
      pdf_Folder.createFile(doc.getAs("application/pdf")).setName(pdf_name);
      // console.log('OK: ' + buff.getName());
    }catch{
      // console.log(buff.getName());
      // console.log(e);
    }
  }
}

function rename(){
  const cname = sheet.getRange("G4").getValue();
  const pname = sheet.getRange("G5").getValue();
  const docID = sheet.getRange("G6").getValue();

  const filename = trim(cname) +'_'+ trim(pname);
  const file = DriveApp.getFileById(docID);
  file.setName(filename);
}

function trim(s){
  s = s.replace(' ', '');
  s = s.replace('　', '');
  return s;
}

function setName(){
  let name = sheet.getActiveCell().getValue();
  sheet.getRange("G5").setValue(name);
}

function getInfomation(){
  sheet.getRange("D5:D50").clearContent();
  sheet.getRange("G4:G6").clearContent();
  let docID = sheet.getActiveCell().getValue();
  if(docID=='') return;

  try{
    let doc = DocumentApp.openById(docID);
    let text = doc.getBody().getText();
    sheet.getRange("G6").setValue(docID);
    // console.log(docID);
    let lines = text.split("\n");
    for(let i=0; i<lines.length; i++){
      sheet.getRange(i+5,4).setValue(lines[i]);
      if(lines[i].indexOf('会社')>0){
        sheet.getRange("G4").setValue(lines[i]);
      }
    }
  }catch{}
}

function allClear(){
  sheet.getRange("D5:D50").clearContent();
  sheet.getRange("G4:G6").clearContent();
  sheet.getRange("B5:B22").clearContent();
}

function createDocument() {
  allClear();

  let files = DriveApp.getFolderById(folderID).getFiles();
  //Googleドキュメントに渡すオプション。OCR設定
  let option = {
    'ocr': true,        // OCRを行う
    'ocrLanguage': 'ja',// OCRを行う言語
  }
  let r=5;
  while(files.hasNext()){
    let file = files.next();
    //Googleドキュメントのファイル名＝画像ファイル名
    subject = file.getName();
    let resource = {
      title: subject
    };
    //画像をGoogleドキュメントで開いて文字起こしをする。
    let image = Drive.Files.copy(resource, file.getId(), option);
    //文字起こししたテキストを取得
    let doc = DocumentApp.openById(image.id);
    let body = doc.getBody();
    body.setFontSize(14);
    body.setForegroundColor('#000000');
    body.setFontFamily("メイリオ");
    body.setBold(false);

    // let text = DocumentApp.openById(image.id).getBody().getText();
    sheet.getRange(r,2).setValue(doc.getId());
    r++;

  }
}
