// スプレッドシートの行列情報
var StartRow = 2;         // 開始行
var NameCol = 1;          // 名前列
var BirthdayCol = 2;      // 誕生日列

//------------------------------
// main処理
//------------------------------
function main() {
  // 今日の日付を取得
  var nowDate = new Date();
  
  // 今日の曜日を取得
  var todayDayOfWeek = nowDate.getDay();

  // 土曜、日曜は投稿しない
  if (isSatOrSun(todayDayOfWeek) === true) {
    return;
  }
  
  // 誕生日リストのスプレッドシートを参照
  var spreadsheet = SpreadsheetApp.openById('XXXXX');
  
  // シートにアクセス
  var sheet = spreadsheet.getSheetByName('誕生日');
  
  // データ取得
  var data = sheet.getDataRange().getValues();
  
  for (var i = 0; i < data.length; i++) {
    if (i < StartRow) {
      continue;
    }
    
    // 名前を取得
    var name = data[i][NameCol];
    
    // 誕生日を取得
    var birthday = data[i][BirthdayCol];
    
    // 誕生日を日付でパースし、日付が正しく入力されているかチェック
    var isDate = !isNaN(Date.parse(birthday));

    // 誕生日入力済みであることをチェック
    if (isDate === true) {
      // 投稿処理
      postProc(name, birthday, nowDate); 
    }
  }
}

//------------------------------
// 投稿処理
//------------------------------
function postProc(name, birthday, nowDate)
{
  var todayDayOfWeek = nowDate.getDay();
  var message = "";

  // 金曜日
  if (todayDayOfWeek === 5){
    
    // 1日後(土曜日)
    var dtSat = new Date();
    dtSat.setDate(dtSat.getDate() + 1);
    
    // 2日後(日曜日)
    var dtSun = new Date();
    dtSun.setDate(dtSun.getDate() + 2);
    
    if (isSend(dtSat, birthday) === true) {
      message = createMessage(name, "明日");
    }
    else if (isSend(dtSun, birthday) === true) {
      message = createMessage(name, "あさって");
    }
    else if (isSend(nowDate, birthday) === true) {
      message = createMessage(name, "今日");
    }
  }
  else
  {
    if (isSend(nowDate, birthday) === true) {
      message = createMessage(name, "今日");
    }      
  }

  // メッセージが生成されていれば、Slackに投稿  
  if (message != "") {
    // おめでとう画像を取得
    var fileId = getRandomImageID();
        
    // Slackへ投稿
    postSlack(message, fileId);
  }
}

//------------------------------
// 土曜日 or 日曜日判定
//------------------------------
function isSatOrSun(todayDayOfWeek)
{
  // 土曜、日曜は投稿しない
  if ((todayDayOfWeek === 0) || (todayDayOfWeek === 6)) {
    return true;
  }
  
  return false;
}

//------------------------------
// 送信有無取得
//------------------------------
function isSend(compareDate, birthdayDate)
{
  // 月、日が一致するかをチェックする
  if ((compareDate.getMonth() === birthdayDate.getMonth()) && 
    (compareDate.getDate() === birthdayDate.getDate()))
    {
      return true;
    }
  
  return false;
}

//------------------------------
// メッセージ作成
//------------------------------
function createMessage(name, dayMessage)
{
  var message = dayMessage + 'は, `' + name + '`さんの誕生日です :birthday:' + '\r\n' + 
    'おめでとうございます :tada: :tada: :tada:'
    
  return message;  
}

//------------------------------
// 誕生日おめでとう画像のIDを取得
//------------------------------
function getRandomImageID() {
  var imageFileArray = [
    "AAAAA",
    "BBBBB",
    "CCCCC",
    "DDDDD",
    "EEEEE",
    "FFFFF",
    "GGGGG",
    "HHHHH",
  ];

  // 0～7の値を取得
  var randomValue = Math.floor( Math.random() * 8 );
  
  return imageFileArray[randomValue];
}

//------------------------------
// Slackへ投稿
//------------------------------
function postSlack(text, fileId) {
  
  var payload = {
    text: text,
    attachments: [
        {
            color: "#2eb886",
            pretext: "",
            image_url: "http://drive.google.com/uc?export=download&id=" + fileId,
        }      
      ]
  };
 
  // Incoming WebhooksのURL
  var url = "YYYYY";
  
  var options = {
    "method" : "POST",
    "headers": {"Content-type": "application/json"},
    "payload": JSON.stringify(payload)
  };
  UrlFetchApp.fetch(url, options); 
}