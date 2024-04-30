const workbook = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty("SPREADSHEET_ID"));
const detail_sheet = workbook.getSheetByName('detail');
let sheet_data_before = detail_sheet.getDataRange().getValues();
const about_sheet = workbook.getSheetByName('about');
let about_sheet_data = about_sheet.getDataRange().getValues();
const headers = sheet_data_before.shift();
const day_array = ['日', '月', '火', '水', '木', '金', '土'];
let header_array = new Array()
let sheet_data = sheet_data_before.map(row => {
  var obj = {};
  headers.forEach((header, index) => {
    obj[header] = row[index];
    header_array.push(header);
  });
  return obj;
});

for (var i = 0; i < sheet_data.length; i++) {
  if (sheet_data[i]['title'] == '') {
    sheet_data.splice(i, sheet_data.length - i);
    break;
  }
}

const token = PropertiesService.getScriptProperties().getProperty("LINE_NOTIFY_API_TOKEN");
const lineNotifyApi = 'https://notify-api.line.me/api/notify';


function messageText(day_name, i) {
  const data = new Date(Number(sheet_data[i]['next_year']), Number(sheet_data[i]['next_month']) - 1, Number(sheet_data[i]['next_day']));
  let text = '\n\n' + sheet_data[i]['title'] + 'は' + day_name + '¥' + sheet_data[i]['price'] + 'での更新です。\n' + 
                'この請求は' + sheet_data[i]['interval'] + sheet_data[i]['unit'] + '毎です。\n\n' +
                '名称: ' + sheet_data[i]['title'] + '\n' +
                '金額: ¥' + sheet_data[i]['price'] + '\n' +
                '請求予定日: ' + sheet_data[i]['next_year'] + '/' + sheet_data[i]['next_month'] + '/' + sheet_data[i]['next_day'] + ' (' + day_array[data.getDay()] + ')';
  if (sheet_data[i]['payment_method']) {
    text += '\n' + '支払い方法: ' + sheet_data[i]['payment_method']
  }
  return text;
}


function notifyToday() {

  const today = new Date();

  for (var i = 0; i < sheet_data.length; i++) {
    if (sheet_data[i]['stop'] === false && sheet_data[i]['next_year'] === today.getFullYear() && sheet_data[i]['next_month'] === today.getMonth() + 1 && sheet_data[i]['next_day'] === today.getDate()) {
      var message_text = messageText('今日', i)
      var options = {"method"  : "post", "payload" : {"message": message_text}, "headers" : {"Authorization":"Bearer " + token}};
      UrlFetchApp.fetch(lineNotifyApi, options);

      var next_date = new Date();

      if (sheet_data[i]['unit'] === '日') {
        next_date.setDate(next_date.getDate() + Number(sheet_data[i]['interval']));
      } else if (sheet_data[i]['unit'] === '週') {
        next_date.setDate(next_date.getDate() + (Number(sheet_data[i]['interval'] * 7)));
      } else if (sheet_data[i]['unit'] === '月') {
        next_date.setMonth(next_date.getMonth() + Number(sheet_data[i]['interval']));
      } else if (sheet_data[i]['unit'] === '年') {
        next_date.setFullYear(next_date.getFullYear() + Number(sheet_data[i]['interval']));
      }

      var sheet_date = new Array()
      sheet_date.push(next_date.getFullYear());
      sheet_date.push(next_date.getMonth() + 1);
      sheet_date.push(next_date.getDate());

      detail_sheet.getRange(i + 2, header_array.indexOf('next_year') + 1, 1, 3).setValues([sheet_date]);
    }
  }
}


function notifyBefore3Days() {

  let today = new Date();
  today.setDate(today.getDate() + 3);

  for (var i = 0; i < sheet_data.length; i++) {
    if (sheet_data[i]['stop'] === false && sheet_data[i]['next_year'] === today.getFullYear() && sheet_data[i]['next_month'] === today.getMonth() + 1 && sheet_data[i]['next_day'] === today.getDate()) {
      var message_text = messageText('3日後', i)
      var options = {"method"  : "post", "payload" : {"message": message_text}, "headers" : {"Authorization":"Bearer " + token}};
      UrlFetchApp.fetch(lineNotifyApi, options);
    }
  }
}


function divideAnnualAmountBy12() {

  var message_text = "\n\n月額支払いを除いた来月の支払額は¥" + Math.ceil(Number(about_sheet_data[4][1])) + "です。"

  var options = {"method"  : "post", "payload" : {"message": message_text}, "headers" : {"Authorization":"Bearer " + token}};
  UrlFetchApp.fetch(lineNotifyApi, options);
}


function deleteTrigger() {
  
  const triggers = ScriptApp.getProjectTriggers();

  for (var i = 1; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
}


function setTriggerAbout() {

  
}


function setTrigger() {

  deleteTrigger();
  
  let timeBefore3Days = new Date();
  timeBefore3Days.setHours(18);
  timeBefore3Days.setMinutes(0);
  timeBefore3Days.setSeconds(0);
  timeBefore3Days.setMilliseconds(0);
  let timeToday = new Date();
  timeToday.setHours(8);
  timeToday.setMinutes(0);
  timeToday.setSeconds(0);
  timeToday.setMilliseconds(0);
  ScriptApp.newTrigger('notifyBefore3Days').timeBased().at(timeBefore3Days).create();
  ScriptApp.newTrigger('notifyToday').timeBased().at(timeToday).create();
}
