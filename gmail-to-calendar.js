const QUERY_YAMATO = 'subject:(受け取り日時変更依頼受付完了のお知らせ) ';
const QUERY_GNAVI = 'subject:(［ぐるなび］予約が確定しました) ';
const QUERY_YUBIN = 'subject:(日本郵便】受付完了のお知らせ)';
const QUERY_HOTPEPPER_BEAUTY = 'subject:ご予約が確定いたしました from:reserve@beauty.hotpepper.jp ';

function main() {
  pickUpMessage(QUERY_YAMATO, function (message) {
    parseYamato(message);
  });
  pickUpMessage(QUERY_GNAVI, function (message) {
     parseGnavi(message);
  });
  pickUpMessage(QUERY_YUBIN, function (message) {
    parseYubin(message);
  });
  pickUpMessage(QUERY_HOTPEPPER_BEAUTY, function (message) {
    parseHotpepperBeauty(message);
  });
}

function pickUpMessage(query, callback) {
  const messages = getMail(query);

  for (var i in messages) {
    for (var j in messages[i]) {
      const message = messages[i][j];
      // starは処理済みとする
      if (message.isStarred()) break;

      callback(message);

      message.star()
    }
  }
}

function getMail(query) {
  var threads = GmailApp.search(query, 0, 5);
  return GmailApp.getMessagesForThreads(threads);
}

function createEvent(title, description, location, year, month, dayOfMonth,
  startTimeHour, startTimeMinutes, endTimeHour, endTimeMinutes) {

  const calendar = CalendarApp.getDefaultCalendar();
  const startTime = new Date(year, month - 1, dayOfMonth, startTimeHour, startTimeMinutes, 0);
  const endTime = new Date(year, month - 1, dayOfMonth, endTimeHour, endTimeMinutes, 0);
  const option = {
    description: description,
    location: location,
  }

  console.log("start time: " + startTime);
  console.log("end time: " + endTime);
  calendar.createEvent(title, startTime, endTime, option);
}

// ヤマト運輸
function parseYamato(message) {
  const strDate = message.getDate();
  const strMessage = message.getPlainBody();

  const datePrefix = "■お受け取りご希望日時 ： ";
  const regexp = RegExp(datePrefix + '.*', 'gi');

  const result = strMessage.match(regexp);
  if (result == null) {
    console.log("This message doesn't have info.");
    return;
  }
  const parsedDate = result[0].replace(datePrefix, '');

  const year = new Date().getFullYear();
  const month = parsedDate.match(/[0-9]{2}\//gi)[0].replace('/', '');
  const dayOfMonth = parsedDate.match(/\/[0-9]{2}/gi)[0].replace('/', '');
  const matchedStartTimeHour = parsedDate.match(/[0-9]{2}時から/gi);
  const matchedEndTimeHour = parsedDate.match(/[0-9]{2}時まで/gi);

  var startTimeHour;
  var endTimeHour;
  if (matchedStartTimeHour == null || matchedEndTimeHour == null) {
    startTimeHour = '9';
    endTimeHour = '12';
  } else {
    startTimeHour = matchedStartTimeHour[0].replace('時から', '');
    endTimeHour = matchedEndTimeHour[0].replace('時まで', '');
  }

  createEvent("ヤマト配達", "mailDate: " + strDate,
    "", year, month, dayOfMonth, startTimeHour, 0, endTimeHour, 0);
}

// 日本郵便
function parseYubin(message) {

  const strDate = message.getDate();
  const strMessage = message.getPlainBody();

  const datePrefix = "【お届け予定日】";
  const dateSuffix = "【お届け希望時間帯】";
  const timeSuffix = "【ご希望配達先】";
  var regexp = RegExp(datePrefix + '[\\s\\S]*?' + dateSuffix, 'gi');

  const dateResult = strMessage.match(regexp);

  regexp = RegExp(dateSuffix + '[\\s\\S]*?' + timeSuffix, 'gi');

  const timeResult = strMessage.match(regexp);

  if (dateResult == null || timeResult == null) {
    console.log("This message doesn't have info.");
    return;
  }
  const parsedDate = dateResult[0].replace(datePrefix, '').replace(dateSuffix, '');
  const parsedTime = timeResult[0].replace(dateSuffix, '').replace(timeSuffix, '');

  const year = new Date().getFullYear();
  const month = parsedDate.match(/[0-9]*月/gi)[0].replace('月', '');
  const dayOfMonth = parsedDate.match(/[0-9]*日/gi)[0].replace('日', '');
  const matchedStartTimeHour = parsedTime.match(/[0-9]{2}～/gi);
  const matchedEndTimeHour = parsedTime.match(/[0-9]{2}時/gi);

  var startTimeHour;
  var endTimeHour;
  if (matchedStartTimeHour == null || matchedEndTimeHour == null) {
    startTimeHour = '9';
    endTimeHour = '12';
  } else {
    startTimeHour = matchedStartTimeHour[0].replace('～', '');
    endTimeHour = matchedEndTimeHour[0].replace('時', '');
  }

  createEvent("郵便配達", "mailDate: " + strDate,
    "", year, month, dayOfMonth, startTimeHour, 0, endTimeHour, 0);
}

// ぐるなび
function parseGnavi(message) {
  const strDate = message.getDate();
  const strMessage = message.getPlainBody();

  const suffix = "のご予約が確定しました。";
  const regexp = RegExp('.*' + suffix, 'gi');

  const result = strMessage.match(regexp);
  if (result == null) {
    console.log("This message doesn't have info.");
    return;
  }

  const baseStr = result[0].replace(suffix, '');

  const year = baseStr.match(/[0-9]{4}年/gi)[0].replace('年', '');
  const month = baseStr.match(/[0-9]{2}月/gi)[0].replace('月', '');
  const dayOfMonth = baseStr.match(/[0-9]{2}日/gi)[0].replace('日', '');
  const startTimeHour = baseStr.match(/[0-9]{2}時/gi)[0].replace('時', '');
  const startTimeMinute = baseStr.match(/[0-9]{2}分/gi)[0].replace('分', '');
  const title = baseStr.match(/「.*」/gi)[0].replace('「', '').replace('」', '');

  // 住所
  const addressPrefix = '- - - - - - - - - -\n';
  const addressRegexp = RegExp(addressPrefix + '.*', 'gi');
  const address = strMessage.match(addressRegexp)[2].replace(addressPrefix, '');

  createEvent(title, "mailDate: " + strDate, address,
    year, month, dayOfMonth, startTimeHour, startTimeMinute, startTimeHour, startTimeMinute);
}

// Hotpepper beauty
function parseHotpepperBeauty(message) {
  const strMessage = message.getPlainBody();

  const salonPrefix = "■サロン名";
  const regexpSalon = RegExp(salonPrefix + '[\\s\\S]*?' + '.+', 'gi');
  const salonName = strMessage.match(regexpSalon)[0]
    .replace(salonPrefix, '')
    .trim();

  const descriptionPrefix = "■メニュー";
  const couponPrefix = "■ご利用クーポン";
  const regexpDescription = RegExp(descriptionPrefix + '[\\s\\S]*?' + '(.+|[\\s\\S]*?)' + couponPrefix, 'gi');
  const description = strMessage.match(regexpDescription)[0]
    .replace(descriptionPrefix, '')
    .replace(couponPrefix, '')
    .trim();

  const locationPrefix = "※上の地図ページが開かない場合こちらの地図をご参照ください";
  const regexpLocation = RegExp(locationPrefix + '[\\s\\S]*?' + '.+', 'gi');
  const location = strMessage.match(regexpLocation)[0]
    .replace(locationPrefix, '')
    .trim();

  const datePrefix = "■来店日時";
  const dateRegex = "[0-9]{4}年(0[1-9]|1[0-2])月(0[1-9]|[12][0-9]|3[01])日（.）([01][0-9]|2[0-3]):[0-5][0-9]";
  const regexpDate = RegExp(datePrefix + '[\\s\\S]*?' + dateRegex, 'gi');
  const date = strMessage.match(regexpDate)[0]
    .replace(datePrefix, '')
    .trim();

  const year = date.match(/[0-9]{4}年/gi)[0].replace('年', '');
  const month = date.match(/(0[1-9]|1[0-2])月/gi)[0].replace('月', '');
  const dayOfMonth = date.match(/(0[1-9]|[12][0-9]|3[01])日/gi)[0].replace('日', '');
  const startTimeHour = date.match(/([01][0-9]|2[0-3]):/gi)[0].replace(':', '');
  const startTimeMinute = date.match(/:[0-5][0-9]/gi)[0].replace(':', '');

  createEvent(salonName, description,
    location, year, month, dayOfMonth, startTimeHour, startTimeMinute, parseInt(startTimeHour) + 1, startTimeMinute);
}
