/**
 * Utils.gs - 日付・時間ユーティリティ、営業日判定
 */

/**
 * 日付と時間文字列を結合して Date を返す
 * @param {Date} date - 基準日
 * @param {string|Date} timeInput - "09:00" 形式の文字列または Date（時刻部分のみ使用）
 * @returns {Date}
 */
function combineDateAndTime(date, timeInput) {
  var d = new Date(date.getTime());
  d.setHours(0, 0, 0, 0);
  var hours = 0;
  var minutes = 0;
  if (typeof timeInput === 'string') {
    var parts = timeInput.trim().match(/^(\d{1,2}):(\d{2})/);
    if (parts) {
      hours = parseInt(parts[1], 10);
      minutes = parseInt(parts[2], 10);
    }
  } else if (timeInput && timeInput.getHours !== undefined) {
    hours = timeInput.getHours();
    minutes = timeInput.getMinutes();
  }
  d.setHours(hours, minutes, 0, 0);
  return d;
}

/**
 * 指定分を加算した日時を返す
 * @param {Date} date - 基準日時
 * @param {number} minutes - 加算する分数
 * @returns {Date}
 */
function addMinutes(date, minutes) {
  return new Date(date.getTime() + minutes * 60 * 1000);
}

/**
 * 営業日かどうか（土日・祝日を設定に応じて除外）
 * @param {Date} date - 判定する日付
 * @param {Object} [settings] - getSettings() の戻り値。省略時は土日のみ判定
 * @returns {boolean}
 */
function isBusinessDay(date, settings) {
  var day = date.getDay();
  if (settings && settings.excludeWeekends) {
    if (day === 0 || day === 6) return false;
  } else if (!settings) {
    if (day === 0 || day === 6) return false;
  }
  if (settings && settings.excludeHolidays) {
    if (isJapaneseHoliday(date)) return false;
  }
  return true;
}

/**
 * 日本の祝日かどうか（静的リストで判定）
 * @param {Date} date
 * @returns {boolean}
 */
function isJapaneseHoliday(date) {
  var y = date.getFullYear();
  var m = date.getMonth() + 1;
  var d = date.getDate();
  var key = m + '/' + d;
  var holidays = getJapaneseHolidaysForYear(y);
  return holidays.indexOf(key) !== -1;
}

/**
 * 指定年の祝日リスト（MM/DD 形式）
 * @param {number} year
 * @returns {string[]}
 */
function getJapaneseHolidaysForYear(year) {
  var list = [];
  list.push('1/1');
  list.push('1/2');
  list.push('1/3');
  var secondMondayJan = getNthWeekday(year, 1, 1, 2);
  if (secondMondayJan) list.push('1/' + secondMondayJan);
  list.push('2/11');
  list.push('2/23');
  var dayOfVernalEquinox = getVernalEquinoxDay(year);
  if (dayOfVernalEquinox) list.push('3/' + dayOfVernalEquinox);
  list.push('4/29');
  list.push('5/3');
  list.push('5/4');
  list.push('5/5');
  var thirdMondayJul = getNthWeekday(year, 7, 1, 3);
  if (thirdMondayJul) list.push('7/' + thirdMondayJul);
  list.push('8/11');
  var thirdMondaySep = getNthWeekday(year, 9, 1, 3);
  if (thirdMondaySep) list.push('9/' + thirdMondaySep);
  var dayOfAutumnalEquinox = getAutumnalEquinoxDay(year);
  if (dayOfAutumnalEquinox) list.push('9/' + dayOfAutumnalEquinox);
  var secondMondayOct = getNthWeekday(year, 10, 1, 2);
  if (secondMondayOct) list.push('10/' + secondMondayOct);
  list.push('11/3');
  list.push('11/23');
  list.push('12/23');
  return list;
}

function getNthWeekday(year, month, weekday, n) {
  var count = 0;
  var date = new Date(year, month - 1, 1);
  while (date.getMonth() === month - 1) {
    if (date.getDay() === weekday) {
      count++;
      if (count === n) return date.getDate();
    }
    date.setDate(date.getDate() + 1);
  }
  return null;
}

function getVernalEquinoxDay(year) {
  if (year <= 2099) return Math.floor(20.8431 + 0.242194 * (year - 1980) - Math.floor((year - 1980) / 4));
  return 20;
}

function getAutumnalEquinoxDay(year) {
  if (year <= 2099) return Math.floor(23.2488 + 0.242194 * (year - 1980) - Math.floor((year - 1980) / 4));
  return 23;
}
