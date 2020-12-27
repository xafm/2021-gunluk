Date.prototype.getWeek = function (dowOffset) {
  dowOffset = typeof dowOffset == 'int' ? dowOffset : 1 //default dowOffset to zero
  var newYear = new Date(this.getFullYear(), 0, 1)
  var day = newYear.getDay() - dowOffset //the day of week the year begins on
  day = day >= 0 ? day : day + 7
  var daynum =
    Math.floor(
      (this.getTime() -
        newYear.getTime() -
        (this.getTimezoneOffset() - newYear.getTimezoneOffset()) * 60000) /
        86400000,
    ) + 1
  var weeknum
  //if the year starts before the middle of a week
  if (day < 4) {
    weeknum = Math.floor((daynum + day - 1) / 7) + 1
    if (weeknum > 52) {
      nYear = new Date(this.getFullYear() + 1, 0, 1)
      nday = nYear.getDay() - dowOffset
      nday = nday >= 0 ? nday : nday + 7
      /*if the next year starts before the middle of
                the week, it is week #1 of that year*/
      weeknum = nday < 4 ? 1 : 53
    }
  } else {
    weeknum = Math.floor((daynum + day - 1) / 7)
  }
  return weeknum
}

const isDateValid = (...val) => !Number.isNaN(new Date(...val).valueOf())

const formatDate = date => {
  if (!date) return
  date = new Date(date)
  let month = date.getMonth() + 1
  month = month < 10 ? '0' + month : month
  let day = date.getDate() < 10 ? '0' + date.getDate() : date.getDate()

  let dayOfWeekText = daysOfWeek[date.getDay()]

  return `${day}.${month}.${date.getFullYear()} ${dayOfWeekText}`
}

const monthNames = {
  0: 'Ocak',
  1: 'Şubat',
  2: 'Mart',
  3: 'Nisan',
  4: 'Mayıs',
  5: 'Haziran',
  6: 'Temmuz',
  7: 'Ağustos',
  8: 'Eylül',
  9: 'Ekim',
  10: 'Kasım',
  11: 'Aralık',
}

const daysOfWeek = {
  0: 'Pazar',
  1: 'Pazartesi',
  2: 'Salı',
  3: 'Çarşamba',
  4: 'Perşembe',
  5: 'Cuma',
  6: 'Cumartesi',
}

const numberThousandSeperator = number => {
  return number.toString().replace(/(\d)(?=(\d{3})+(?!\d))/g, '$1,');
};

module.exports = {
  isDateValid,
  formatDate,
  numberThousandSeperator,
  monthNames,
  daysOfWeek
}

