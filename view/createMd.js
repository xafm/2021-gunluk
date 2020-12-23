const ExcelJS = require('exceljs')
const json2md = require('json2md')
const fs = require('fs')
const path = require('path')
const workbook = new ExcelJS.Workbook()
const excelPath = path.join(__dirname, '../', 'data.xlsx')

let activityDefinitions = []
let dailyActivities = []

exports.createMd = async function () {
  activityDefinitions = []
  dailyActivities = []

  try {
    await workbook.xlsx.readFile(excelPath)
  } catch (error) {
    throw new Error(`Excel dosyası ${excelPath} dizininden okunamadı`)
  }

  let ws = workbook.worksheets[0]
  if (!ws) {
    throw new Error(
      `Excel dosyasındaki 1. sayfa okunamadı (Aktivite Tanımları)`,
    )
  }

  let activityDefinitionsTemp = []

  ws.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return

    const newActivityDefinition = {}
    row.eachCell((cell, colNumber) => {
      switch (colNumber) {
        case 1:
          newActivityDefinition.code = cell.value
          newActivityDefinition.rowNumber = rowNumber
          break
        case 2:
          newActivityDefinition.name = cell.value
          newActivityDefinition.rowNumber = rowNumber
          break
        case 3:
          newActivityDefinition.unit = cell.value
          newActivityDefinition.rowNumber = rowNumber
          break
        default:
          break
      }
    })
    activityDefinitionsTemp.push(newActivityDefinition)
  })

  if (!activityDefinitionsTemp.length) {
    throw new Error(
      `Excel dosyasında "Aktivite Tanımları" sayfasını doldurmalısınız`,
    )
  }

  activityDefinitionsTemp.forEach(row => {
    if (!row.code) {
      throw new Error(
        `Aktivite kodunu girin. Aktivite tanımları satır: ${row.rowNumber}`,
      )
    }

    // if (!row.name) {
    //   throw new Error(
    //     `Aktivite tanımını girin. Aktivite tanımları satır: ${row.rowNumber}`,
    //   )
    // }

    // if (!row.unit) {
    //   throw new Error(
    //     `Aktivite birimini girin. Aktivite tanımları satır: ${row.rowNumber}`,
    //   )
    // }

    if (activityDefinitions.find(item => item.code === row.code)) {
      throw new Error(
        `Aynı aktivite kodu 2 defa kullanılamaz. Aktivite tanımları satır: ${row.rowNumber} (${row.code})`,
      )
    }

    if (activityDefinitions.find(item => item.name === row.name)) {
      throw new Error(
        `Aynı aktivite tanımı 2 defa kullanılamaz. Aktivite tanımları satır: ${row.rowNumber} (${row.name})`,
      )
    }

    activityDefinitions.push({
      code: row.code,
      name: row.name,
      unit: row.unit,
    })
  })
  activityDefinitionsTemp = null

  // Günlük Aktiviteler
  ws = workbook.worksheets[1]
  if (!ws) {
    throw new Error(
      `Excel dosyasındaki 2. sayfa okunamadı (Günlük Aktiviteler)`,
    )
  }

  let dailyActivitiesTemp = []
  ws.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return

    const newDailyActivity = {}
    row.eachCell((cell, colNumber) => {
      switch (colNumber) {
        case 1:
          if (!isDateValid(cell.value)) {
            throw new Error(
              `Hatalı tarih formatı (Günlük Aktiviteler satır: ${rowNumber}).`,
            )
          }
          cell.value = new Date(String(cell.value))
          newDailyActivity.date = cell.value
          newDailyActivity.month = cell.value.getMonth()
          newDailyActivity.weekNumber = cell.value.getWeek()
          newDailyActivity.formattedDate = formatDate(cell.value)
          newDailyActivity.rowNumber = rowNumber
          break
        case 2:
          newDailyActivity.code = cell.value
          newDailyActivity.rowNumber = rowNumber
          break
        case 3:
          newDailyActivity.text = cell.value
          newDailyActivity.rowNumber = rowNumber
          break
        case 4:
          newDailyActivity.count = cell.value
          newDailyActivity.rowNumber = rowNumber
          break
        default:
          break
      }
    })
    dailyActivitiesTemp.push(newDailyActivity)
  })

  if (!dailyActivitiesTemp.length) {
    throw new Error(
      `Excel dosyasında "Günlük Aktiviteler" sayfasını doldurmalısınız`,
    )
  }

  dailyActivitiesTemp.forEach(row => {
    if (!row.code) {
      throw new Error(
        `Aktivite kodunu girin. (Günlük Aktiviteler, satır: ${row.rowNumber}).`,
      )
    }

    let activityDefinition = activityDefinitions.find(
      actDef => actDef.code === row.code,
    )

    if (!activityDefinition) {
      throw new Error(
        `Aktivite kodu "${row.code}" tanımlanmamış. Aktivite Tanımları sayfasında tanımlayın (Günlük Aktiviteler, satır: ${row.rowNumber}).`,
      )
    }

    dailyActivities.push({
      rowNumber: row.rowNumber,
      date: row.date,
      month: row.month,
      weekNumber: row.weekNumber,
      formattedDate: row.formattedDate,
      code: row.code,
      activityDefinition: activityDefinition.name,
      unit: activityDefinition.unit,
      text: row.text,
      count: row.count,
    })
  })

  const md = []
  md.push({h1: '2021'})

  let previousRecordDay
  let previousRecordWeekNumber
  let previousRecordMonthNumber

  let monthMD = []
  let dailyActivitiesTexts = []
  let weeklyTotal = {}
  let monthlyTotal = {}
  let yearlyTotal = {}

  previousRecordDay = dailyActivities[0].date
  previousRecordWeekNumber = dailyActivities[0].weekNumber
  previousRecordMonthNumber = dailyActivities[0].month
  monthMD.push({h1: monthNames[previousRecordMonthNumber]})

  dailyActivities.forEach(dailyActivity => {
    // Gün değiştiyse, günün verilerini tutan (önceki günü tutuyor olacak) array'i yazdır ardından array'i temizle
    // Hafta değiştiyse, hafta toplam verilerini tutan array'i yazdır. Ardından array'i temizle
    // Ay değiştiyse, ay toplamını tutan array'i yazdır array'i yazdır ve array'i temizle. Ardından yeni ayın başlığını at

    if (
      previousRecordDay.toLocaleString() !== dailyActivity.date.toLocaleString()
    ) {
      writeDay()
    }

    if (previousRecordWeekNumber !== dailyActivity.weekNumber) {
      // önceki hafta tamamlandı yazdır obje sıfırla
      writeWeeklyTotals()
    }

    if (previousRecordMonthNumber !== dailyActivity.month) {
      // önceki ay tamamlandı toplamı yazdır obje sıfırla
      writeMonthlyTotals()

      // yeni ay ocak değilse yeni ayın başlığını at
      // if (!yeniyıl) {
      monthMD.push({h1: monthNames[dailyActivity.month]})
      // }
    }

    // if (yeniyıl) {
    //   // eski yıl toplamı
    //   // yeni  yıl başlığı
    //   monthMD.push({h1: monthNames[dailyActivity.month]})
    // }

    // Aktivite textleri doldurulur
    // Daha sonra aylık mdnin içerisine atılacak
    // bir gün = {blockquote: [day.date, {ul: [...activities]}]}
    let activityText = buildDayText(dailyActivity)
    dailyActivitiesTexts.push(activityText)

    if (weeklyTotal[dailyActivity.code]) {
      weeklyTotal[dailyActivity.code] += dailyActivity.count
    } else {
      weeklyTotal[dailyActivity.code] = dailyActivity.count
    }

    if (monthlyTotal[dailyActivity.code]) {
      monthlyTotal[dailyActivity.code] += dailyActivity.count
    } else {
      monthlyTotal[dailyActivity.code] = dailyActivity.count
    }

    previousRecordDay = dailyActivity.date
    previousRecordWeekNumber = dailyActivity.weekNumber
    previousRecordMonthNumber = dailyActivity.month
  })

  // gün hafta ay yıl yaz
  writeDay()
  writeWeeklyTotals()
  writeMonthlyTotals()

  writeMDFile(md)

  function writeDay() {
    monthMD.push({
      blockquote: [
        formatDate(previousRecordDay),
        {ul: [...dailyActivitiesTexts]},
      ],
    })
    dailyActivitiesTexts = []
  }

  function writeWeeklyTotals() {
    monthMD.push({h2: `&nbsp; ${previousRecordWeekNumber + 1}. hafta toplamı`})
    let totalsTexts = []
    for (const act in weeklyTotal) {
      let actDefinition = activityDefinitions.find(
        actDef => actDef.code === act,
      )
      totalsTexts.push(
        `${actDefinition.name}: ${weeklyTotal[act]} ${actDefinition.unit}`,
      )
    }
    monthMD.push({ul: [...totalsTexts]})
    weeklyTotal = {}
  }

  function writeMonthlyTotals() {
    monthMD.push({
      h2: `&nbsp; ${monthNames[previousRecordMonthNumber]} ayı toplamı`,
    })
    let totalsTexts = []
    for (const act in monthlyTotal) {
      let actDefinition = activityDefinitions.find(
        actDef => actDef.code === act,
      )
      totalsTexts.push(
        `${actDefinition.name}: ${monthlyTotal[act]} ${actDefinition.unit}`,
      )
    }
    monthMD.push({ul: [...totalsTexts]})
    monthlyTotal = {}
    md.push({blockquote: [...monthMD, '&nbsp;']})
    md.push('&nbsp;')
    monthMD = []
  }
}

function buildDayText(dailyActivity) {
  let actCount = dailyActivity.count
    ? ` / (${dailyActivity.count} ${dailyActivity.unit})`
    : ''
  if (!dailyActivity.text) {
    actCount = actCount.replace(' / ', '')
  }

  let ret = `${dailyActivity.activityDefinition}: `
  if (dailyActivity.text) {
    ret += dailyActivity.text
  }

  if (actCount) {
    ret += actCount
  }

  return ret
}

function formatDate(date) {
  if (!date) return
  date = new Date(date)
  let month = date.getMonth() + 1
  month = month < 10 ? '0' + month : month
  let day = date.getDate() < 10 ? '0' + date.getDate() : date.getDate()

  let dayOfWeekText = daysOfWeek[date.getDay()]

  return `${day}.${month}.${date.getFullYear()} ${dayOfWeekText}`
}

async function writeMDFile(arrMd) {
  const md = json2md(arrMd)

  const search = '>  -'
  const replaceWith = '> *'

  const replacedMd = md.split(search).join(replaceWith)

  try {
    const mdPath = path.join(__dirname, '../', '2021.md')
    await fs.writeFileSync(mdPath, replacedMd, e => console.log(e))

    activityDefinitions = []
    dailyActivities = []
  } catch (error) {
    console.log(error.message)
  }
}

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

var isDateValid = (...val) => !Number.isNaN(new Date(...val).valueOf())

var monthNames = {
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

var daysOfWeek = {
  0: 'Pazar',
  1: 'Pazartesi',
  2: 'Salı',
  3: 'Çarşamba',
  4: 'Perşembe',
  5: 'Cuma',
  6: 'Cumartesi',
}
