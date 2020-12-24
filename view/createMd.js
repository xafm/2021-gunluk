const ExcelJS = require('exceljs')
const json2md = require('json2md')
const fs = require('fs')
const path = require('path')
const workbook = new ExcelJS.Workbook()
const excelPath = path.join(__dirname, '../', 'data.xlsx')
const {isDateValid, formatDate, monthNames, daysOfWeek} = require('./helper')

let activityDefinitions = []
let dailyActivities = []

exports.createMd = async function () {
  activityDefinitions = []
  dailyActivities = []

  try {
    await parseExcel()
    const md = prepareMd()
    await writeMDFile(md)
  } catch (error) {
    throw new Error(error.message)
  }
}

async function parseExcel() {
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
          newDailyActivity.year = cell.value.getFullYear()
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
    if (!row.date) {
      throw new Error(
        `Tarih girin. (Günlük Aktiviteler satır: ${row.rowNumber}).`,
      )
    }

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
      year: row.year,
      formattedDate: row.formattedDate,
      code: row.code,
      activityDefinition: activityDefinition.name,
      unit: activityDefinition.unit,
      text: row.text,
      count: row.count,
    })
  })
}

function prepareMd() {
  const md = []

  let previousRecordDay
  let previousRecordWeekNumber
  let previousRecordMonthNumber
  let previousRecordYear

  let monthMD = []
  let dailyActivitiesTexts = []
  let weeklyTotal = {}
  let monthlyTotal = {}
  let yearlyTotal = {}

  previousRecordDay = dailyActivities[0].date
  previousRecordWeekNumber = dailyActivities[0].weekNumber
  previousRecordMonthNumber = dailyActivities[0].month
  previousRecordYear = dailyActivities[0].year

  md.push({h1: previousRecordYear})
  monthMD.push({h1: monthNames[previousRecordMonthNumber]})

  dailyActivities.forEach(dailyActivity => {
    // Gün değiştiyse, günün verilerini tutan (önceki günü tutuyor olacak) array'i yazdır ardından array'i temizle
    // Hafta değiştiyse, hafta toplam verilerini tutan array'i yazdır. Ardından array'i temizle
    // Ay değiştiyse, ay toplamını tutan array'i yazdır array'i yazdır ve array'i temizle. Ardından yeni ayın başlığını at
    console.log(previousRecordYear)
    if (
      previousRecordDay.toLocaleString() !== dailyActivity.date.toLocaleString()
    ) {
      writeDay()
    }

    if (previousRecordWeekNumber !== dailyActivity.weekNumber) {
      // önceki hafta tamamlandı, toplamı yazdır. weeklyTotal objesini sıfırla
      writeWeeklyTotals()
    }

    if (previousRecordMonthNumber !== dailyActivity.month) {
      // önceki ay tamamlandı, toplamı yazdır. monthlyTotal objesini sıfırla
      writeMonthlyTotals()

      // yıl değişmediyse yeni ayı yazdır
      if (previousRecordYear === dailyActivity.year) {
        monthMD.push({h1: monthNames[dailyActivity.month]})
      }
    }

    if (previousRecordYear !== dailyActivity.year) {
      // eski yıl toplamı
      // yeni  yıl başlığı
      writeYearlyTotals()
      md.push({h1: dailyActivity.year})
      monthMD.push({h1: monthNames[dailyActivity.month]})
    }

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

    if (yearlyTotal[dailyActivity.code]) {
      yearlyTotal[dailyActivity.code] += dailyActivity.count
    } else {
      yearlyTotal[dailyActivity.code] = dailyActivity.count
    }

    previousRecordDay = dailyActivity.date
    previousRecordWeekNumber = dailyActivity.weekNumber
    previousRecordMonthNumber = dailyActivity.month
    previousRecordYear = dailyActivity.year
  })

  // gün hafta ay yıl yaz
  writeDay()
  writeWeeklyTotals()
  writeMonthlyTotals()
  writeYearlyTotals()

  return md

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

  function writeYearlyTotals() {
    // md.push({
    //   h2: `&nbsp; ${previousRecordYear} yılı toplamı 🎉`,
    // })
    let totalsTexts = []
    for (const act in yearlyTotal) {
      let actDefinition = activityDefinitions.find(
        actDef => actDef.code === act,
      )
      totalsTexts.push(
        `${actDefinition.name}: ${yearlyTotal[act]} ${actDefinition.unit}`,
      )
    }
    // md.push({ul: [...totalsTexts]})
    yearlyTotal = {}
    md.push({
      blockquote: [
        {h2: `&nbsp; ${previousRecordYear} yılı toplamı 🎉`},
        {ul: [...totalsTexts]},
        // '&nbsp;',
      ],
    })
    md.push('&nbsp;')
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
