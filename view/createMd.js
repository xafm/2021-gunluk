const ExcelJS = require('exceljs');
const json2md = require('json2md');
const fs = require('fs');
const path = require('path');

let activityDefinitions = [];
let dailyActivities = [];

exports.createMd = async function () {
  const workbook = new ExcelJS.Workbook();
  const excelPath = path.join(__dirname, '../', 'data.xlsx')
  console.log(excelPath)
  await workbook.xlsx.readFile(excelPath)
  // try {
  //   console.log('amına koymaya geliyorum')
  //   console.log(__dirname)
  //   await workbook.xlsx.readFile('../../../data.xlsx');
  //   console.log(3)
  // } catch (error) {
  //   try {
  //     await workbook.xlsx.readFile('../../data.xlsx');
  //     console.log(2)
  //   } catch (error) {
  //     try {
  //       await workbook.xlsx.readFile('../data.xlsx');
  //       console.log(1)
  //     } catch (error) {
  //       try {
  //         await workbook.xlsx.readFile('../../../../data.xlsx');
  //         console.log(4)
  //       } catch (error) {
  //         try {
  //           await workbook.xlsx.readFile('/data.xlsx');
  //           console.log(4)
  //         } catch (error) {
  //           console.log('anasını sikim')
  //         }
  //       }
  //     }
  //   }
  // }
  let ws = workbook.worksheets[0];
  ws.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;

    const newActivityDefinition = {};
    row.eachCell((cell, colNumber) => {
      switch (colNumber) {
        case 1:
          newActivityDefinition.code = cell.value;
          break;
        case 2:
          newActivityDefinition.name = cell.value;
          break;
        case 3:
          newActivityDefinition.unit = cell.value;
          break;
        default:
          break;
      }
    });
    activityDefinitions.push(newActivityDefinition);
  });

  ws = workbook.worksheets[1];
  ws.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;

    const newDailyActivity = {};
    row.eachCell((cell, colNumber) => {
      switch (colNumber) {
        case 1:
          newDailyActivity.date = cell.value;
          newDailyActivity.month = cell.value.getMonth();
          newDailyActivity.weekNumber = cell.value.getWeek();
          newDailyActivity.formattedDate = formatDate(cell.value);
          break;
        case 2:
          newDailyActivity.code = cell.value;
          break;
        case 3:
          newDailyActivity.text = cell.value;
          break;
        case 4:
          newDailyActivity.count = cell.value;
          break;
        default:
          break;
      }
    });
    dailyActivities.push(newDailyActivity);
  });

  const md = [];
  md.push({h1: '2021'});

  let previousRecordDay;
  let previousRecordWeekNumber;
  let previousRecordMonthNumber;

  let monthMD = [];
  let dailyActivitiesTexts = [];
  let weeklyTotal = {};
  let monthlyTotal = {};
  dailyActivities.forEach(dailyActivity => {
    // gün geçtiyse arrayi yazdır arrayi temizle
    // hafta geçtiyse arrayi yazdır arrayi temizle
    // ay geçtiyse geçtiyse arrayi yazdır arrayi temizle, yeni ayın başlığını at

    if (previousRecordDay === undefined) {
      previousRecordDay = dailyActivity.date;
    }

    if (previousRecordWeekNumber === undefined) {
      previousRecordWeekNumber = dailyActivity.weekNumber;
    }

    // ilk kayda özel ay başlığı atılsın
    if (previousRecordMonthNumber === undefined) {
      monthMD.push({h1: monthNames[dailyActivity.month]});
      previousRecordMonthNumber = dailyActivity.month;
    }

    if (
      previousRecordDay.toLocaleString() !== dailyActivity.date.toLocaleString()
    ) {
      // önceki gün tamamlandı yazdır array sıfırla
      monthMD.push({
        blockquote: [
          formatDate(previousRecordDay),
          {ul: [...dailyActivitiesTexts]},
        ],
      });
      dailyActivitiesTexts = [];
    }

    if (previousRecordWeekNumber !== dailyActivity.weekNumber) {
      // önceki hafta tamamlandı yazdır obje sıfırla
      writeWeekTotals();
    }

    if (previousRecordMonthNumber !== dailyActivity.month) {
      // önceki ay tamamlandı toplamı yazdır obje sıfırla
      writeMonthlyTotals();
      md.push({blockquote: [...monthMD, '&nbsp;']});
      md.push('&nbsp;');
      monthMD = [];

      // yeni ayın başlığını at
      monthMD.push({h1: monthNames[dailyActivity.month]});
    }

    // Aktivite textleri doldurulur
    // Daha sonra aylık mdnin içerisine atılacak
    // bir gün = {blockquote: [day.date, {ul: [...activities]}]}
    let activityText = buildDayText(dailyActivity);
    dailyActivitiesTexts.push(activityText);

    if (weeklyTotal[dailyActivity.code]) {
      weeklyTotal[dailyActivity.code] += dailyActivity.count;
    } else {
      weeklyTotal[dailyActivity.code] = dailyActivity.count;
    }

    if (monthlyTotal[dailyActivity.code]) {
      monthlyTotal[dailyActivity.code] += dailyActivity.count;
    } else {
      monthlyTotal[dailyActivity.code] = dailyActivity.count;
    }

    previousRecordDay = dailyActivity.date;
    previousRecordWeekNumber = dailyActivity.weekNumber;
    previousRecordMonthNumber = dailyActivity.month;
  });

  monthMD.push({
    blockquote: [
      formatDate(previousRecordDay),
      {ul: [...dailyActivitiesTexts]},
    ],
  });

  // gün hafta ay yıl yaz
  writeWeekTotals();
  writeMonthlyTotals();
  md.push({blockquote: [...monthMD, '&nbsp;']});

  writeMDFile(md);

  function writeWeekTotals() {
    monthMD.push({h2: `&nbsp; ${previousRecordWeekNumber + 1}. hafta toplamı`});
    let totalsTexts = [];
    for (const act in weeklyTotal) {
      let actDefinition = activityDefinitions.find(
        actDef => actDef.code === act,
      );
      totalsTexts.push(
        `${actDefinition.name}: ${weeklyTotal[act]} ${actDefinition.unit}`,
      );
    }
    monthMD.push({ul: [...totalsTexts]});
    weeklyTotal = {};
  }

  function writeMonthlyTotals() {
    monthMD.push({
      h2: `&nbsp; ${monthNames[previousRecordMonthNumber]} ayı toplamı`,
    });
    let totalsTexts = [];
    for (const act in monthlyTotal) {
      let actDefinition = activityDefinitions.find(
        actDef => actDef.code === act,
      );
      totalsTexts.push(
        `${actDefinition.name}: ${monthlyTotal[act]} ${actDefinition.unit}`,
      );
    }
    monthMD.push({ul: [...totalsTexts]});
    monthlyTotal = {};
  }
};

function buildDayText(dailyActivity) {
  let activityDefinition = activityDefinitions.find(
    actDef => actDef.code === dailyActivity.code,
  );

  let actCount = dailyActivity.count
    ? ` / (${dailyActivity.count} ${activityDefinition.unit})`
    : '';
  if (!dailyActivity.text) {
    actCount = actCount.replace(' / ', '');
  }

  let ret = `${activityDefinition.name}: `;
  if (dailyActivity.text) {
    ret += dailyActivity.text;
  }

  if (actCount) {
    ret += actCount;
  }

  return ret;
}

function formatDate(date) {
  if (!date) return;
  date = new Date(date);
  let month = date.getMonth() + 1;
  month = month < 10 ? '0' + month : month;
  let day = date.getDate() < 10 ? '0' + date.getDate() : date.getDate();

  let dayOfWeekText = daysOfWeek[date.getDay()];

  return `${day}.${month}.${date.getFullYear()} ${dayOfWeekText}`;
}

async function writeMDFile(arrMd) {
  const md = json2md(arrMd);

  const search = '>  -';
  const replaceWith = '> *';

  const replacedMd = md.split(search).join(replaceWith);

  try {
    await fs.writeFileSync('2021.md', replacedMd, e => console.log(e));

    activityDefinitions = [];
    dailyActivities = [];

    // converter = new showdown.Converter();
    // const html = converter.makeHtml(replacedMd)
    // await fs.writeFile('../2021.html', html, (e) => console.log(e))
  } catch (error) {
    console.log(error.message);
  }
}

Date.prototype.getWeek = function (dowOffset) {
  dowOffset = typeof dowOffset == 'int' ? dowOffset : 1; //default dowOffset to zero
  var newYear = new Date(this.getFullYear(), 0, 1);
  var day = newYear.getDay() - dowOffset; //the day of week the year begins on
  day = day >= 0 ? day : day + 7;
  var daynum =
    Math.floor(
      (this.getTime() -
        newYear.getTime() -
        (this.getTimezoneOffset() - newYear.getTimezoneOffset()) * 60000) /
        86400000,
    ) + 1;
  var weeknum;
  //if the year starts before the middle of a week
  if (day < 4) {
    weeknum = Math.floor((daynum + day - 1) / 7) + 1;
    if (weeknum > 52) {
      nYear = new Date(this.getFullYear() + 1, 0, 1);
      nday = nYear.getDay() - dowOffset;
      nday = nday >= 0 ? nday : nday + 7;
      /*if the next year starts before the middle of
                the week, it is week #1 of that year*/
      weeknum = nday < 4 ? 1 : 53;
    }
  } else {
    weeknum = Math.floor((daynum + day - 1) / 7);
  }
  return weeknum;
};

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
};

var daysOfWeek = {
  0: 'Pazar',
  1: 'Pazartesi',
  2: 'Salı',
  3: 'Çarşamba',
  4: 'Perşembe',
  5: 'Cuma',
  6: 'Cumartesi',
};
