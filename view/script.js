const fs = require('fs')
const marked = require('marked')
const remote = require('electron').remote
const path = require('path')
const {createMd} = require('./createMd')
const {shell} = require('electron')

const excelPath = path.join(__dirname, '/data.xlsx')
const mdPath = path.join(__dirname, '/2021.md')

const readFile = file => {
  if (!fs.existsSync(file)) {
    throw new Error('başaramadım. neyi başaramadın a..?')
  }

  fs.readFile(file, (err, data) => {
    if (!err) {
      document.querySelector('.md').innerHTML = marked(data.toString())
    }
  })
}

const refreshDocument = async () => {
  try {
    await createMd()
    let mdPath = path.join(__dirname, '/../../../2021.md')
    readFile(mdPath)
  } catch (error) {
    console.log(error.message)
    try {
      let mdPath = path.join(__dirname, '/../2021.md')
      readFile(mdPath)
    } catch (error) {
      console.log(error.message)
    }
  }
}

const openDataExcel = async () => {
  try {
    let excelPath = path.join(__dirname, '../data.xlsx')
    if (fs.existsSync(excelPath)) {
      shell.openPath(excelPath)
    } else {
      throw new Error()
    }
  } catch (error) {
    console.log(error.message)
    let excelPath = path.join(__dirname, '/data.xlsx')
    if (fs.existsSync(excelPath)) {
      shell.openPath(excelPath)
    } else {
      throw new Error()
    }
  }
}

const close = e => {
  const window = remote.getCurrentWindow()
  window.close()
}

const showMessage = (message, type) => {
  message = 'Excel dosyasında hata! '
  const messageClass = message === 's' ? 'successMessage' : 'errorMessage'
  document.querySelector('.message').innerHTML = `
   <div class="${messageClass}">
    <h3> ${message}</h3>
    </div>
  `
  // setTimeout(() => {
  //   document.querySelector('.message').innerHTML = ''
  // }, 1000);
}

document.querySelector('.close').addEventListener('click', close)
document.querySelector('.refresh').addEventListener('click', showMessage)
document.querySelector('.open-excel').addEventListener('click', openDataExcel)
;(function init() {
  refreshDocument()
})()
