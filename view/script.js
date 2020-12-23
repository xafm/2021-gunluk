const fs = require('fs')
const marked = require('marked')
const remote = require('electron').remote
const path = require('path')
const {createMd} = require('./createMd')
const {shell} = require('electron')

const excelPath = path.join(__dirname, '../data.xlsx')
const mdPath = path.join(__dirname, '../2021.md')

const readFile = (file, showSuccessMessage) => {
  fs.access(file, fs.F_OK, err => {
    if (err) {
      showMessage({
        message: `Teknik hata: Oluşturulan markdown dosyası ${mdPath} dizininde bulunamadı`,
        type: 'e',
      })
      return
    }
    fs.readFile(file, (err, data) => {
      if (err) {
        showMessage({
          message: `Teknik hata: Markdown dosyası ${mdPath} dizininden okunamadı`,
          type: 'e',
        })
        return
      }
      document.querySelector('.md').innerHTML = marked(data.toString())
      if (showSuccessMessage) {
        showMessage({
          message: 'Yenilendi!',
          type: 's',
        })
      }
    })
  })
}

const refreshDocument = async (showSuccessMessage) => {
  try {
    await createMd()
    readFile(mdPath, showSuccessMessage)
  } catch (error) {
    showMessage({
      message: error.message,
      type: 'e',
    })
  }
}

const openDataExcel = async () => {
  fs.access(excelPath, fs.F_OK, err => {
    if (err) {
      showMessage({
        message: `Excel dosyası ${excelPath} dizininde bulunamadı`,
        type: 'e',
      })
      return
    }
    shell.openPath(excelPath)
  })
}

const close = e => {
  const window = remote.getCurrentWindow()
  window.close()
}

const timeoutList = []
const showMessage = ({message, type}) => {
  if (!message || !type) {
    return
  }

  let messageClass = ''
  switch (type) {
    case 's':
      messageClass = 'successMessage'
      break
    case 'e':
      messageClass = 'errorMessage'
      break
    default:
      break
  }

  if (timeoutList.length) {
    while (timeoutList.length) {
      clearTimeout(timeoutList.pop())
    }
  }

  document.querySelector('.message').innerHTML = `
   <div  id="msg" class="${messageClass}">
    <h3> ${message}</h3>
    </div>
  `

  const timeout = setTimeout(() => {
    document.querySelector('.message').innerHTML = ''
  }, 5000)
  timeoutList.push(timeout)
}

document.querySelector('.close').addEventListener('click', close)
document.querySelector('.refresh').addEventListener('click', () => refreshDocument(true))
document.querySelector('.open-excel').addEventListener('click', openDataExcel)
;(function init() {
  refreshDocument(false)
})()
