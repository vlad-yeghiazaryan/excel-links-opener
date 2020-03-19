const excelInput = document.querySelector('#excel_input')
const linksTable = document.querySelector('#linksTable')

// Excel to json reader class with parseExcel method
class ExcelToJSON {
  constructor () {
    this.parseExcel = function (file) {
      var reader = new FileReader()
      reader.onload = function (e) {
        var data = e.target.result
        // parsing data
        var workbook = XLSX.read(data, {
          type: 'binary'
        })
        // All main run here
        const sheets = workbook.Sheets
        const jsonData = makeJson(sheets)
        jsonData.forEach(sheet => {
          makeTable(sheet, linksTable)
        })
      }
      reader.onerror = function (ex) {
        console.log(ex)
      }
      reader.readAsBinaryString(file)
    }
  }
}

// range function from python
function range (size, startAt = 0) {
  return [...Array(size).keys()].map(i => i + startAt)
}

// Excel column iterator
class StringIdGenerator {
  constructor (chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') {
    this._chars = chars
    this._nextId = [0]
  }

  next () {
    const r = []
    for (const char of this._nextId) {
      r.unshift(this._chars[char])
    }
    this._increment()
    return r.join('')
  }

  _increment () {
    for (let i = 0; i < this._nextId.length; i++) {
      const val = ++this._nextId[i]
      if (val >= this._chars.length) {
        this._nextId[i] = 0
      } else {
        return
      }
    }
    this._nextId.push(0)
  }

  * [Symbol.iterator] () {
    while (true) {
      yield this.next()
    }
  }
}

// The table maker
const makeJson = (sheets) => {
  // Going over each sheet
  const sheetsList = []
  for (const sheet in sheets) {
    const sheetData = sheets[sheet]
    console.log(sheetData)
    const sheetRange = sheetData['!ref']
    // getting the number of cells to make for each sheet
    const rowStart = parseInt(sheetRange.split(':')[0].replace(/\D/g, '')) + 1
    const rowEnd = parseInt(sheetRange.split(':')[1].replace(/\D/g, ''))
    const columnStart = sheetRange.split(':')[0].replace(/[0-9]/g, '')
    const columnEnd = sheetRange.split(':')[1].replace(/[0-9]/g, '')
    const columnIterator = new StringIdGenerator()
    const columnRangeIterator = new StringIdGenerator()
    let columnRangeEnd = 0
    let columnRangeStart = 0
    let checkingRange = true
    // finding column range
    while (checkingRange) {
      if (columnRangeStart !== 0) {
        columnIterator.next()
      }
      columnRangeStart += 1
      if (columnRangeIterator.next() === columnStart) {
        columnRangeEnd = columnRangeStart
        while (columnRangeIterator.next() !== columnEnd) {
          columnRangeEnd += 1
        }
        checkingRange = false
      }
    }
    // Looping through rows
    const sheetrows = []
    for (const rowIndex in range(rowEnd, rowStart)) {
      const columnIterator = new StringIdGenerator()
      const sheetrow = {}
      for (const fieldIndex in range(columnRangeEnd + 1, columnRangeStart)) {
        const column = columnIterator.next()
        const ref = column + parseInt(parseInt(rowIndex) + 1)
        sheetrow[column] = sheetData[ref]
      }
      sheetrows.push(sheetrow)
    }
    sheetsList.push(sheetrows)
  }
  return sheetsList
}
// function that checks if a string is a valid url
function validURL (str) {
  var pattern = new RegExp('^(https?:\\/\\/)?' + // protocol
    '((([a-z\\d]([a-z\\d-]*[a-z\\d])*)\\.)+[a-z]{2,}|' + // domain name
    '((\\d{1,3}\\.){3}\\d{1,3}))' + // OR ip (v4) address
    '(\\:\\d+)?(\\/[-a-z\\d%_.~+]*)*' + // port and path
    '(\\?[;&a-z\\d%_.~+=-]*)?' + // query string
    '(\\#[-a-z\\d_]*)?$', 'i') // fragment locator
  return !!pattern.test(str)
}
// Make a table from json
function makeTable (jsonData, table) {
  // Making the headers
  const tr = document.createElement('tr')
  Object.keys(jsonData[0]).forEach(key => {
    const th = document.createElement('th')
    th.innerHTML = key
    th.setAttribute('scope', 'col')
    tr.appendChild(th)
  })
  // Adding Button
  const th = document.createElement('th')
  th.innerHTML = 'Button'
  th.setAttribute('scope', 'col')
  tr.appendChild(th)
  // Button added

  table.appendChild(tr)
  jsonData.forEach(record => {
    // Getting object keys
    const keys = Object.keys(record)
    const tr = document.createElement('tr')

    // Adding main names
    const columnRange = Object.keys(jsonData[0]).length
    const td = document.createElement('td')
    const span = document.createElement('span')

    // Finding the first row that has data
    let mainRow = 0
    while (record[keys[mainRow]] === undefined) {
      mainRow += 1
      if (mainRow > columnRange) {
        break
      }
    }
    if (record[keys[mainRow]] !== undefined) {
      span.innerHTML = record[keys[mainRow]].v
      td.appendChild(span)
      td.setAttribute('scope', 'row')
      tr.appendChild(td)
    } else {
      span.innerHTML = '-'
      td.appendChild(span)
      td.setAttribute('scope', 'row')
      tr.appendChild(td)
    }
    // Main names added

    // Adding feilds
    for (let index = mainRow + 1; index < keys.length + mainRow; index++) {
      const td = document.createElement('td')
      const a = document.createElement('a')
      const span = document.createElement('span')
      const fieldElement = record[keys[index]]
      const isNotUndefined = fieldElement !== undefined

      if (isNotUndefined) {
        const field = fieldElement.v
        const isValidURL = validURL(field)

        if (isValidURL) {
          a.href = field
          a.innerHTML = field
          a.setAttribute('target', '_blank')
          td.appendChild(a)
          td.setAttribute('scope', 'row')
          tr.appendChild(td)
        } else {
          span.innerHTML = field
          td.appendChild(span)
          td.setAttribute('scope', 'row')
          tr.appendChild(td)
        }
      } else {
        span.innerHTML = '-'
        td.appendChild(span)
        td.setAttribute('scope', 'row')
        tr.appendChild(td)
      }
    }
    // Adding button element
    const btn = document.createElement('button')
    btn.setAttribute('type', 'button')
    btn.setAttribute('class', 'btn btn-success')
    btn.innerHTML = 'Open'
    btn.addEventListener('click', openWindows)
    tr.appendChild(btn)
    // Button element added

    // making the record
    table.appendChild(tr)
  })
}
// function for opening windows
const openWindows = async (e) => {
  e.preventDefault()
  const links = e.target.parentElement.children
  for (let index = 1; index < links.length - 1; index++) {
    const link = links[index].firstElementChild.href
    if (link !== undefined) { window.open(link) }
  }
}
// excelTable runs when data is inputed
const excelTable = async (e) => {
  var xl2json = new ExcelToJSON()
  const excelFile = e.target.files[0]
  xl2json.parseExcel(excelFile)
}

excelInput.addEventListener('change', excelTable)
