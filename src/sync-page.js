import sketch from 'sketch'
import fs from '@skpm/fs'
import dialog from '@skpm/dialog'
import BrowserWindow from 'sketch-module-web-view'

var constants = require('./constants')
var UI = require('sketch/ui')
var path = require('path')
var Settings = require('sketch/settings')
// var csv = require('csvtojson')
var XLSX = require('xlsx')

// documentation: https://developer.sketchapp.com/reference/api/
// Based on: https://github.com/DWilliames/Google-sheets-content-sync-sketch-plugin/blob/master/Google%20sheets%20content%20sync.sketchplugin/Contents/Sketch/main.js

var document = sketch.getSelectedDocument()

var contentDictionary = {}
var languageOptions = []
var selectedLanguage

// var contentFile
// if (fs.existsSync(defaultContentFileXLSX)) {
//   contentFile = defaultContentFileXLSX
// } else if (fs.existsSync(defaultContentFileCSV)) {
//   contentFile = defaultContentFileCSV
// }

export function panelTest (context) {
  console.log('panelTest4')

  const options = {
    identifier: 'unique.id'
  }
  const browserWindow = new BrowserWindow(options)
  browserWindow.loadURL(require('./my-screen.html'))
}

export function syncCurrentPage (context) {
  console.log('syncCurrentPage')
  if (contentDocumentExists()) {
    console.log('contentDocumentExists true')
    var contentFile = Settings.documentSettingForKey(document, 'excelTranslateContentFile')
    loadData(contentFile)
    populatePage()
    context.document.reloadInspector()
  }
}

export function syncAllPages (context) {
  console.log('syncAllPages')
  if (contentDocumentExists() && document.pages) {
    var contentFile = Settings.documentSettingForKey(document, 'excelTranslateContentFile')
    loadData(contentFile)
    for (let page of document.pages) {
      // Don't add symbols page
      if (page.name !== 'Symbols') {
        populatePage(page)
      }
    }
    context.document.reloadInspector()
  } else {
    console.log('Document contains no pages, or content document does not exist.')
    sketch.UI.message('Document contains no pages, or content document does not exist.')
  }
}

export function selectContentDocument (context) {
  console.log('selectContentDocument')
  var filePaths = dialog.showOpenDialog({
    properties: ['openFile'],
    defaultPath: 'directory',
    filters: [{
      name: 'Excel or CSV',
      extensions: ['xlsx', 'xls', 'csv']
    }]
  })
  if (filePaths.length) {
    var contentFile = filePaths[0]
    console.log('set contentFile: ', contentFile)
    // set as document key
    Settings.setDocumentSettingForKey(document, 'excelTranslateContentFile', contentFile)
    console.log('DocumentKey: ', Settings.documentSettingForKey(document, 'excelTranslateContentFile'))
    sketch.UI.message('Content document set! Now you can translate the current page or all pages.')
  } else {
    console.log('no file selected')
    sketch.UI.message('No file selected. Select Excel or CSV file to continue. Generate a file if you don\'t have one.')
  }
}

function loadData (contentFile) {
  // check filetype
  let fileType = path.extname(contentFile)
  switch (fileType.toLowerCase()) {
    case '.csv':
      console.log('csv')
      // loadCSVData(contentFile)
      break
    // eslint-disable-next-line no-sequences
    case '.xls', '.xlsx':
      console.log('Excel')
      loadExcelData(contentFile)
      break
    default:
      console.log('File format not supported.')
      sketch.UI.message('File format not supported.')
  }
}

function contentDocumentExists () {
  var contentFile = Settings.documentSettingForKey(document, 'excelTranslateContentFile')
  console.log('contentFile: ', contentFile)
  if (contentFile) {
    if (fs.existsSync(contentFile)) {
      console.log('file exists: ', contentFile)
      return true
    }
  } else {
    console.log('No content file.')
    sketch.UI.message('Select a content document first.')
    return false
  }
}

function showLanguageSelectionPopup (languageOptions) {
  let fileName = Settings.documentSettingForKey(document, 'excelTranslateContentFile')
  UI.getInputFromUser(
    `Sync to language?\n\nContent file: ${fileName}`,
    {
      type: UI.INPUT_TYPE.selection,
      possibleValues: languageOptions
    },
    (err, value) => {
      if (err) {
        // most likely the user canceled the input
        return
      }
      console.log('selected language')
      console.log(value)
      selectedLanguage = value // languageOptions[selection[1]]
    }
  )
}

function loadExcelData (contentFile) {
  let xlsData = fs.readFileSync(contentFile)

  var workbook = XLSX.read(xlsData, {
    type: 'buffer'
  })
  /* Get worksheet. Only support one sheet at the moment. */
  var firstSheetName = workbook.SheetNames[0]
  var worksheet = workbook.Sheets[firstSheetName]

  var excelJson = XLSX.utils.sheet_to_json(worksheet, { range: 0, defval: '' })
  console.log(excelJson)
  let rowNumber = 2 // Excel row starts at 0, and 1st row is key/value

  // get language options
  var keyAndLanguageOptions = Object.keys(excelJson[0]) // get language options from first row
  if (!keyAndLanguageOptions) {
    console.log('File format not supported. Language options not found')
    sketch.UI.message('File format not supported. Language options not found')
    return
  }
  console.log(keyAndLanguageOptions)
  keyAndLanguageOptions.shift() // remove 'key'
  for (let language of keyAndLanguageOptions) {
    languageOptions.push(language)
  }

  // ask for language first so we don't load all language data into the object.
  console.log('showLanguageSelectionPopup')
  console.log(languageOptions)
  showLanguageSelectionPopup(languageOptions)
  if (!selectedLanguage) {
    console.log('loadExcelData() aborted. No language selected.')
    return
  }
  var currentArboardName = ''
  for (var row in excelJson) {
    // skip empty content
    if (excelJson[row]['key'].includes(constants.artboardPrefix)) {
      currentArboardName = excelJson[row]['key']
      console.log('updated current arboard' + currentArboardName)
    }
    if (excelJson[row][selectedLanguage]) {
      console.log('rowNumber: ' + rowNumber, currentArboardName)
      contentDictionary[String(currentArboardName + '.' + excelJson[row]['key'])] = String(excelJson[row][selectedLanguage])
    } else {
      console.log('skipped rowNumber: ' + rowNumber)
    }
    rowNumber += 1
  }
  console.log(contentDictionary)
  onComplete()
}

function analyzeLayer (layers) {
  for (let layer of layers) {
    console.log(layer.name, layer.type)
    switch (layer.type) {
      case String(sketch.Types.Group):
        console.log('group layer')
        analyzeLayer(layer.layers)
        break
      case String(sketch.Types.SymbolInstance):
        console.log('symbol layer')
        updateSymbolLayer(layer)
        break
      case String(sketch.Types.Text):
        console.log('text layer')
        updateTextLayer(layer)
        break
      case String(sketch.Types.Artboard):
        console.log('artboard layer')
        analyzeLayer(layer.layers)
        break
    }
  }
}

function populatePage (page) {
  // abort if no language is chosen
  if (!selectedLanguage) {
    console.log('populatePage() aborted. No language selected.')
    return
  }

  // Use selected page if no page is set
  if (!page) {
    page = document.selectedPage
    console.log('page: ' + page.name)
  }

  console.log('page layers: ' + page.layers.length)
  analyzeLayer(page.layers)
  onComplete()
}

// Load CSV File
/*
function loadCSVData (contentFile) {
  let csvData = fs.readFileSync(contentFile)

  csv({
    noheader: false
    // output: "csv"
  })
    .fromString(csvData.toString())
    .subscribe((json, lineNumber) => {
      console.log(lineNumber)
      // console.log(json)
      updateContent(json['key'], json[selectedLanguage])
    }, onError, onComplete)
}
*/

function onComplete () {
  console.log('Completed')
  sketch.UI.message('Completed')
}

function updateTextLayer (layer) {
  console.log('updateTextLayer')
  
  // TODO: Add Artboard
  let artboardAndLayerName = constants.artboardPrefix + layer.getParentArtboard().name + '.' + layer.name
  console.log(artboardAndLayerName)
  if (contentDictionary[artboardAndLayerName]) {
  // if (contentDictionary[layer.name]) {
    layer.text = contentDictionary[artboardAndLayerName]
    console.log('new value', layer.name, artboardAndLayerName, layer.text)
  }
  console.log('updateTextLayer done')
}

function updateSymbolLayer (symbol) {
  console.log('updateSymbolLayer')
  console.log(symbol.name)
  // console.log(symbol)

  for (let override of symbol.overrides) {
    if (override.property === 'stringValue') {
      let layerNameAndOverride = symbol.name + constants.excelDivider + layerNamesFromPath(override.path)

      if (contentDictionary[layerNameAndOverride]) {
        override.value = contentDictionary[layerNameAndOverride]
      }
    }
  }
  console.log('updateSymbolLayer done')
}

// function updateArtboardLayer (artboard) {
//   console.log('updateArtboardLayer')
//   console.log('page layers: ' + artboard.layers.length)
//   for (let layer of artboard.layers) {
//     console.log(layer.name, layer.type)
//     switch (layer.type) {
//       case String(sketch.Types.SymbolInstance):
//         updateSymbolLayer(layer)
//         break
//       case String(sketch.Types.Text):
//         updateTextLayer(layer)
//         break
//     }
//   }
//   console.log('updateArtboardLayer done')
// }

// **********************
//   Helper methods
// **********************
// TODO: function duplicated
function layerNamesFromPath (path) {
  var layerNames = []
  let layerIDs = path.split(constants.sketchSymbolDivider)
  for (let layerID of layerIDs) {
    let layer = document.getLayerWithID(layerID)

    // TODO: Sketch libraries not supported yet.
    if (layer) {
      let layerName = layer.name
      layerNames.push(layerName)
    }
  }
  return layerNames.join(constants.excelDivider)
}
