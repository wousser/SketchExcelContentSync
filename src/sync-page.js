import sketch from 'sketch'
import fs from '@skpm/fs'
import dialog from '@skpm/dialog'
const constants = require('./constants.js')
const UI = require('sketch/ui')
const path = require('path')
const csv = require('csvtojson')
const XLSX = require('xlsx')

// documentation: https://developer.sketchapp.com/reference/api/
// Based on: https://github.com/DWilliames/Google-sheets-content-sync-sketch-plugin/blob/master/Google%20sheets%20content%20sync.sketchplugin/Contents/Sketch/main.js

const document = sketch.getSelectedDocument()
const directory = path.dirname(document.path)
const defaultContentFileCSV = path.join(directory, "content.csv")
const defaultContentFileXLSX = path.join(directory, "content.xlsx")

var contentDictionary = {}
var languageOptions = []
var selectedLanguage

var contentFile
if (fs.existsSync(defaultContentFileXLSX)) {
  contentFile = defaultContentFileXLSX
} else if (fs.existsSync(defaultContentFileCSV)) {
  contentFile = defaultContentFileCSV
}

export function syncAllPages(context) {
  console.log("syncAllPages")
}

export function syncCurrentPage(context) {
  console.log("syncCurrentPage")

  //check if default file exist or ask for file input
  if (fs.existsSync(contentFile)) {
    console.log("file exists: " + contentFile);
  } else {
    console.log("showOpenDialog");
    var filePaths = dialog.showOpenDialog({
      properties: ['openFile'],
      defaultPath: 'directory',
      filters: [
        { name: 'Excel or CSV', extensions: ['xlsx', 'csv'] }
      ]
    })
    if (filePaths.length) {
      contentFile = filePaths[0]
    } else {
      console.log("no file selected")
      sketch.UI.message("No file selected. Select an Excel or CSV file to continue.")
    }
  }

  //check filetype
  var fileType = contentFile.split(".").pop()

  switch(fileType.toLowerCase()) {
    case "csv":
      console.log("csv")
      loadCSVData(contentFile)
      populatePage()
      break
    case "xls", "xlsx":
      console.log("Excel")
      loadExcelData(contentFile)
      populatePage()
      break
    default:
      console.log("File format not supported.")
      sketch.UI.message("File format not supported.")
  }
}

function showLanguageSelectionPopup(languageOptions) {
  var selection = UI.getSelectionFromUser(
    "Language?",
    languageOptions
  )

  var ok = selection[2]
  if (ok) {
    selectedLanguage = languageOptions[selection[1]]
    console.log(selectedLanguage)
  } else {
    console.log("showLanguageSelectionPopup() pressed cancel.")
    return
  }
}

function loadExcelData(contentFile) {
  const xlsData = fs.readFileSync(contentFile)

  var workbook = XLSX.read(xlsData, {type: "buffer"})
  /* Get worksheet. Only support one sheet at the moment. */
  var first_sheet_name = workbook.SheetNames[0]
  var worksheet = workbook.Sheets[first_sheet_name]

  var excelJson = XLSX.utils.sheet_to_json(worksheet)

  let rowNumber = 2 //Excel row starts at 0, and 1st row is key/value

  //get language options
  var keyAndLanguageOptions = Object.keys(excelJson[0]) //get language options from first row
  if (!keyAndLanguageOptions) {
    console.log("File format not supported. Language options not found")
    sketch.UI.message("File format not supported. Language options not found")
    return
  }
  keyAndLanguageOptions.shift() //remove 'key'
  for (let language of keyAndLanguageOptions) {
    languageOptions.push(language)
  }

  //ask for language first so we don't load all language data into the object.
  showLanguageSelectionPopup(languageOptions)
  if (!selectedLanguage) {
    console.log("loadExcelData() aborted. No language selected.")
    return
  }

  for (var row in excelJson) {
        //skip empty content
        if (excelJson[row][selectedLanguage]) {
          console.log("rowNumber: " + rowNumber)
          contentDictionary[String(excelJson[row]['key'])] = String(excelJson[row][selectedLanguage])
        } else {
          console.log("skipped rowNumber: " + rowNumber)
        }
        rowNumber += 1
    }
    onComplete()
}

function populatePage(page) {
  //abort if no language is chosen
  if (!selectedLanguage) {
    console.log("populatePage() aborted. No language selected.")
    return
  }

  // Use selected page if no page is set
  if (!page) {
    page = document.selectedPage
    console.log("page: " + page.name)
  }

  console.log("page layers: " + page.layers.length)
  for (let layer of page.layers) {
    console.log(layer.name, layer.type)
    switch(layer.type) {
      case String(sketch.Types.SymbolInstance):
        updateSymbolLayer(layer)
        break
      case String(sketch.Types.Text):
        updateTextLayer(layer)
        break
      case String(sketch.Types.Artboard):
        updateArtboardLayer(layer)
        console.log("artboard");
        break
    }
  }
  context.document.reloadInspector()
  onComplete()
}

//Load CSV File
function loadCSVData(contentFile) {
  const csvData = fs.readFileSync(contentFile)

  csv({
    noheader: false
    // output: "csv"
  })
  .fromString(csvData.toString())
  .subscribe((json,lineNumber)=>{
    console.log(lineNumber)
    // console.log(json)
    updateContent(json['key'], json[selectedLanguage])

}, onError, onComplete)
}

function onError(err) {
    console.log("Error: " + err)
    sketch.UI.message("An error occured: " + err)
}

function onComplete() {
  console.log("Completed")
  sketch.UI.message("Completed")
}

function updateTextLayer(layer) {
  console.log("updateTextLayer")
  console.log(layer.name)
  if (contentDictionary[layer.name]) {
    layer.text = contentDictionary[layer.name]
  }
  console.log("updateTextLayer done")
}

function updateSymbolLayer(symbol) {
  console.log("updateSymbolLayer")
  console.log(symbol.name)
  // console.log(symbol)

  for (let override of symbol.overrides) {
    if (override.property == "stringValue") {
      let layerNameAndOverride = symbol.name + constants.excelDivider + layerNamesFromPath(override.path)

      if (contentDictionary[layerNameAndOverride]) {
        override.value = contentDictionary[layerNameAndOverride]
      }
    }
  }
  console.log("updateSymbolLayer done")
}

function updateArtboardLayer(artboard) {
  console.log("updateArtboardLayer")
  console.log("page layers: " + artboard.layers.length)
  for (let layer of artboard.layers) {
    console.log(layer.name, layer.type)
    switch(layer.type) {
      case String(sketch.Types.SymbolInstance):
        updateSymbolLayer(layer)
        break
      case String(sketch.Types.Text):
        updateTextLayer(layer)
        break
    }
  }
  console.log("updateArtboardLayer done")
}

// **********************
//   Helper methods
// **********************

function layerNamesFromPath(path) {
    var layerNames = []
    let layerIDs = path.split(constants.sketchSymbolDivider)
    for (let layerID of layerIDs) {
      let layer = document.getLayerWithID(layerID)
      let layerName = layer.name
      layerNames.push(layerName)
    }
    return layerNames.join(constants.excelDivider)
}
