import sketch from 'sketch'
import fs from '@skpm/fs'
import dialog from '@skpm/dialog'
const utilities = require('./utilities')
const constants = require('./constants')
const path = require('path')
const XLSX = require('xlsx')

// documentation: https://developer.sketchapp.com/reference/api/
// Based on: https://github.com/DWilliames/Google-sheets-content-sync-sketch-plugin/blob/master/Google%20sheets%20content%20sync.sketchplugin/Contents/Sketch/main.js

const document = sketch.getSelectedDocument()
// const page = document.selectedPage

const directory = path.dirname(document.path)
// const contentFileName = 'content.csv'

//Excel header
var generatedFileData = []

var duplicateKeys = 0
// const excelHeader = new ExcelContent('key', 'en-US')
// generatedFileData.push(excelHeader)

export default function() {

  if (document.pages) {

    for (let page of document.pages) {
      //Don't add symbols page
      if (page.name != "Symbols") {
        // console.log(page.name)

        for (let layer of page.layers) {
          // console.log(layer.name)
          getLayers(layer)
        }
      }
    }
    saveToFile()
  } else {
    console.log("Document contains no pages")
  }
}

function getLayers(layer) {
  if (layer.layers) {
    // console.log("still has layers")
    for (let layer of layer.layers) {
      getLayers(layer)
    }
  } else {
    // console.log("No more layers " + layer.name, layer.type )
    if (layer.type === String(sketch.Types.Text)) {
      // console.log("Text layer")
      addToSheet(layer.name, layer.text)
    }

    if (layer.type === String(sketch.Types.SymbolInstance)) {
      console.log("SymbolInstance layer")
      console.log(layer.name)
      // console.log(layer)
      for (let override of layer.overrides) {
        // console.log("override:")
        // console.log(override)
        if (override.property == "stringValue") {
          console.log("stringValue affectedLayer: ")
          console.log(override.id)
          console.log(override.path)
          console.log(layer.name)

          console.log(layerNamesFromPath(override.path))

          let key = layer.name + constants.excelDivider + layerNamesFromPath(override.path)
          addToSheet(key, override.value)
        }
      }
    }
  }
}

function ExcelContent(key, value) {
  this.key = key
  this.value = value
}

function addToSheet(key, value) {
  //check if key already exists
  if (generatedFileData.filter(excelContent => (excelContent.key === key)).length) {
    //skip
    duplicateKeys += 1
  } else {
    //add to array
    const keyValue = new ExcelContent(key, value)
    generatedFileData.push(keyValue)
  }
  console.log("Adding to sheet: " + key, value)
}

function saveToFile() {
  var date = new Date()
  var dateFormat = date.getFullYear() + "" + (date.getMonth() + 1) + "" + date.getDate()
  console.log(dateFormat)
  var sketchFileName = "Sketch"
  var defaultPath = path.join(directory, 'sketchFileName-content-'+ dateFormat +'.xlsx')
  console.log(defaultPath)

  var filePath = dialog.showSaveDialog({
    filters: [
      { name: 'Excel', extensions: ['xlsx'] }
    ],
    defaultPath: defaultPath
  })
  console.log(filePath);
  console.log(generatedFileData);
  const book = XLSX.utils.book_new()
  const sheet = XLSX.utils.json_to_sheet(generatedFileData)
  XLSX.utils.book_append_sheet(book, sheet, "content");

  const content = XLSX.write(book, { type: 'buffer', bookType: 'xlsx', bookSST: false });
  fs.writeFileSync(filePath, content, { encoding: 'binary' });
  console.log("File created.")
  onComplete()
}

function onComplete() {
  console.log("Completed")
  sketch.UI.message("Completed. Duplicates: " + duplicateKeys + " File saved as .xlsx")
}

// **********************
//   Helper methods
// **********************
//TODO: function duplicated
function layerNamesFromPath(path) {

    var layerNames = []
    let layerIDs = path.split(constants.sketchSymbolDivider)
    for (let layerID of layerIDs) {

      let layer = document.getLayerWithID(layerID)

      //TODO: Sketch libraries not supported yet.
      if (layer) {
        let layerName = layer.name
        layerNames.push(layerName)
      }
    }
    return layerNames.join(constants.excelDivider)
}
