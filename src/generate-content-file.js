import sketch from 'sketch'
import fs from '@skpm/fs'
import dialog from '@skpm/dialog'
const constants = require('./constants')
const path = require('path')
const XLSX = require('xlsx')

// documentation: https://developer.sketchapp.com/reference/api/
// Based on: https://github.com/DWilliames/Google-sheets-content-sync-sketch-plugin/blob/master/Google%20sheets%20content%20sync.sketchplugin/Contents/Sketch/main.js

const document = sketch.getSelectedDocument()

class ExcelContent {
  constructor (key, value) {
    this.key = key
    this.value = value
  }
}
var generatedFileData = []
var duplicateKeys = 0

export default function () {
  if (document.pages) {
    for (let page of document.pages) {
      // Don't add Symbols page
      if (page.name !== 'Symbols') {
        getPageContent(page)
      }
    }
    saveToFile()
  } else {
    console.log('Document contains no pages')
    sketch.UI.message('Document contains no pages')
  }
}

function getPageContent (page) {
  console.log('getPageContent: ', page.name)

  for (let layer of page.layers) {
    // console.log(layer.name)
    getContent(layer)
  }
}

function getContent (layer) {
  switch (layer.type) {
    case String(sketch.Types.SymbolInstance):
      contentFromSymbolLayer(layer)
      break
    case String(sketch.Types.Text):
      contentFromTextLayer(layer)
      break
    case String(sketch.Types.Artboard):
      contentFromArtboardLayer(layer)
      break
    default:
      // check for groups and combined layers
      if (layer.layers) {
        for (let individualLayer of layer.layers) {
          getContent(individualLayer)
        }
      }
  }
}

function contentFromSymbolLayer (symbol) {
  console.log('contentFromSymbolLayer: ', symbol.name)

  for (let override of symbol.overrides) {
    if (override.property === 'stringValue') {
      // console.log('stringValue')
      console.log(symbol.id, override.id, override.path, symbol.name)

      let key = symbol.name + constants.excelDivider + layerNamesFromPath(override.path)
      addToSheet(key, override.value)
    }
  }
}

function contentFromTextLayer (layer) {
  console.log('contentFromTextLayer: ', layer.name)
  addToSheet(layer.name, layer.text)
}

function contentFromArtboardLayer (artboard) {
  console.log('contentFromArtboardLayer: ', artboard.name)
  // add artboard market to sheet
  let artboardDivider = 'ARTBOARD: ' + artboard.name
  addToSheet('', '') // add empty row
  addToSheet(artboardDivider, '')

  // console.log('artboard layers: ' + artboard.layers.length)
  for (let layer of artboard.layers) {
    // console.log(layer.name, layer.type)
    getContent(layer)
  }
}

function addToSheet (key, value) {
  // check if key already exists, except for empty row
  if (generatedFileData.filter(excelContent => (excelContent.key === key)).length && key !== '') {
    // skip
    duplicateKeys += 1
  } else {
    // add to array
    let keyValue = new ExcelContent(key, value)
    generatedFileData.push(keyValue)
  }
  // console.log('Adding to sheet: ' + key, value)
}

function saveToFile () {
  var date = new Date()
  var dateFormat = date.getFullYear() + '' + (date.getMonth() + 1) + '' + date.getDate()

  let name = decodeURI(path.basename(document.path, '.sketch'))
  let contentFileName = name + '-content-' + dateFormat + '.xlsx'
  var defaultPath = path.join(path.dirname(document.path), contentFileName)
  console.log(defaultPath)
  var filePath = dialog.showSaveDialog({
    filters: [
      { name: 'Excel', extensions: ['xlsx'] }
    ],
    defaultPath: defaultPath
  })

  // check if user want to save the file
  if (filePath) {
    // console.log(generatedFileData)
    const book = XLSX.utils.book_new()
    const sheet = XLSX.utils.json_to_sheet(generatedFileData)
    XLSX.utils.book_append_sheet(book, sheet, 'content')

    const content = XLSX.write(book, { type: 'buffer', bookType: 'xlsx', bookSST: false })
    fs.writeFileSync(filePath, content, { encoding: 'binary' })
    console.log('File created as:', filePath)

    // done
    console.log('Completed. Duplicates: ' + duplicateKeys + ' File saved as ' + path.basename(filePath))
    sketch.UI.message('Content file generated. Found ' + duplicateKeys + ' duplicated keys. File saved as ' + decodeURI(path.basename(filePath)))
  }
}

// **********************
//   Helper methods
// **********************
// TODO: function duplicated
function layerNamesFromPath (path) {
  var layerNames = []
  let layerIDs = path.split(constants.sketchSymbolDivider)
  for (let layerID of layerIDs) {
    let layer = document.getLayerWithID(layerID)

    // TODO: Sketch libraries not supported.
    if (layer) {
      let layerName = layer.name
      layerNames.push(layerName)
    }
  }
  // console.log(layerNames.join(constants.excelDivider))
  return layerNames.join(constants.excelDivider)
}
