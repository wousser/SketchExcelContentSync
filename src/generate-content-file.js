import sketch from 'sketch'
import fs from '@skpm/fs'
import dialog from '@skpm/dialog'
const constants = require('./constants')
const path = require('path')
const XLSX = require('xlsx')

// documentation: https://developer.sketchapp.com/reference/api/
// Based on: https://github.com/DWilliames/Google-sheets-content-sync-sketch-plugin/blob/master/Google%20sheets%20content%20sync.sketchplugin/Contents/Sketch/main.js

const document = sketch.getSelectedDocument()

// const directory = path.dirname(document.path)

// Excel header
var generatedFileData = []

var duplicateKeys = 0

export default function () {
  if (document.pages) {
    for (let page of document.pages) {
      // Don't add symbols page
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
    getLayers(layer)
  }
}

function getLayers (layer) {
  if (layer.layers) {
    // console.log("still has layers")
    for (let indLayer of layer.layers) {
      getLayers(indLayer)
    }
  } else {
    console.log(layer.name, layer.type)
    switch (layer.type) {
      case String(sketch.Types.SymbolInstance):
        contentFromSymbolLayer(layer)
        break
      case String(sketch.Types.Text):
        contentFromTextLayer(layer)
        break
      case String(sketch.Types.Artboard):
        contentFromArtboardLayer(layer)
        console.log('artboard')
        break
    }
  }
}

function contentFromSymbolLayer (symbol) {
  console.log('contentFromSymbolLayer')
  // console.log(layer.name)
  // console.log(layer)
  for (let override of symbol.overrides) {
    // console.log("override:")
    // console.log(override)
    if (override.property === 'stringValue') {
      console.log('stringValue')
      console.log(symbol.id)
      console.log(symbol.symbolId)
      console.log(override.id)
      // console.log(symbol.name)
      // console.log(symbol)
      console.log(override.id)
      // console.log(override.path)

      // if (symbol.id == 'F5DA5A57-72E6-4A02-8048-0827032405B7') {
      //   console.log(symbol)
      // }

      console.log(layerNamesFromPath(override.path))

      let key = symbol.name + constants.excelDivider + layerNamesFromPath(override.path)
      addToSheet(key, override.value)
    }
  }
}

function contentFromTextLayer (layer) {
  console.log('contentFromTextLayer')
  addToSheet(layer.name, layer.text)
}

function contentFromArtboardLayer (artboard) {
  console.log('contentFromArtboardLayer')
  console.log('artboard layers: ' + artboard.layers.length)
  for (let layer of artboard.layers) {
    console.log(layer.name, layer.type)
    switch (layer.type) {
      case String(sketch.Types.SymbolInstance):
        contentFromSymbolLayer(layer)
        break
      case String(sketch.Types.Text):
        contentFromTextLayer(layer)
        break
    }
  }
}

class ExcelContent {
  constructor (key, value) {
    this.key = key
    this.value = value
  }
}

function addToSheet (key, value) {
  // check if key already exists
  if (generatedFileData.filter(excelContent => (excelContent.key === key)).length) {
    // skip
    duplicateKeys += 1
  } else {
    // add to array
    let keyValue = new ExcelContent(key, value)
    generatedFileData.push(keyValue)
  }
  console.log('Adding to sheet: ' + key, value)
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
    sketch.UI.message('Completed. Duplicates: ' + duplicateKeys + ' File saved as ' + decodeURI(path.basename(filePath)))
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

    // TODO: Sketch libraries not supported yet.
    if (layer) {
      let layerName = layer.name
      layerNames.push(layerName)
    }
  }
  return layerNames.join(constants.excelDivider)
}
