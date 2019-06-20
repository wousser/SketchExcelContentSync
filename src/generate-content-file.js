import sketch from 'sketch'
import fs from '@skpm/fs'
import dialog from '@skpm/dialog'
var UI = require('sketch/ui')
var constants = require('./constants')
var path = require('path')
var XLSX = require('xlsx')

// documentation: https://developer.sketchapp.com/reference/api/
// Based on: https://github.com/DWilliames/Google-sheets-content-sync-sketch-plugin/blob/master/Google%20sheets%20content%20sync.sketchplugin/Contents/Sketch/main.js

let document = sketch.getSelectedDocument()
var contentLanguage = 'en-US' // set en-US as standard

var renameTextLayersFlag = false

console.log('2057')

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
    // ask for language
    if (askContentLanguage() && askRenameTextLayers()) {
      for (let page of document.pages) {
        // Don't add Symbols page
        if (page.name !== 'Symbols') {
          if (renameTextLayersFlag) {
            renameTextLayers(page)
          }
          getPageContent(page)
        }
      }
      saveToFile()
    }
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
        for (let eachLayer of layer.layers) {
          getContent(eachLayer)
        }
      }
  }
}

function contentFromSymbolLayer (symbol) {
  console.log('contentFromSymbolLayer: ', symbol.name)
  if (symbol.name.charAt(0) === constants.translateLayerPrefix) {
    for (let override of symbol.overrides) {
      if (override.property === 'stringValue') {
        // console.log('stringValue')
        console.log(symbol.id, override.id, override.path, symbol.name)
  
        let key = symbol.name + constants.excelDivider + layerNamesFromPath(override.path)
        addToSheet(key, override.value)
      }
    }
  }
}

function contentFromTextLayer (layer) {
  console.log('contentFromTextLayer', layer.name, layer.getParentArtboard().name)
  if (layer.name.charAt(0) === constants.translateLayerPrefix) {
    addToSheet(layer.name, layer.text)
  }
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
  // if (generatedFileData.filter(excelContent => (excelContent.key === key)).length && key !== '') {
  if (false) {
    // skip
    duplicateKeys += 1
  } else {
    // add to array
    let keyValue = new ExcelContent(key, value)
    generatedFileData.push(keyValue)
  }
  // console.log('Adding to sheet: ' + key, value)
}

function renameTextLayers (page) {
  console.log('renameTextLayers: ', page.name)
}

function askRenameTextLayers () {
  var returnValue = false
  UI.getInputFromUser(
    `Prefix text and symbol layers with '${constants.translateLayerPrefix}'?`,
    {
      type: UI.INPUT_TYPE.selection,
      possibleValues: ['yes', 'no'],
      initialValue: 'no'
    },
    (err, value) => {
      if (err) {
        console.log('pressed cancel')
        // most likely the user canceled the input
        return
      }
      if (value !== 'null' && value.length > 1) {
        renameTextLayersFlag = value === 'yes'
        console.log('set renameTextLayersFlag', renameTextLayersFlag)
        returnValue = true
      }
    }
  )
  return returnValue
}

function askContentLanguage () {
  var returnValue = false
  UI.getInputFromUser(
    'Current document language?',
    {
      initialValue: 'en-US'
    },
    (err, value) => {
      if (err) {
        console.log('pressed cancel')
        // most likely the user canceled the input
        return
      }
      if (value !== 'null' && value.length > 1) {
        contentLanguage = value
        console.log('set contentLanguage', contentLanguage)
        returnValue = true
      }
    }
  )
  return returnValue
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
    let book = XLSX.utils.book_new()
    let sheet = XLSX.utils.json_to_sheet(generatedFileData)
    sheet['B1'].v = contentLanguage
    XLSX.utils.book_append_sheet(book, sheet, 'content')

    let content = XLSX.write(book, { type: 'buffer', bookType: 'xlsx', bookSST: false })
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
