import fs from "@skpm/fs";
import dialog from "@skpm/dialog";

const XLSX = require("xlsx");
var _ = require("lodash");

var sketch = require("sketch/dom");
var constants = require("./constants");
var UI = require("sketch/ui");
var path = require("path");
var Settings = require("sketch/settings");

var document = sketch.getSelectedDocument();

var contentDictionary = [];
var languageOptions = [];
var selectedLanguage;
var contentLanguage = "en-US"; // set en-US as standard
var renameTextLayersFlag = false;
var generatedFileData = [];
var duplicateKeys = 0;
var excelJson;

class ExcelContent {
  constructor(key, value) {
    this.key = key;
    this.value = value;
  }
}

//
// Generate Content
//
export function generateContentFile() {
  if (document.pages) {
    // ask for language
    if (askContentLanguage() && askRenameTextLayers()) {
      for (let page of document.pages) {
        // Don't add Symbols page
        if (page.name !== "Symbols") {
          generateContentForPage(page);
        }
      }
      saveToFile();
    }
  } else {
    console.log("Document contains no pages");
    UI.message("Sketch document contains no pages");
  }
}

function generateContentForPage(page) {
  console.time("generateContent");

  //add page to excel
  addToSheet("", ""); // add empty row
  addToSheet("", ""); // add empty row
  const pageDivider = `${constants.pagePrefix}${page.name}`;
  addToSheet(pageDivider, "");

  // all artboards
  const artBoardLayers = sketch.find("Artboard", page);
  console.log("ArtBoard layers", artBoardLayers.length);

  if (artBoardLayers.length > 0) {
    artBoardLayers.forEach((artBoard) => {
      const artBoardName = artBoard.name;

      //add artBoard to excel
      const artboardDivider = `${constants.artboardPrefix}${artBoardName}`;
      addToSheet("", ""); // add empty row
      addToSheet(artboardDivider, "");

      // Text
      const textLayers = sketch.find("Text", artBoard);
      console.log("Text layers", textLayers.length);
      if (textLayers.length > 0) {
        textLayers.forEach((textLayer) => {
          //rename with #
          renameLayer(textLayer);

          //add to sheet buffer
          addToSheet(textLayer.name, textLayer.text);
        });
      }

      // Symbol
      const symbolLayers = sketch.find("SymbolInstance", artBoard);
      console.log("Symbol layers", symbolLayers.length);
      if (symbolLayers.length > 0) {
        symbolLayers.forEach((symbolLayer) => {
          //rename with #
          renameLayer(symbolLayer);

          //add to sheet buffer
          let textOverrides = extractTextOverrides(symbolLayer);
          console.log("textOverrides", textOverrides);
          textOverrides.forEach((textOverride) => {
            addToSheet(textOverride.fullPath, textOverride.value);
          });
        });
      }
    });
  }
  console.timeEnd("generateContent");
}

function renameLayer(layer) {
  if (renameTextLayersFlag && layer.name.charAt(0) !== "#") {
    layer.name = constants.translateLayerPrefix + layer.name;
    console.log("layer renamed", layer.name);
  }
}

function extractTextOverrides(symbol) {
  console.log("extractTextOverrides");
  if (symbol.overrides && symbol.overrides.length > 0) {
    const symbolName = `${symbol.name}`;

    var result = [];

    symbol.overrides.forEach((override) => {
      if (
        override.affectedLayer.type === sketch.Types.Text &&
        override.property === "stringValue"
      ) {
        console.log(symbol.name);
        const fullPath = `${symbolName}${
          constants.excelDivider
        }${layerNamesFromPath(override.path, symbol)}`;
        console.log("path", fullPath, override.value);
        console.log(symbol);
        result.push({
          fullPath: fullPath,
          value: override.value,
        });
        // return { fullPath: fullPath, value: override.value };
        // addToSheet(fullPath, override.value);
      }
    });
    return result;
  }
}

function addToSheet(key, value) {
  // TODO: Add checkbox to skip duplicate keys
  if (false) {
    // skip
    duplicateKeys += 1;
  } else {
    // add to array
    let keyValue = new ExcelContent(key, value);
    generatedFileData.push(keyValue);
  }
  // console.log('Adding to sheet: ' + key, value)
}

function askRenameTextLayers() {
  var returnValue = false;
  UI.getInputFromUser(
    `Prefix text and symbol layers with '${constants.translateLayerPrefix}'?`,
    {
      type: UI.INPUT_TYPE.selection,
      possibleValues: ["yes", "no"],
      initialValue: "yes",
    },
    (err, value) => {
      if (err) {
        console.log("pressed cancel");
        // most likely the user canceled the input
        return;
      }
      if (value !== "null" && value.length > 1) {
        renameTextLayersFlag = value === "yes";
        console.log("set renameTextLayersFlag", renameTextLayersFlag);
        returnValue = true;
      }
    }
  );
  return returnValue;
}

function askContentLanguage() {
  var returnValue = false;
  UI.getInputFromUser(
    "Current document language?",
    {
      initialValue: "en-US",
    },
    (err, value) => {
      if (err) {
        console.log("pressed cancel");
        // most likely the user canceled the input
        return;
      }
      if (value !== "null" && value.length > 1) {
        contentLanguage = value;
        console.log("set contentLanguage", contentLanguage);
        returnValue = true;
      }
    }
  );
  return returnValue;
}

function saveToFile() {
  var date = new Date();
  var dateFormat = `${date.getFullYear()}-${
    date.getMonth() + 1
  }-${date.getDate()} ${date.getHours()}-${date.getMinutes()}`;

  let name = decodeURI(path.basename(document.path, ".sketch"));
  let contentFileName = name + "-content-" + dateFormat + ".xlsx";
  var defaultPath = path.join(path.dirname(document.path), contentFileName);
  console.log(defaultPath);
  var filePath = dialog.showSaveDialogSync({
    filters: [{ name: "Excel", extensions: ["xlsx"] }],
    defaultPath: defaultPath,
  });

  try {
    // check if user want to save the file
    if (filePath) {
      // console.log(generatedFileData)
      let book = XLSX.utils.book_new();
      let sheet = XLSX.utils.json_to_sheet(generatedFileData);
      sheet["B1"].v = contentLanguage;
      XLSX.utils.book_append_sheet(book, sheet, "content");

      let content = XLSX.write(book, {
        type: "buffer",
        bookType: "xlsx",
        bookSST: false,
      });
      fs.writeFileSync(filePath, content, { encoding: "binary" });
      console.log("File created as:", filePath);

      Settings.setDocumentSettingForKey(
        document,
        "excelTranslateContentFile",
        filePath
      );

      // done
      console.log(
        "Completed. Duplicates: " +
          duplicateKeys +
          " File saved as " +
          path.basename(filePath)
      );
      UI.message(
        "Content file generated. Found " +
          duplicateKeys +
          " duplicated keys. File saved as " +
          decodeURI(path.basename(filePath))
      );
    }
  } catch (error) {
    console.log(error);
    UI.message("error");
  }
}

//
// Sync page
//

export function syncCurrentPage(context) {
  console.time("syncCurrentPage(context)");
  console.log("syncCurrentPage");
  if (contentDocumentExists()) {
    console.log("contentDocumentExists true");
    var contentFile = Settings.documentSettingForKey(
      document,
      "excelTranslateContentFile"
    );

    loadData(contentFile);

    syncContentForPage(document.selectedPage);

    context.document.reloadInspector();
  } else {
    console.log("Content file not found.");
    UI.message("Content file not found.");
  }

  console.timeEnd("syncCurrentPage(context)");
}

export function syncAllPages(context) {
  console.log("syncAllPages");
  if (contentDocumentExists() && document.pages) {
    var contentFile = Settings.documentSettingForKey(
      document,
      "excelTranslateContentFile"
    );
    loadData(contentFile);
    for (let page of document.pages) {
      // Don't add symbols page
      if (page.name !== "Symbols") {
        syncContentForPage(page);
      }
    }
    context.document.reloadInspector();
  } else {
    console.log("Document contains no pages, or content file not found.");
    UI.message("Document contains no pages, or content file not found.");
  }
}

function syncContentForPage(page) {
  console.time("syncContentForPage");
  console.log("syncContentForPage");
  UI.message("Syncing content to page.", page);

  //page
  const pageName = page.name;
  console.log("pageName", pageName);

  // all artboards
  const artBoardLayers = sketch.find("Artboard", page);
  console.log("ArtBoard layers", artBoardLayers.length);

  if (artBoardLayers.length > 0) {
    artBoardLayers.forEach((artBoard) => {
      const artBoardName = artBoard.name;
      console.log("artBoardName", artBoardName);
      UI.message("Syncing content to page.", page, artBoardName);

      // Text
      const textLayers = sketch.find("Text", artBoard);
      console.log("Text layers", textLayers.length);
      if (textLayers.length > 0) {
        textLayers.forEach((textLayer, index) => {
          //sync content
          console.log("textLayer name", textLayer.name);
          UI.message(
            "Syncing content to page.",
            page,
            artBoardName,
            `${index + 1}/${textLayers.length}`
          );

          //1 find Sketch Content
          // console.log("finding", pageName, artBoardName, textLayer.name);

          let result = _.find(contentDictionary, {
            page: pageName,
            artboard: artBoardName,
            key: textLayer.name,
          });
          // console.log("result", result);

          //2 replace text
          if (result) {
            textLayer.text = result.value;
          } else {
            console.log("skipped text", textLayer.name);
          }
        });
      }

      // Symbol
      const symbolLayers = sketch.find("SymbolInstance", artBoard);
      console.log("Symbol layers", symbolLayers.length);
      if (symbolLayers.length > 0) {
        symbolLayers.forEach((symbolLayer, index) => {
          console.log(symbolLayer.name);
          UI.message(
            "Syncing content to page.",
            page,
            artBoardName,
            `${index + 1}/${symbolLayers.length}`
          );

          //for each override
          symbolLayer.overrides.forEach((override) => {
            if (
              override.affectedLayer.type === sketch.Types.Text &&
              override.property === "stringValue"
            ) {
              const textLayerPath = layerNamesFromPath(
                override.path,
                symbolLayer
              );
              console.log("textLayerPath", textLayerPath);

              //2 find Sketch Content
              let key =
                symbolLayer.name + constants.excelDivider + textLayerPath;
              console.log("finding", pageName, artBoardName, key);

              let result = _.find(contentDictionary, {
                page: pageName,
                artboard: artBoardName,
                key: key,
              });
              console.log("result", result);

              //3 replace text
              if (result) {
                override.value = result.value;
              } else {
                console.log("skipped symbol", textLayerPath);
              }

              //4 auto layout
              symbolLayer.resizeWithSmartLayout();
            }
          });
        });
      }
    });
  }
  console.log("pageName", pageName, "Done 🎉");
  UI.message("Syncing content to page.", page, "Done 🎉");
  console.timeEnd("syncContentForPage");
}

export function selectContentFile(context) {
  console.log("selectContentFile");
  var filePaths = dialog.showOpenDialogSync({
    properties: ["openFile"],
    defaultPath: "directory",
    filters: [
      {
        name: "Excel or CSV",
        extensions: ["xlsx", "xls", "csv"],
      },
    ],
  });
  console.log("filePaths", filePaths, filePaths[0], filePaths.length);
  if (filePaths.length) {
    var contentFile = filePaths[0];
    console.log("set contentFile: ", contentFile);
    // set as document key
    Settings.setDocumentSettingForKey(
      document,
      "excelTranslateContentFile",
      contentFile
    );
    UI.message(
      "Content file set! Now you can translate the current page or all pages."
    );
  } else {
    console.log("no file selected");
    UI.message(
      "No file selected. Select Excel or CSV file to continue. Generate a file if you don't have one."
    );
  }
}

function loadData(contentFile) {
  // check filetype
  let fileType = path.extname(contentFile);
  switch (fileType.toLowerCase()) {
    case ".csv":
      console.log("csv");
      // TODO: Support csv files
      // loadCSVData(contentFile)
      break;
    // eslint-disable-next-line no-sequences
    case (".xls", ".xlsx"):
      console.log("Excel");
      loadExcelData(contentFile);
      break;
    default:
      console.log("File format not supported.", fileType.toLowerCase());
      UI.message("File format not supported.");
  }
}

function contentDocumentExists() {
  var contentFile = Settings.documentSettingForKey(
    document,
    "excelTranslateContentFile"
  );
  console.log("contentFile: ", contentFile);
  if (contentFile) {
    if (fs.existsSync(contentFile)) {
      // console.log("file exists: ", contentFile);
      return true;
    }
  } else {
    console.log("No content file.");
    UI.message("Select a content document first.");
    return false;
  }
}

function loadExcelData(contentFile) {
  console.time("loadExcelData");
  console.log("loadExcelData");

  let xlsData = fs.readFileSync(contentFile);

  var workbook = XLSX.read(xlsData, {
    type: "buffer",
  });

  /* Get worksheet. Only support one sheet at the moment. */

  var firstSheetName = workbook.SheetNames[0];

  var worksheet = workbook.Sheets[firstSheetName];

  excelJson = XLSX.utils.sheet_to_json(worksheet, { range: 0, defval: "" });

  console.log("excelJson");
  console.log(excelJson);

  // get language options

  var keyAndLanguageOptions = Object.keys(excelJson[0]); // get language options from first row
  if (!keyAndLanguageOptions) {
    console.log("Key not found");
    UI.message(
      "Language options not found. Ensure the top row contains a 'key' and at least one data 'column'"
    );
    return;
  }
  console.log(keyAndLanguageOptions);

  languageOptions = _.without(keyAndLanguageOptions, "__EMPTY", "key");

  console.timeEnd("loadExcelData");

  // ask for language first so we don't load all language data into the object.
  console.log("showLanguageSelectionPopup");
  console.log(languageOptions);

  let contentFileName = path.basename(contentFile);
  UI.getInputFromUser(
    `Select content column:\n\nContent file:\n${contentFileName}`,
    {
      type: UI.INPUT_TYPE.selection,
      possibleValues: languageOptions,
    },
    (err, value) => {
      if (err) {
        // most likely the user canceled the input
        console.log(err);
        return;
      }
      console.log("selected language", value);

      selectedLanguage = value; // languageOptions[selection[1]]
    }
  );
  if (selectedLanguage) {
    processData();
  }
}

function processData() {
  console.time("processData");
  console.log("processData");

  if (!selectedLanguage) {
    console.log("loadExcelData() aborted. No data columnn selected.");
    return;
  }

  var currentPage = "";
  var currentArboard = "";

  excelJson.forEach((contentObject) => {
    console.log("contentObject", contentObject);
    console.log("contentObject", contentObject.key);
    if (contentObject.key.includes(constants.pagePrefix)) {
      currentPage = contentObject.key.replace(constants.pagePrefix, "");
    }

    if (contentObject.key.includes(constants.artboardPrefix)) {
      currentArboard = contentObject.key.replace(constants.artboardPrefix, "");
    }

    // console.log("contentObject", contentObject.key);
    if (contentObject[`${selectedLanguage}`]) {
      contentDictionary.push({
        page: currentPage,
        artboard: currentArboard,
        key: contentObject.key,
        value: String(contentObject[`${selectedLanguage}`]),
      });
    }
  });

  console.log("contentDictionary");
  console.log(contentDictionary);

  console.timeEnd("processData");
}

// **********************
//   Helper methods
// **********************
// TODO: function duplicated

function layerNameFromLibrary(symbol, layerID) {
  console.log("layerNameFromLibrary", symbol, layerID);
  var nameFromLibary = null;
  if (symbol.overrides && symbol.overrides.length > 0) {
    // const symbolName = `${symbol.name}`;

    symbol.overrides.forEach((override) => {
      if (
        override.affectedLayer.id === layerID &&
        (override.property === "stringValue" ||
          override.property === "symbolID") &&
        override.editable === true
      ) {
        console.log(symbol.name, override.affectedLayer.name);
        nameFromLibary = override.affectedLayer.name;
      }
    });
  }
  console.log("nameFromLibary", nameFromLibary);
  return nameFromLibary;
}

function layerNamesFromPath(path, symbol) {
  console.log("layerNamesFromPath path:", path);
  var layerNames = [];
  let layerIDs = path.split(constants.sketchSymbolDivider);
  for (let layerID of layerIDs) {
    let layer = document.getLayerWithID(layerID);

    if (layer) {
      let layerName = layer.name;
      layerNames.push(layerName);
    } else {
      //library, try find in overrides
      layerNames.push(layerNameFromLibrary(symbol, layerID));
    }
  }
  // filter empty string
  console.log("layerNamesFromPath", layerNames);
  layerNames = layerNames.filter(Boolean);
  return layerNames.join(constants.excelDivider);
}
