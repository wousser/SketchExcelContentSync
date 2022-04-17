import fs from "@skpm/fs";
import dialog from "@skpm/dialog";

const XLSX = require("xlsx");
var _ = require("lodash");

var sketch = require("sketch/dom");
var constants = require("./constants");
var UI = require("sketch/ui");
var Page = require("sketch/dom").Page;
var path = require("path");
var Settings = require("sketch/settings");

var document = sketch.getSelectedDocument();

var contentDictionary = [];
var languageOptions = [];
var selectedLanguage;
var contentLanguage = "en-US"; // set en-US as standard
var contentFileType = "Excel";
var renameTextLayersFlag = false;
var generatedFileData = [];
var duplicateData = [];
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
  console.log("generateContentFile");
  if (document.pages) {
    // ask for language
    if (askContentLanguage() && askRenameTextLayers() && askFileType()) {
      UI.alert(
        "Creating content document",
        "Press OK to start creating content document\n\nDepending on the number of pages, artboards and content this might take a while..."
      );

      const symbolsPage = Page.getSymbolsPage(document);
      const symbolsPageName = symbolsPage ? symbolsPage.name : "Symbols";
      console.log("symbolsPageName", symbolsPageName);

      for (let page of document.pages) {
        // Don't add Symbols page

        if (page.name !== symbolsPageName) {
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

function getGroupPath(layer) {}

function generateContentForPage(page) {
  console.time("generateContent");
  console.log("page: ", page.name);

  //add page to excel
  addToSheet("", ""); // add empty row
  addToSheet("", ""); // add empty row
  const pageDivider = `${constants.pagePrefix}${page.name}`;
  addToSheet(pageDivider, "");

  // all artboards
  const artBoardLayers = sketch.find("Artboard", page);

  if (artBoardLayers) {
    console.log("ArtBoards: ", artBoardLayers.length);
    artBoardLayers.forEach((artBoard) => {
      const artBoardName = artBoard.name;
      console.log("ArtBoard: ", artBoardName);

      //add artBoard to excel
      const artboardDivider = `${constants.artboardPrefix}${artBoardName}`;
      addToSheet("", ""); // add empty row
      addToSheet(artboardDivider, "");

      // Text
      const textLayers = sketch.find("Text", artBoard);

      if (textLayers) {
        console.log("Text layers: ", textLayers.length);
        textLayers.forEach((textLayer) => {
          console.log("Text layer: ", textLayer.name);
          console.log(textLayer);
          if (textLayer.parent) {
            console.log("Group:", textLayer.parent.name);
          }

          //duplicateCheck
          let isDuplicate = duplicateData.includes(textLayer.name);
          console.log(
            "isDuplicate:",
            isDuplicate,
            textLayer.name,
            duplicateData
          );

          if (isDuplicate) {
            //add ID to layer name
            addToSheet(
              textLayer.name + constants.LayerIdPrefix + textLayer.id,
              textLayer.text
            );
            duplicateData.push(
              textLayer.name + constants.LayerIdPrefix + textLayer.id
            );
            duplicateKeys += 1;
          } else {
            //add to sheet buffer
            addToSheet(textLayer.name, textLayer.text);
            duplicateData.push(textLayer.name);
          }

          //rename with #
          renameLayer(textLayer, isDuplicate);
        });
      }

      // Symbol
      const symbolLayers = sketch.find("SymbolInstance", artBoard);

      if (symbolLayers) {
        console.log("Symbols: ", symbolLayers.length);
        symbolLayers.forEach((symbolLayer) => {
          console.log("Symbol: ", symbolLayer.name);
          //rename with #
          renameLayer(symbolLayer);

          //add to sheet buffer
          let textOverrides = extractTextOverrides(symbolLayer);
          // console.log("textOverrides", textOverrides);
          if (textOverrides) {
            textOverrides.forEach((textOverride) => {
              addToSheet(textOverride.fullPath, textOverride.value);
            });
          }
        });
      }
    });
  }
  console.timeEnd("generateContent");
}

function renameLayer(layer, isDuplicate) {
  if (renameTextLayersFlag) {
    if (isDuplicate) {
      layer.name = layer.name + constants.LayerIdPrefix + layer.id;
    }

    if (layer.name.charAt(0) !== constants.translateLayerPrefix) {
      layer.name = constants.translateLayerPrefix + layer.name;
    }

    console.log("layer renamed", layer.name);
  } else {
    console.log("Rename Layers disabled");
  }
}

function extractTextOverrides(symbol) {
  // console.log("extractTextOverrides");

  const symbolName = `${symbol.name}`;

  var result = [];
  if (symbol.overrides) {
    symbol.overrides.forEach((override) => {
      if (
        override.affectedLayer.type === sketch.Types.Text &&
        override.property === "stringValue"
      ) {
        // console.log(symbol.name);
        const fullPath = `${symbolName}${
          constants.excelDivider
        }${layerNamesFromPath(override.path, symbol)}`;
        // console.log("path", fullPath, override.value);
        // console.log(symbol);
        result.push({
          fullPath: fullPath,
          value: override.value,
        });
      }
    });
  }

  return result;
}

function addToSheet(key, value) {
  // add to array
  let keyValue = new ExcelContent(key, value);
  generatedFileData.push(keyValue);
}

function askRenameTextLayers() {
  var returnValue = false;
  UI.getInputFromUser(
    `Prefix text and symbol layers with '${constants.translateLayerPrefix}', and rename duplicates?`,
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

function askFileType() {
  var returnValue = false;
  UI.getInputFromUser(
    `Filetype: Save as Excel or CSV file?`,
    {
      type: UI.INPUT_TYPE.selection,
      possibleValues: ["Excel", "CSV"],
      initialValue: "Excel",
    },
    (err, value) => {
      if (err) {
        console.log("pressed cancel");
        // most likely the user canceled the input
        return;
      }
      if (value !== "null" && value.length > 1) {
        contentFileType = value;
        console.log("set contentFileType", contentFileType);
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

  var defaultPath = "/";
  if (document.path) {
    let name = decodeURI(path.basename(document.path, ".sketch"));
    let contentFileName = name + "-content-" + dateFormat;
    defaultPath = path.join(
      path.dirname(document.path),
      contentFileName
      // + ".xlsx"
    );
    switch (contentFileType) {
      case "Excel":
        defaultPath += ".xslx";
        break;
      case "CSV":
        defaultPath += ".csv";
        break;
    }
  }

  console.log(defaultPath);
  var filePath = dialog.showSaveDialogSync({
    title: "Export as:",
    message: "Export Sketch content as Excel or CSV file.",
    filters: [
      { name: "Excel", extensions: ["xlsx"] },
      { name: "CSV", extensions: ["csv"] },
    ],
    defaultPath: defaultPath,
  });

  try {
    // check if user want to save the file
    if (filePath) {
      console.log("filePath", filePath);
      let fileType = path.extname(filePath);
      console.log("fileType", fileType);

      let book = XLSX.utils.book_new();
      let sheet = XLSX.utils.json_to_sheet(generatedFileData);
      sheet["B1"].v = contentLanguage;
      XLSX.utils.book_append_sheet(book, sheet, "content");
      var content;
      switch (fileType.toLowerCase()) {
        case ".csv":
          console.log("fileType csv");
          content = CSVFile(book);
          break;
        case (".xls", ".xlsx"):
          console.log("fileType excel");
          content = excelFile(book);
          break;
      }

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

function excelFile(book) {
  // console.log(generatedFileData)

  return XLSX.write(book, {
    type: "buffer",
    bookType: "xlsx",
    bookSST: false,
  });
}

function CSVFile(book) {
  return XLSX.write(book, {
    type: "buffer",
    bookType: "csv",
    bookSST: false,
  });
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

    //ask language / column
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
          console.log("user cancelled input");
          console.log(err, value);
          return;
        }
        console.log("selected language", value);

        selectedLanguage = value; // languageOptions[selection[1]]
        UI.alert(
          "Syncing content",
          "Press OK to start syncing content\n\nDepending on the number of pages, artboards and content this might take a while..."
        );

        if (selectedLanguage) {
          syncContentForPage(document.selectedPage);

          context.document.reloadInspector();
        }
      }
    );
    //ask language / column
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

    //ask language / column
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
          console.log("user cancelled input");
          console.log(err, value);
          return;
        }
        console.log("selected language", value);

        selectedLanguage = value; // languageOptions[selection[1]]
        UI.alert(
          "Syncing content",
          "Press OK to start syncing content\n\nDepending on the number of pages, artboards and content this might take a while..."
        );

        if (selectedLanguage) {
          const symbolsPage = Page.getSymbolsPage(document);
          const symbolsPageName = symbolsPage ? symbolsPage.name : "Symbols";
          for (let page of document.pages) {
            // Don't add symbols page
            if (page.name !== symbolsPageName) {
              syncContentForPage(page);
            }
          }
          context.document.reloadInspector();
        }
      }
    );
    //ask language / column
  } else {
    console.log("Document contains no pages, or content file not found.");
    UI.message("Document contains no pages, or content file not found.");
  }
}

function syncContentForPage(page) {
  console.time("syncContentForPage");
  console.log("syncContentForPage");
  UI.message("Syncing content:", page);

  //page
  const pageName = page.name;
  console.log("Page:", pageName);

  // all artboards
  const artBoardLayers = sketch.find("Artboard", page);

  if (artBoardLayers) {
    console.log("ArtBoards:", artBoardLayers.length);
    artBoardLayers.forEach((artBoard) => {
      const artBoardName = artBoard.name;
      console.log("ArtBoard:", artBoardName);
      UI.message("Syncing content:", page, artBoardName);

      // Text
      const textLayers = sketch.find("Text", artBoard);

      if (textLayers) {
        console.log("Text layers:", textLayers.length);

        console.log(contentDictionary);

        textLayers.forEach((textLayer, index) => {
          //sync content
          console.log("textLayer:", textLayer.name);
          UI.message(
            "Syncing content:",
            page,
            artBoardName,
            `${index + 1}/${textLayers.length}`
          );

          //1 find Sketch Content
          console.log("finding", pageName, artBoardName, textLayer.name);

          let result = _.find(contentDictionary, {
            page: pageName,
            artboard: artBoardName,
            key: textLayer.name,
          });
          // console.log("result", result);

          //2 replace text
          if (result) {
            const resultValue = result.content[`${selectedLanguage}`];
            if (resultValue) {
              textLayer.text = resultValue;
            }
          } else {
            console.log("Skipped text layer:", textLayer.name);
          }
        });
      }

      // Symbol
      const symbolLayers = sketch.find("SymbolInstance", artBoard);

      if (symbolLayers) {
        console.log("Symbols:", symbolLayers.length);
        symbolLayers.forEach((symbolLayer, index) => {
          console.log("Symbol:", symbolLayer.name);
          UI.message(
            "Syncing content:",
            page,
            artBoardName,
            `${index + 1}/${symbolLayers.length}`
          );

          //for each override
          if (symbolLayer.overrides) {
            symbolLayer.overrides.forEach((override) => {
              if (
                override.affectedLayer.type === sketch.Types.Text &&
                override.property === "stringValue"
              ) {
                const textLayerPath = layerNamesFromPath(
                  override.path,
                  symbolLayer
                );
                // console.log("textLayerPath", textLayerPath);

                //2 find Sketch Content
                let key =
                  symbolLayer.name + constants.excelDivider + textLayerPath;
                // console.log("finding", pageName, artBoardName, key);

                let result = _.find(contentDictionary, {
                  page: pageName,
                  artboard: artBoardName,
                  key: key,
                });
                // console.log("result", result);

                //3 replace text
                if (result) {
                  const resultValue = result.content[`${selectedLanguage}`];
                  if (resultValue) {
                    override.value = resultValue;
                  }
                } else {
                  console.log(
                    "Skipped symbol: ",
                    symbolLayer.name,
                    textLayerPath
                  );
                }

                //4 auto layout
                symbolLayer.resizeWithSmartLayout();
              }
            });
          }
        });
      }
    });
  }
  console.log("Page:", pageName, "Done 🎉");
  UI.message("Syncing content:", page, "Done 🎉");
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
  if (filePaths) {
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
      loadExcelData(contentFile);
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

  languageOptions = _.without(keyAndLanguageOptions, "key");
  languageOptions = _.remove(languageOptions, function (languageOption) {
    return !languageOption.includes("__EMPTY");
  });
  // languageOptions = languageOptions.filter(
  //   ({ languageOption }) => !languageOption.includes("__EMPTY")
  // );

  console.timeEnd("loadExcelData");

  // ask for language first so we don't load all language data into the object.
  // console.log("showLanguageSelectionPopup");
  // console.log(languageOptions);

  processData();
}

function processData() {
  console.time("processData");
  console.log("processData");

  var currentPage = "";
  var currentArboard = "";
  if (excelJson) {
    excelJson.forEach((contentObject) => {
      // console.log("contentObject", contentObject);
      // console.log("contentObject", contentObject.key);
      if (contentObject.key.includes(constants.pagePrefix)) {
        currentPage = contentObject.key.replace(constants.pagePrefix, "");
      }

      if (contentObject.key.includes(constants.artboardPrefix)) {
        currentArboard = contentObject.key.replace(
          constants.artboardPrefix,
          ""
        );
      }

      //all content
      var allContent = {};
      languageOptions.forEach((languageOption) => {
        if (contentObject[languageOption]) {
          allContent[languageOption] = String(contentObject[languageOption]);
        }
      });

      // console.log("contentObject", contentObject.key);

      contentDictionary.push({
        page: currentPage,
        artboard: currentArboard,
        key: contentObject.key,
        content: allContent,
      });
    });
  }

  console.log("contentDictionary");
  console.log(contentDictionary);

  console.timeEnd("processData");
}

// **********************
//   Helper methods
// **********************
// TODO: function duplicated

function layerNameFromLibrary(symbol, layerID) {
  // console.log("layerNameFromLibrary", symbol, layerID);
  var nameFromLibary = null;
  // const symbolName = `${symbol.name}`;

  if (symbol.overrides) {
    symbol.overrides.forEach((override) => {
      if (
        override.affectedLayer.id === layerID &&
        (override.property === "stringValue" ||
          override.property === "symbolID") &&
        override.editable === true
      ) {
        // console.log(symbol.name, override.affectedLayer.name);
        nameFromLibary = override.affectedLayer.name;
      }
    });
  }

  // console.log("nameFromLibary", nameFromLibary);
  return nameFromLibary;
}

function layerNamesFromPath(path, symbol) {
  // console.log("layerNamesFromPath path:", path);
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
  // console.log("layerNamesFromPath", layerNames);
  layerNames = layerNames.filter(Boolean);
  return layerNames.join(constants.excelDivider);
}
