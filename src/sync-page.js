import fs from "@skpm/fs";
import dialog from "@skpm/dialog";

const XLSX = require("xlsx");

var sketch = require("sketch/dom");
var constants = require("./constants");
var UI = require("sketch/ui");
var path = require("path");
var Settings = require("sketch/settings");

var document = sketch.getSelectedDocument();

var contentDictionary = {};
var languageOptions = [];
var selectedLanguage;

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

    populatePage();
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
        populatePage(page);
      }
    }
    context.document.reloadInspector();
  } else {
    console.log("Document contains no pages, or content file not found.");
    UI.message("Document contains no pages, or content file not found.");
  }
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
      console.log("file exists: ", contentFile);
      return true;
    }
  } else {
    console.log("No content file.");
    UI.message("Select a content document first.");
    return false;
  }
}

function showLanguageSelectionPopup(languageOptions) {
  let contentFile = Settings.documentSettingForKey(
    document,
    "excelTranslateContentFile"
  );
  let contentFileName = path.basename(contentFile);
  UI.getInputFromUser(
    `Sync to language?\n\nContent file:\n${contentFileName}`,
    {
      type: UI.INPUT_TYPE.selection,
      possibleValues: languageOptions,
    },
    (err, value) => {
      if (err) {
        // most likely the user canceled the input
        return;
      }
      console.log("selected language");
      console.log(value);
      selectedLanguage = value; // languageOptions[selection[1]]
    }
  );
}

function loadExcelData(contentFile) {
  console.time("loadExcelData");
  console.log("loadExcelData");

  console.time("loadExcelData 1");
  let xlsData = fs.readFileSync(contentFile);
  console.timeEnd("loadExcelData 1");

  console.time("loadExcelData 2");
  var workbook = XLSX.read(xlsData, {
    type: "buffer",
  });
  console.timeEnd("loadExcelData 2");

  /* Get worksheet. Only support one sheet at the moment. */
  console.time("loadExcelData 3");
  var firstSheetName = workbook.SheetNames[0];
  console.timeEnd("loadExcelData 3");

  console.time("loadExcelData 4");
  var worksheet = workbook.Sheets[firstSheetName];
  console.timeEnd("loadExcelData 4");

  console.time("loadExcelData 5");
  var excelJson = XLSX.utils.sheet_to_json(worksheet, { range: 0, defval: "" });
  console.timeEnd("loadExcelData 5");

  console.log("excelJson");
  console.log(excelJson);
  let rowNumber = 2; // Excel row starts at 0, and 1st row is key/value

  // get language options
  console.time("loadExcelData 6");
  var keyAndLanguageOptions = Object.keys(excelJson[0]); // get language options from first row
  if (!keyAndLanguageOptions) {
    console.log("Key not found");
    UI.message(
      "Language options not found. Ensure the top row contains a 'key' and at least one data 'column'"
    );
    return;
  }
  console.log(keyAndLanguageOptions);
  console.timeEnd("loadExcelData 6");

  console.time("loadExcelData 7");
  languageOptions = keyAndLanguageOptions;

  // remove 'key', first item
  languageOptions.shift();

  //remove xlsx empty options
  languageOptions = languageOptions.filter((ele) => !ele.includes("__EMPTY"));
  console.timeEnd("loadExcelData 7");

  console.timeEnd("loadExcelData");
  // ask for language first so we don't load all language data into the object.
  console.log("showLanguageSelectionPopup");
  console.log(languageOptions);
  showLanguageSelectionPopup(languageOptions);
  if (!selectedLanguage) {
    console.log("loadExcelData() aborted. No data columnn selected.");
    return;
  }

  console.time("contentDictionary");

  var currentArboardName = "";

  for (var row in excelJson) {
    const rowData = excelJson[row];
    console.log("row", rowData);

    // Save current artboard
    if (rowData["key"].includes(constants.artboardPrefix)) {
      currentArboardName = rowData["key"];
      console.log("updated current arboard" + currentArboardName);
    } else {
      // skip empty content
      if (rowData[selectedLanguage]) {
        // console.log("rowNumber: " + rowNumber, currentArboardName);
        contentDictionary[String(currentArboardName + "." + rowData["key"])] =
          String(rowData[selectedLanguage]);
      } else {
        console.log("skipped rowNumber: " + rowNumber);
      }
    }

    rowNumber += 1;
  }
  console.log("contentDictionary");
  // console.log(contentDictionary);
  onComplete();
  console.timeEnd("contentDictionary");
}

function analyzeLayer(layers) {
  for (let layer of layers) {
    console.log(layer.name, layer.type);
    switch (layer.type) {
      case String(sketch.Types.Group):
        console.log("group layer");
        analyzeLayer(layer.layers);
        break;
      case String(sketch.Types.SymbolInstance):
        console.log("symbol layer");
        updateSymbolLayer(layer);
        break;
      case String(sketch.Types.Text):
        console.log("text layer");
        updateTextLayer(layer);
        break;
      case String(sketch.Types.Artboard):
        console.log("artboard layer");
        analyzeLayer(layer.layers);
        break;
    }
  }
}

function populatePage(page) {
  console.time("populatePage(page)");
  // abort if no language is chosen
  if (!selectedLanguage) {
    console.log("populatePage() aborted. No language selected.");
    return;
  }

  // Use selected page if no page is set
  if (!page) {
    page = document.selectedPage;
    console.log("page: " + page.name);
  }

  console.log("page layers: " + page.layers.length);
  analyzeLayer(page.layers);
  onComplete();
  console.timeEnd("populatePage(page)");
}

function onComplete() {
  console.log("Completed");
  UI.message("Completed");
}

function updateTextLayer(layer) {
  console.time("updateTextLayer");
  console.log("updateTextLayer");

  // Check for layers outside artboard
  if (layer.getParentArtboard() === undefined) {
    console.log("getParentArtboard() undefined");
  } else {
    let artboardAndLayerName =
      constants.artboardPrefix +
      layer.getParentArtboard().name +
      "." +
      layer.name;
    console.log(artboardAndLayerName);
    if (contentDictionary[artboardAndLayerName]) {
      layer.text = contentDictionary[artboardAndLayerName];
      console.log("new value", layer.name, artboardAndLayerName, layer.text);
    }
    console.log("updateTextLayer done");
  }
  console.timeEnd("updateTextLayer");
}

function updateSymbolLayer(symbol) {
  console.time("updateSymbolLayer");
  console.log("updateSymbolLayer");
  if (symbol.getParentArtboard() === undefined) {
    console.log("getParentArtboard() undefined");
  } else {
    let artboardAndSymbolName =
      constants.artboardPrefix +
      symbol.getParentArtboard().name +
      "." +
      symbol.name;
    console.log(artboardAndSymbolName);

    for (let override of symbol.overrides) {
      if (override.property === "stringValue") {
        let layerNameAndOverride =
          artboardAndSymbolName +
          constants.excelDivider +
          layerNamesFromPath(override.path);

        if (contentDictionary[layerNameAndOverride]) {
          override.value = contentDictionary[layerNameAndOverride];
        }
      }
    }
    console.log("updateSymbolLayer done");
  }
  console.timeEnd("updateSymbolLayer");
}

// **********************
//   Helper methods
// **********************
// TODO: function duplicated
function layerNamesFromPath(path) {
  var layerNames = [];
  let layerIDs = path.split(constants.sketchSymbolDivider);
  for (let layerID of layerIDs) {
    let layer = document.getLayerWithID(layerID);

    // TODO: Sketch libraries not supported yet.
    if (layer) {
      let layerName = layer.name;
      layerNames.push(layerName);
    }
  }

  return layerNames.join(constants.excelDivider);
}
