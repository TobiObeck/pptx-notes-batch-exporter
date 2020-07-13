const projectName = "pptx-notes-to-text";
const fs = require("fs").promises;
const path = require("path");
const presentationsPath = path.join(__dirname, "presentations");
const xml2js = require("xml2js");

// ## presentation
// ### slide

fs.readdir(presentationsPath, { encoding: "utf8" })
  .then(async (files) => {
    processFiles(files);
  })
  .catch((error) => {
    throw (
      "Directory must be called 'presentations'. " +
      "Put your powerpoint (pptx) files in the 'presentations' folder!" +
      error
    );
  });

function processFiles(files) {
  const fileName = files[0]; //files.forEach((fileName) => {

  const tempPath = path.join(presentationsPath, fileName);
  processExtractedDirectory(tempPath);
  //});
}

async function processExtractedDirectory(tempPath) {
  const stat = await fs.lstat(tempPath);
  if (stat.isDirectory() === true) {
    //console.log(fileName);
    const pathToNotes = path.join(tempPath, "ppt", "notesSlides");

    fs.readdir(pathToNotes, { encoding: "utf8" })
      .then(async (notesSlideXmls) => {
        //console.log(pathToNotes);
        //console.log(notesSlideXmls); // notes of the first presentation
        //console.log(notesSlideXmls.length) // how many slides with notes?

        parseNoteSlides(pathToNotes, notesSlideXmls);
      })
      .catch((error) => {
        throw (
          "Extracting (unzipped) some_presentation.pptx should create" +
          "folders with a path like " +
          `'somewhere\\on\\your\\harddrive\\${projectName}\\presentations\\ppt\\notesSlides'.\n` +
          "The notesSlides folder should contain files like 'notesSlide1.xml', 'notesSlide2.xml'...\n" +
          "This folder was not found! Unzip the pptx files properly" +
          error
        );
      });
  }
}

async function parseNoteSlides(pathToNotes, notesSlideXmlFiles) {

  let fullPresentationText = "";

  // sort files by trailing number
  notesSlideXmlFiles.sort(function(current, next){
    const currNumber = parseInt(extractNumberFromFileName(current), 10)
    const nextNumber = parseInt(extractNumberFromFileName(next), 10)

    if(currNumber < nextNumber) { return -1; }
    if(currNumber > nextNumber) { return 1; }
    return 0;
  })

  // extract text from every xml file
  for (const notesSlideXmlFile of notesSlideXmlFiles) {
    const notesSlideXmlPath = path.join(pathToNotes, notesSlideXmlFile);
    console.log(notesSlideXmlPath);
    try {
      const stat = await fs.lstat(notesSlideXmlPath);
      if (stat.isFile() === true) {
        const xmlString = await fs.readFile(notesSlideXmlPath, {
          encoding: "utf-8",
        });

        const cleanedXmlString = replaceSpecialCharacters(xmlString);
        const notesJsObj = await convertXmlToJsObject(cleanedXmlString);
        const extractedText = extractTextFromJsObj(
          notesJsObj,
          notesSlideXmlFile
        );
        // console.log(extractedText);

        fullPresentationText += extractedText + "\n";
      }
    } catch (error) {
      throw "error while reading some notesSlide{NUMBER}.xml file!!!" + error;
    } 
  }

  writeFile(fullPresentationText);
}

/**
 * while extracting the pptx files with 7zip special characters
 * like & are replaced with an encoded version &amp;
 * The xml to js object library replace
 * maybe it would be better to decode the string with
 * urldecode or decodeURIComponent() or sth like this
 */
function replaceSpecialCharacters(xmlString) {
  return xmlString.replace("&amp;", "&amp;");
  //return xmlString.replace("&amp;", "'&amp;'")
  //return xmlString
}

async function convertXmlToJsObject(xmlString) {
  try {
    return await xml2js
      .parseStringPromise(xmlString /*, options */)
      .then(function (result) {
        return result;
      });
  } catch (error) {
    throw error;
  }
}

function extractTextFromJsObj(notesJsObj, notesSlideXmlFile) {
  const paragraphs =
    notesJsObj["p:notes"]["p:cSld"][0]["p:spTree"][0]["p:sp"][1]["p:txBody"][0][
      "a:p"
    ];

  // console.log(paragraphs);

  const slideNumber = extractNumberFromFileName(notesSlideXmlFile);
  let slideText = `### ${slideNumber}\n\n`;

  console.log(slideText)

  paragraphs.forEach((para) => {
    if (para.hasOwnProperty("a:r") === true) {
      const rows = para["a:r"];
      const textRow = rows.reduce((acc, row) => acc + row["a:t"][0], "");
      
      slideText += textRow;
    }
    //if (para.hasOwnProperty("a:endParaRPr") === true) {    
    //}
    slideText += "\n";
  });

  // slideText += "\n";

  return slideText;
}

function extractNumberFromFileName(notesSlideXmlFile) {
  return notesSlideXmlFile.replace("notesSlide", "").replace(".xml", "");
}

function writeFile(extractedText) {
  const outputPath = path.join(__dirname, "output");
  const outputFilePath = path.join(outputPath, "output.txt");
  outputPath;

  fs.mkdir(outputPath, { recursive: true });

  fs.writeFile(outputFilePath, extractedText, "utf8").catch((error) => {
    throw "error while writing file" + error;
  });
}
