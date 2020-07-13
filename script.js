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
  processUnzippedDirectory(tempPath);
  //});
}

async function processUnzippedDirectory(tempPath) {
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

function parseNoteSlides(pathToNotes, notesSlideXmls) {
  const i = 1
  const notesSlideXmlPath = path.join(pathToNotes, notesSlideXmls[i]); // FIRST SLIDE ONLY
  console.log(notesSlideXmlPath);

  fs.readFile(notesSlideXmlPath, { encoding: "utf-8" })
    .then(async (xmlString) => {
      const stat = await fs.lstat(notesSlideXmlPath);
      if (stat.isFile() === true) {
        const cleanedXmlString = replaceSpecialCharacters(xmlString);
        const notesJsObj = await convertXmlToJsObject(cleanedXmlString);
        const extractedText = extractTextFromJsObj(notesJsObj, notesSlideXmls[i]);
        console.log(extractedText);
      }
    })
    .catch((error) => {
      throw error;
    });
}

/**
 * while extracting the pptx files with 7zip special characters 
 * like & are replaced with an encoded version &amp;
 * The xml to js object library replace
 * maybe it would be better to decode the string with
 * urldecode or decodeURIComponent() or sth like this
 */
function replaceSpecialCharacters(xmlString){
  return xmlString.replace("&amp;", '\&amp;')
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

function extractTextFromJsObj(notesJsObj, notesSlideXml) {
  const paragraphs =
    notesJsObj["p:notes"]["p:cSld"][0]["p:spTree"][0]["p:sp"][1]["p:txBody"][0][
      "a:p"
    ];

  console.log(paragraphs);

  let slideText = `### ${notesSlideXml}\n\n`
  
  paragraphs.forEach((para) => {
    if (para.hasOwnProperty("a:r") === true) {
      const rows = para["a:r"];
      const textRow = rows.reduce((acc, row) => acc + row["a:t"][0], "");
      //console.log("textRow", textRow);
      slideText += textRow // + '\n'
    }
    if (para.hasOwnProperty("a:endParaRPr") === true) {
       // slideText += '\n'
    }
    slideText += '\n'
  });

  slideText += "\n"

  return slideText

  // //a:r[1]/a:t[1]/text()[1]
}

