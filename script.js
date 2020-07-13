const projectName = "pptx-notes-to-text"
const fs = require("fs").promises;
const path = require("path");
const presentationsPath = path.join(__dirname, "presentations");

/*
const filePromise = readFile(presentationsPath, {encoding: 'utf-8'})
filePromise.then((file)=>{

})
.catch((error) =>{
  throw "Directory must be called 'presentations'. " + error
})
*/

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
  const fileName = files[0]//files.forEach((fileName) => {
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
      .then(async (notesSlidesXml) => {
        console.log(pathToNotes)
        console.log(notesSlidesXml) // notes of the first presentation
        //console.log(notesSlidesXml.length) // how many slides with notes?

        parseNoteSlides(notesSlidesXml)
      })
      .catch((error) => {
        throw (
          "Extracting (unzipped) some_presentation.pptx should create"+
          "folders with a path like "+
          `'somewhere\\on\\your\\harddrive\\${projectName}\\presentations\\ppt\\notesSlides'.\n`+
          "The notesSlides folder should contain files like 'notesSlide1.xml', 'notesSlide2.xml'...\n" +
          "This folder was not found! Unzip the pptx files properly" +
          error
        );
      });
  }
}

function parseNoteSlides(notesSlidesXml){
  notesSlidesXml.
}