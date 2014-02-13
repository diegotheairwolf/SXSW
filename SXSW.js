//var ALLMOVIES = DocumentApp.create("All Movies");
var ALLMOVIES = new Array();


/**
 * A special function that inserts a custom menu when the spreadsheet opens.
 */
function onOpen() {
  var menu = [{name: 'Run', functionName: 'runSXSW'}];
  SpreadsheetApp.getActive().addMenu('SXSW', menu);
}


function runSXSW() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];

  // Get the range of cells that store movie data.
  var movieDataRange = ss.getRangeByName("movies");
  //var headerDataRange = ss.getRangeByName("header")

  // For every row of movie data, generate an movie object.
  var movieObjects = getRowsData(sheet, movieDataRange);
    
  // Create folders if inexistent.
  var foldersExist = false;
  var folderList = DriveApp.getFolders();
  while (folderList.hasNext()) { 
   var temp = folderList.next();
   if (temp.getName() === 'Forms' || temp === 'Docs' || temp === 'Sheets'){
     foldersExist = true;
   }
  }
  if(!foldersExist){
    var folders = createFolders();
    
  }
  
  var allMoviesDocId = createAllMovies();
  
  /*
  // If file is inexistent, create is
  // If file exists, delete the existing one, and create a new one
  var fileExists = false;
  var fileList = DriveApp.getFilesByType(MimeType.GOOGLE_DOCS);
  while(fileList.hasNext()){
    var tempFile = fileList.next();
    Logger.log(tempFile.getName())
    var tempFileString = tempFile.toString();
    if (tempFile.getName() === 'All Movies'){
      fileExists = true;
      var allMoviesDocId = createAllMovies();
      var tempFileTwo = DocsList.getFileById(tempFile.getId());
      tempFileTwo.removeFromFolder(DocsList.getRootFolder());
    }
  }
  if(!fileExists){
    var allMoviesDocId = createAllMovies();
  }
  
  */
  
  
  
  //This is the loop to create the form for all movies
  for(var i in movieObjects){
    var movie = movieObjects[i];
    setUpForm(movie, sheet, i, folders);
  }
  
  
  // Sorts ALLMOVIES
  ALLMOVIES.sort(function(a,b){
    var textA = a.filmName.toUpperCase();
    var textB = b.filmName.toUpperCase();
    return (textA < textB) ? -1 : (textA > textB) ? 1 : 0;
  });
  
      
  // Populate ALLMOVIES DOC
  var allMoviesDoc = DocumentApp.openById(allMoviesDocId);
  var body = allMoviesDoc.getBody();
  var numChildren = body.getNumChildren();
  
 /* for(var i = numChildren; i>0; i--){
    var tempChild = body.getChild(i-1);
    body.removeChild(tempChild);
  }
  */
  
  var child = body.getChild(0);
  
  
  for(var i in ALLMOVIES){
      
   // Append a image
   var blob = UrlFetchApp.fetch(ALLMOVIES[i].qrCodeUrl);
    
   // Append a document header paragraph. 
   var para2 = body.appendParagraph(ALLMOVIES[i].filmName + '\n\n');
   var para2text = para2.editAsText();
    para2text.setBold(true);
    para2text.setFontSize(20);
    para2.appendInlineImage(blob);
   
   // Append a section header paragraph.
   var para3 = body.appendParagraph(shortenUrl(ALLMOVIES[i].url));
   var para3text = para3.editAsText();
    para3text.setFontSize(16);    
  
   body.appendPageBreak();
    
  }
 
  
}



/**
 * Creates a Google Form that allows respondents to select which conference
 * sessions they would like to attend, grouped by date and start time.
 *
 * @param {Object} movie contains the data per row of spreadsheet.
 */
function setUpForm(movie, sheet, row, folders) {
  
 ALLMOVIES.push(movie);

 var form;
 var temporaryForm;
 var startRow = 2;
  
  
 /******* 
 *
 * Evrything that is done for the first time and just once!
 *
 ********/
 
  if(!movie.destinationId){
    
   
   
   // Create form, sheet, and doc
   temporaryForm = FormApp.create(movie.filmName);
   var temporarySheet = SpreadsheetApp.create(movie.filmName);
   var temporaryDoc = DocumentApp.create(movie.filmName);
   
   // Parse form, sheet, and doc into file type
   var tempForm = DocsList.getFileById(temporaryForm.getId());
   var tempSheet = DocsList.getFileById(temporarySheet.getId());
   var tempDoc = DocsList.getFileById(temporaryDoc.getId());
   
   // Make a copy of file
   form = tempForm.makeCopy(movie.filmName);
   var ss = tempSheet.makeCopy(movie.filmName);  
   var doc = tempDoc.makeCopy(movie.filmName); 
   
   //Add folder, sheet, and doc copies to respective folder and remove from root
   form.addToFolder(DocsList.getFolderById(folders.forms.getId()));
   form.removeFromFolder(DocsList.getRootFolder());
   tempForm.removeFromFolder(DocsList.getRootFolder());
   ss.addToFolder(DocsList.getFolderById(folders.sheets.getId()));
   ss.removeFromFolder(DocsList.getRootFolder());
   tempSheet.removeFromFolder(DocsList.getRootFolder());
   doc.addToFolder(DocsList.getFolderById(folders.docs.getId()));
   doc.removeFromFolder(DocsList.getRootFolder());
   tempDoc.removeFromFolder(DocsList.getRootFolder());
   
   // Parse files back into form, sheet, and doc
   form = FormApp.openById(form.getId());
   ss = SpreadsheetApp.openById(ss.getId());
   doc = DocumentApp.openById(doc.getId());
      
   form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
   form.setPublishingSummary(true);
        
   // Delete first sheet so that we only have the response form sheet
   ss.setActiveSheet(ss.getSheets()[0]);
   ss.deleteActiveSheet();
 
   movie.destinationId = form.getDestinationId();
   movie.url = form.getPublishedUrl();
   movie.formId = form.getId();
   movie.qrCodeUrl = make_QR(movie.url);
 
 
   /*  
   **************************************
   Set info into the database spreadsheet
   **************************************
   */
   sheet.getRange(+row + +startRow, 11).setValue(movie.destinationId);
   sheet.getRange(+row + +startRow, 12).setValue(movie.url);
   sheet.getRange(+row + +startRow, 13).setValue(movie.formId);
   sheet.getRange(+row + +startRow, 14).setValue(movie.qrCodeUrl);
   //sheet.insertImage(movie.qrCodeUrl, 10, +row + +startRow);
   //sheet.getRange(+row + +startRow, 10).setFormula('=image(\"' + movie.qrCodeUrl + '\",1)');
  
   
   
   /*
   ****************************
   Set info into the google doc
   ****************************
   */
   
   doc.setName(movie.filmName);
      
   // Append a document header paragraph.
   var header = doc.appendParagraph(movie.filmName);
   header.setHeading(DocumentApp.ParagraphHeading.HEADING1).setBold(true);
   //var headerALLMOVIES = doc.appendParagraph(movie.filmName);
   //header.setHeading(DocumentApp.ParagraphHeading.HEADING1).setBold(true);
   
   // Append a section header paragraph.
   var section1 = doc.appendParagraph(shortenUrl(movie.url));
   section1.setHeading(DocumentApp.ParagraphHeading.HEADING2);
   
   // Append a image
   var blob = UrlFetchApp.fetch(movie.qrCodeUrl);
   doc.getChild(0).asParagraph().appendInlineImage(blob);
   
    
    
    
   /*******
   *
   *  Fill form with information
   *
   *******/
  
   form.setDescription(' \n ' + 
                       ' \n Format: ' + movie.format +  
                       ' \n DCP Drive Type: ' + movie.dcpDriveType + 
                       ' \n DCP File Size: ' + movie.dcpFileSize +
                       ' \n DCP Encription: ' + movie.dcpEncryption + 
                       ' \n Aspect Ratio: ' + movie.aspectRatio + 
                       ' \n Sound Format: ' + movie.soundFormat +
                       ' \n Runtime: ' + movie.runtime +                        
                       ' \n Exhibition Notes: ' + movie.exhibitionNotes +
                       ' \n Backup: ' + movie.backup
                      );
 
   //form.addTextItem().setTitle("Confirmed Technical Info");
 
   var action = form.addMultipleChoiceItem()
    .setRequired(true)
    .setTitle('Action')
    .setChoiceValues(['Transit', 'Received', 'Ingested', 'Played',
                    'Deleted', 'ERROR!'])
    .showOtherOption(true);

   var location = form.addMultipleChoiceItem()
    .setRequired(true)
    .setTitle('Location')
    .setChoiceValues(['Tech Center', 'Roadies', 'TBD Post', 'Vimeo ACC', 'Paramount'
                     , 'Stateside', 'Ritz', 'AMC @VCC', 'Topfer'
                     , 'Rollins', 'Slaughter', 'Village', 'Marchesa'])
    .showOtherOption(true);
  
   var comments = form.addParagraphTextItem();
    comments.setTitle('Comments'); 

 }
  
 /*
  
  /**********
  *
  *  Everything that happens after the first time
  *
  ***********/
  
  else{
    
   //ALLMOVIES.push(movie);
    
   form = FormApp.openById(movie.formId);
   form.setDescription(' \n "' + 
                       ' \n Format: ' + movie.format +  
                       ' \n Backup Format: ' + movie.backupFormat + 
                       ' \n DCP Drive Type: ' + movie.dcpDriveType + 
                       ' \n DCP File Size: ' + movie.dcpFileSice +
                       ' \n DCP Encription?: ' + movie.dcpEncryption + 
                       ' \n DCP Notes: ' + movie.dcpNotes +
                       ' \n Aspect Ratio: ' + movie.aspectRatio + 
                       ' \n Sound Format: ' + movie.soundFormat +
                       ' \n Runtime: '
                      );
   
 }
    
 
  
 
}



/**
 * A trigger-driven function.
 *
 * @param {Object} e The event parameter for form submission to a spreadsheet;
 *     see https://developers.google.com/apps-script/understanding_events
 */
function onFormSubmit(e) {
  var response = e.values();
  
  var values = SpreadsheetApp.getActive().getSheetByName('Form Response 1')
     .getDataRange().getValues();
  
  Browser.msgBox(values.toString());
  
  
  

}





/**
 * 
 * @param {String} urlToShorten - contains the string with the url to be shortened
 * return url.id - contains the shortened url
 * see https://developers.google.com/apps-script/advanced/url-shortener
 */
function shortenUrl(urlToShorten) {
  var url = UrlShortener.Url.insert({
    longUrl: urlToShorten
  });
  return url.id;
}

/**
 *
 * @param {String} url - Contains a string with the url to be converted into qrCode
 * return image_url    - qrCode image
 * see https://google-developers.appspot.com/chart/infographics/docs/overview
 */
function make_QR( url ) {
  var size = 150 // The height and width needed.
  var encoded_url = encodeURIComponent( url )  
  var image_url = "http://chart.googleapis.com/chart?chs=" + size + "x" + size + "&cht=qr&chl=" + encoded_url
  return image_url
}



/**
 * Creates three folders: 'forms', 'sheets', 'docs', where files are going to be stored
 * 
 * @param {} 
 * return 
 * see https://google-developers.appspot.com/chart/infographics/docs/overview
 */
function createFolders(){
  var folderForms = DriveApp.createFolder("Digital STR's");
  var folderSheets = DriveApp.createFolder("Response Sheets");
  var folderDocs = DriveApp.createFolder("QR Codes");
  
  var folders = new Object();
    folders.forms = folderForms;
    folders.sheets = folderSheets;
    folders.docs = folderDocs;
  
  return folders;
}


/**
 * Creates one doc where all the docs are to be stored
 * 
 * @param {} 
 * return document Id
 * see https://google-developers.appspot.com/chart/infographics/docs/overview
 */
function createAllMovies(){
  var document = DocumentApp.create("All Movies");
  
  return document.getId();
}






// getRowsData iterates row by row in the input range and returns an array of objects.
// Each object contains all the data for a given row, indexed by its normalized column name.
// Arguments:
//   - sheet: the sheet object that contains the data to be processed
//   - range: the exact range of cells where the data is stored
//   - columnHeadersRowIndex: specifies the row number where the column names are stored.
//       This argument is optional and it defaults to the row immediately above range; 
// Returns an Array of objects.
function getRowsData(sheet, range, columnHeadersRowIndex) {
  columnHeadersRowIndex = columnHeadersRowIndex || range.getRowIndex() - 1;
  var numColumns = range.getLastColumn() - range.getColumn() + 1;
  var headersRange = sheet.getRange(columnHeadersRowIndex, range.getColumn(), 1, numColumns);
  var headers = headersRange.getValues()[0];
  //Browser.msgBox(headers.toSource());
  return getObjects(range.getValues(), normalizeHeaders(headers));
  
}


// For every row of data in data, generates an object that contains the data. Names of
// object fields are defined in keys.
// Arguments:
//   - data: JavaScript 2d array
//   - keys: Array of Strings that define the property names for the objects to create
function getObjects(data, keys) {
  var objects = [];
  for (var i = 0; i < data.length; ++i) {
    var object = {};
    var hasData = false;
    for (var j = 0; j < data[i].length; ++j) {
      var cellData = data[i][j];
      if (isCellEmpty(cellData)) {
        continue;
      }
      object[keys[j]] = cellData;
      hasData = true;
    }
    if (hasData) {
      objects.push(object);
    }
  }
  return objects;
}

// Returns an Array of normalized Strings.
// Arguments:
//   - headers: Array of Strings to normalize
function normalizeHeaders(headers) {
  var keys = [];
  for (var i = 0; i < headers.length; ++i) {
    var key = normalizeHeader(headers[i]);
    if (key.length > 0) {
      keys.push(key);
    }
  }
  return keys;
}

// Normalizes a string, by removing all alphanumeric characters and using mixed case
// to separate words. The output will always start with a lower case letter.
// This function is designed to produce JavaScript object property names.
// Arguments:
//   - header: string to normalize
// Examples:
//   "First Name" -> "firstName"
//   "Market Cap (millions) -> "marketCapMillions
//   "1 number at the beginning is ignored" -> "numberAtTheBeginningIsIgnored"
function normalizeHeader(header) {
  var key = "";
  var upperCase = false;
  for (var i = 0; i < header.length; ++i) {
    var letter = header[i];
    if (letter == " " && key.length > 0) {
      upperCase = true;
      continue;
    }
    if (!isAlnum(letter)) {
      continue;
    }
    if (key.length == 0 && isDigit(letter)) {
      continue; // first character must be a letter
    }
    if (upperCase) {
      upperCase = false;
      key += letter.toUpperCase();
    } else {
      key += letter.toLowerCase();
    }
  }
  return key;
}

// Returns true if the cell where cellData was read from is empty.
// Arguments:
//   - cellData: string
function isCellEmpty(cellData) {
  return typeof(cellData) == "string" && cellData == "";
}

// Returns true if the character char is alphabetical, false otherwise.
function isAlnum(char) {
  return char >= 'A' && char <= 'Z' ||
    char >= 'a' && char <= 'z' ||
    isDigit(char);
}

// Returns true if the character char is a digit, false otherwise.
function isDigit(char) {
  return char >= '0' && char <= '9';
}

// Given a JavaScript 2d Array, this function returns the transposed table.
// Arguments:
//   - data: JavaScript 2d Array
// Returns a JavaScript 2d Array
// Example: arrayTranspose([[1,2,3],[4,5,6]]) returns [[1,4],[2,5],[3,6]].
function arrayTranspose(data) {
  if (data.length == 0 || data[0].length == 0) {
    return null;
  }

  var ret = [];
  for (var i = 0; i < data[0].length; ++i) {
    ret.push([]);
  }

  for (var i = 0; i < data.length; ++i) {
    for (var j = 0; j < data[i].length; ++j) {
      ret[j][i] = data[i][j];
    }
  }

  return ret;
}
