// -- Trip Choices - By achiu@vsa.edu.hk --
// This App allows registered students to indicate their preferences for trips and trip buddy.
//

// -- GLOBALS --
var myListDocID = '1oZLuCFaxrcNgnd8ASTw-13tiCDm_OGhCXneNjC3X86c'; 

const emailTemplateId = '1bzdo9Uxe22nyamlETnXr2wqIE702Z9nV-XGBxzDuqDg'; // For sending Allocation Emails

//Spreadsheet sheet names for Get Lists...
var myListSheetName = 'Choices';
//var myStudentDataSheetName = 'StudentList'; //Sheet for authorising students
var myApproverListSheetName = 'Approver';

//Spreadsheet for saving student choices and Teacher Recommendations. 
var mySurveyCollector = myListDocID; 
//var mySurveyCollector = '1LsiB1BFZgc-RjmQ0zFUw4cNvAkyz7nSnKqLmPT70tCY'; //Trial DEV Spreadsheet


var mySurveySheetName = 'LiveResults'; //Results sheet 
//var mySurveySheetName = 'Results-V2'; //Current results sheet with 2 Adv List Choices (obsolete - moved over to "results")
var mySurveyTripCounts = 'LiveCounts';
var buddyListSheetName = "BuddyList";

//Get User


var thisUser = Session.getActiveUser().getEmail(); //Logged In User
//var thisUser = "frwong@vsa.edu.hk" //Logged In User

//Globals for displaying course options in each block

var DisableDateLocking = false; //turns on or off the date lock out.
var showSecondChoice = false; //turns on or off for displaying a second choice.
var maxTripChoices = 3; //How many trip choices / checkboxes can users indicate as chosen?
var extraDataFields = ["T25HK", "chosenBuddyId"]; //Field names to be included and collected from user

//For HTMLService app (from template)
//var userSheetName = 'Members';
var challengeSheetName = 'Trip Preference';
var appTitle = "Indicating your trip preference";

  var approverEmailCol = 0;
  var studentEmailCol = 0;
  var parentEmailCol = 4;
  var advisorEmailCol = 10;
  var endDateCol = 9; //Survey end date col on student sheet
  var startDateCol = 8; //survey start date col on student sheet
  var canPostCol = 7;

//
// -----
//  doGet - main function for web app
// -----
function doGet(){
  if(thisUser){
    var myDoc = 'index';  
  } else {
    var myDoc = 'notLoggedIn';
  }
  
  return HtmlService.createTemplateFromFile(myDoc).evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME);
}
  
// ----
// function loadGInfo() - Gets all info from spreadsheets and passes it back to the client
// -----
function loadGInfo() {

  var ListDoc = SpreadsheetApp.openById(myListDocID);
  var SurveyDoc = ListDoc;
  //var SurveyDoc = SpreadsheetApp.openById(mySurveyCollector);
  
  //Check are we adventure week approver, HRM teacher or student.
  var userType = {isStudent: false, isAdvisor: false, isApprover: false, isParent: false};
  var error = {status: false, class: "bg-danger", msg: "Error"};
  
  var approverList = ListDoc.getSheetByName(myApproverListSheetName).getDataRange().getValues(); 
  //var studentList = ListDoc.getSheetByName(myStudentDataSheetName).getDataRange().getValues();
  var choicesList = SurveyDoc.getSheetByName(mySurveySheetName).getDataRange().getValues();
  var buddySheet = ListDoc.getSheetByName(buddyListSheetName);
  var buddyList = arrayToObjects(buddySheet.getRange(1,1,buddySheet.getLastRow(),5).getValues()); 


  var myApprover = getRowsMatching(approverList, approverEmailCol, thisUser);
  var myAdvisor = getRowsMatching(choicesList, advisorEmailCol, thisUser);
  var myStudentInfo = getRowsMatching(choicesList, studentEmailCol, thisUser);
  var myChildInfo = getRowsMatching(choicesList, parentEmailCol, thisUser);
  var studentInfo = [];
  var choiceData = [];
  var canPost = false; 
  
  userType.isApprover = (myApprover.length > 0);//We are an approver
  userType.isAdvisor = (myAdvisor.length > 0); //We are an advisor
  userType.isStudent = (myStudentInfo.length > 0); //We are a student
  userType.isParent = (myChildInfo.length > 0); //We are a parent
      Logger.log(myStudentInfo);
  
  // Check if user has no role at all
  if (!userType.isApprover && !userType.isAdvisor && !userType.isStudent && !userType.isParent) {
    error = {status: true, class: "bg-dark text-light", msg: "You do not have access to this trip system. Please contact the administrators if you believe this is an error."};
  }
  
  if (userType.isApprover) { //We are an approver

    studentInfo = arrayToObjects(choicesList);
    //Logger.log(studentInfo);
    
  } else if (userType.isAdvisor) { //we are an advisor
    studentInfo = arrayToObjects(choicesList, "advisorEmail", thisUser);
    //canPost = checkCanPost(myStudentInfo); 
  } else if (userType.isStudent) { //we are a student
    canPost = checkCanPost(myStudentInfo); 
    studentInfo = arrayToObjects(choicesList,"email",thisUser);
    //Logger.log(canPost);
    //Logger.log(studentInfo);
    //if(!myStudentInfo[0][canPostCol]) error = {status: true, class: "bg-info", msg: "This survey is not open to you at this time."};
  } else if (userType.isParent) { //we are a parent
    studentInfo = arrayToObjects(choicesList,"parentEmail", thisUser);
    
  }
  
  // Check date locking and update canPost values for each student
  if (!DisableDateLocking && studentInfo.length > 0) {
    for (var i = 0; i < studentInfo.length; i++) {
      var student = studentInfo[i];
      if (student.postStartDate && student.postEndDate) {
        var timeOpen = timeCheck(new Date(student.postStartDate), new Date(student.postEndDate));
        student.canPost = student.canPost && timeOpen; // Only allow posting if both original canPost and time window are valid
      }
    }
  }
  
  //Get a list of possible choices for this user
  var possibleChoices = [];
  var possibleChoiceList = textifyDates(ListDoc.getSheetByName(myListSheetName).getDataRange().getValues().splice(1));
  /* for (var pc = 0; pc < possibleChoiceList.length; pc++){
    possibleChoices.push(possibleChoiceList[pc][0]);
  } */
  var tripCounts = arrayToObjects(SurveyDoc.getSheetByName(mySurveyTripCounts).getDataRange().getValues());
  
  var refreshTime = Utilities.formatDate(new Date(), "GMT+08:00", "dd-MMM-yyyy hh:mm:ss")
  
  //Logger.log({studentInfo: studentInfo, tripCounts: tripCounts, buddyList: buddyList, userType: userType, canPost: canPost, error: error,  refreshed: refreshTime, maxSelections: maxTripChoices, user: thisUser, dataFields: extraDataFields})
  return {studentInfo: studentInfo, tripCounts: tripCounts, buddyList: buddyList, userType: userType, canPost: canPost, error: error,  refreshed: refreshTime, maxSelections: maxTripChoices, user: thisUser, dataFields: extraDataFields};
}

function getTripCounts(){
  var SurveyDoc = SpreadsheetApp.openById(mySurveyCollector);
  
  
  return tripCounts;
  
}

function checkCanPost(studentInfo){
  try{
    var timeOpen = timeCheck(new Date(studentInfo[0][startDateCol]), new Date(studentInfo[0][endDateCol]));
    var canPost = (studentInfo[0][canPostCol] && timeOpen); //Check "Survey Permitted" column for this student, and if timeOpen is TRUE (line above)
    return canPost;
  } catch(e){
    return false;
  }
  
}
//-----
// function getStudentInfo(myID) returns the student listing info for thisUser
//-----
function getStudentInfo(ListDoc, myID){

  var myStudentListsheet = ListDoc.getSheetByName(myStudentDataSheetName);
  var studentInfo = getRowsMatching(myStudentListsheet.getDataRange().getValues(),1,myID);
  Logger.log(studentInfo);
  return studentInfo;
}

//-----
// function getStudentInfo(myID) returns all student listing info if thisUser is in Principals list
//-----
function getPrincipalInfo(myID){

  var studentInfo = [];
  var myPrincipalList = SpreadsheetApp.openById(myListDocID).getSheetByName(myApproverListSheetName).getDataRange().getValues();
  
  for (var i =0; i < myPrincipalList.length; i++){
    if (myID === myPrincipalList[i][0]) studentInfo = SpreadsheetApp.openById(myListDocID).getSheetByName(myStudentDataSheetName).getDataRange().getValues().splice(1);  
  }
  return studentInfo;
}


//--------
// studentPostData(data) - expects choice data from the interface - posts choices to sheet
//--------
function studentPostData(rowindex, postedStudentId, currentChoice, currentExtraData, sindex){
  var dataStartingCol = 19 // Which column does the data to be written start?
  var ListDoc = SpreadsheetApp.openById(myListDocID);
  var SurveyDoc = ListDoc;
  var msg = "";
  var problemEnc = false;

  try{ // Log the post
    var logOutput = [new Date(), rowindex, postedStudentId, currentChoice, currentExtraData, sindex, thisUser];
    var logSheet = ListDoc.getSheetByName("Logs");
    logSheet.appendRow(logOutput);
  } catch(e){
    Logger.log(e);
  }

  var lock = LockService.getScriptLock(); //Get Lock on script
  lock.waitLock(30000);

  //Get all results sheet
  var sheet = SurveyDoc.getSheetByName(mySurveySheetName);
  var data = sheet.getDataRange().getValues();

  //read result row given by rowindex
  var rowStudentId = data[rowindex][0];
  var rowParentId = data[rowindex][parentEmailCol];
  if ((rowStudentId == postedStudentId)&&(rowParentId == thisUser)){ //check the expected student row is the same as the expected student row (same studentIds)
    if (checkCanPost([data[rowindex]])){ //check the survey is open for this student 
      //map headers to cols
      var headers = data[0].slice(dataStartingCol); // Get the headers of the columns to be written
      var headerMap = {};
      for (var h=0; h<headers.length; h++){
        headerMap[headers[h]] = h;
      }
      var writeData = data[rowindex].slice(dataStartingCol); //Get the existing data
      for (const key in currentChoice){ //loop through and prepare to write all currentChoice data
        var thisValue = currentChoice[key];
        var dataCol = headerMap[key];
        if (writeData[dataCol] != -1){ //if it is not -1 in the system
          writeData[headerMap[key]] = thisValue ? 1 : 0;
        }
      }
      for (const key2 in currentExtraData){
        var thisValue = currentExtraData[key2];
        var dataCol = headerMap[key2];
        writeData[headerMap[key2]] = thisValue;
      }
      //write data
      //Get Lock
      
      var targetRange = sheet.getRange(rowindex+1, dataStartingCol+1, 1, writeData.length).setValues([writeData]);
      SpreadsheetApp.flush();
      // clean up and release the lock
      lock.releaseLock();
      var newData = {};
      var readData = targetRange.getValues()[0];
      for (var r=0; r<readData.length;r++){
        newData[headers[r]] = readData[r];
      }

      return{error: {status: false, class: "bg-success", msg: msg}, update:  newData, sindex: sindex }
    } else {
      return{error: {status: true, class: "bg-danger", msg: "You don't have permission to post at this time."}, update:  {}, sindex: sindex}
    }
  } else {
    return{error: {status: true, class: "bg-danger", msg: "Record line mismatch! Unexpected user in data saving. Please try again."}, sindex: sindex}
  }
}
 
//--------
// approverPostData(data) - expects choice data from the interface - posts choices to sheet
//--------
function approverPostData(data, sindex){
  
  var passTest = true;
  //Check we have permission to approve
  passTest = getPrincipalInfo(thisUser).length > -1; 
  //check choices data exists and has been filled 

  if(passTest){   
    var mySurveyFile = SpreadsheetApp.openById(mySurveyCollector);
      var myC = new Array();  
      myC.push([data.id, 
                  data.email,
                  data.timeDate,
                  data.hrm,
                  data.principalChecked,
                  data.choice1,
                  data.choice2,
                  thisUser,
                  new Date(),
                  data.principalChecked,
                  data.paid,
                  data.house]);
      return postData(myC, mySurveyFile, false, sindex, false);
  } else {
      return {error: {status: true, msg: "No posting permissions", class: "bg-danger"}};
  }
}

//-------
//DO POST
//-------

function postData(myC, SurveyDoc, checkQuota, sindex, canPost) {

  var sheet = SurveyDoc.getSheetByName(mySurveySheetName);
  var countSheet = SurveyDoc.getSheetByName(mySurveyTripCounts);
  var statusMessage = "";
  var problemEnc = false;
  
  //Get Lock
  var lock = LockService.getPublicLock();
  lock.waitLock(30000);

  //Get the survey data.
  var lastRow = sheet.getLastRow();
  var sheetData = sheet.getRange(1,1,lastRow,14).getValues();
  var thisRow = getExistingRow(sheetData,1,myC[0][1]);  
  var nowTime = new Date();
    
  //If record already exists, clear it first (this is to ensure that we get accurate course spot counts without including this user's previous choice)
  if (thisRow < lastRow){   
    var clearRange = sheet.getRange(thisRow+1, 1, 1, 7).setValues([[myC[0][0], myC[0][1], myC[0][2], myC[0][3], myC[0][4], '', '']] );    
    //sheet.deleteRow(thisRow+1);
  }
  
  if(checkQuota){//If we are checking quota
    //Check not over quota. Get quota info...
    var myCounts = countSheet.getDataRange().getValues();
    // For each course entered, if spots taken greater or equal spots available, blank that course and add error message.
    for (var x=5; x <= 5; x++){
      var RowX = ArrayLib.indexOf(myCounts, 0, myC[0][x]);
      if (RowX > -1){ 
        if (myCounts[RowX][2] >= myCounts[RowX][1]) {
          statusMessage += myC[0][x] + ' already full! Please select another choice!'; 
          problemEnc = true;
          myC[0][x] = '';
        }
      }
    } 
  }

  var targetRange = sheet.getRange(thisRow+1, 1, 1, myC[0].length).setValues(myC);    
    
  SpreadsheetApp.flush();
  // clean up and release the lock
  lock.releaseLock();
  
  //var canPost = checkCanPost(thisUser);
  
  var choice = {};
  var myChoices = textifyDates(getRowsMatching(sheet.getDataRange().getValues(),0,myC[0][0]));
  if (myChoices.length > 0) {
    choice = { id: myChoices[0][0],
                       email: myChoices[0][1],
                       timeDate: myChoices[0][2],
                       hrm: myChoices[0][3],
                       advisorChecked: myChoices[0][4],
                       choice1: myChoices[0][5],
                       choice2: myChoices[0][6],
                       approvalChange: myChoices[0][7],
                       approvalDate: myChoices[0][8],
                       principalChecked: myChoices[0][9],
                       paid: myChoices[0][10],
                       canPost: canPost,
                       house: myChoices[0][11]
                        }
  } else {
    problemEnc = true;
    statusMessage = "Could not reload your data. Try reloading this page.";
  }

  return {error: {status: problemEnc, msg: statusMessage, class: "bg-danger", log: myChoices}, choice: choice, sindex: sindex};
}





//-----
// function getStudentInfo(myID) returns the student listing info for thisUser
//-----
function getStudentInfo(myID){

  var myStudentListsheet = SpreadsheetApp.openById(myListDocID).getSheetByName(myStudentDataSheetName);  
  var LastSsRow = myStudentListsheet.getLastRow()-1;
  var studentInfo = getRowsMatching(myStudentListsheet.getRange(2, 1, LastSsRow,myStudentListsheet.getLastColumn()).getValues(),1,myID);

  return studentInfo;
}


//-----
// getExistingRow - returns the row that contains the data matching the criteria, or returns the next row in the spreadsheet.
//
//-----
function getExistingRow(myList,checkCol,checkCriteria){
  var myRow = 0;
  while (myRow < myList.length){
    if (myList[myRow][checkCol] == checkCriteria) {
      return myRow;
    }
    myRow++;
  } 
  return myRow;
}


function myClickHandler(e) {
  var app = UiApp.getActiveApplication();

  var label = app.getElementById('statusLabel');
  label.setVisible(true);

  app.close();
  return app;
}

function usableColValues(coldata, lastrow) {
  for( var i = (lastrow - 1) ; i > 0; i--){

    if(coldata[i] != "") {
      return coldata.slice(0,i+1);
      };
  };
  return coldata;
}



//getRowsMatching takes a data list and searches the sortIndex for all values that match valueToFind, returning the rows that match this value

function getRowsMatching(DataList, sortIndex, valueToFind){
  
  var foundList = new Array();
  var myDataList = DataList.slice(0); //Added this line and changed the function parameter from myDataList to DataList - hoping to stop changes to original array
  myDataList.sort(function(a, b){ //Sort the items by sortIndex
    var x = a[sortIndex];
    var y = b[sortIndex];
    return (x < y ? -1 : (x > y ? 1 : 0));});
  
  var cdr = 0;
  var found = false; 

  while ( cdr < myDataList.length){
    if (myDataList[cdr][sortIndex] == valueToFind) {
      found=true;
      foundList.push(myDataList[cdr])
    }
    else if (found){
      return foundList;
    }
    cdr++;
  }

  return foundList;
  
}

//getRowsIncluding takes a data list and searches the sortIndex for values that include valueToFind, returning the rows that match this value. It may be less efficient than getRowsMatching because it goes through all rows.

function getRowsIncluding(DataList, sortIndex, valueToFind){
  
  var foundList = new Array();
  var myDataList = DataList.slice(0); //Added this line and changed the function parameter from myDataList to DataList - hoping to stop changes to original array
  
  for (var cdr = 0; cdr < myDataList.length; cdr++){
    if (myDataList[cdr][sortIndex].includes(valueToFind)) {
      found=true;
      foundList.push(myDataList[cdr])
    }
  }
  return foundList;
  
}


//-----
// function timeCheck (studentInfo)
//-----
function timeCheck(openTime, closeTime) {
  //Logger.log([openTime, closeTime, (new Date() >= openTime)]);
  return DisableDateLocking || ((new Date() >= openTime) && (new Date() <= closeTime));
}

// -----
// include - include files
// -----
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename)
      .getContent();
}

// ------------------------------------ USEFUL FUNCTIONS--------------------------------------------------
//

//-----
// function textifyDates(myArr) - converts all dates into text format - assumes a 2D array as an input, returns the array.
//-----
function textifyDates(myArr){
  
  
  for(var r=0; r < myArr.length; r++){
    for(var c=0; c < myArr[r].length; c++){
      if (Object.prototype.toString.call(myArr[r][c]) === '[object Date]'){
        try {           
          myArr[r][c] = Utilities.formatDate(myArr[r][c], "GMT+08:00", "dd-MMM-yyyy hh:mm:ss")
        } 
        catch(err) { myArr[r][c] = err};
      }
    }
  }
  return myArr;
}

//-----
// function textifyDates(myArr) - converts all dates into text format - assumes a 2D array as an input, returns the array.
//-----
function textifyDates1D(myArr){
  
  for(var r=0; r < myArr.length; r++){
      if (Object.prototype.toString.call(myArr[r]) === '[object Date]'){
        try {
          //myArr[r] = myArr[r].toString();
          myArr[r] = Utilities.formatDate(myArr[r], "GMT+08:00", "dd-MMM-yyyy")
          } 
        catch(err) { myArr[r] = err};
      }
    
  }
  return myArr;
}


/**
 * Converts a 2D array into an array of objects, using the first row as the keys.
 * Optionally, it can filter the resulting array based on a specified key and value.
 * Automatically formats any date values in the resulting objects using the user's or script's timezone.
 *
 * @param {Array<Array>} data - The 2D array to be converted.
 * @param {string} [filterKey] - The key to filter the resulting array by.
 * @param {any} [filterValue] - The value to filter the resulting array by.
 * @param {string} [dateTimeZone] - The time zone to use for formatting date values.
 * @returns {Array<Object>} - An array of objects, with the first row as the keys.
 */
function arrayToObjects(data, filterKey, filterValue, dateTimeZone = getDefaultTimezone()) {
  if (data.length === 0) {
    return [];
  }
  const headers = data[0];
  const result = [];

  for (let i = 1; i < data.length; i++) {
    const obj = {row: i};
    for (let j = 0; j < headers.length; j++) {
      const value = data[i][j];
      if (Object.prototype.toString.call(value) === "[object Date]") {
        obj[headers[j]] = Utilities.formatDate(value, dateTimeZone, "dd-MMM-yyyy HH:mm:ss");
      } else {
        obj[headers[j]] = value;
      }
    }

    // Filter the object if filterKey and filterValue are provided
    if (filterKey && filterValue !== undefined) {
      if (obj[filterKey] === filterValue) {
        result.push(obj);
      }
    } else {
      result.push(obj);
    }
  }
  return result;
}

/**
 * Determines the default timezone to use for the script or user.
 * @returns {string} - The default timezone in the format "GMT+hh:mm".
 */
function getDefaultTimezone() {
  const tz = Session.getScriptTimeZone();
  const offset = new Date().getTimezoneOffset();
  const sign = offset > 0 ? "-" : "+";
  const hours = Math.floor(Math.abs(offset) / 60);
  const minutes = Math.abs(offset) % 60;
  return `GMT${sign}${hours.toString().padStart(2, "0")}:${minutes.toString().padStart(2, "0")}`;
}

function myTest(){
  Logger.log(loadGInfo());
}




function sendAllocationEmails() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ALLOCATION EMAIL LIST");
  const data = sheet.getDataRange().getValues();
  const doc = DocumentApp.openById(emailTemplateId);
  var ui = SpreadsheetApp.getUi();

  const emailSendFlag = 8; //Column index for "Yes" if row should be sent.
  const emailSentColumn = 9; // Column index for "Sent" status
  const emailBodyColumn = 11; // Column index for email body
  const timestampColumn = 12; // Column index for timestamp
  const senderEmailColumn = 13; // Column index for sender email

  const emailsToSend = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[emailSendFlag] === 'Yes') { // Check if "Yes" is in column I
      const tripName = row[0];
      const bodyText = doc.getBody().getText().replace('{TripName}', tripName);
      const toAddresses = [row[4], row[5]];
      const subject = 'Y7-8 Trip Allocation';

      emailsToSend.push({
        row: i + 1,
        toAddresses: toAddresses,
        subject: subject,
        bodyText: bodyText,
      });
    }
  }

  if (emailsToSend.length === 0) {
    ui.alert('No emails to send.');
    return;
  } else {
    var result = ui.alert(
    "Please confirm",
    'Send ' + emailsToSend.length + ' emails?',
    ui.ButtonSet.YES_NO,
    );
  }
  if (result === ui.Button.YES) {
    const senderEmail = Session.getActiveUser().getEmail();
    for (const email of emailsToSend) {
      try{
        const now = new Date();
        GmailApp.sendEmail(email.toAddresses, email.subject, email.bodyText);
        sheet.getRange(email.row, emailSentColumn, 1, 4).setValues([[
         'Sent', 
         email.body, 
         now, 
          senderEmail
        ]]);
      } catch(e){
        Logger.log(e);
      }
      
    }
    
    ui.alert('Sent ' + emailsToSend.length + ' emails.');
  }
}

/**
 * Function to populate student data from the master student database
 * Prompts user for year groups and confirms before writing data
 */
function populateStudentData() {
  const sourceSpreadsheetId = "1B_DenT-IBnDh3rEaI-D2n6yaKVaTlshWd7PCGEOCkAM";
  const sourceSheetName = "Portal_Order";
  const ui = SpreadsheetApp.getUi();
  
  try {
    // Prompt for year groups
    const yearGroupResponse = ui.prompt(
      'Enter Year Groups',
      'Please enter year groups separated by commas (e.g., 7,8,9):',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (yearGroupResponse.getSelectedButton() !== ui.Button.OK) {
      return; // User cancelled
    }
    
    const yearGroupsText = yearGroupResponse.getResponseText().trim();
    if (!yearGroupsText) {
      ui.alert('No year groups entered. Operation cancelled.');
      return;
    }
    
    // Parse year groups
    const yearGroups = yearGroupsText.split(',').map(yg => yg.trim()).filter(yg => yg);
    
    if (yearGroups.length === 0) {
      ui.alert('No valid year groups entered. Operation cancelled.');
      return;
    }
    
    // Get source data
    const sourceSpreadsheet = SpreadsheetApp.openById(sourceSpreadsheetId);
    const sourceSheet = sourceSpreadsheet.getSheetByName(sourceSheetName);
    const sourceData = sourceSheet.getDataRange().getValues();
    
    if (sourceData.length <= 1) {
      ui.alert('No student data found in source sheet.');
      return;
    }
    
    // Find column indices in source data
    const sourceHeaders = sourceData[0];
    const getColIndex = (headerName) => {
      const index = sourceHeaders.indexOf(headerName);
      if (index === -1) {
        throw new Error(`Column '${headerName}' not found in source sheet`);
      }
      return index;
    };
    
    const sourceIndices = {
      reg: getColIndex('Reg'),
      surname: getColIndex('Surname'),
      firstName: getColIndex('First Name'),
      preferredName: getColIndex('Preferred Name'),
      studentEmail: getColIndex(' Student Email'),
      familyEmail: getColIndex(' Family Email'),
      house: getColIndex('House'),
      gender: getColIndex('Gender'),
      yearGroup: getColIndex('Year Group')
    };
    
    // Filter students by year group and prepare data
    const filteredStudents = [];
    for (let i = 1; i < sourceData.length; i++) {
      const row = sourceData[i];
      const studentYearGroup = String(row[sourceIndices.yearGroup]).trim();
      
      if (yearGroups.includes(studentYearGroup)) {
        // Build name from Surname, First Name, Preferred Name
        const surname = row[sourceIndices.surname] || '';
        const firstName = row[sourceIndices.firstName] || '';
        const preferredName = row[sourceIndices.preferredName] || '';
        const fullName = [preferredName, firstName, surname].filter(name => name).join(' ').trim();
        //id, 	email, 	name, 	hrm, 	parentEmail, 	house, 	gender, 
        const studentData = [
          row[sourceIndices.studentEmail], // col 0: id (student email)
          row[sourceIndices.studentEmail], // col 1: email (student email)
          fullName, // col 2: name
          row[sourceIndices.reg], // col 3: hrm (Reg)
          row[sourceIndices.familyEmail], // col 4: parentEmail
          row[sourceIndices.house], // col 5: house
          row[sourceIndices.gender], // col 6: gender
          true, // col 7: canPost (canPostCol)
          new Date(new Date().getTime() + (10 * 24 * 60 * 60 * 1000)), // col 8: startDate (startDateCol) - 10 days from now
          new Date(new Date().getTime() + (30 * 24 * 60 * 60 * 1000)), // col 9: endDate (endDateCol) - 30 days from now
          '',
          row[sourceIndices.yearGroup]
        ];
        
        filteredStudents.push(studentData);
      }
    }
    
    if (filteredStudents.length === 0) {
      ui.alert(`No students found for year groups: ${yearGroups.join(', ')}`);
      return;
    }
    
    // Show confirmation
    const confirmResponse = ui.alert(
      'Confirm Student Data Import',
      `Found ${filteredStudents.length} students for year groups: ${yearGroups.join(', ')}\n\nThis will add datato LiveResults sheet. Continue?`,
      ui.ButtonSet.YES_NO
    );

    
    
    if (confirmResponse !== ui.Button.YES) {
      ui.alert('Operation cancelled.');
      return;
    }
    
    // Write to destination sheet
    const destSpreadsheet = SpreadsheetApp.openById(myListDocID);
    const destSheet = destSpreadsheet.getSheetByName(mySurveySheetName); // "LiveResults"
    destSheet.copyTo(destSpreadsheet); //Make a backup copy of old LiveResults
    
    // Prepare headers and data
    /*
    // id, 	email, 	name, 	hrm, 	parentEmail, 	house, 	gender, 
    const headers = ['id', 'email', 'name', 'hrm', 'parentEmail', 'col5', 'col6', 'canPost', 'startDate', 'endDate', 'advisorEmail', 'house', 'gender', 'col13', 'col14', 'col15', 'col16', 'col17', 'col18'];
    
    // Clear existing data and write new data
    destSheet.clear();
    
    // Write headers
    destSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    */
    // Write student data
    if (filteredStudents.length > 0) {
      destSheet.getRange(destSheet.getLastRow()+1, 1, filteredStudents.length, filteredStudents[0].length).setValues(filteredStudents);
    }
    
    ui.alert('Success!', `Successfully imported ${filteredStudents.length} students to the LiveResults sheet.`, ui.ButtonSet.OK);
    
  } catch (error) {
    Logger.log('Error in populateStudentData: ' + error.toString());
    ui.alert('Error', 'An error occurred while importing student data: ' + error.toString(), ui.ButtonSet.OK);
  }
}

function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('VSA Trips')
      .addItem('Send Allocation Emails', 'sendAllocationEmails')
      .addItem('Populate Student Data', 'populateStudentData')
      .addToUi();
}