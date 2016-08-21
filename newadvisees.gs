function onOpen () {
  var ss = SpreadsheetApp.getActive(), options = [{name: "Run Notifications", functionName: "makeFacultyNotifications"}];
  ss.addMenu("Advisor Folders", options);
  
}

function buildStudentFolders() {
  
  var appSheet, shtOne;
  var fldrAdvisees, fldrAdvising;
  var rngStudents, arrayStudents;
  var i, leng;
  var fldrName, fldrStudent, fldrRoot;
  //folderString, flderAdvising, flderStud;
  
  //Get the main sheet and folder information
  sa = SpreadsheetApp.getActiveSpreadsheet();
  shtOne = sa.getSheetByName("Students");
  fldrAdvisees = DriveApp.getFolderById('0B2SNMLPe9seBbngzeVRHVDMtR0E');
  fldrAdvising = DriveApp.getFolderById('0B2SNMLPe9seBMzJOdEpoWlE1RVU');
  fldrRoot = DriveApp.getRootFolder();
  
  //Get the student information from the sheet titled Students
  rngStudents = shtOne.getRange(2, 1, shtOne.getLastRow(), 2); //Load the data into a range
  arrayStudents = rngStudents.getValues(); //Load the range into an array
  
  //Loop through each student in the sheet and create a folder in the Advisees folder
  i = 0;
  while(i < arrayStudents.length)
  {
    fldrName = arrayStudents[i][0] + " - " + arrayStudents[i][1]; //Build the folder name
    fldrStudent = DriveApp.createFolder(fldrName); //Create the folder with the given name
    //fldrAdvisees.addFolder(fldrStudent); //Add the folder to the Advisees folder
    //fldrRoot.removeFolder(fldrStudent); //Remove the folder from the root drive
    i++;
  }
}

function buildAdvisorFolders(AdvisingGroup) {
  
  var appSheet, shtOne;
  var fldrAdvisingRoot, fldrRoot;
  var rngAdvisors, arrayAdvisors;
  var i, fldrName, fldrAdvisor;
  
  //Get the main sheet and folder information
  sa = SpreadsheetApp.getActiveSpreadsheet();
  shtOne = sa.getSheetByName("Advisors");
  fldrAdvisingRoot = DriveApp.getFolderById('0B-BMvh_WRnHkcnZONkIyRENWUW8');
  fldrRoot = DriveApp.getRootFolder();
  
  //Get the advisor information for the sheet
  rngAdvisors = shtOne.getRange(2, 1, shtOne.getLastRow(), 2); //Put the advisors into a range
  arrayAdvisors = rngAdvisors.getValues(); //put the range into an array
  
  //Loop through each advisor in the sheet and create a folder
  i = 0;
  while(i < arrayAdvisors.length) //while there are still advisors
  {
    fldrName = "Fall 16 Incoming Advising - " + arrayAdvisors[i][0]; //generate the folder name
    fldrAdvisor = DriveApp.createFolder(fldrName); //create the advisor folder
    //fldrAdvisingRoot.addFolder(fldrAdvisor); //Add the folder to the Class of folder
    //fldrRoot.removeFolder(fldrAdvisor); //Remove the advisor folder from the root
    i++;
  }
}

function processExports()
{
  var shtOne, rngFolders, arrayFolders;
  var fldrEssays, fldrTranscripts, fldrAdvising, fldrAcademic, fldrFYPEssays;
  
  //Get the Folders and files in the folders
  fldrEssays = DriveApp.getFolderById('0B2SNMLPe9seBMUFqSzkzS0tPam8');
  fldrTranscripts = DriveApp.getFolderById('0B2SNMLPe9seBN3R4bUViTW1yUU0');
  fldrAdvising = DriveApp.getFolderById('0B2SNMLPe9seBMzJOdEpoWlE1RVU');
  fldrAcademic = DriveApp.getFolderById('0B2SNMLPe9seBeVNBS3luSGJYRFU');
  fldrFYPEssays = DriveApp.getFolderById('0B2SNMLPe9seBMGtLNzdkYmxudGc');

  //For each export folder, process the values  
  //renamingFiles(fldrEssays);
  renamingFiles(fldrAdvising, "Supplemental Info for ");
  //renamingFiles(fldrAcademic, "Academic Merge for ");
  //renamingFiles(fldrTranscripts);
  //renamingFiles(fldrFYPEssays);
}

function renamingFiles(fldrExports, prefix) //Rename the files in the folder passed in to the function
{
  var files, token, PageSize;
  var result, helper;
  var strName, strNewName;
  var strPrefix, type;
  var fldrAdvisees, strSearch, rtnFolder, folder;
  
  //Get all of the files in the folder sent into the function
  files = fldrExports.getFilesByType(MimeType.PDF);
  fldrAdvisees = DriveApp.getFolderById('0B2SNMLPe9seBbngzeVRHVDMtR0E');
  
  //Loop through the files in the folder
  while(files.hasNext()) {
    var file = files.next();
    strName = file.getName(); //Get the name of the file
    //strNewName = strName.substring(strName.length-14, strName.length-5).trim() + ".pdf"; //New Name of the file
    strNewName = prefix + strName.trim();
    file.setName(strNewName);
    
    strSearch = "fullText contains '" + strName.substring(strName.length-14, strName.length-5) + "'";
    rtnFolder = fldrAdvisees.searchFolders(strSearch);
    if(rtnFolder.hasNext()){
      folder = rtnFolder.next();
      helper = folder.getName();
      folder.addFile(file);
      fldrExports.removeFile(file);
    }
  }
}

/* No longer needed...merged with rename function above
function moveFiles(fldrExports) //Moves the files from the holding folder to the right student folder
{
  var folders;
  var files, rtnFile;
  var fldrAdvisees, rtnFolder;
  var strFile, strBannerID, k, strSearch, strPrefix;
  
  //Get the files from the folder you would like to process
  files = fldrExports.getFiles();
  fldrAdvisees = DriveApp.getFolderById('0B2SNMLPe9seBbngzeVRHVDMtR0E');
  
  switch(fldrExports.getName().trim()){
    case "Summer Reading Essays":
      strPrefix = "Summer Reading Essay for ";
      break;
    case "Academic Merges":
      strPrefix = "Academic Merge for ";
      break;
    case "Advising Merge":
      strPrefix = "Advising Info for ";
      break;
    case "Transcripts":
      strPrefix = "Transcripts for ";
      break;
    case "FYP Essays":
      strPrefix = "FYP Essay for ";
      break;
  }
  Logger.log(strPrefix);
    
  //Loop through each file in the folder
  while(files.hasNext()) {
    var file = files.next(); //Get the next file
    
    //Get the file name and build the search string
    strFile = file.getName();
    strSearch = "fullText contains '" + strFile.substring(0, 9) + "'";
    //Locate the matching folder in the Advisees folder
    rtnFolder = fldrAdvisees.searchFolders(strSearch);
    var folder = rtnFolder.next();
    Logger.log(file.setName(strPrefix + folder.getName() + ".pdf"));
    folder.addFile(file);
    fldrExports.removeFile(file);
  }
  Logger.log("Move files on " + fldrExports.getName() + " is complete.");
}
*/

function moveAdviseeFolders() {
  var sa, shtStudents, rngStudents, arrayStudents;
  var fldrAdvisees, fldrAdvisee;
  var fldrAdvising, fldrAdvisor;
  var studentSearch, advisorSearch, i;
  
  //Get an array of students and their advisors to process
  sa = SpreadsheetApp.getActiveSpreadsheet();
  shtStudents = sa.getSheetByName("Students with Advisors2");
  fldrAdvisees = DriveApp.getFolderById('0B2SNMLPe9seBbngzeVRHVDMtR0E'); //Folder with all of the Advisees in it
  fldrAdvising = DriveApp.getFolderById('0B-BMvh_WRnHkcnZONkIyRENWUW8'); //Class of folder

  //Go through the array
  rngStudents = shtStudents.getRange(2, 1, shtStudents.getLastRow(), 11); //Load the data into a range
  arrayStudents = rngStudents.getValues(); //Load the range into an array
  
  //Loop through all of the students in the array
  i = 0;
  while(i < arrayStudents.length) {
    //Build the search criteria
    studentSearch = "fullText contains '" + arrayStudents[i][0] + "'";
    advisorSearch = "fullText contains '" + arrayStudents[i][10] + "'";
    
    //Find the student and advisor folder
    fldrAdvisee = fldrAdvisees.searchFolders(studentSearch);
    fldrAdvisor = fldrAdvising.searchFolders(advisorSearch);
    while(fldrAdvisee.hasNext()) {
      var folder = fldrAdvisee.next();
      Logger.log(folder.getName());
      fldrAdvisor.next().addFolder(folder); //Put the advisee folder in the advisor folder
      fldrAdvisees.removeFolder(folder); //Remove the advisee folder from the Advisees folder
    }
      
    //Logger.log(arrayStudents[i][0]);
    i++;
  }
}

function makeFacultyNotifications() {
  var sa, shtAdvisors, i;
  var rngAdvisors, arrayAdvisors;
  var fldrAdvisors, fldAdvisor;
  var advisorSearch;
  
  //Get the advisors from the Advisor sheet
  sa = SpreadsheetApp.getActiveSpreadsheet();
  shtAdvisors = sa.getSheetByName("Advisors");
  fldrAdvisors = DriveApp.getFolderById('0B-BMvh_WRnHkcnZONkIyRENWUW8');
  
  rngAdvisors = shtAdvisors.getRange(2, 1, shtAdvisors.getLastRow(), shtAdvisors.getLastColumn());
  arrayAdvisors = rngAdvisors.getValues(); //Make an array of advisors
  
  //Loop through all of the advisors
  i = 0
  while(i < arrayAdvisors.length && arrayAdvisors[i][0] != "") {
    advisorSearch = "title contains '" + arrayAdvisors[i][0] + "'"; //Make String: "title contains '[ADVISOR NAME]' "
    fldrAdvisor = fldrAdvisors.searchFolders(advisorSearch); //Search for the advisor folder with the name from above
    while(fldrAdvisor.hasNext()) {
      var folder = fldrAdvisor.next();
      //Set the edit permissions
      Logger.log(arrayAdvisors[i][3] + " added to " + folder.getName());
      folder.addEditor(arrayAdvisors[i][3]); //Lets the advisors edit their folders
    }
    i++;
  }
}

function checkPermissions(){
  var fldrAdvisors, fldrAllAdvisors, fldrStudents;
  var allFiles, arrayEditors, advisor;
  var i, j, strEditors;
  var sa, shtAdvisors, rngAdvisors, arrayAdvisors;

  //Write it to a spreadsheet
  
  //Set the advisor folder
  fldrAdvisors = DriveApp.getFolderById('0B-BMvh_WRnHkcnZONkIyRENWUW8');
  fldrAllAdvisors = fldrAdvisors.getFolders();
  i = 0;
  arrayEditors = new Array(1000);
  while(fldrAllAdvisors.hasNext()){
    var folder = fldrAllAdvisors.next();
    arrayEditors[i] = folder.getName();
    
    fldrStudents = folder.getFolders();
    while(fldrStudents.hasNext()){
      j = i + 1;
      var fldrStudent = fldrStudents.next();
      //Logger.log(fldrStudent.getName());
      arrayEditors[j] = fldrStudent.getName();
      
      allFiles = fldrStudent.getFiles();
      while(allFiles.hasNext()) {
        k = j + 1;
        var file = allFiles.next();
        var editors = file.getEditors();
        strEditors = "";
        for(var r = 0; r < editors.length; r++)
        {
          strEditors = strEditors + ", " + editors[r].getEmail();
        }
       arrayEditors[k] = strEditors;
       k++;
      }
      j = k + 1;
    }
    i = j + 1;
  }

  sa = SpreadsheetApp.getActiveSpreadsheet();
  sa.insertSheet("Advisor Check", sa.getSheets().length + 1);
  sa.setActiveSheet(sa.getSheetByName("Advisor Check"));
  var rngValues = sa.getRange("A1:A2000");
  //for(t = 1; t <= rngValues.length; t++)
  //{
    rngValues.setValue(arrayEditors);
    //rng.setValue(rngValues[t]);
  //}
  
}
