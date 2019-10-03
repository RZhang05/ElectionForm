//Created by Raymond Zhang (zhangraymond05@gmail.com)
//global variables (not very important can just be local)
Logger.log("Initializing");
var d = new Date();
var year = d.getFullYear();

//when the sheet is opened
function onOpen() {
  //add "generate form" option to taskbar on top (file, edit, view, etc)
  var ss = SpreadsheetApp.getActive();
  var menu = [{name: 'Generate Form', functionName: 'generate'}]; //run the function called generate
  menu.push(null);
  menu.push({name: 'Generate Results', functionName: 'tallyVotes'});
  ss.addMenu('Generator', menu);
}

//makes the form
function generate() {
  //variables
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName('Candidates');
  var endpoint = ss.getSheetByName('Results');
  var range = sheet.getDataRange();
  var values = range.getValues();
  var parentFolder = DriveApp.getFolderById('1_J5QVewVApOJKh5FGNszah0wt9dlkzW7');
  
  var ui = SpreadsheetApp.getUi();
  var grade = "";
  
  do {
    var response = ui.prompt('Grade Selection', 'Enter a grade', ui.ButtonSet.OK_CANCEL);
    grade = response.getResponseText().trim();
  } while(response.getSelectedButton() == ui.Button.CANCEL || response.getSelectedButton() == ui.Button.CLOSE);
  
  
  var newFolder = parentFolder.createFolder(year + ' Grade ' + grade + ' Election');
  
  //stores old sheet data and resets to base
  var sheets = ss.getSheets();
  if(sheets.length>2) {
    Logger.log('Archived ' + (year-1) + ' Grade ' + grade + ' Election');
    var archive = SpreadsheetApp.create((year-1) + ' Archive');
    sheets[1].copyTo(archive);
    sheets[2].copyTo(archive);
    moveForm((year-1) + ' Archive', newFolder.getId());
    
    var formURL = sheets[0].getFormUrl();
    FormApp.openByUrl(formURL).removeDestination();
    ss.deleteSheet(sheets[0]);
  }
  
  //delete names
  var h = endpoint.getDataRange().getHeight()-1;
  if(h>0) endpoint.deleteRows(2, h);
  
  //add new names
  var names = [];
  for(i=1;i<values.length;i++) {
    names.push(values[i][0]);
    endpoint.appendRow([values[i][0],0]);
  }
    
  //make the form
  makeOurForm(names, grade, newFolder.getId());
}

//function which creates the form
function makeOurForm(names, grade, fileId) {
  var ss = SpreadsheetApp.getActive();
  Logger.log('Starting to create form');
  
  var name = 'Grade ' + grade + ' Student Senator Ballot';
  var form = FormApp.create(name);
  
  //make sure only one response per user is allowed
  form.setCollectEmail(true);
  form.setLimitOneResponsePerUser(true);
  form.setRequireLogin(true);
  
  var item = form.addCheckboxItem();
  item.setTitle('Vote for at most two candidates');
  item.setChoiceValues(names);
  item.setRequired(true);
  
  //ensures that max two is allowed
  var checkBoxValidation = FormApp.createCheckboxValidation()
  .setHelpText('Select at most two candidates.')
  .requireSelectAtMost(2)
  .build();
  item.setValidation(checkBoxValidation);
  
  //make the results go to the master sheet
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
  
  //create an event that runs when the spreadsheet recieves a result from the form
  deleteTriggers();
  ScriptApp.newTrigger('onFormSubmit').forSpreadsheet(ss).onFormSubmit().create();
  
  //move form into archive folder
  moveForm(name, fileId);
  
  Logger.log('Form successfully created');
  Logger.log('Published URL: ' + form.getPublishedUrl());
  Logger.log('Editor URL: ' + form.getEditUrl());
}

//function which moves files around
function moveForm(name, fileId) {
  Logger.log('Moving ' + name + ' to correct location');
  var rootFolder = DriveApp.getRootFolder();
  var archiveFolder = DriveApp.getFolderById(fileId);
  
  var newForms = rootFolder.getFilesByName(name);
  while(newForms.hasNext()) {
    var newForm = newForms.next();
    archiveFolder.addFile(newForm);
    rootFolder.removeFile(newForm);
  }
  Logger.log(name + ' successfully moved');
}

//function which runs when a response is recieved
function onFormSubmit(e) {
  var ss = SpreadsheetApp.getActive();
  ss.getSheets()[0].setName('Responses');
  var sheet = ss.getSheetByName('Candidates');
  var range = sheet.getDataRange();
  var values = range.getValues();
  
  //retrieve student numbers of applicants
  var invalid = [];
  for(i=1;i<values.length;i++) {
    invalid.push(values[i][1]);
  }
  
  //retrieve response data
  var email = e.namedValues['Email Address'][0];
  
  //check if response is invalid (email is one of the applicants)
  if(invalid.indexOf(email)>-1) {
    Logger.log(email + '\'s vote was removed');
    e.range.deleteCells(SpreadsheetApp.Dimension.COLUMNS);
  }
}

//count up all the votes
function tallyVotes() {
  //variables for manipulating spreadsheet
  var ss = SpreadsheetApp.getActive();
  var output = ss.getSheetByName('Results');
  var r = output.getDataRange();
  var v = r.getValues();
  var input = ss.getSheetByName('Responses');
  var r2 = input.getDataRange();
  var v2 = r2.getValues();
  
  //clear spreadsheet
  output.clear();
  output.appendRow(['Name', 'Votes']);
  
  //count votes for every person
  for(i=1;i<v.length;i++) {
    var name = v[i][0];
    var num = v[i][1];
    for(j=1;j<v2.length;j++) {
      var names = v2[j][2].split(", ");
      for(k=0;k<names.length;k++) {
        if(names[k]===name) {
          num++;
          break;
        }
      }
    }
    output.appendRow([name,num]);
  }
  
  //output results
  output.sort(2, false);
}

//get rid of previous triggers (a new one is created everytime this code is run)
function deleteTriggers(){
  var triggers = ScriptApp.getProjectTriggers();

  triggers.forEach(function(trigger){

    try{
      ScriptApp.deleteTrigger(trigger);
    } catch(e) {
      throw e.message;
    };

    Utilities.sleep(1000);

  });

};