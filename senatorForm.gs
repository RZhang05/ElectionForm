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
  var newFolder = parentFolder.createFolder(year + ' Exec Election');
  
  //stores old sheet data and resets to base
  var sheets = ss.getSheets();
  if(sheets.length>2) {
    Logger.log('Archived previous election');
    var archive = SpreadsheetApp.create('[rename this] Previous exec election archive');
    sheets[1].copyTo(archive);
    sheets[2].copyTo(archive);
    moveForm('[rename this] Previous exec election archive', newFolder.getId());
    
    var formURL = sheets[0].getFormUrl();
    FormApp.openByUrl(formURL).removeDestination();
    ss.deleteSheet(sheets[0]);
  }
  
  //delete names
  var h = endpoint.getDataRange().getHeight()-1;
  if(h>0) endpoint.deleteRows(2, h);
  
  //add new names
  var p = [];
  var vp = [];
  var t = [];
  var pd = [];
  var prd = [];
  var socd = [];
  var spird = [];
  var cc = [];
  
  for(i=1;i<values.length;i++) {
    var line = [];
    var cnt = 0;
    for(j=0;j<values[i].length;j++) {
      var a = values[i][j];
      if(a!=="") {
        if(j==0) p.push(a);
        if(j==1) vp.push(a);
        if(j==2) t.push(a);
        if(j==3) pd.push(a);
        if(j==4) prd.push(a);
        if(j==5) socd.push(a);
        if(j==6) spird.push(a);
        if(j==7) cc.push(a);
      } else cnt++;
      line.push(a);
      if(a!=="") line.push(0);
      else line.push("");
    }
    if(cnt == values[i].length) break;
    endpoint.appendRow(line);
  }
  p.push('No Vote'); vp.push('No Vote'); t.push('No Vote'); 
  pd.push('No Vote'); prd.push('No Vote'); socd.push('No Vote');
  spird.push('No Vote'); cc.push('No Vote');
  //make the form
  makeOurForm(p,vp,t,pd,prd,socd,spird,cc, newFolder.getId());
}

//function which creates the form
function makeOurForm(p,vp,t,pd,prd,socd,spird,cc, fileId) {
  var ss = SpreadsheetApp.getActive();
  Logger.log('Starting to create form');
  
  var name = 'BOSS Exec Position Ballot';
  var form = FormApp.create(name);
  
  //make sure only one response per user is allowed
  form.setCollectEmail(true);
  form.setLimitOneResponsePerUser(true);
  form.setRequireLogin(true);
  
  //put in all positions and candidates
  var item = form.addMultipleChoiceItem();
  item.setTitle('President');
  item.setChoiceValues(p);
  
  var item2 = form.addMultipleChoiceItem();
  item2.setTitle('Vice President');
  item2.setChoiceValues(vp);
  
  var item3 = form.addMultipleChoiceItem();
  item3.setTitle('Treasurer');
  item3.setChoiceValues(t);
  
  var item4 = form.addMultipleChoiceItem();
  item4.setTitle('Publicity Director');
  item4.setChoiceValues(pd);
  
  var item5 = form.addMultipleChoiceItem();
  item5.setTitle('Public Relation Director');
  item5.setChoiceValues(prd);
  
  var item6 = form.addMultipleChoiceItem();
  item6.setTitle('Social Director');
  item6.setChoiceValues(socd);
  
  var item7 = form.addMultipleChoiceItem();
  item7.setTitle('Spirit Director');
  item7.setChoiceValues(spird);
  
  var item8 = form.addMultipleChoiceItem();
  item8.setTitle('Clubs Coordinator');
  item8.setChoiceValues(cc);
  
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
  var sheets = ss.getSheets();
  if(sheets[0].getName()!=='Responses') {
    ss.getSheets()[0].setName('Responses');
  }
}

//count up all the votes
function tallyVotes() {
  //variables for manipulating spreadsheet
  var ss = SpreadsheetApp.getActive();
  var output = ss.getSheetByName('Results');
  var input = ss.getSheetByName('Responses');
  var r = output.getDataRange();
  var v = r.getValues();
  var r2 = input.getDataRange();
  var v2 = r2.getValues();
  
  //clear spreadsheet
  output.clear();
  output.appendRow(['President', 'Votes','Vice President','Votes',
                    'Treasurer','Votes','Publicity Director','Votes',
                    'Public Relation Director','Votes','Social Director','Votes',
                    'Spirit Director','Votes','Clubs Coordinator','Votes']);
  
  //count votes for every candidate
  for(i=1;i<v.length;i++) {
    var rowOutput = [];
    
    //pres
    var name = v[i][0];
    var num = v[i][1];
    if(name !== "") {
      for(j=1;j<v2.length;j++) {
        var name1 = v2[j][2];
        if(name1===name) {
          num++;
        }
      }
    } else num = "";
    rowOutput.push(name);
    rowOutput.push(num);
    
    //vice pres
    var name = v[i][2];
    var num = v[i][3];
    if(name !== "") {
      for(j=1;j<v2.length;j++) {
        var name1 = v2[j][3];
        if(name1===name) {
          num++;
        }
      }
    } else num = "";
    rowOutput.push(name);
    rowOutput.push(num);
    
    //treasurer
    var name = v[i][4];
    var num = v[i][5];
    if(name !== "") {
      for(j=1;j<v2.length;j++) {
        var name1 = v2[j][4];
        if(name1===name) {
          num++;
        }
      }
    } else num = "";
    rowOutput.push(name);
    rowOutput.push(num);
    
    //public direc
    var name = v[i][6];
    var num = v[i][7];
    if(name !== "") {
      for(j=1;j<v2.length;j++) {
        var name1 = v2[j][5];
        if(name1===name) {
          num++;
        }
      }
    } else num = "";
    rowOutput.push(name);
    rowOutput.push(num);
    
    //public relation direc
    var name = v[i][8];
    var num = v[i][9];
    if(name !== "") {
      for(j=1;j<v2.length;j++) {
        var name1 = v2[j][6];
        if(name1===name) {
          num++;
        }
      }
    } else num = "";
    rowOutput.push(name);
    rowOutput.push(num);
    
    //social direc
    var name = v[i][10];
    var num = v[i][11];
    if(name !== "") {
      for(j=1;j<v2.length;j++) {
        var name1 = v2[j][7];
        if(name1===name) {
          num++;
        }
      }
    } else num = "";
    rowOutput.push(name);
    rowOutput.push(num);
    
    //spirit direc
    var name = v[i][12];
    var num = v[i][13];
    if(name !== "") {
      for(j=1;j<v2.length;j++) {
        var name1 = v2[j][8];
        if(name1===name) {
          num++;
        }
      }
    } else num = "";
    rowOutput.push(name);
    rowOutput.push(num);
    
    //clubs direc
    var name = v[i][14];
    var num = v[i][15];
    if(name !== "") {
      for(j=1;j<v2.length;j++) {
        var name1 = v2[j][9];
        if(name1===name) {
          num++;
        }
      }
    } else num = "";
    rowOutput.push(name);
    rowOutput.push(num);
    
    output.appendRow(rowOutput);
  }
  
  //output results
  var range = output.getRange('A1:B100');
  range.sort({column: 2, ascending: false});
  
  var range = output.getRange('C1:D100');
  range.sort({column: 4, ascending: false});
  
  var range = output.getRange('E1:F100');
  range.sort({column: 6, ascending: false});
  
  var range = output.getRange('G1:H100');
  range.sort({column: 8, ascending: false});
  
  var range = output.getRange('I1:J100');
  range.sort({column: 10, ascending: false});
  
  var range = output.getRange('K1:L100');
  range.sort({column: 12, ascending: false});
  
  var range = output.getRange('M1:N100');
  range.sort({column: 14, ascending: false});
  
  var range = output.getRange('O1:P100');
  range.sort({column: 16, ascending: false});
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