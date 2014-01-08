

//NEW CODE

//new function
//anytime spreadsheet is edited (AKA form submit), start mail merge

//set a trigger to start this function on every form submit
function formSubmitReply(e){
 var ss = SpreadsheetApp.getActiveSpreadsheet();
 var dataSheet = ss.getSheets()[0];
 if(dataSheet.getRange(1,dataSheet.getLastColumn()).getValue() != 'Automatic Response Status'){
   dataSheet.getRange(1,dataSheet.getLastColumn()+1).setValue('Automatic Response Status');
 }
 var headers = dataSheet.getRange(1, 1, 1, dataSheet.getLastColumn()).getValues();
 var emailColumnFound = false;

 //FIND THE COLUMN THAT CONTAINS EMAIL ADDRESSES
 for(i in headers[0]){
   if(headers[0][i] == "Email Address"){
     emailColumnFound = true;
   }
 }
 //IF CAN'T FIND, ASK USER TO MANUALLY INPUT
 if(!emailColumnFound){
   var emailColumn = Browser.inputBox("Which column contains the recipient's email address ? (A, B,...)");
   dataSheet.getRange(emailColumn+''+1).setValue("Email Address");
 }

 var dataRange = dataSheet.getRange(2, 1, dataSheet.getLastRow() - 1, dataSheet.getLastColumn());


//GET THE CORRECT DRAFT
 var selectedTemplate = GmailApp.search("in:drafts")[0].getMessages()[0];
 var emailTemplate = selectedTemplate.getBody();
 var attachments = selectedTemplate.getAttachments();
 var cc = selectedTemplate.getCc();
 var bcc = "";
 //if (e.parameter.bcc == "true") {
   bcc = "info@bouncemediagroup.com";
 //}

//GENERATE AND PROPEGATE THE EMAIL
  var regMessageId = new RegExp(selectedTemplate.getId(), "g");
  
  if (emailTemplate.match(regMessageId) != null) {
    var inlineImages = {};
    var nbrOfImg = emailTemplate.match(regMessageId).length;
    var imgVars = emailTemplate.match(/<img[^>]+>/g);
    var imgToReplace = [];
    for (var i = 0; i < imgVars.length; i++) {
      if (imgVars[i].search(regMessageId) != -1) {
        var id = imgVars[i].match(/Inline\simages?\s(\d)/);
        imgToReplace.push([parseInt(id[1]), imgVars[i]]);
      }
    }
    imgToReplace.sort(function (a, b) {
      return a[0] - b[0];
    });
    for (var i = 0; i < imgToReplace.length; i++) {
      var attId = (attachments.length - nbrOfImg) + i;
      var title = 'inlineImages' + i;
      inlineImages[title] = attachments[attId].copyBlob().setName(title);
      attachments.splice(attId, 1);
      var newImg = imgToReplace[i][1].replace(/src="[^\"]+\"/, "src=\"cid:" + title + "\"");
      emailTemplate = emailTemplate.replace(imgToReplace[i][1], newImg);
    }
  }

 objects = getRowsData(dataSheet, dataRange);
 for (var i = 0; i < objects.length; ++i) {   
   var rowData = objects[i];
    //AF: changed row to automaticRepsonseStatus
   if(rowData.automaticResponseStatus != "EMAIL_SENT"){
     
     // Replace markers (for instance ${"First Name"}) with the 
     // corresponding value in a row object (for instance rowData.firstName).
     
     var emailText = fillInTemplateFromObject(emailTemplate, rowData);     
     var emailSubject = fillInTemplateFromObject(selectedTemplate.getSubject(), rowData);
      //AF: SEND EMAIL  
     GmailApp.sendEmail(rowData.emailAddress, emailSubject, emailText,
                        {name: selectedTemplate.name, attachments: attachments, htmlBody: emailText, cc: cc, bcc: bcc, inlineImages: inlineImages});      

  
     //AF: get current date/time to use as timestamp
     var dt = new DateTime();


     //AF: PRINT TIMESTAMP
     dataSheet.getRange(i+2,dataSheet.getLastColumn()).setValue("EMAIL_SENT: "+ dt.formats.pretty.b);
     
   }  
 }
  
 var app = UiApp.getActiveApplication();
 app.close();
 return app;

}


function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var menu = [ 
    {name: "Clear Canvas (Reset)", functionName: "labnolReset"},
    {name: "Start Mail Merge", functionName: "fnMailMerge"}
    ];  
  ss.addMenu("Mail Merge", menu);
  ss.toast("Please click the Mail Merge menu above to continue..", "", 5);
}

function labnolReset() {  
  var mySheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();   
  mySheet.getRange(2, 1, mySheet.getMaxRows() - 1, mySheet.getMaxColumns()).clearContent();
}

function fnMailMerge() {
  var threads = GmailApp.search('in:draft', 0, 10);
  if (threads.length === 0) {
    Browser.msgBox("We found no templates in Gmail. Please save a template as a draft message in your Gmail mailbox and re-run the Mail Merge program.");
    return;
  }
  var myapp = UiApp.createApplication().setTitle('Mail Merge HD').setHeight(160).setWidth(300);
  var top_panel = myapp.createFlowPanel();   
  top_panel.add(myapp.createLabel("Please select your Mail Merge template"));
  var lb = myapp.createListBox(false).setWidth(250).setName('templates').addItem("Select template...").setVisibleItemCount(1);
  
  for (var i = 0; i < threads.length; i++) {
    //finding correct draft
    if(threads[i].getFirstMessageSubject() == "Thank You For Your Interest"){
      lb.addItem((i+1)+'- '+threads[i].getFirstMessageSubject().substr(0, 40));
    }
  }

  top_panel.add(lb);
  top_panel.add(myapp.createLabel("").setHeight(10));
  top_panel.add(myapp.createLabel("Please write the sender's full name"));
  var name_box = myapp.createTextBox().setName("name").setWidth(250);
  top_panel.add(name_box);  
  top_panel.add(myapp.createLabel("").setHeight(10));
  var bcc_box = myapp.createCheckBox().setName("bcc").setText("BCC yourself?").setWidth(250);
  top_panel.add(bcc_box);
  top_panel.add(myapp.createLabel("").setHeight(5));
  var ok_btn = myapp.createButton("Start Mail Merge"); 
  top_panel.add(ok_btn);
  myapp.add(top_panel);
  
  var handler = myapp.createServerClickHandler('startMailMerge').addCallbackElement(lb).addCallbackElement(name_box).addCallbackElement(bcc_box);
  ok_btn.addClickHandler(handler);
  
  SpreadsheetApp.getActiveSpreadsheet().show(myapp);
}


function startMailMerge(e) {
 var ss = SpreadsheetApp.getActiveSpreadsheet();
 var dataSheet = ss.getSheets()[0];
 if(dataSheet.getRange(1,dataSheet.getLastColumn()).getValue() != 'Automatic Response Status'){
   dataSheet.getRange(1,dataSheet.getLastColumn()+1).setValue('Automatic Response Status');
 }
 var headers = dataSheet.getRange(1, 1, 1, dataSheet.getLastColumn()).getValues();
 var emailColumnFound = false;

 //FIND THE COLUMN THAT CONTAINS EMAIL ADDRESSES
 for(i in headers[0]){
   if(headers[0][i] == "Email Address"){
     emailColumnFound = true;
   }
 }
 //IF CAN'T FIND, ASK USER TO MANUALLY INPUT
 if(!emailColumnFound){
   var emailColumn = Browser.inputBox("Which column contains the recipient's email address ? (A, B,...)");
   dataSheet.getRange(emailColumn+''+1).setValue("Email Address");
 }

 var dataRange = dataSheet.getRange(2, 1, dataSheet.getLastRow() - 1, dataSheet.getLastColumn());

//POP-UP NOTIFICATION FOR USER
 ss.toast('Starting mail merge, please wait...','Mail Merge',-1);
  

 var selectedTemplate = GmailApp.search("in:drafts")[(parseInt(e.parameter.templates.substr(0, 2))-1)].getMessages()[0];
 var emailTemplate = selectedTemplate.getBody();
 var attachments = selectedTemplate.getAttachments();
 var cc = selectedTemplate.getCc();
 var bcc = "";
 if (e.parameter.bcc == "true") {
   bcc = selectedTemplate.getFrom();
 }

  var regMessageId = new RegExp(selectedTemplate.getId(), "g");
  
  if (emailTemplate.match(regMessageId) != null) {
    var inlineImages = {};
    var nbrOfImg = emailTemplate.match(regMessageId).length;
    var imgVars = emailTemplate.match(/<img[^>]+>/g);
    var imgToReplace = [];
    for (var i = 0; i < imgVars.length; i++) {
      if (imgVars[i].search(regMessageId) != -1) {
        var id = imgVars[i].match(/Inline\simages?\s(\d)/);
        imgToReplace.push([parseInt(id[1]), imgVars[i]]);
      }
    }
    imgToReplace.sort(function (a, b) {
      return a[0] - b[0];
    });
    for (var i = 0; i < imgToReplace.length; i++) {
      var attId = (attachments.length - nbrOfImg) + i;
      var title = 'inlineImages' + i;
      inlineImages[title] = attachments[attId].copyBlob().setName(title);
      attachments.splice(attId, 1);
      var newImg = imgToReplace[i][1].replace(/src="[^\"]+\"/, "src=\"cid:" + title + "\"");
      emailTemplate = emailTemplate.replace(imgToReplace[i][1], newImg);
    }
  }

 objects = getRowsData(dataSheet, dataRange);
 for (var i = 0; i < objects.length; ++i) {   
   var rowData = objects[i];
    //AF: changed row to automaticRepsonseStatus
   if(rowData.automaticResponseStatus != "EMAIL_SENT"){
     
     // Replace markers (for instance ${"First Name"}) with the 
     // corresponding value in a row object (for instance rowData.firstName).
     
     var emailText = fillInTemplateFromObject(emailTemplate, rowData);     
     var emailSubject = fillInTemplateFromObject(selectedTemplate.getSubject(), rowData);
         
     GmailApp.sendEmail(rowData.emailAddress, emailSubject, emailText,
                        {name: e.parameter.name, attachments: attachments, htmlBody: emailText, cc: cc, bcc: bcc, inlineImages: inlineImages});      

  
     //AF: get current date/time to use as timestamp
     var dt = new DateTime();



     dataSheet.getRange(i+2,dataSheet.getLastColumn()).setValue("EMAIL_SENT: "+ dt.formats.pretty.b);
     
   }  
 }
  
 ss.toast('You can reach the script developer at amit@labnol.org for support and customization.','Mail Merge Complete',-1);
  
 var app = UiApp.getActiveApplication();
 app.close();
 return app;
}

// Replaces markers in a template string with values define in a JavaScript data object.
// Arguments:
//   - template: string containing markers, for instance ${"Column name"}
//   - data: JavaScript object with values to that will replace markers. For instance
//           data.columnName will replace marker ${"Column name"}
// Returns a string without markers. If no data is found to replace a marker, it is
// simply removed.
function fillInTemplateFromObject(template, data) {
 var email = template;
 // Search for all the variables to be replaced, for instance ${"Column name"}
 var templateVars = template.match(/\$\%[^\%]+\%/g);
 if(templateVars!= null){
   // Replace variables from the template with the actual values from the data object.
   // If no value is available, replace with the empty string.
   for (var i = 0; i < templateVars.length; ++i) {
     // normalizeHeader ignores ${"} so we can call it directly here.
     var variableData = data[normalizeHeader(templateVars[i])];
     email = email.replace(templateVars[i], variableData || "");
   }
 }
 return email;
}


/* This code is reused from the 'Reading Spreadsheet data using JavaScript Objects' tutorial */

function getRowsData(sheet, range, columnHeadersRowIndex) {
 columnHeadersRowIndex = columnHeadersRowIndex || range.getRowIndex() - 1;
 var numColumns = range.getEndColumn() - range.getColumn() + 1;
 var headersRange = sheet.getRange(columnHeadersRowIndex, range.getColumn(), 1, numColumns);
 var headers = headersRange.getValues()[0];
 return getObjects(range.getValues(), normalizeHeaders(headers));
}

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

function isCellEmpty(cellData) {
 return typeof(cellData) == "string" && cellData == "";
}

function isAlnum(char) {
 return char >= 'A' && char <= 'Z' ||
   char >= 'a' && char <= 'z' ||
   isDigit(char);
}

function isDigit(char) {
 return char >= '0' && char <= '9';
}

function DateTime() {
    function getDaySuffix(a) {
        var b = "" + a,
            c = b.length,
            d = parseInt(b.substring(c-2, c-1)),
            e = parseInt(b.substring(c-1));
        if (c == 2 && d == 1) return "th";
        switch(e) {
            case 1:
                return "st";
                break;
            case 2:
                return "nd";
                break;
            case 3:
                return "rd";
                break;
            default:
                return "th";
                break;
        };
    };

    this.getDoY = function(a) {
        var b = new Date(a.getFullYear(),0,1);
    return Math.ceil((a - b) / 86400000);
    }

    this.date = arguments.length == 0 ? new Date() : new Date(arguments);

    this.weekdays = new Array('Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday');
    this.months = new Array('January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December');
    this.daySuf = new Array( "st", "nd", "rd", "th" );

    this.day = {
        index: {
            week: "0" + this.date.getDay(),
            month: (this.date.getDate() < 10) ? "0" + this.date.getDate() : this.date.getDate()
        },
        name: this.weekdays[this.date.getDay()],
        of: {
            week: ((this.date.getDay() < 10) ? "0" + this.date.getDay() : this.date.getDay()) + getDaySuffix(this.date.getDay()),
            month: ((this.date.getDate() < 10) ? "0" + this.date.getDate() : this.date.getDate()) + getDaySuffix(this.date.getDate())
        }
    }

    this.month = {
        index: (this.date.getMonth() + 1) < 10 ? "0" + (this.date.getMonth() + 1) : this.date.getMonth() + 1,
        name: this.months[this.date.getMonth()]
    };

    this.year = this.date.getFullYear();

    this.time = {
        hour: {
            meridiem: (this.date.getHours() > 12) ? (this.date.getHours() - 12) < 10 ? "0" + (this.date.getHours() - 12) : this.date.getHours() - 12 : (this.date.getHours() < 10) ? "0" + this.date.getHours() : this.date.getHours(),
            military: (this.date.getHours() < 10) ? "0" + this.date.getHours() : this.date.getHours(),
            noLeadZero: {
                meridiem: (this.date.getHours() > 12) ? this.date.getHours() - 12 : this.date.getHours(),
                military: this.date.getHours()
            }
        },
        minute: (this.date.getMinutes() < 10) ? "0" + this.date.getMinutes() : this.date.getMinutes(),
        seconds: (this.date.getSeconds() < 10) ? "0" + this.date.getSeconds() : this.date.getSeconds(),
        milliseconds: (this.date.getMilliseconds() < 100) ? (this.date.getMilliseconds() < 10) ? "00" + this.date.getMilliseconds() : "0" + this.date.getMilliseconds() : this.date.getMilliseconds(),
        meridiem: (this.date.getHours() > 12) ? "PM" : "AM"
    };

    this.sym = {
        d: {
            d: this.date.getDate(),
            dd: (this.date.getDate() < 10) ? "0" + this.date.getDate() : this.date.getDate(),
            ddd: this.weekdays[this.date.getDay()].substring(0, 3),
            dddd: this.weekdays[this.date.getDay()],
            ddddd: ((this.date.getDate() < 10) ? "0" + this.date.getDate() : this.date.getDate()) + getDaySuffix(this.date.getDate()),
            m: this.date.getMonth() + 1,
            mm: (this.date.getMonth() + 1) < 10 ? "0" + (this.date.getMonth() + 1) : this.date.getMonth() + 1,
            mmm: this.months[this.date.getMonth()].substring(0, 3),
            mmmm: this.months[this.date.getMonth()],
            yy: (""+this.date.getFullYear()).substr(2, 2),
            yyyy: this.date.getFullYear()
        },
        t: {
            h: (this.date.getHours() > 12) ? this.date.getHours() - 12 : this.date.getHours(),
            hh: (this.date.getHours() > 12) ? (this.date.getHours() - 12) < 10 ? "0" + (this.date.getHours() - 12) : this.date.getHours() - 12 : (this.date.getHours() < 10) ? "0" + this.date.getHours() : this.date.getHours(),
            hhh: this.date.getHours(),
            m: this.date.getMinutes(),
            mm: (this.date.getMinutes() < 10) ? "0" + this.date.getMinutes() : this.date.getMinutes(),
            s: this.date.getSeconds(),
            ss: (this.date.getSeconds() < 10) ? "0" + this.date.getSeconds() : this.date.getSeconds(),
            ms: this.date.getMilliseconds(),
            mss: Math.round(this.date.getMilliseconds()/10) < 10 ? "0" + Math.round(this.date.getMilliseconds()/10) : Math.round(this.date.getMilliseconds()/10),
            msss: (this.date.getMilliseconds() < 100) ? (this.date.getMilliseconds() < 10) ? "00" + this.date.getMilliseconds() : "0" + this.date.getMilliseconds() : this.date.getMilliseconds()
        }
    };

    this.formats = {
        compound: {
            commonLogFormat: this.sym.d.dd + "/" + this.sym.d.mmm + "/" + this.sym.d.yyyy + ":" + this.sym.t.hhh + ":" + this.sym.t.mm + ":" + this.sym.t.ss,
            exif: this.sym.d.yyyy + ":" + this.sym.d.mm + ":" + this.sym.d.dd + " " + this.sym.t.hhh + ":" + this.sym.t.mm + ":" + this.sym.t.ss,
            /*iso1: "",
            iso2: "",*/
            mySQL: this.sym.d.yyyy + "-" + this.sym.d.mm + "-" + this.sym.d.dd + " " + this.sym.t.hhh + ":" + this.sym.t.mm + ":" + this.sym.t.ss,
            postgreSQL1: this.sym.d.yyyy + "." + this.getDoY(this.date),
            postgreSQL2: this.sym.d.yyyy + "" + this.getDoY(this.date),
            soap: this.sym.d.yyyy + "-" + this.sym.d.mm + "-" + this.sym.d.dd + "T" + this.sym.t.hhh + ":" + this.sym.t.mm + ":" + this.sym.t.ss + "." + this.sym.t.mss,
            //unix: "",
            xmlrpc: this.sym.d.yyyy + "" + this.sym.d.mm + "" + this.sym.d.dd + "T" + this.sym.t.hhh + ":" + this.sym.t.mm + ":" + this.sym.t.ss,
            xmlrpcCompact: this.sym.d.yyyy + "" + this.sym.d.mm + "" + this.sym.d.dd + "T" + this.sym.t.hhh + "" + this.sym.t.mm + "" + this.sym.t.ss,
            wddx: this.sym.d.yyyy + "-" + this.sym.d.m + "-" + this.sym.d.d + "T" + this.sym.t.h + ":" + this.sym.t.m + ":" + this.sym.t.s
        },
        constants: {
            atom: this.sym.d.yyyy + "-" + this.sym.d.mm + "-" + this.sym.d.dd + "T" + this.sym.t.hhh + ":" + this.sym.t.mm + ":" + this.sym.t.ss,
            cookie: this.sym.d.dddd + ", " + this.sym.d.dd + "-" + this.sym.d.mmm + "-" + this.sym.d.yy + " " + this.sym.t.hhh + ":" + this.sym.t.mm + ":" + this.sym.t.ss,
            iso8601: this.sym.d.yyyy + "-" + this.sym.d.mm + "-" + this.sym.d.dd + "T" + this.sym.t.hhh + ":" + this.sym.t.mm + ":" + this.sym.t.ss,
            rfc822: this.sym.d.ddd + ", " + this.sym.d.dd + " " + this.sym.d.mmm + " " + this.sym.d.yy + " " + this.sym.t.hhh + ":" + this.sym.t.mm + ":" + this.sym.t.ss,
            rfc850: this.sym.d.dddd + ", " + this.sym.d.dd + "-" + this.sym.d.mmm + "-" + this.sym.d.yy + " " + this.sym.t.hhh + ":" + this.sym.t.mm + ":" + this.sym.t.ss,
            rfc1036: this.sym.d.ddd + ", " + this.sym.d.dd + " " + this.sym.d.mmm + " " + this.sym.d.yy + " " + this.sym.t.hhh + ":" + this.sym.t.mm + ":" + this.sym.t.ss,
            rfc1123: this.sym.d.ddd + ", " + this.sym.d.dd + " " + this.sym.d.mmm + " " + this.sym.d.yyyy + " " + this.sym.t.hhh + ":" + this.sym.t.mm + ":" + this.sym.t.ss,
            rfc2822: this.sym.d.ddd + ", " + this.sym.d.dd + " " + this.sym.d.mmm + " " + this.sym.d.yyyy + " " + this.sym.t.hhh + ":" + this.sym.t.mm + ":" + this.sym.t.ss,
            rfc3339: this.sym.d.yyyy + "-" + this.sym.d.mm + "-" + this.sym.d.dd + "T" + this.sym.t.hhh + ":" + this.sym.t.mm + ":" + this.sym.t.ss,
            rss: this.sym.d.ddd + ", " + this.sym.d.dd + " " + this.sym.d.mmm + " " + this.sym.d.yy + " " + this.sym.t.hhh + ":" + this.sym.t.mm + ":" + this.sym.t.ss,
            w3c: this.sym.d.yyyy + "-" + this.sym.d.mm + "-" + this.sym.d.dd + "T" + this.sym.t.hhh + ":" + this.sym.t.mm + ":" + this.sym.t.ss
        },
        pretty: {
            a: this.sym.t.hh + ":" + this.sym.t.mm + "." + this.sym.t.ss + this.time.meridiem + " " + this.sym.d.dddd + " " + this.sym.d.ddddd + " of " + this.sym.d.mmmm + ", " + this.sym.d.yyyy,
            b: this.sym.t.hh + ":" + this.sym.t.mm + " " + this.sym.d.dddd + " " + this.sym.d.ddddd + " of " + this.sym.d.mmmm + ", " + this.sym.d.yyyy
        }
    };
};