/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Excel, Office, Word, Promise */

export function writeDataToOfficeDocument(result) {
  return new Promise(function (resolve, reject) {
    try {
      switch (Office.context.host) {
        case Office.HostType.Excel:
          writeDataToExcel(result);
          break;
        case Office.HostType.Outlook:
          writeDataToOutlook(result);
          break;
        case Office.HostType.PowerPoint:
          writeDataToPowerPoint(result);
          break;
        case Office.HostType.Word:
          writeDataToWord(result);
          break;
        default:
          throw "Unsupported Office host application: This add-in only runs on Excel, Outlook, PowerPoint, or Word.";
      }
      resolve();
    } catch (error) {
      reject(Error("Unable to write data to document. " + error.toString()));
    }
  });
}



export function writeOfficeTimeData(result) {
  return new Promise(function (resolve, reject) {
    try {
        var filteredres = result.value.filter(function (el) {        
        return el.fields.Title === $("input#PlaceHolderMain_txtLocation").val();
      });
      writeofficetimeDataToOutlook(filteredres);
      resolve();
    } catch (error) {
      reject(Error("Unable to write office time data to outlook. " + error.toString()));
    }
  });
}


function filterUserProfileInfo(result) {
  let userProfileInfo = [];
  userProfileInfo.push(result["givenName"]);
  userProfileInfo.push(result["surname"]);
  $("input#PlaceHolderMain_txtName").val(result["givenName"] + " " + result["surname"]);
  userProfileInfo.push(result["jobTitle"]);
  $("input#PlaceHolderMain_txtJobtitle").val(result["jobTitle"]);
  userProfileInfo.push(result["mail"]);
  userProfileInfo.push(result["mobilePhone"]);
  $("input#PlaceHolderMain_txtMobile").val(result["mobilePhone"]);
 $("input#PlaceHolderMain_txtPhone").val(result["businessPhones"][0]);
 if (result["officeLocation"].trim() === "HQ-EN074") {  
  $("input#PlaceHolderMain_txtLocation").val("Headquarter"); 
 }
 else {
  $("input#PlaceHolderMain_txtLocation").val(result["officeLocation"]); 
 }
 $("input#PlaceHolderMain_txtAddress").val(result["streetAddress"] + ", " + result["postalCode"] + " " + result["city"]); 
 $("input#PlaceHolderMain_txtFax").val(result["faxNumber"]);
  userProfileInfo.push(result["officeLocation"]);
  return userProfileInfo;
}


function writeofficetimeDataToOutlook(result) {  
 if (result){  
$("textarea#PlaceHolderMain_tbOpeningHours").val(result[0].fields.OpeningHours);
}
}


function writeDataToExcel(result) {
  return Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    let data = [];
    let userProfileInfo = filterUserProfileInfo(result);

    for (let i = 0; i < userProfileInfo.length; i++) {
      if (userProfileInfo[i] !== null) {
        let innerArray = [];
        innerArray.push(userProfileInfo[i]);
        data.push(innerArray);
      }
    }
    const rangeAddress = `B5:B${5 + (data.length - 1)}`;
    const range = sheet.getRange(rangeAddress);
    range.values = data;
    range.format.autofitColumns();

    return context.sync();
  });
}

function writeDataToOutlook(result) {
  let data = [];
  let userProfileInfo = filterUserProfileInfo(result);

  for (let i = 0; i < userProfileInfo.length; i++) {
    if (userProfileInfo[i] !== null) {
      data.push(userProfileInfo[i]);
    }
  }

  let userInfo = "";
  for (let i = 0; i < data.length; i++) {
    userInfo += data[i] + "\n";
  }

  //Office.context.mailbox.item.body.setSelectedDataAsync(userInfo, { coercionType: Office.CoercionType.Html });
}

function writeDataToPowerPoint(result) {
  let data = [];
  let userProfileInfo = filterUserProfileInfo(result);

  for (let i = 0; i < userProfileInfo.length; i++) {
    if (userProfileInfo[i] !== null) {
      data.push(userProfileInfo[i]);
    }
  }

  let userInfo = "";
  for (let i = 0; i < data.length; i++) {
    userInfo += data[i] + "\n";
  }
  Office.context.document.setSelectedDataAsync(userInfo, function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      throw asyncResult.error.message;
    }
  });
}

function writeDataToWord(result) {
  return Word.run(function (context) {
    let data = [];
    let userProfileInfo = filterUserProfileInfo(result);

    for (let i = 0; i < userProfileInfo.length; i++) {
      if (userProfileInfo[i] !== null) {
        data.push(userProfileInfo[i]);
      }
    }

    const documentBody = context.document.body;
    for (let i = 0; i < data.length; i++) {
      if (data[i] !== null) {
        documentBody.insertParagraph(data[i], "End");
      }
    }
    return context.sync();
  });
}
