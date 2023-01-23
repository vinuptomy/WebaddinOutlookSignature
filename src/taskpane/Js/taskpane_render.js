// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

let _display_name;
let _job_title;
let _email_id;
let _phone_number;
let _office_location;
let _office_address;
let _office_times;
let _fax;
let _fix_line;
let _message;
var fixchecked = false;
var faxchecked = false;
var officetimechecked = false;
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
/* global document, Office, require */

const ssoAuthHelper = require("../../helpers/ssoauthhelper");

Office.initialize = function(reason)
{
  on_initialization_complete();
}

function on_initialization_complete()
{
	$(document).ready
	(
		function()
		{
      _output = $("textarea#output");     
      _display_name = $("input#PlaceHolderMain_txtName");
      _email_id = $("input#PlaceHolderMain_txtEmail");
      _job_title = $("input#PlaceHolderMain_txtJobtitle");
      _phone_number = $("input#PlaceHolderMain_txtMobile"); 
      _office_location = $("input#PlaceHolderMain_txtLocation"); 
      _office_address = $("input#PlaceHolderMain_txtAddress");      
      _fax = $("input#PlaceHolderMain_txtFax");
       _fix_line = $("input#PlaceHolderMain_txtPhone");
      _office_times =$("input#PlaceHolderMain_tbOpeningHours");
      _message = $("p#message");

      prepopulate_from_userprofile();
      load_saved_user_info();
      ssoAuthHelper.getGraphData();      
		}
	);
}

function prepopulate_from_userprofile()
{
  _display_name.val($("input#PlaceHolderMain_txtName"));
  _email_id.val(Office.context.mailbox.userProfile.emailAddress); 
}


function load_saved_user_info()
{
  let user_info_str = localStorage.getItem('user_info');
  if (!user_info_str)
  {
    user_info_str = Office.context.roamingSettings.get('user_info');
  }

  if (user_info_str)
  {
    const user_info = JSON.parse(user_info_str);

    $("input#PlaceHolderMain_txtName").val(user_info.name);
    $("input#PlaceHolderMain_txtEmail").val(user_info.email);
    $("input#PlaceHolderMain_txtJobtitle").val(user_info.job);
    $("input#PlaceHolderMain_txtMobile").val(user_info.phone);
    $("input#PlaceHolderMain_txtLocation").val(user_info.location);
    $("input#PlaceHolderMain_txtAddress").val(user_info.address);
    if(user_info.fixflag)
    {
      $('td input#PlaceHolderMain_chkPhone').prop('checked', true); 

    }
    $("input#PlaceHolderMain_txtPhone").val(user_info.fix);
    if(user_info.faxflag)
    {
      $('td input#PlaceHolderMain_chkFax').prop('checked', true);
    }
    if(user_info.officetimeflag)
    {
      $('td input#PlaceHolderMain_chkOpeningHours').prop('checked', true);
    }
    $("input#PlaceHolderMain_txtFax").val(user_info.fax);
    $("textarea#PlaceHolderMain_tbOpeningHours").val(user_info.officetime);

  }
}

function display_message(msg)
{
   $("p#message").text(msg);
}

function clear_message()
{  
  $("p#message").text("");
}

function is_not_valid_text(text)
{
  return text.length <= 0;
}

function is_not_valid_email_address(email_address)
{
  let email_address_regex = /^\w+([\.-]?\w+)*@\w+([\.-]?\w+)*(\.\w{2,3})+$/;
  return is_not_valid_text(email_address) || !(email_address_regex.test(email_address));
}

function form_has_valid_data(email,name)
{
  if (is_not_valid_text(name))
  {
    display_message("Please enter a valid name.");
    return false;
  }

  if (is_not_valid_email_address(email))
  {
    display_message("Please enter a valid email address.");
    return false;
  }

  return true;
}
function updateFixfieldFlag()
{
  fixchecked = $('td input#PlaceHolderMain_chkPhone').prop('checked');
  
}

function updateFaxfieldFlag()
{
  faxchecked = $('td input#PlaceHolderMain_chkFax').prop('checked');  
  
}  

function updateOfficeTimefieldFlag()
{
  officetimechecked =  $('td input#PlaceHolderMain_chkOpeningHours').prop('checked'); 
  
} 
function navigate_to_taskpane_assignsignature()
{
  window.location.href = '/assignsignature.html';
}

function create_user_info()
{
 
  let email =  $("input#PlaceHolderMain_txtEmail").val().trim();  
  let name =  $("input#PlaceHolderMain_txtName").val().trim();  
  clear_message();

  if (form_has_valid_data(email,name))
  {
    clear_message();

    let user_info = {};

    user_info.name = name;
    user_info.email = email;
    user_info.job = $("input#PlaceHolderMain_txtJobtitle").val().trim();
    user_info.phone =  $("input#PlaceHolderMain_txtMobile").val().trim();
    user_info.location =  $("input#PlaceHolderMain_txtLocation").val().trim();
    user_info.address =  $("input#PlaceHolderMain_txtAddress").val().trim();
    user_info.fixflag =  $('td input#PlaceHolderMain_chkPhone').prop('checked');   
    console.log("fixflag " +user_info.fixflag);
    user_info.fix =  $("input#PlaceHolderMain_txtPhone").val().trim();  
    user_info.faxflag =  $('td input#PlaceHolderMain_chkFax').prop('checked');
    console.log("faxfalg " +user_info.faxflag); 
    user_info.fax =  $("input#PlaceHolderMain_txtFax").val().trim();
    user_info.officetimeflag =  $('td input#PlaceHolderMain_chkOpeningHours').prop('checked'); 
    console.log("officetimeflag " +user_info.officetimeflag); 
    user_info.officetime = $("textarea#PlaceHolderMain_tbOpeningHours").val().trim();
    console.log(user_info);
    localStorage.setItem('user_info', JSON.stringify(user_info));
    navigate_to_taskpane_assignsignature();
  }
}

function clear_all_fields()
{   
    $("input#PlaceHolderMain_txtName").val("");
    $("input#PlaceHolderMain_txtEmail").val("");
    $("input#PlaceHolderMain_txtJobtitle").val("");
    $("input#PlaceHolderMain_txtMobile").val("");      
    $("input#PlaceHolderMain_txtFax").val("");
    $("input#PlaceHolderMain_txtPhone").val("");
    $("input#PlaceHolderMain_txtLocation").val("");
    $("input#PlaceHolderMain_txtAddress").val("");
    $("textarea#PlaceHolderMain_tbOpeningHours").val("");
    $("select#PlaceHolderMain_ddlRegion").prop('selectedIndex',0);
    $('td input#PlaceHolderMain_chkPhone').prop('checked', false);
    $('td input#PlaceHolderMain_chkFax').prop('checked', false);
    $('td input#PlaceHolderMain_chkOpeningHours').prop('checked', false);
}

function clear_all_localstorage_data()
{
  localStorage.removeItem('user_info');
  localStorage.removeItem('newMail');
  localStorage.removeItem('reply');
  localStorage.removeItem('forward');
  localStorage.removeItem('override_olk_signature');
}

function clear_roaming_settings()
{
  Office.context.roamingSettings.remove('user_info');
  Office.context.roamingSettings.remove('newMail');
  Office.context.roamingSettings.remove('reply');
  Office.context.roamingSettings.remove('forward');
  Office.context.roamingSettings.remove('override_olk_signature');

  Office.context.roamingSettings.saveAsync
  (
    function (asyncResult)
    {
      console.log("clear_roaming_settings - " + JSON.stringify(asyncResult));

      let message = "All settings reset successfully! This add-in won't insert any signatures. You can close this pane now.";
      if (asyncResult.status === Office.AsyncResultStatus.Failed)
      {
        message = "Failed to reset. Please try again.";
      }

      display_message(message);
    }
  );
}

function clear_signature_in_body()
{
  Office.context.mailbox.item.body.setSignatureAsync
  (
	"",

	{
		coercionType: Office.CoercionType.Html
	},

	function (asyncResult)
	{
	  console.log("clear_signature - " + JSON.stringify(asyncResult));
	}
  );
}

function save_select_template()
{
  create_user_info();
}
function reset_all_configuration()
{
  clear_all_fields();
  clear_all_localstorage_data();
  clear_roaming_settings();
  clear_signature_in_body();
}
