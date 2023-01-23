// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// Contains code for event-based activation on Outlook on web, on Windows, and on Mac (new UI preview).

/**
 * Checks if signature exists.
 * If not, displays the information bar for user.
 * If exists, insert signature into new item (appointment or message).
 * @param {*} eventObj Office event object
 * @returns
 */
function checkSignature(eventObj) {
  var user_info_str = Office.context.roamingSettings.get("user_info");
  console.log("check signature - " + user_info_str);

  if (!user_info_str) {
    display_insight_infobar();
  } else {
    var user_info = JSON.parse(user_info_str);

    if (Office.context.mailbox.item.getComposeTypeAsync) {
      //Find out if the compose type is "newEmail", "reply", or "forward" so that we can apply the correct template.
      Office.context.mailbox.item.getComposeTypeAsync({
        asyncContext: {
          user_info: user_info,
          eventObj: eventObj
        }
      }, function (asyncResult) {
        if (asyncResult.status === "succeeded") {
          insert_auto_signature(asyncResult.value.composeType, asyncResult.asyncContext.user_info, asyncResult.asyncContext.eventObj);
        }
      });
    } else {
      // Appointment item. Just use newMail pattern
      var _user_info = JSON.parse(user_info_str);

      insert_auto_signature("newMail", _user_info, eventObj);
    }
  }
}
/**
 * For Outlook on Windows and on Mac only. Insert signature into appointment or message.
 * Outlook on Windows and on Mac can use setSignatureAsync method on appointments and messages.
 * @param {*} compose_type The compose type (reply, forward, newMail)
 * @param {*} user_info Information details about the user
 * @param {*} eventObj Office event object
 */


function insert_auto_signature(compose_type, user_info, eventObj) {
  var template_name = get_template_name(compose_type);
  var signature_info = get_signature_info(template_name, user_info);
  addTemplateSignature(signature_info, eventObj);
}
/**
 * 
 * @param {*} signatureDetails object containing:
 *  "signature": The signature HTML of the template,
    "logoBase64": The base64 encoded logo image,
    "logoFileName": The filename of the logo image
 * @param {*} eventObj 
 * @param {*} signatureImageBase64 
 */


function addTemplateSignature(signatureDetails, eventObj, signatureImageBase64) {
  if (is_valid_data(signatureDetails.logoBase64) === true) {
    //If a base64 image was passed we need to attach it.
    Office.context.mailbox.item.addFileAttachmentFromBase64Async(signatureDetails.logoBase64, signatureDetails.logoFileName, {
      isInline: true
    }, function (result) {
      //After image is attached, insert the signature
      Office.context.mailbox.item.body.setSignatureAsync(signatureDetails.signature, {
        coercionType: "html",
        asyncContext: eventObj
      }, function (asyncResult) {
        asyncResult.asyncContext.completed();
      });
    });
  } else {
    //Image is not embedded, or is referenced from template HTML
    Office.context.mailbox.item.body.setSignatureAsync(signatureDetails.signature, {
      coercionType: "html",
      asyncContext: eventObj
    }, function (asyncResult) {
      asyncResult.asyncContext.completed();
    });
  }
}
/**
 * Creates information bar to display when new message or appointment is created
 */


function display_insight_infobar() {
  console.log('its  enetered display infobar');
  Office.context.mailbox.item.notificationMessages.addAsync("fd90eb33431b46f58a68720c36154b4a", {
    type: "insightMessage",
    message: "Please set your signature with Drei Outlook add-in.",
    icon: "Icon.16x16",
    actions: [{
      actionType: "showTaskPane",
      actionText: "Signatures Generator",
      commandId: get_command_id(),
      contextData: "{''}"
    }]
  });
}
/**
 * Gets template name  mapped based on the compose type, in our case there is no mappin for compose per type
 * @param {*} compose_type The compose type (reply, forward, newMail)
 * @returns Name of the template to use for the compose type
 */


function get_template_name(compose_type) {
  if (compose_type === "reply") return Office.context.roamingSettings.get("newMail");
  if (compose_type === "forward") return Office.context.roamingSettings.get("newMail");
  return Office.context.roamingSettings.get("newMail");
}
/**
 * Gets HTML signature in requested template format for given user
 * @param {\} template_name Which template format to use (A,B,C)
 * @param {*} user_info Information details about the user
 * @returns HTML signature in requested template format
 */


function get_signature_info(template_name, user_info) {
  if (template_name === "templateB") return get_template_B_info(user_info);
  if (template_name === "templateW") return get_template_W_info(user_info);
  return get_template_S_info(user_info);
}
/**
 * Gets correct command id to match to item type (appointment or message)
 * @returns The command id
 */


function get_command_id() {
  if (Office.context.mailbox.item.itemType == "appointment") {
    return "MRCS_TpBtn1";
  }

  return "MRCS_TpBtn0";
}
/**
 * Gets HTML string for template Standard
 * References the signature logo image from the HTML
 * @param {*} user_info Information details about the user
 * @returns Object containing:
 *  "signature": The signature HTML of template standard,
    "logoBase64": null since this template references the image and does not embed it ,
    "logoFileName": null since this template references the image and does not embed it
 */


function get_template_S_info(user_info) {
  var str = "";
  str += "<table style='color: black; font-family: Arial, sans-serif; font-size: 8pt;'>";
  str += "<tr>";
  str += "<td style='padding-left: 5px;'>";
  str += "<span style='color: rgb(0, 160, 210); font-family: Arial, sans-serif;font-size: 10pt;font-weight: bold;'>" + user_info.name + "</span>";
  str += "<br/>";

  if (is_valid_data(user_info.job)) {
    str += "<div style='color: black; font-family: Arial, sans-serif; font-size: 8pt;'>" + user_info.job + "</div>";
  }

  str += "<br/>";

  if (is_valid_data(user_info.phone)) {
    str += "<div>Mobil: " + user_info.phone + "</div>";
  }

  if (user_info.fixflag && is_valid_data(user_info.fix)) {
    str += "<div>Fix: " + user_info.fix + "</div>";
  }

  if (user_info.faxflag && is_valid_data(user_info.fax)) {
    str += "<div>Fax: " + user_info.fax + "</div>";
  }

  str += "<div style='color: black; font-family: Arial, sans-serif; font-size: 8pt;' ><a href=mailto:" + user_info.email + ">" + user_info.email + "</a></div>";
  str += "<div style='color: black; font-family: Arial, sans-serif; font-size: 8pt;' ><a href='https://www.drei.at'>www.drei.at</a></div>";
  str += "</td>";
  str += "</tr>";
  str += "<td style='padding-left: 5px;'>";
  str += "<br/>";
  str += "<img src='https://localhost:3000/assets/DreiLogo-64.png' alt='Logo' />";
  str += "<br/>";
  str += "<br/>";
  str += "</td>";
  str += "</tr>";
  str += "<tr>";
  str += "<td style='padding-left: 5px;'>";
  str += "<span style='color: black; font-family: Arial, sans-serif; font-size: 10pt; font-weight: bold;'>Drei.</span>";
  str += "<span>Macht's einfach.</span>";
  str += "</td>";
  str += "</tr>";
  str += "<tr>";
  str += "<td style='padding-left: 5px;'>";
  str += "<br/>";
  str += "<div style='color: black; font-family: Arial, sans-serif; font-size: 8pt; font-weight: bold;'>Hutchison Drei Austria GmbH</div>";

  if (user_info.location.toLowerCase().startsWith("drei shop")) {
    str += "<div style='color: black; font-family: Arial, sans-serif; font-size: 8pt; font-weight: bold;'>" + user_info.location + "</div>";
  }

  str += "<div style='color: black; font-family: Arial, sans-serif; font-size: 8pt;'>" + user_info.address + "</div>";
  str += "<div style='color: black; font-family: Arial, sans-serif; font-size: 8pt;'>FN 140132b, Handelsgericht Wien</div>";
  str += "</td>";
  str += "</tr>";

  if (user_info.officetimeflag && is_valid_data(user_info.officetime)) {
    var formatedofficetime = user_info.officetime.replaceAll("Uhr", "Uhr <br />");
    console.log(formatedofficetime);
    str += "<tr>";
    str += "<td style='padding-left: 5px;'>";
    str += "<br/>";
    str += "<div style='color: black; font-family: Arial, sans-serif; font-size: 8pt;'>" + formatedofficetime + "</div>";
    str += "</td>";
    str += "</tr>";
  }

  str += "</table>";
  return {
    signature: str,
    logoBase64: null,
    logoFileName: null
  };
}
/**
 * Gets HTML string for template Business
 * References the signature logo image from the HTML
 * @param {*} user_info Information details about the user
 * @returns Object containing:
 *  "signature": The signature HTML of template Business,
    "logoBase64": null since this template references the image and does not embed it ,
    "logoFileName": null since this template references the image and does not embed it
 */


function get_template_B_info(user_info) {
  var str = "";
  str += "<table style='color: black; font-family: Arial, sans-serif; font-size: 8pt;'>";
  str += "<tr>";
  str += "<td style='padding-left: 5px;'>";
  str += "<span style='color: rgb(0, 160, 210); font-family: Arial, sans-serif;font-size: 10pt;font-weight: bold;'>" + user_info.name + "</span>";
  str += "<br/>";

  if (is_valid_data(user_info.job)) {
    str += "<div style='color: black; font-family: Arial, sans-serif; font-size: 8pt;'>" + user_info.job + "</div>";
  }

  str += "<br/>";

  if (is_valid_data(user_info.phone)) {
    str += "<div>Mobil: " + user_info.phone + "</div>";
  }

  if (user_info.fixflag && is_valid_data(user_info.fix)) {
    str += "<div>Fix: " + user_info.fix + "</div>";
  }

  if (user_info.faxflag && is_valid_data(user_info.fax)) {
    str += "<div>Fax: " + user_info.fax + "</div>";
  }

  str += "<div style='color: black; font-family: Arial, sans-serif; font-size: 8pt;' ><a href=mailto:" + user_info.email + ">" + user_info.email + "</a></div>";
  str += "<div style='color: black; font-family: Arial, sans-serif; font-size: 8pt;' ><a href='https://www.drei.at'>www.drei.at</a></div>";
  str += "</td>";
  str += "</tr>";
  str += "<td style='padding-left: 5px;'>";
  str += "<br/>";
  str += "<img src='https://localhost:3000/assets/DreiLogo-64.png' alt='Logo' />";
  str += "<br/>";
  str += "<br/>";
  str += "</td>";
  str += "</tr>";
  str += "<tr>";
  str += "<td style='padding-left: 5px;'>";
  str += "<span style='color: black; font-family: Arial, sans-serif; font-size: 10pt; font-weight: bold;'>Drei</span>";
  str += "<span style='color: rgb(0, 160, 210); font-family: Arial, sans-serif; font-size: 10pt;'>Business.</span>";
  str += "<div><span>Macht's einfach.</span></div>";
  str += "</td>";
  str += "</tr>";
  str += "<tr>";
  str += "<td style='padding-left: 5px;'>";
  str += "<br/>";
  str += "<div style='color: black; font-family: Arial, sans-serif; font-size: 8pt; font-weight: bold;'>Hutchison Drei Austria GmbH</div>";

  if (user_info.location.toLowerCase().startsWith("drei shop")) {
    str += "<div style='color: black; font-family: Arial, sans-serif; font-size: 8pt; font-weight: bold;'>" + user_info.location + "</div>";
  }

  str += "<div style='color: black; font-family: Arial, sans-serif; font-size: 8pt;'>" + user_info.address + "</div>";
  str += "<div style='color: black; font-family: Arial, sans-serif; font-size: 8pt;'>FN 140132b, Handelsgericht Wien</div>";
  str += "</td>";
  str += "</tr>";

  if (user_info.officetimeflag && is_valid_data(user_info.officetime)) {
    var formatedofficetime = user_info.officetime.replaceAll("Uhr", "Uhr <br />");
    console.log(formatedofficetime);
    str += "<tr>";
    str += "<td style='padding-left: 5px;'>";
    str += "<br/>";
    str += "<div style='color: black; font-family: Arial, sans-serif; font-size: 8pt;'>" + formatedofficetime + "</div>";
    str += "</td>";
    str += "</tr>";
  }

  str += "</table>";
  return {
    signature: str,
    logoBase64: null,
    logoFileName: null
  };
}
/**
 * Gets HTML string for template Wholesale
 * @param {*} user_info Information details about the user
 * @returns Object containing:
 *  "signature": The signature HTML of template Wholesale,
    "logoBase64": null since there is no image,
    "logoFileName": null since there is no image
 */


function get_template_W_info(user_info) {
  var str = "";
  str += "<table style='color: black; font-family: Arial, sans-serif; font-size: 8pt;'>";
  str += "<tr>";
  str += "<td style='padding-left: 5px;'>";
  str += "<span style='color: rgb(0, 160, 210); font-family: Arial, sans-serif;font-size: 10pt;font-weight: bold;'>" + user_info.name + "</span>";
  str += "<br/>";

  if (is_valid_data(user_info.job)) {
    str += "<div style='color: black; font-family: Arial, sans-serif; font-size: 8pt;'>" + user_info.job + "</div>";
  }

  str += "<br/>";

  if (is_valid_data(user_info.phone)) {
    str += "<div>Mobil: " + user_info.phone + "</div>";
  }

  if (user_info.fixflag && is_valid_data(user_info.fix)) {
    str += "<div>Fix: " + user_info.fix + "</div>";
  }

  if (user_info.faxflag && is_valid_data(user_info.fax)) {
    str += "<div>Fax: " + user_info.fax + "</div>";
  }

  str += "<div style='color: black; font-family: Arial, sans-serif; font-size: 8pt;' ><a href=mailto:" + user_info.email + ">" + user_info.email + "</a></div>";
  str += "<div style='color: black; font-family: Arial, sans-serif; font-size: 8pt;' ><a href='https://www.drei.at'>www.drei.at</a></div>";
  str += "</td>";
  str += "</tr>";
  str += "<td style='padding-left: 5px;'>";
  str += "<br/>";
  str += "<img src='https://localhost:3000/assets/DreiLogo-64.png' alt='Logo' />";
  str += "<br/>";
  str += "<br/>";
  str += "</td>";
  str += "</tr>";
  str += "<tr>";
  str += "<td style='padding-left: 5px;'>";
  str += "<span style='color: black; font-family: Arial, sans-serif; font-size: 10pt; font-weight: bold;'>Drei</span>";
  str += "<span style='color: rgb(0, 160, 210); font-family: Arial, sans-serif; font-size: 10pt;'>Wholesale.</span>";
  str += "<div><span>Macht's einfach.</span></div>";
  str += "</td>";
  str += "</tr>";
  str += "<tr>";
  str += "<td style='padding-left: 5px;'>";
  str += "<br/>";
  str += "<div style='color: black; font-family: Arial, sans-serif; font-size: 8pt; font-weight: bold;'>Hutchison Drei Austria GmbH</div>";

  if (user_info.location.toLowerCase().startsWith("drei shop")) {
    str += "<div style='color: black; font-family: Arial, sans-serif; font-size: 8pt; font-weight: bold;'>" + user_info.location + "</div>";
  }

  str += "<div style='color: black; font-family: Arial, sans-serif; font-size: 8pt;'>" + user_info.address + "</div>";
  str += "<div style='color: black; font-family: Arial, sans-serif; font-size: 8pt;'>FN 140132b, Handelsgericht Wien</div>";
  str += "</td>";
  str += "</tr>";

  if (user_info.officetimeflag && is_valid_data(user_info.officetime)) {
    var formatedofficetime = user_info.officetime.replaceAll("Uhr", "Uhr <br />");
    console.log(formatedofficetime);
    str += "<tr>";
    str += "<td style='padding-left: 5px;'>";
    str += "<br/>";
    str += "<div style='color: black; font-family: Arial, sans-serif; font-size: 8pt;'>" + formatedofficetime + "</div>";
    str += "</td>";
    str += "</tr>";
  }

  str += "</table>";
  return {
    signature: str,
    logoBase64: null,
    logoFileName: null
  };
}
/**
 * Validates if str parameter contains text.
 * @param {*} str String to validate
 * @returns true if string is valid; otherwise, false.
 */


function is_valid_data(str) {
  return str !== null && str !== undefined && str !== "";
}

Office.actions.associate("checkSignature", checkSignature);