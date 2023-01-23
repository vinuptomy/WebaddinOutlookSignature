// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
function get_template_Standard_str(user_info) {
  var str = "";
  var template_name = $("#new_mail option:selected").val();
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

  if (template_name === "templateS") {
    str += "<span style='color: black; font-family: Arial, sans-serif; font-size: 10pt; font-weight: bold;'>Drei.</span>";
    str += "<span>Macht's einfach.</span>";
  }

  if (template_name === "templateB") {
    str += "<span style='color: black; font-family: Arial, sans-serif; font-size: 10pt; font-weight: bold;'>Drei</span>";
    str += "<span style='color: rgb(0, 160, 210); font-family: Arial, sans-serif; font-size: 10pt;'>Business.</span>";
    str += "<div><span>Macht's einfach.</span></div>";
  }

  if (template_name === "templateW") {
    str += "<span style='color: black; font-family: Arial, sans-serif; font-size: 10pt; font-weight: bold;'>Drei</span>";
    str += "<span style='color: rgb(0, 160, 210); font-family: Arial, sans-serif; font-size: 10pt;'>Wholesale.</span>";
    str += "<div><span>Macht's einfach.</span></div>";
  }

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
  return str;
}