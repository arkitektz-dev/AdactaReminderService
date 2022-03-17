/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("app-submit").onclick = closeItem;
    document.getElementById("app-refresh").onclick = saveAppointmentInfo;
    document.getElementById("email").onkeyup = validateEmail;
    document.getElementById("sms").onkeyup = validateSms;
    //document.getElementById("run").onclick = run;
    //document.getElementById("email").onclick = email;
    run();
  }
});

export async function run() {
  var mailbox = Office.context.mailbox;

  document.querySelector("#autosave").checked = true;
  var item = await mailbox.item;
  item.loadCustomPropertiesAsync((result) => {
    const props = result.value;
    //const email = props.get("ARS-Email");
    if (props.get("ARS-Email") !== undefined) document.getElementById("email").value = props.get("ARS-Email");
    if (props.get("ARS-Phone") !== undefined) document.getElementById("sms").value = props.get("ARS-Phone");
    if (props.get("ARS-Host") !== undefined) document.getElementById("host").value = props.get("ARS-Host");
    if (props.get("ARS-Location") !== undefined) document.getElementById("location").value = props.get("ARS-Location");
    if (props.get("ARS-MeetingPhone") !== undefined)
      document.getElementById("Phone").value = props.get("ARS-MeetingPhone");
    if (props.get("ARS-Enabled") !== undefined)
      document.querySelector("#enableReminder").checked = Boolean(props.get("ARS-Enabled"));
    if (props.get("ARS-Status") !== undefined)
      document.getElementById("ReminderStatus").innerHTML = props.get("ARS-Status");
    else document.getElementById("ReminderStatus").innerHTML = "Not enabled";
    item.organizer.getAsync((host) => {
      if (props.get("ARS-Host") === undefined) document.getElementById("host").value = host.value.displayName;
    });
  });
}

export async function closeItem() {
  await saveAppointmentInfo();
  //Office.context.mailbox.item.close();
}

export async function saveAppointmentInfo() {
  var item = Office.context.mailbox.item;
  var email = document.getElementById("email").value;
  var sms = document.getElementById("sms").value;
  var host = document.getElementById("host").value;
  var location = document.getElementById("location").value;
  var phone = document.getElementById("Phone").value;
  var enableReminder = document.querySelector("#enableReminder").checked;
  var autosave = document.querySelector("#autosave").checked;

  item.loadCustomPropertiesAsync((result) => {
    const props = result.value;
    props.set("ARS-Email", email);
    props.set("ARS-Phone", sms);
    props.set("ARS-Host", host);
    props.set("ARS-Location", location);
    props.set("ARS-MeetingPhone", phone);
    props.set("ARS-Enabled", enableReminder);
    props.set("ARS-Status", "PENDING");
    props.saveAsync((saveResult) => {
      console.log("SAVE_CUSTOM_PROP", saveResult);
    });
    if (autosave) {
      Office.context.mailbox.item.saveAsync().then(() => console.log("called from inside"));
    }
  });
}

export function validateEmail(e) {
  const regex = /^(([^<>()[\]\.,;:\s@\"]+(\.[^<>()[\]\.,;:\s@\"]+)*)|(\".+\"))@(([^<>()[\]\.,;:\s@\"]+\.)+[^<>()[\]\.,;:\s@\"]{2,})$/i;
  const textValue = e.srcElement.value;
  if (textValue.match(regex)) document.getElementById("emailProvided").src = "../../assets/checked.png";
  else document.getElementById("emailProvided").src = "../../assets/cancel.png";
}

export function validateSms(e) {
  const regex = /^((((0{2}?)|(\+){1})46)|0)7[\d]{8}[0-9]*$/i;
  const zeroRegex = /^0/i;
  var textValue = e.srcElement.value;
  textValue = textValue.replace(zeroRegex, "+46");
  if (textValue.match(regex)) document.getElementById("smsProvided").src = "../../assets/checked.png";
  else document.getElementById("smsProvided").src = "../../assets/cancel.png";
}