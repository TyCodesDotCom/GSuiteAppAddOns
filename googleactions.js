// *************************************************************
// ** Re-Enable & Disable Chromebook Form Functionality
// ** G-Suit Google App Script Application
// **
// ** Author: Ty Christian
// ** Date: 5/13/2019
// **
// ************************************************************


function onEdit(e) {

    // Gather Initial Information
    var customerId = "*******";
    var SN = getSN();
    var action = getAction();

    // Test to verify device exists
    var testing = isADevice(customerId, SN);

    // If exists, continue on, if not, send email of error
    if (testing) {
        var chromebook = getDeviceBySN(customerId, SN);
        var id = getDeviceId(chromebook);
        var user = getLastKnownUser(chromebook);
        runAction(action, customerId, id, chromebook);
        Logger.log("It Worked");
    } else {
        customEmail("serialNumberError");
    }
}

// Get Email of Form Submitter
function getSubmitter() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[0];
    var range = sheet.getRange("G2:J2");
    var name = range.getCell(1, 1).getValue();
    return name;
}

// Gather SN from Spreadsheet
function getSN() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[0];
    var range = sheet.getRange("G2:J2");
    var SN = range.getCell(1, 2).getValue();
    var new_SN = SN.trim();

    return new_SN;
}

// Gather action from spreadsheet
function getAction() {
    var tt = SpreadsheetApp.getActiveSpreadsheet();
    var shet = tt.getSheets()[0];
    var ranges = shet.getRange("G2:J2");
    var action = ranges.getCell(1, 3).getValue();
    return action.toLowerCase();
}

// Pull Device info using SN from API
function getDeviceBySN(customerId, sn) {
    var optionalArgs = {
        customer: customerId,
        projection: 'FULL',
        query: sn
    };

    var device = AdminDirectory.Chromeosdevices.list(customerId, optionalArgs);
    var chromebook = JSON.parse(device.chromeosdevices);

    return chromebook;
}

// Get Device ID, different then SN
function getDeviceId(chromebook) {

    var id = chromebook["deviceId"]

    return id;
}

// Get Status of Chromebook from Google Device Manager
function getStatus(chromebook) {

    var status = chromebook["status"]

    return status;
}

// Get Last Known user of Chromebook
function getLastKnownUser(chromebook) {

    var users = chromebook["recentUsers"]
    var user = users[0];
    var email = user["email"]

    return email;
}

// Check if SN has a device
function isADevice(customerId, sn) {
    var optionalArgs = {
        customer: customerId,
        projection: 'FULL',
        query: sn
    };

    var returnValue = true;

    try {
        var device = AdminDirectory.Chromeosdevices.list(customerId, optionalArgs);
        var chromebook = JSON.parse(device.chromeosdevices);

    } catch (e) {
        returnValue = false;
    }

    return returnValue;
}


function runAction(action, customerId, deviceId, chromebook) {

    var resource = {
        "action": action
    };
    var alert = "";
    var status = getStatus(chromebook);

    if (action === "disable") {
        if (status === "ACTIVE") {
            AdminDirectory.Chromeosdevices.action(resource, customerId, deviceId);
            Logger.log("Device has been Disabled!")
            alert = "disableSuccess";
            action = "done";

        } else {
            alert = "disableError";
            action = "done";
        }
    }
    if (action === "reenable") {
        if (status === "DISABLED") {
            AdminDirectory.Chromeosdevices.action(resource, customerId, deviceId);
            Logger.log("Device has been Enabled!")
            alert = "enableSuccess";
            action = "done";
        } else {
            alert = "enableError";
            action = "done";
        }
    }
    customEmail(alert, chromebook);
}

function customEmail(alert, chromebook) {

    chromebook = chromebook || 0;

    // init msg and subject variables
    var msg = "";
    var subject = "";

    // Get Submitter Email
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[0];
    var range = sheet.getRange("G2:J2");
    var name = range.getCell(1, 1).getValue();

    // Get last known user, and SN
    var sn = getSN();

    //Check Case
    switch (alert) {
        case "enableError":
            subject = "Alert! Device Already Enabled"
            msg = "Hello, this device (" + sn + ") has already been enabled. Please make sure you selected the correct field on the Chromebook Re-Enable/Disable Form. If you continue to recieve this error please notify your EduTek Department. This is a automated message, please do not reply.";
            break;
        case "disableError":
            subject = "Alert! Device Already Disabled"
            msg = "Hello, this device (" + sn + ") has already been disabled. Please make sure you selected the correct field on the Chromebook Re-Enable/Disable Form. If you continue to recieve this error please notify your EduTek Department. This is a automated message, please do not reply.";
            break;
        case "serialNumberError":
            subject = "Alert! Serial Number Does Not Exist."
            msg = "Hello, the Serial Number (" + sn + ") entered does not exist in our system. Please make sure you entered in the correct Serial Number on the Chromebook Re-Enable/Disable Form.  If you continue to recieve this error please notify your EduTek Department. This is a automated message, please do not reply.";
            break;
        case "enableSuccess":
            subject = "Success! Device has been Enabled";
            msg = "Hello, this device (" + sn + ") has been enabled. This is a automated message, please do not reply.";
            break;
        case "disableSuccess":
            var student = getLastKnownUser(chromebook);
            subject = "Success! Device has been Disabled"
            msg = "Hello, this device (" + sn + ") has been disabled. The last student to use this device was: " + student + ". This is a automated message, please do not reply.";
            break;
        default:
            subject = "Alert! Something Went Wrong!";
            msg = "Please notify your Edutek Department of this error. (" + sn + ")";
    }

    MailApp.sendEmail(name, subject, msg);

}