function sync() {
  syncSheet("1i7y_tFpeO68SetmsU2t-C6LsFETuZtkJGY5AVZ2PHW8", );
}

function doGet(e) {
  try {
    const sheetId = e.parameter.sheetId;
    const groups = e.parameter.groups;
    Logger.log("Web deployment called with " + sheetId + " and " + groups);
    if (sheetId && groups) {
      return syncForWeb(sheetId, groups);
    } else if (sheetId) {
      const storedGroups = getControls(sheetId).Groups;
      if (storedGroups) {
        Logger.log("Found groups stored in the sheet: " + storedGroups);
        return syncForWeb(sheetId, storedGroups);
      } else {
        Logger.log("Form accessed with sheetId as " + sheetId + ", will display groups form");
        const form = HtmlService.createTemplateFromFile('groupsForm');
        form.sheetId = sheetId;
        form.formUrl = ScriptApp.getService().getUrl();
        return form.evaluate();
      }
    } else {
      Logger.log("Form accessed without sheetId, will display sheet form");
      const form = HtmlService.createTemplateFromFile('sheetForm');
      form.formUrl = ScriptApp.getService().getUrl();
      return form.evaluate();
    }
  } catch (e) {
    Logger.log(e);
    return HtmlService.createHtmlOutput(e);
  }
}

function syncForWeb(sheetId, groups) {
  try {
    Logger.log("Web call with sheetId as " + sheetId + " and groups as " + groups + ", will run sync");
    syncSheet(sheetId, groups);
    return HtmlService.createHtmlOutput("Contacts downloaded to sheet!");
  } catch (e) {
    Logger.log(e);
    return HtmlService.createHtmlOutput(e);
  }
}

function doPost(e) {
  const sheetId = e.parameter.sheetId;
  const groups = e.parameter.groups;
  Logger.log("Got sheetId as " + sheetId + " and groups as " + groups);
  const url = ScriptApp.getService().getUrl() +
      '?sheetId=' + encodeURIComponent(sheetId) +
      '&groups=' + encodeURIComponent(groups);

  // Redirect back to the form with the parameters
  return HtmlService.createHtmlOutput('<html><head>' +
      '<meta http-equiv="refresh" content="0; url=' + url + '">' +
      '</head></html>');
}

function syncSheet(sheetId, groups) {
  const controls = getControls(sheetId, groups);
  if (!controls.Groups) {
    //If no groups specified, cannot sync any data
    throw new Error("Require at least one google group to find and download contacts to sheet");
  }
  let sheet = SpreadsheetApp.openById(sheetId).getSheetByName(controls.SheetName);
  if (!sheet) {
    Logger.log("Sheet " + controls.SheetName + " doesn't exist, creating it");
    sheet = SpreadsheetApp.openById(sheetId).insertSheet(controls.SheetName, 1);
  }
  if (sheet.getLastRow() === 0) {
    Logger.log("No header row present, creating one");
    sheet.appendRow(["Name", "Kerberos", "Email", "Employee ID", "Job Title", "Location", "Manager", "Termination"]);
    sheet.getRange(1,1,1,8).setFontWeight("bold");
    sheet.setFrozenRows(1);
    sheet.setFrozenColumns(1);
  }
  let storedUserIds;
  if (sheet.getLastRow() > 1) {
    storedUserIds = sheet.getRange(2, 2, sheet.getLastRow() - 1, 1).getValues().map(array => array[0]);
  } else {
    storedUserIds = [];
  }

  const users = getAllUsers(controls.Groups.split(","));

  //Delete associates that have left the specified groups from the sheet:
  const storedUserSet = new Set();
  for (let i = storedUserIds.length - 1; i >= 0; i--) {
    storedUserSet.add(storedUserIds[i]);
    if (!users.has(storedUserIds[i])) {
      if (controls.DeleteOnTermination === true) {
        Logger.log("User " + storedUserIds[i] + " no longer found in specified groups, deleting their entry from the spreadsheet");
        sheet.deleteRow(i + 2);
      } else {
        //If not already marked, mark them as terminated
        if (!sheet.getRange(i + 2, 8, 1, 1).getValue()) {
          Logger.log("User " + storedUserIds[i] + " no longer found in specified groups, marking them as terminated");
          sheet.getRange(i + 2, 8, 1, 1).setValue(new Date().toDateString());
        }
      }
    }
  }

  const response = getContacts(users, controls.SyncToken, new Set(controls.ExcludedTitles.split(",")));

  //Add new associates returned by the group queries that are not in the sheet:
  for (let user of users) {
    if (!storedUserSet.has(user)) {
      const contact = response.contacts.get(user);
      //Add it if it's not an excluded entry, for example based on title:
      if (contact) {
        Logger.log("User " + user + " not found in the spreadsheet, adding their entry");
        sheet.appendRow(getContactRow(contact));
      }
    }
  }

  let unchangedContacts = [];
  for (let i = 0; i < storedUserIds.length; i++) {
    const contact = response.contacts.get(storedUserIds[i]);
    if (contact) {
      const updatedRow = getContactRow(contact);
      const currentRow = sheet.getRange(i + 2, 1, 1, 7).getValues()[0];
      currentRow[3] = currentRow[3].toString();
      if (JSON.stringify(currentRow) === JSON.stringify(updatedRow)) {
        unchangedContacts.push(updatedRow[1]);
      } else {
        Logger.log("User " + storedUserIds[i] + " returned from the corporate directory, likely updated, will overwrite their entry in the spreadsheet");
        sheet.getRange(i + 2, 1, 1, 7).setValues([updatedRow]);
      }
    }
  }
  if (unchangedContacts.length > 0) {
    Logger.log("Corporate directory returned updates in the following contacts, but no relevant fields had changed: " + JSON.stringify(unchangedContacts));
  }
  if (response.nextSyncToken) {
    //Sometimes the response does not contain any info, possibly because there are no changes at all and previous sync token remains valid
    setNextSyncToken(sheetId, response.nextSyncToken);
  }
}

function getControls(sheetId, groups) {
  const controls = {};
  const sheetName = "Controls";
  let sheet;
  try {
    sheet = SpreadsheetApp.openById(sheetId).getSheetByName(sheetName);
  } catch (e) {
    Logger.log("Caught error, will throw it right back up: " + e);
    throw new Error("Unable to open the specified sheet at https://docs.google.com/spreadsheets/d/" + sheetId);
  }
  if (!sheet) {
    Logger.log("Sheet called " + sheetName + " doesn't exist, creating it");
    sheet = SpreadsheetApp.openById(sheetId).insertSheet(sheetName, 0);
  }
  if (sheet.getLastRow() === 0) {
    Logger.log("No content present, creating defaults");
    sheet.appendRow(["Groups", groups]);
    sheet.appendRow(["SheetName", "Roster"]);
    sheet.appendRow(["ExcludedTitles", ""]);
    sheet.appendRow(["DeleteOnTermination", ""]);
    sheet.appendRow(["SyncToken", ""]);
    sheet.getRange(1,1,5,1).setFontWeight("bold");
  }
  let values = sheet.getRange(1, 1, sheet.getLastRow(), 2).getValues();
  for (row of values) {
    controls[row[0]] = row[1];
  }
  if (groups) {
    //override potential groups entry in sheet with provided parameter
    controls.Groups = groups;
  }
  return controls;
}

function setNextSyncToken(sheetId, syncToken) {
  let sheet = SpreadsheetApp.openById(sheetId).getSheetByName("Controls");
  let values = sheet.getRange(1, 1, sheet.getLastRow(), 2).getValues();
  for (let i = 0; i < values.length; i++) {
    if (values[i][0] === "SyncToken") {
      Logger.log("Will set SyncToken to " + syncToken + " for next run");
      sheet.getRange(i + 1, 2, 1, 1).setValue(syncToken);
    }
  }
}

function getContactRow(contact) {
  const row = [];
  row.push(contact.name);
  row.push(contact.kerberos);
  row.push(contact.email);
  row.push(contact.employeeId);
  row.push(contact.title);
  row.push(contact.location);
  row.push(contact.manager);
  return row;
}

function getAllUsers(groups) {
  const users = new Set();
  for (let group of groups) {
    const members = getMembers(group);
    members.forEach(member => users.add(member));
  }
  return users;
}

function getMembers(group) {
  const groupObj = GroupsApp.getGroupByEmail(group);
  const users = groupObj.getUsers();
  return users.map(user => user.getUsername());
}

function getContacts(users, syncToken, excludedTitles) {
  const contacts = new Map();
  let nextPageToken;
  const result = {
    contacts: contacts
  }
  do {
    const response = People.People.listDirectoryPeople({
      readMask: 'emailAddresses,locations,names,organizations,relations,externalIds',
      sources: [
        'DIRECTORY_SOURCE_TYPE_DOMAIN_PROFILE'
      ],
      pageSize: 1000,
      pageToken: nextPageToken,
      syncToken: syncToken,
      requestSyncToken: true
    });
    if (!response.people) {
      Logger.log("Got response with no people in it");
      return result;
    }
    Logger.log("Got " + response.people.length + " contacts back");
    for (let person of response.people) {
      const emailObject = getPrimary(person.emailAddresses);
      if (!emailObject) {
        continue;
      }
      const email = emailObject.value;
      const kerberos = email.split("@")[0];
      if (users.has(kerberos)) {
        const work = getPrimary(person.organizations);
        if (excludedTitles.has(work.title)) {
          continue;
        }
        contacts.set(kerberos, {
          name: getPrimary(person.names).displayName,
          kerberos: kerberos,
          email: email,
          employeeId: getPrimary(person.externalIds).value,
          manager: getManager(person.relations),
          title: work.title,
          location: work.location
        });
      }
    }
    nextPageToken = response.nextPageToken;
    result.nextSyncToken = response.nextSyncToken;
  } while (nextPageToken);
  return result;
}

function getPrimary(repeatingObject) {
  if (!repeatingObject) {
    return undefined;
  }
  for (let object of repeatingObject) {
    if (object.metadata.primary && object.metadata.primary === true) {
      return object;
    }
  }
}

function getManager(relations) {
  for (let relation of relations) {
    if (relation.type && relation.type === "manager") {
      return relation.person.split("@")[0];
    }
  }
}
