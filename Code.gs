function getSheetId() {
  return "10IVTc2hmZroOQ4xcXJ4cbR0gDXNIzpuoG2qXeYK-PHM";
}

function getControls() {
  const controls = {};
  let sheet = SpreadsheetApp.openById(getSheetId()).getSheetByName("Controls");
  let values = sheet.getRange(1, 1, sheet.getLastRow(), 2).getValues();
  for (row of values) {
    controls[row[0]] = row[1];
  }
  return controls;
}

function setNextSyncToken(syncToken) {
  let sheet = SpreadsheetApp.openById(getSheetId()).getSheetByName("Controls");
  let values = sheet.getRange(1, 1, sheet.getLastRow(), 2).getValues();
  for (let i = 0; i < values.length; i++) {
    if (values[i][0] === "SyncToken") {
      Logger.log("Will set SyncToken to " + syncToken + " for next run");
      sheet.getRange(i + 1, 2, 1, 1).setValue(syncToken);
    }
  }
}

function sync() {
  const controls = getControls();
  let sheet = SpreadsheetApp.openById(getSheetId()).getSheetByName(controls.SheetName);
  if (!sheet) {
    Logger.log("Sheet " + controls.SheetName + " doesn't exist, creating it");
    sheet = SpreadsheetApp.openById(getSheetId()).insertSheet(controls.SheetName, 1);
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

  for (let i = 0; i < storedUserIds.length; i++) {
    const contact = response.contacts.get(storedUserIds[i]);
    if (contact) {
      Logger.log("User " + storedUserIds[i] + " returned from the corporate directory, likely updated, will overwrite their entry in the spreadsheet");
      sheet.getRange(i + 1, 1, 1, 8).setValues([getContactRow(contact)]);
    }
  }
  setNextSyncToken(response.nextSyncToken);
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
  row.push(""); //no termination date
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

  let nextSyncToken;
  let nextPageToken;
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
    nextSyncToken = response.nextSyncToken;
  } while (nextPageToken);
  return {
    nextSyncToken: nextSyncToken,
    contacts: contacts
  };
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
