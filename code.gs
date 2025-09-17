function getFormData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const form = ss.getSheetByName("Task Form");

  const values = form.getRange("B2:B10").getValues().flat();

  // Check required fields B1 to B7 (index 0 to 6)
  for (let i = 0; i < 7; i++) {
    if (!values[i]) {
      SpreadsheetApp.getUi().alert("⚠️ Please complete all required fields (all except Comments).");
      return null;
    }
  }

  return values;
}

function ccreateTask() {
  const data = getFormData();
  if (!data) return;

  const userEmail = Session.getActiveUser().getEmail();
  if (!userEmail) {
    SpreadsheetApp.getUi().alert("User email not found. Make sure you're logged in.");
    return;
  }

  const backend = getCurrentSprintSheet();

  // ✅ Get max Task ID from column A
  let maxTaskId = 0;
  const dataRowCount = backend.getLastRow() - 1;
  if (dataRowCount > 0) {
    const taskIdValues = backend.getRange(2, 1, dataRowCount).getValues().flat();
    taskIdValues.forEach(id => {
      const numericId = Number(id);
      if (!isNaN(numericId) && numericId > maxTaskId) {
        maxTaskId = numericId;
      }
    });
  }
  const newTaskId = maxTaskId + 1;

  // ✅ Add new row
  const rowData = [
    newTaskId,        // A → Task ID
    data[0],          // B → Task Name
    data[1],          // C → Task Type
    userEmail,        // D → Created By
    data[2],          // E → Story
    data[3],          // F → Team
    data[4],          // G → Assignee
    data[5],          // H → Status
    data[6],          // I → Committed to Deliver
    data[7],          // J - Dependent on
    "",               // K - Assignee QA
    "",               // L → Delivered At
    "",               // M → QA Status
    "",               // N → QA Committed to Done
    data[8] || ""     // O → Comments
  ];

  backend.appendRow(rowData);
  clearForm();
  const rowIdx = backend.getLastRow();


  // ✅ Protect columns for the creator:
  // A (Task ID)
  // B–D (Task Name, Task Type, Created By)
  // F–J (Story → Committed to Deliver)
  // N–O (Dependent On, Comments)
  const protectedCols = [
    { startCol: 1, numCols: 1 },  // A → Task ID
    { startCol: 2, numCols: 3 },  // B–D
    { startCol: 6, numCols: 5 },  // F–J
    { startCol: 14, numCols: 2 }  // N–O
  ];

  protectedCols.forEach(({ startCol, numCols }) => {
    const range = backend.getRange(rowIdx, startCol, 1, numCols);
    const protection = range.protect().setDescription(`Task fields locked for others`);
    protection.addEditor(userEmail);
    protection.removeEditors(protection.getEditors().filter(e => e.getEmail() !== userEmail));
    if (protection.canDomainEdit()) protection.setDomainEdit(false);
  });


  SpreadsheetApp.getUi().alert(`✅ Task #${newTaskId} created and protected!`);
  // clearForm();
  
  // SpreadsheetApp.getActiveSpreadsheet()
  // .getSheetByName("Flags")
  // .getRange("A1")
  // .setValue(new Date());  // Any changing value
}


function updateTask() {
  const data = getFormData();
  if (!data) return;

  const userEmail = Session.getActiveUser().getEmail();
  if (!userEmail) {
    SpreadsheetApp.getUi().alert("User email not found. Make sure you're logged in.");
    return;
  }

  const [taskName] = data;
  const backend = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sprint 44");
  const rows = backend.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    const existingTaskName = row[0];
    const createdBy = row[1];

    if (existingTaskName === taskName) {
      if (createdBy !== userEmail) {
        SpreadsheetApp.getUi().alert("❌ You are not allowed to update this task. It was created by another user.");
        return;
      }

      // Replace only data columns (starting from column 3 onward)
      const updatedRow = [taskName, userEmail, ...data.slice(1)];
      backend.getRange(i + 1, 1, 1, updatedRow.length).setValues([updatedRow]);
      SpreadsheetApp.getUi().alert("✅ Task updated!");
      clearForm();
      return;
    }
  }

  SpreadsheetApp.getUi().alert("Task not found.");
}


function deleteTask() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const form = ss.getSheetByName("Task Form");
  const taskName = form.getRange("B1").getValue();
  if (!taskName) {
    SpreadsheetApp.getUi().alert("Please enter Task Name to delete.");
    return;
  }
  const backend = ss.getSheetByName("Sprint 44");
  const rows = backend.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === taskName) {
      backend.deleteRow(i + 1);
      SpreadsheetApp.getUi().alert("✅ Task deleted.");
      clearForm();
      return;
    }
  }
  SpreadsheetApp.getUi().alert("Task not found to delete.");
}


function clearForm() {
  const form = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Task Form");
  form.getRange("B2:B10").clearContent();
}

function getQAFormData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const form = ss.getSheetByName("QA Form");
  const values = form.getRange("B2:B5").getValues().flat();

  // Only Task ID (B2) and QA Status (B5) are required
  if (!values[0] || !values[3]) {
    SpreadsheetApp.getUi().alert("⚠️ Please provide Task ID and QA Status.");
    return null;
  }

  return values; // deliveredAt and qaCommitted may be empty
}

function submitQA() {
  const data = getQAFormData();
  if (!data) return;

  const [taskId, deliveredAt, qaCommitted, qaStatus] = data;
  const backend = getCurrentSprintSheet();
  const userEmail = Session.getActiveUser().getEmail();

  // 🔍 Find task row
  const taskRange = backend.getRange(2, 1, backend.getLastRow() - 1, 1).getValues();
  let rowIndex = -1;
  for (let i = 0; i < taskRange.length; i++) {
    if (String(taskRange[i][0]) === String(taskId)) {
      rowIndex = i + 2; // +2 = offset for header
      break;
    }
  }
  if (rowIndex === -1) {
    SpreadsheetApp.getUi().alert("❌ Task ID not found.");
    return;
  }

  // 📝 Write QA values
  backend.getRange(rowIndex, 11).setValue(userEmail);   // K: Assignee QA
  backend.getRange(rowIndex, 12).setValue(deliveredAt); // L: Delivered At
  backend.getRange(rowIndex, 13).setValue(qaStatus); // M: QA Status
  backend.getRange(rowIndex, 14).setValue(qaCommitted);    // N: QA commited to done

  // 🔐 Protect QA-related fields (K, L, M, N)
  const qaCols = [11, 12, 13, 14];
  qaCols.forEach(col => {
    const range = backend.getRange(rowIndex, col);
    const protection = range.protect().setDescription(`Protected QA field by ${userEmail}`);
    protection.addEditor(userEmail);
    protection.removeEditors(protection.getEditors().filter(e => e.getEmail() !== userEmail));
    if (protection.canDomainEdit()) protection.setDomainEdit(false);
  });

  SpreadsheetApp.getUi().alert(`✅ QA info submitted for Task #${taskId}`);
  clearQAForm();
}


function clearQAForm() {
  const form = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("QA Form");
  form.getRange("B2:B5").clearContent();
}


function protectExistingRowsByOwner() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sprint 44");
  const data = sheet.getDataRange().getValues();

  const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  protections.forEach(p => p.remove()); // remove old protections

  for (let i = 1; i < data.length; i++) { // skip header
    const userEmail = data[i][1]; // "Created By" column
    if (userEmail) {
      const range = sheet.getRange(i + 1, 1, 1, sheet.getLastColumn());
      const protection = range.protect().setDescription(`Protected task row for ${userEmail}`);
      protection.addEditor(userEmail);
      protection.removeEditors(protection.getEditors().filter(e => e.getEmail() !== userEmail));
      if (protection.canDomainEdit()) protection.setDomainEdit(false);
    }
  }
}

function createSprint() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const formSheet = ss.getSheetByName("Create Sprint");

  const sprintNumber = formSheet.getRange("B2").getValue().toString().trim();
  const startDate = formSheet.getRange("B3").getValue();
  const endDate = formSheet.getRange("B4").getValue();

  if (!sprintNumber || !startDate || !endDate) {
    SpreadsheetApp.getUi().alert("⚠️ Please fill all fields: Sprint ID, Start Date, End Date.");
    return;
  }

  // ✅ Format dates as dd-MM-yy
  const tz = ss.getSpreadsheetTimeZone();
  const formattedStart = Utilities.formatDate(new Date(startDate), tz, "dd-MM-yy");
  const formattedEnd = Utilities.formatDate(new Date(endDate), tz, "dd-MM-yy");

  // ✅ Create sprint name with dates
  const sprintId = `Sprint ${sprintNumber} [${formattedStart} to ${formattedEnd}]`;

  if (ss.getSheetByName(sprintId)) {
    SpreadsheetApp.getUi().alert(`❌ Sheet "${sprintId}" already exists.`);
    return;
  }

  const baseSheet = ss.getSheetByName("Base Sheet");
  const qaFormSheet = ss.getSheetByName("QA Form");

  if (!baseSheet || !qaFormSheet) {
    SpreadsheetApp.getUi().alert("❌ 'Base Sheet' or 'QA Form' sheet not found.");
    return;
  }

  // Copy Base Sheet to new sprint sheet
  const copiedSheet = baseSheet.copyTo(ss);
  copiedSheet.setName(sprintId);

  // Move new sheet after QA Form
  const sheets = ss.getSheets();
  const qaFormIndex = sheets.findIndex(sheet => sheet.getName() === "QA Form");
  if (qaFormIndex !== -1) {
    ss.setActiveSheet(copiedSheet);
    ss.moveActiveSheet(qaFormIndex + 2);
  }

  // In case of temporary "_copy" name issue
  if (copiedSheet.getName() !== sprintId) {
    try {
      copiedSheet.setName(sprintId);
    } catch (e) {
      SpreadsheetApp.getUi().alert(`❌ Unable to rename sheet to "${sprintId}". Try again with a unique name.`);
      return;
    }
  }

  // Log the sprint in "Sprint Info"
  const infoSheet = ss.getSheetByName("Sprint Info");
  if (!infoSheet) {
    SpreadsheetApp.getUi().alert("❌ 'Sprint Info' sheet not found.");
    return;
  }

  // Append to log
  const lastRow = infoSheet.getLastRow();
  infoSheet.getRange(lastRow + 1, 1, 1, 3).setValues([[sprintNumber, startDate, endDate]]);
  infoSheet.getRange(2, 1, infoSheet.getLastRow() - 1, 3).sort({ column: 1, ascending: false });

  // Protect Sprint Info Sheet
  let protection = infoSheet.getProtections(SpreadsheetApp.ProtectionType.SHEET)[0];
  if (!protection) {
    protection = infoSheet.protect();
    protection.setDescription('Protect Sprint Info');
  }

  protection.setWarningOnly(false);

  // Copy editors from Create Sprint protection
  const formProtections = formSheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
  if (formProtections.length > 0) {
    const formProtection = formProtections[0];
    const authorizedEditors = formProtection.getEditors();

    // Ensure script owner is included
    const scriptOwner = Session.getEffectiveUser();
    const emails = authorizedEditors.map(e => e.getEmail());
    if (!emails.includes(scriptOwner.getEmail())) {
      emails.push(scriptOwner.getEmail());
    }

    // Apply editors to new protection
    protection.addEditors(emails);
  }

  // Store current sprint ID in script properties
  PropertiesService.getScriptProperties().setProperty("currentSprintSheet", sprintId);

  // Update references in forms
  updateSprintNameInForms(sprintId);

  clearSprintForm();

  SpreadsheetApp.getUi().alert(`✅ Sprint "${sprintId}" created successfully.`);
}



function clearSprintForm() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Create Sprint");
  sheet.getRange("B2:B4").clearContent(); // Clears Sprint ID, Start Date, End Date
}


function updateSprintNameInForms(sprintName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // ✅ Define authorized editors here
  const authorizedEditors = [
    "mdabunadif@gmail.com",
    "tasfiq.kamran@gmail.com",
    Session.getEffectiveUser().getEmail() // Optional: include current user
  ];

  // Utility function to protect and assign editors
  function protectRange(range, description) {
    const protection = range.protect().setDescription(description);
    protection.removeEditors(protection.getEditors()); // Remove previous

    // Assign only authorized editors
    authorizedEditors.forEach(email => protection.addEditor(email));

    if (protection.canDomainEdit()) protection.setDomainEdit(false);
  }

  // Clear previous form data
  clearForm();
  clearQAForm();

  // 📝 Update Task Form
  const taskForm = ss.getSheetByName("Task Form");
  if (taskForm) {
    taskForm.showColumns(4, 5);  // D to H
    taskForm.showRows(6, 1);     // Row 6

    const taskRange = taskForm.getRange("D6:H6");
    taskRange.clearContent();
    taskRange.merge();
    taskRange.setValue(sprintName)
             .setFontWeight("bold")
             .setFontSize(14)
             .setHorizontalAlignment("center")
             .setVerticalAlignment("middle");
    protectRange(taskRange, "Protect Sprint Name Task Form");
  }

  // 📝 Update QA Form
  const qaForm = ss.getSheetByName("QA Form");
  if (qaForm) {
    qaForm.showColumns(4, 5);
    qaForm.showRows(3, 1);

    const qaRange = qaForm.getRange("D3:H3");
    qaRange.clearContent();
    qaRange.merge();
    qaRange.setValue(sprintName)
           .setFontWeight("bold")
           .setFontSize(14)
           .setHorizontalAlignment("center")
           .setVerticalAlignment("middle");
    protectRange(qaRange, "Protect Sprint Name QA Form");
  }

  // 📝 Update Sprint Creation Form
  const sprintForm = ss.getSheetByName("Create Sprint");
  if (sprintForm) {
    sprintForm.showColumns(4, 5);
    sprintForm.showRows(3, 1);

    const sprintRange = sprintForm.getRange("D3:H3");
    sprintRange.clearContent();
    sprintRange.merge();
    sprintRange.setValue(sprintName)
               .setFontWeight("bold")
               .setFontSize(14)
               .setHorizontalAlignment("center")
               .setVerticalAlignment("middle");
    protectRange(sprintRange, "Protect Sprint Name Sprint Form");
  }
}


function sanitizeRangeName(name) {
  return name.replace(/[^A-Za-z0-9_]/g, "_");
}


function sanitizeTableName(name) {
  return name.replace(/[^A-Za-z0-9_]/g, "_");
}


function getCurrentSprintSheet() {
  const sheetName = PropertiesService.getScriptProperties().getProperty("currentSprintSheet");
  if (!sheetName) {
    throw new Error("No current sprint sheet set.");
  }
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) {
    throw new Error(`Sheet "${sheetName}" not found.`);
  }
  return sheet;
}




