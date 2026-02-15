/** -----------------------
 *  HELPERS
 *  -----------------------
 */
// Extract ID from full Google Form/Sheet URL or accept plain ID
function extractId(urlOrId) {
  if (!urlOrId) return "";
  var match = urlOrId.match(/[-\w]{25,}/);
  return match ? match[0] : urlOrId.trim();
}

// Shuffle helper
function shuffleArray(array) {
  for (var i = array.length - 1; i > 0; i--) {
    var j = Math.floor(Math.random() * (i + 1));
    var tmp = array[i]; array[i] = array[j]; array[j] = tmp;
  }
}

function requireValidId(id, label) {
  if (!id || !/[-\w]{25,}/.test(id)) {
    throw new Error("❌ Invalid or missing " + label + ". Please paste a full Google Sheet URL or ID.");
  }
}


/** -----------------------
 *  UI
 *  -----------------------
 */


function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("📋 Quiz Manager")
    .addItem("Open Dashboard", "showSidebar")
    .addSeparator()
    .addItem("Classify PSM Topics", "classifyQuestionsPSM_FINAL")
    .addToUi();
}










function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile("Sidebar")
    .setTitle("📋 Quiz Manager");
  SpreadsheetApp.getUi().showSidebar(html);
}

/** -----------------------
 *  CONFIG SAVE/LOAD
 *  -----------------------
 */
function saveConfig(data) {
  var props = PropertiesService.getUserProperties();
  props.setProperty("formId", extractId(data.formId));
  props.setProperty("sheetId", extractId(data.sheetId));
  props.setProperty("numQ", data.numQ);
  props.setProperty("formTitles", data.formTitles || "");
  logHistory("Config Saved", "Form & Sheet IDs saved.");
  return "💾 Configuration saved!";
}

function loadConfig() {
  var props = PropertiesService.getUserProperties();
  return {
    formId: props.getProperty("formId") || "",
    sheetId: props.getProperty("sheetId") || "",
    numQ: props.getProperty("numQ") || "10",
    formTitles: props.getProperty("formTitles") || ""
  };
}

function clearConfiguration() {
  var props = PropertiesService.getUserProperties();
  props.deleteAllProperties();
  return "🗑️ Configuration cleared!";
}

/** -----------------------
 *  MAIN FUNCTIONS
 *  -----------------------
 */
function runCreateQuiz(inputs) {
  var sheetId = extractId(inputs.sheetId);
  requireValidId(sheetId, "Sheet ID");

  var numQ = parseInt(inputs.numQ, 10) || 10;
  var filterTitles = (inputs.formTitles || "")
    .split(",")
    .map(t => t.trim())
    .filter(t => t);

  var ss = SpreadsheetApp.openById(sheetId);
  var sheet = ss.getSheetByName("Questions");
  var data = sheet.getDataRange().getValues();
  data.shift();

  // Filter by formTitle column if specified
  if (filterTitles.length > 0) {
    data = data.filter(r => filterTitles.indexOf(r[6]) !== -1);
  }

  data = data.filter(r => r[0]);
  shuffleArray(data);

  var form = FormApp.create("Quiz - " + new Date().toLocaleDateString());
  form.setIsQuiz(true);

  // Participant info
  form.addTextItem().setTitle(inputs.participantName || "Full Name").setRequired(true);
  form.addTextItem().setTitle(inputs.participantBatch || "Batch Number").setRequired(true);
  form.addTextItem().setTitle(inputs.participantRoll || "Roll Number").setRequired(true);

  form.addSectionHeaderItem().setTitle("Quiz Questions");

  var validQuestions = 0;
  var currentIndex = 0;

  while (validQuestions < numQ && currentIndex < data.length) {
    var row = data[currentIndex];
    currentIndex++;

    var q = row[0];
    var options = row.slice(1, 5).filter(o => o);
    var correctRaw = (row[5] || "").toString().trim();

    if (!q || !correctRaw) continue;

    // ✅ FIX: detect if correctRaw is a full option (even with commas inside)
    var correctAnswers;
    if (options.indexOf(correctRaw) !== -1) {
      correctAnswers = [correctRaw];
    } else {
      correctAnswers = correctRaw.split(",").map(c => c.trim()).filter(c => c);
    }

    // Remove duplicates
    var uniqueOptions = [];
    var seenOptions = {};
    options.forEach(function (opt) {
      if (opt && !seenOptions[opt]) {
        seenOptions[opt] = true;
        uniqueOptions.push(opt);
      }
    });

    if (uniqueOptions.length < 2) continue;

    // Create a Checkbox item if multiple correct answers, else MultipleChoice
    var item;
    if (correctAnswers.length > 1) {
      item = form.addCheckboxItem();
    } else {
      item = form.addMultipleChoiceItem();
    }

    var choices = uniqueOptions.map(function (o) {
      var isCorrect = correctAnswers.indexOf(o) !== -1;
      return item.createChoice(o, isCorrect);
    });

    item.setTitle(q).setChoices(choices).setRequired(true).setPoints(1);
    validQuestions++;
  }

  logHistory("Create Quiz", "Quiz created with " + validQuestions + " questions", form.getEditUrl());
  return {
    message: "✅ Quiz created with " + validQuestions + " questions!",
    url: form.getEditUrl(),
    expected: numQ,
    actual: validQuestions
  };
}


function runExportSheet(inputs) {
  var formId = extractId(inputs.formId);
  requireValidId(formId, "Form ID");
  var sheetId = extractId(inputs.sheetId);
  requireValidId(sheetId, "Sheet ID");

  var form = FormApp.openById(formId);

  // Prefer the form title; if empty, use Drive file name
  var formTitle = String(form.getTitle() || "").trim();
  if (!formTitle) {
    try {
      var file = DriveApp.getFileById(formId);
      formTitle = String(file.getName() || "").trim();
    } catch (e) {
      formTitle = "";
    }
  }

  var ss = SpreadsheetApp.openById(sheetId);
  var sheet = ss.getSheetByName("Questions") || ss.insertSheet("Questions");

  // Ensure header includes Form Title column
  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      "Question",
      "Option 1",
      "Option 2",
      "Option 3",
      "Option 4",
      "Correct",
      "Form Title"
    ]);
  } else {
    var hdrRange = sheet.getRange(1, 1, 1, Math.max(1, sheet.getLastColumn()));
    var hdr = hdrRange.getValues()[0];
    if (hdr.length < 7) {
      var newHdr = [
        hdr[0] || "Question",
        hdr[1] || "Option 1",
        hdr[2] || "Option 2",
        hdr[3] || "Option 3",
        hdr[4] || "Option 4",
        hdr[5] || "Correct",
        hdr[6] || "Form Title"
      ];
      sheet.getRange(1, 1, 1, 7).setValues([newHdr]);
    }
  }

  // Gather existing questions
  var existingQuestions = {};
  var dataRange = sheet.getDataRange();
  if (dataRange.getNumRows() > 1) {
    var data = dataRange.getValues();
    for (var i = 1; i < data.length; i++) {
      var question = data[i][0];
      if (question) {
        existingQuestions[String(question).trim()] = true;
      }
    }
  }

  var items = form.getItems();
  var addedCount = 0;
  var skippedCount = 0;

  items.forEach(function (it) {
    var itType = it.getType();
    if (itType === FormApp.ItemType.MULTIPLE_CHOICE ||
      itType === FormApp.ItemType.LIST ||
      itType === FormApp.ItemType.CHECKBOX) {

      var questionText = String(it.getTitle() || "").replace(/^\s*\d+\s*\.?\s*/, "").trim();
      if (!questionText) return;
      if (existingQuestions[questionText]) {
        skippedCount++;
        return;
      }

      var row = [questionText];
      var correct = [];

      var choices = [];
      try {
        if (itType === FormApp.ItemType.MULTIPLE_CHOICE) choices = it.asMultipleChoiceItem().getChoices();
        else if (itType === FormApp.ItemType.LIST) choices = it.asListItem().getChoices();
        else if (itType === FormApp.ItemType.CHECKBOX) choices = it.asCheckboxItem().getChoices();
      } catch (e) {
        choices = [];
      }

      for (var k = 0; k < choices.length && k < 4; k++) {
        var choiceVal = String(choices[k].getValue() || "");
        row.push(choiceVal);
        try {
          if (typeof choices[k].isCorrectAnswer === "function" && choices[k].isCorrectAnswer()) {
            correct.push(choiceVal);
          }
        } catch (e) { }
      }

      while (row.length < 5) row.push("");
      row.push(correct.join(","));
      row.push(formTitle);

      sheet.appendRow(row);
      existingQuestions[questionText] = true;
      addedCount++;
    }
  });

  logHistory("Sync Quiz",
    "Exported " + addedCount + " questions to sheet, " + skippedCount + " duplicates skipped",
    ss.getUrl());

  return {
    message: "✅ Quiz synced to sheet! " + addedCount + " questions added, " + skippedCount + " duplicates skipped.",
    url: ss.getUrl(),
    added: addedCount,
    skipped: skippedCount,
    formTitle: formTitle
  };
}


function runExportWord(inputs) {
  var sheetId = extractId(inputs.sheetId);
  requireValidId(sheetId, "Sheet ID");

  var numQ = parseInt(inputs.numQ, 10) || 10;
  var filterTitles = (inputs.formTitles || "")
    .split(",")
    .map(t => t.trim())
    .filter(t => t);

  var ss = SpreadsheetApp.openById(sheetId);
  var sheet = ss.getSheetByName("Questions");
  var data = sheet.getDataRange().getValues();
  data.shift();

  if (filterTitles.length > 0) {
    data = data.filter(r => filterTitles.indexOf(r[6]) !== -1);
  }

  shuffleArray(data);
  var selected = data.slice(0, numQ);

  var doc = DocumentApp.create("Random Quiz - " + new Date().toLocaleDateString());
  var body = doc.getBody();
  var ans = [];

  selected.forEach(function (row, i) {
    body.appendParagraph((i + 1) + ". " + row[0]);
    var opts = row.slice(1, 5).filter(o => o);

    var correctRaw = (row[5] || "").toString().trim();
    var correctAnswers;

    // ✅ FIX: detect if correctRaw is a single option (even with commas)
    if (opts.indexOf(correctRaw) !== -1) {
      correctAnswers = [correctRaw];
    } else {
      correctAnswers = correctRaw.split(",").map(c => c.trim()).filter(c => c);
    }

    // Map answers to letters
    var correctParts = [];
    correctAnswers.forEach(function (ansText) {
      var idx = opts.findIndex(opt => opt === ansText);
      var letter = idx >= 0 ? String.fromCharCode(97 + idx).toUpperCase() : "?";
      correctParts.push(letter + ") " + ansText);
    });

    opts.forEach(function (opt, j) {
      body.appendParagraph("   " + String.fromCharCode(97 + j) + ") " + opt);
    });

    ans.push((i + 1) + " → " + correctParts.join(", "));
    body.appendParagraph("");
  });

  body.appendParagraph("Answer Key").setHeading(DocumentApp.ParagraphHeading.HEADING2);
  ans.forEach(a => body.appendParagraph(a));
  doc.saveAndClose();

  logHistory("Export Word", "Exported " + selected.length + " Qs", doc.getUrl());
  return { message: "✅ Exported to Word!", url: doc.getUrl() };
}
/***********************
  Analytics & Visit Logging additions
  - add these functions into your existing Code.gs
  - does not remove or modify your other functions (they still work)
************************/

/**
 * Return email of current active user if available.
 * If not available (anonymous), returns 'anonymous'.
 */
function getCurrentUserEmail() {
  try {
    var email = Session.getActiveUser().getEmail();
    if (email && email !== '') return email;
  } catch (e) {
    // ignore
  }
  try {
    var eff = Session.getEffectiveUser().getEmail();
    if (eff && eff !== '') return eff;
  } catch (e) { }
  return 'anonymous';
}

/**
 * Enhanced logHistory: keeps original behavior but appends user info to detail.
 * If you want the user in its own column later, we can change sheet layout.
 */
function logHistory(action, detail, link) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("History") || ss.insertSheet("History");
  // Prepend row with timestamp, action, detail (with user), link
  var user = getCurrentUserEmail();
  var detailWithUser = (detail || "") + " | user: " + user;
  sheet.insertRowBefore(1);
  sheet.getRange(1, 1, 1, 4).setValues([[new Date(), action, detailWithUser, link || ""]]);
}

/**
 * Log a visit from the web app.
 * The client should call google.script.run.logVisit({path:..., ua:...});
 */
function logVisit(data) {
  try {
    var ua = data && data.ua ? String(data.ua) : "";
    var path = data && data.path ? String(data.path) : "";
    var detail = "Opened Web App";
    if (path) detail += " | path: " + path;
    if (ua) detail += " | ua: " + ua.substring(0, 200);
    logHistory("Visit", detail, "");
    return { ok: true };
  } catch (e) {
    return { ok: false, error: e.toString() };
  }
}

/**
 * Compute simple analytics from History sheet
 * returns: { visits, createQuiz, sync, exportWord, uniqueUsers: [...], last: [rows...] }
 */
function getAnalytics() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("History");
  if (!sheet) return {
    visits: 0,
    createQuiz: 0,
    sync: 0,
    exportWord: 0,
    uniqueUsers: [],
    last: []
  };

  var data = sheet.getDataRange().getValues(); // includes header-like rows
  // count events
  var visits = 0, createQuiz = 0, sync = 0, exportWord = 0;
  var usersSet = {};
  var last = [];

  // iterate rows (newest first since you insert at top)
  for (var r = 0; r < data.length; r++) {
    var row = data[r];
    // Expect row[0] = timestamp, row[1]=action, row[2]=detailWithUser, row[3]=link
    if (!row || row.length < 2) continue;
    var action = row[1] ? String(row[1]).toLowerCase() : "";
    var detail = row[2] ? String(row[2]) : "";

    if (action.indexOf("visit") !== -1) visits++;
    if (action.indexOf("create") !== -1) createQuiz++;
    if (action.indexOf("sync") !== -1 || action.indexOf("export") !== -1 && detail.indexOf("sheet") !== -1) sync++;
    if (action.indexOf("export word") !== -1 || action.indexOf("export") !== -1 && detail.indexOf("Word") !== -1) exportWord++;

    // extract user from detail (we appended " | user: EMAIL")
    var userMatch = detail.match(/\| user:\s*(\S+)/);
    if (userMatch && userMatch[1]) usersSet[userMatch[1]] = true;

    // collect last 10 rows (top-down)
    if (last.length < 10) {
      // format each row for client: [timestamp, action, detail, link]
      last.push([
        row[0] ? row[0].toString() : "",
        row[1] ? row[1].toString() : "",
        detail,
        row[3] ? row[3].toString() : ""
      ]);
    }
  }

  var uniqueUsers = Object.keys(usersSet);
  return {
    visits: visits,
    createQuiz: createQuiz,
    sync: sync,
    exportWord: exportWord,
    uniqueUsers: uniqueUsers,
    last: last
  };
}
