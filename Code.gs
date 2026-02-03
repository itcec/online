// Apps Script: exam webapp (FIXED - numeric ID-first, robust question loader + grading)
// Paste this whole file into your Apps Script project and deploy as a Web App.
// Expected sheet layout (columns):
// A: numeric id (1,2,3...) OR code like Q001 (optional, numeric recommended)
// B: question text
// C: answer (A/B/True/False etc.)
// D: optional per-question timer seconds (number)

function doGet(e) {
  return doPost(e);
}

function doPost(e) {
  const action = e.parameter.action;
  const sheetCode = e.parameter.code;
  const lastName = e.parameter.lastName;
  const firstName = e.parameter.firstName;
  const submittedAnswers = e.parameter.submittedAnswers;
  const startTime = e.parameter.startTime;
  const endTime = e.parameter.endTime;
  const date = e.parameter.date;

  // 1) duplicate check for login flow
  if (action === 'checkDuplicate') {
    if (!lastName || !firstName || !sheetCode) {
      return ContentService.createTextOutput(JSON.stringify({exists: false, error: "Missing lastName, firstName, or code"}))
        .setMimeType(ContentService.MimeType.JSON);
    }
    return checkDuplicate(sheetCode, lastName, firstName);
  }

  // 2) legacy final grade writer (kept for compatibility with older clients)
  if (action === 'recordGrade') {
    if (!lastName || !firstName || !sheetCode || !submittedAnswers || !startTime || !endTime || !date) {
      return ContentService.createTextOutput(JSON.stringify({success: false, error: "Missing required parameters"}))
        .setMimeType(ContentService.MimeType.JSON);
    }
    return recordGrade(sheetCode, lastName, firstName, submittedAnswers, startTime, endTime, date);
  }

  // 3) partial/aborted submission writer
  if (action === 'recordPartial') {
    // submittedAnswers may be empty; status and timestamp optional
    const status = e.parameter.status || 'partial';
    const timestamp = e.parameter.timestamp || null;
    return recordPartial(sheetCode, lastName, firstName, submittedAnswers || '{}', status, timestamp);
  }

  // 4) client-side batched loader
  if (action === 'getAllQuestionsAndAnswers' && sheetCode) {
    return getAllQuestionsAndAnswers(sheetCode);
  }

  // 5) NEW: final-results writer strictly for columns F..M
  // F..M order required: lastName, firstName, score, correct, mistakes, startTime, endTime, date
  if (action === 'recordResultsFM') {
    const score = e.parameter.score || '';
    const correct = e.parameter.correct || '';
    const mistakes = e.parameter.mistakes || '';
    return recordResultsFM(sheetCode, lastName, firstName, score, correct, mistakes, startTime || '', endTime || '', date || '');
  }

  return ContentService.createTextOutput(JSON.stringify({error: "Invalid action"}))
    .setMimeType(ContentService.MimeType.JSON);
}

/* Helpers (include in the same project) */
// Normalize question ids to QNNN (Q001, Q014, etc.)
function normalizeToQ(raw) {
  var s = String(raw || '').trim();
  if (!s) return '';
  var u = s.toUpperCase();
  var mQ = u.match(/^Q0*(\d+)$/i);
  if (mQ) return 'Q' + ('000' + mQ[1]).slice(-3);
  var mN = u.match(/^0*(\d+)$/);
  if (mN) return 'Q' + ('000' + mN[1]).slice(-3);
  return 'Q' + Utilities.getUuid().slice(0,6).toUpperCase();
}

/**
 * recordPartial: write an aborted/partial submission to 'Submissions' at row 2.
 * Keeps older behavior for compatibility (stores JSON/text in SubmittedAnswers column).
 */
function recordPartial(sheetCode, lastName, firstName, submittedAnswersJson, status, timestamp) {
  // Use LockService to prevent concurrent write conflicts
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (e) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: 'Lock timeout' }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var testSheet = ss.getSheetByName(sheetCode);
    // Still create Submissions if not existing
    var subName = 'Submissions';
    var sub = ss.getSheetByName(subName);
    if (!sub) {
      sub = ss.insertSheet(subName);
      var headers = ['Timestamp', 'SheetCode', 'LastName', 'FirstName', 'Status', 'SubmittedAnswers', 'StartTime', 'EndTime', 'Date'];
      for (var h = 0; h < 8; h++) headers.push('TeacherCol' + (h + 1));
      sub.getRange(1, 1, 1, headers.length).setValues([headers]);
    }

    // Insert row at 2
    sub.insertRowBefore(2);

    // Try to copy teacher view if available (row 2, cols F..M)
    var teacherView = [];
    if (testSheet) {
      try { teacherView = testSheet.getRange(2, 6, 1, 8).getValues()[0]; } catch (e) { teacherView = []; }
    }

    var ts = timestamp || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

    var row = [];
    row.push(ts);                      // Timestamp
    row.push(sheetCode || '');
    row.push(lastName || '');
    row.push(firstName || '');
    row.push(status || 'partial');     // Status
    row.push(submittedAnswersJson || ''); // SubmittedAnswers raw JSON/text
    // leave Start/End/Date blanks if not provided
    row.push(''); row.push(''); row.push('');
    for (var i = 0; i < 8; i++) row.push(teacherView[i] || '');

    sub.getRange(2, 1, 1, row.length).setValues([row]);

    lock.releaseLock();
    return ContentService.createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    lock.releaseLock();
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/* ------------------ Duplicate check (existing behavior) ------------------ */
function checkDuplicate(sheetCode, lastName, firstName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const testSheet = ss.getSheetByName(sheetCode);
    if (!testSheet) {
      return ContentService.createTextOutput(JSON.stringify({exists: false, error: "Sheet not found for code: " + sheetCode}))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const lastRow = testSheet.getLastRow();
    if (lastRow < 2) {
      return ContentService.createTextOutput(JSON.stringify({exists: false})).setMimeType(ContentService.MimeType.JSON);
    }

    const data = testSheet.getRange(1, 1, lastRow, 7).getValues();
    const normLast = String(lastName || '').toLowerCase().trim();
    const normFirst = String(firstName || '').toLowerCase().trim();

    let submissionStartRow = 2;
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) {
        submissionStartRow = i + 1;
        break;
      }
    }
    if (submissionStartRow > data.length) submissionStartRow = data.length + 1;

    for (let i = submissionStartRow - 1; i < data.length; i++) {
      if (data[i][0]) continue;
      const sheetLast = String(data[i][5] || '').toLowerCase().trim();
      const sheetFirst = String(data[i][6] || '').toLowerCase().trim();
      if (sheetLast === normLast && sheetFirst === normFirst) {
        return ContentService.createTextOutput(JSON.stringify({exists: true})).setMimeType(ContentService.MimeType.JSON);
      }
    }

    return ContentService.createTextOutput(JSON.stringify({exists: false})).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({exists: false, error: err.toString()})).setMimeType(ContentService.MimeType.JSON);
  }
}

/* ------------------ Record final grade (compatible with QNNN keys) ------------------ */
/* recordGrade: writes a full grade submission to 'Submissions' sheet at row 2 */
function recordGrade(sheetCode, lastName, firstName, submittedAnswersJson, startTime, endTime, date) {
  // Use LockService to prevent concurrent write conflicts
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (e) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: 'Lock timeout' }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var testSheet = ss.getSheetByName(sheetCode);
    if (!testSheet) {
      lock.releaseLock();
      return ContentService.createTextOutput(JSON.stringify({ success: false, error: 'Sheet not found' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // Read question rows to build the correct-answer map (assumes answers in column C / index 2)
    var lastRow = testSheet.getLastRow();
    var data = testSheet.getRange(1, 1, Math.max(lastRow, 2), 4).getValues(); // A:D
    var correctMap = {};
    for (var r = 1; r < data.length; r++) {
      var rawId = String(data[r][0] || '').trim();
      // fallback to row index if ID missing so we preserve question order
      var qCode = normalizeToQ(rawId || r);
      var correct = String(data[r][2] || '').trim();
      if (qCode && correct) correctMap[qCode] = correct;
    }

    // Parse submitted answers safely
    var submittedAnswers = {};
    try {
      submittedAnswers = JSON.parse(submittedAnswersJson || '{}');
    } catch (err) {
      lock.releaseLock();
      return ContentService.createTextOutput(JSON.stringify({ success: false, error: 'Invalid submittedAnswers JSON' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // Compare and compute score
    var scoreCount = 0;
    var totalQuestions = Object.keys(correctMap).length;
    var correctStr = '';
    var mistakesStr = '';

    Object.keys(correctMap).forEach(function(qCode) {
      var has = Object.prototype.hasOwnProperty.call(submittedAnswers, qCode);
      if (!has) return; // skipped by student
      var userAnsRaw = submittedAnswers[qCode] || '';
      var userAns = String(userAnsRaw).toLowerCase().trim();
      var correctAns = String(correctMap[qCode] || '').toLowerCase().trim();
      if (userAns === correctAns) {
        scoreCount++;
        if (correctStr) correctStr += ', ';
        correctStr += qCode + ' ' + String(userAnsRaw);
      } else if (userAns) {
        if (mistakesStr) mistakesStr += ', ';
        mistakesStr += qCode + ' ' + String(userAnsRaw);
      }
    });

    var score = (totalQuestions > 0) ? (scoreCount + '/' + totalQuestions) : '0/0';

    // Prepare Submissions sheet (create if missing)
    var subName = 'Submissions';
    var sub = ss.getSheetByName(subName);
    if (!sub) {
      sub = ss.insertSheet(subName);
      // optional header row for teacher clarity
      var headers = ['Timestamp', 'SheetCode', 'LastName', 'FirstName', 'Score', 'CorrectList', 'MistakesList', 'StartTime', 'EndTime', 'Date'];
      // Reserve 8 cells for teacher-view F..M after that
      for (var h = 0; h < 8; h++) headers.push('TeacherCol' + (h + 1));
      sub.getRange(1, 1, 1, headers.length).setValues([headers]);
    }

    // Insert row at row 2 to keep newest on top
    sub.insertRowBefore(2);

    // Copy teacher-view columns F..M from the test sheet row 2 (if present)
    var teacherView = [];
    try {
      // try to grab 8 columns (F..M => 6..13). If sheet too small, this will gracefully fail to shorter array.
      teacherView = testSheet.getRange(2, 6, 1, 8).getValues()[0];
    } catch (e) {
      // fallback to empty cells
      teacherView = [];
    }

    var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

    // Build the row — zero-based then set via setValues
    var row = [];
    row.push(timestamp);             // A
    row.push(sheetCode || '');       // B
    row.push(lastName || '');        // C
    row.push(firstName || '');       // D
    row.push(score);                 // E
    row.push(correctStr);            // F
    row.push(mistakesStr);           // G
    row.push(startTime || '');       // H
    row.push(endTime || '');         // I
    row.push(date || '');            // J

    // Place teacher-view copies into subsequent columns (K..R)
    for (var i = 0; i < 8; i++) {
      row.push(teacherView[i] || '');
    }

    // Write the row into Submissions row 2
    sub.getRange(2, 1, 1, row.length).setValues([row]);

    lock.releaseLock();
    return ContentService.createTextOutput(JSON.stringify({ success: true, score: score }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    lock.releaseLock();
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * UPDATED: recordResultsFM
 * Writes to columns F..M in the SAME sheet as questions (e.g., TEST001)
 * Layout: A-D questions, E blank, F-M results
 * F: lastName, G: firstName, H: score, I: correct, J: mistakes, K: startTime, L: endTime, M: date
 * Finds the first empty row after questions and writes there.
 */
function recordResultsFM(sheetCode, lastName, firstName, score, correct, mistakes, startTime, endTime, date) {
  // Use LockService to prevent concurrent write conflicts when multiple students submit simultaneously
  var lock = LockService.getScriptLock();
  try {
    // Wait up to 30 seconds for other submissions to complete
    lock.waitLock(30000);
  } catch (e) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: 'Could not acquire lock - too many simultaneous submissions. Please try again.' }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(sheetCode);
    
    if (!sheet) {
      lock.releaseLock(); // Release lock before returning
      return ContentService
        .createTextOutput(JSON.stringify({ success: false, error: 'Sheet not found: ' + sheetCode }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // Ensure headers exist in row 1 (optional - only adds if F1 is empty)
    try {
      var headerCheck = sheet.getRange(1, 6, 1, 1).getValue();
      if (!headerCheck || String(headerCheck).trim() === '') {
        var headers = [['ID', 'Question', 'Answer', 'Timer', '', 'LastName', 'FirstName', 'Score', 'Correct', 'Mistakes', 'StartTime', 'EndTime', 'Date']];
        sheet.getRange(1, 1, 1, 13).setValues(headers);
        // Make header row bold
        sheet.getRange(1, 1, 1, 13).setFontWeight('bold');
      }
    } catch (e) {
      // Ignore header setup errors
    }

    // Find the next available row for student results
    // Strategy: Find first row where column F (LastName) is empty
    var lastRow = sheet.getLastRow();
    var targetRow = 2; // Default to row 2
    
    // Check column F (LastName) to find first empty slot
    if (lastRow >= 2) {
      var columnFData = sheet.getRange(2, 6, lastRow - 1, 1).getValues(); // Column F, starting from row 2
      
      // Find first empty cell in column F
      var foundEmptySlot = false;
      for (var i = 0; i < columnFData.length; i++) {
        if (!columnFData[i][0] || String(columnFData[i][0]).trim() === '') {
          targetRow = i + 2; // +2 because array is 0-based and we started at row 2
          foundEmptySlot = true;
          break;
        }
      }
      
      // If all slots filled, append at the end
      if (!foundEmptySlot) {
        targetRow = lastRow + 1;
      }
    }

    // Format times as "9:32:45 AM" with seconds
    var formattedStartTime = '';
    var formattedEndTime = '';
    
    try {
      if (startTime) {
        var startDate = new Date(startTime);
        formattedStartTime = Utilities.formatDate(startDate, Session.getScriptTimeZone(), 'h:mm:ss a');
      }
    } catch (e) {
      formattedStartTime = String(startTime || '');
    }
    
    try {
      if (endTime) {
        var endDate = new Date(endTime);
        formattedEndTime = Utilities.formatDate(endDate, Session.getScriptTimeZone(), 'h:mm:ss a');
      }
    } catch (e) {
      formattedEndTime = String(endTime || '');
    }

    // Write only F..M (columns 6-13). Range: targetRow, column 6, 1 row, 8 columns.
    var valuesFM = [[
      String(lastName || ''),      // F
      String(firstName || ''),     // G
      String(score || ''),         // H
      String(correct || ''),       // I
      String(mistakes || ''),      // J
      formattedStartTime,          // K (formatted time with seconds)
      formattedEndTime,            // L (formatted time with seconds)
      String(date || '')           // M
    ]];

    sheet.getRange(targetRow, 6, 1, 8).setValues(valuesFM);

    // Release lock after successful write
    lock.releaseLock();

    return ContentService
      .createTextOutput(JSON.stringify({ success: true, row: targetRow }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    // Release lock on error
    lock.releaseLock();
    
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/* ------------------ getAllQuestionsAndAnswers: FIXED - numeric-first, robust loader ------------------ */
function getAllQuestionsAndAnswers(sheetCode) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const testSheet = ss.getSheetByName(sheetCode);
    if (!testSheet) {
      return ContentService.createTextOutput(JSON.stringify({
        questions: [], questionsMap: {}, defaultTimerSeconds: 30, error: "Sheet not found for code: " + sheetCode
      })).setMimeType(ContentService.MimeType.JSON);
    }

    const lastRow = testSheet.getLastRow();
    // read columns A-D
    const data = testSheet.getRange(1, 1, Math.max(lastRow, 2), 4).getValues();
    let defaultTimerSeconds = 30;

    if (data.length > 1) {
      const d2Value = data[1][3];
      if (d2Value !== '' && !isNaN(d2Value)) {
        defaultTimerSeconds = Math.max(10, Math.min(300, parseInt(d2Value, 10)));
      }
    }

    Logger.log('Raw sheet data (first 5 rows): ' + JSON.stringify(data.slice(0,5)));

    const questionsList = [];
    let expectedNum = 1;

    for (let r = 1; r < data.length; r++) {
      const rawIdCell = String(data[r][0] || '').trim();
      const rawQuestion = String(data[r][1] || '').trim();
      const rawAnswer = String(data[r][2] || '').trim();
      const rawTimer = data[r][3];

      // skip fully empty rows
      if (!rawIdCell && !rawQuestion && !rawAnswer && (rawTimer === '' || rawTimer == null)) {
        continue;
      }

      // Determine numeric id - FIXED REGEX (was \\d, now \d)
      let num = null;
      if (/^\d+$/.test(rawIdCell)) {
        num = parseInt(rawIdCell, 10);
      } else {
        // if the ID column isn't numeric, fall back to expectedNum to keep contiguous numbering
        num = expectedNum;
      }

      // placeholders for missing question/answer to avoid silent drop
      let questionText = rawQuestion;
      if (!questionText) {
        questionText = `[MISSING QUESTION at row ${r+1}]`;
      }
      let answerText = rawAnswer;
      if (!answerText) {
        answerText = '-';
      }

      // timer
      let timerSeconds = defaultTimerSeconds;
      if (rawTimer !== '' && !isNaN(rawTimer)) {
        timerSeconds = Math.max(10, Math.min(300, parseInt(rawTimer, 10)));
      }

      // formatting - FIXED REGEX (was \\s, now \s)
      let formatted = questionText;
      try {
        if (typeof formatQuestion === 'function') {
          formatted = formatQuestion(questionText);
        } else {
          formatted = questionText.replace(/\s+/g, ' ').trim();
        }
      } catch (e) {
        formatted = questionText.replace(/\s+/g, ' ').trim();
      }

      const code = 'Q' + String(num).padStart(3, '0');

      questionsList.push({
        num: num,
        code: code,
        question: formatted,
        answer: answerText,
        timerSeconds: timerSeconds
      });

      expectedNum = num + 1;
    }

    // Build map for backward compatibility
    const questionsMap = {};
    questionsList.forEach(q => {
      if (questionsMap[q.code]) {
        Logger.log(`Duplicate numeric id detected: ${q.code} (overwriting previous)`);
      }
      questionsMap[q.code] = {
        code: q.code,
        question: q.question,
        answer: q.answer,
        timerSeconds: q.timerSeconds
      };
    });

    Logger.log('Processed questions: ' + questionsList.length);
    Logger.log('QuestionsMap keys: ' + Object.keys(questionsMap).join(', '));

    const response = {
      questions: questionsList,
      questionsMap: questionsMap,
      defaultTimerSeconds: defaultTimerSeconds
    };
    return ContentService.createTextOutput(JSON.stringify(response)).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    Logger.log('getAllQuestionsAndAnswers error: ' + err);
    return ContentService.createTextOutput(JSON.stringify({
      questions: [], questionsMap: {}, defaultTimerSeconds: 30, error: err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

/* ------------------ Helpers (question type & formatting) - FIXED REGEX ------------------ */
function getQuestionType(rawQuestion) {
  const lowerQ = String(rawQuestion || '').toLowerCase();
  const hasA = lowerQ.includes('a.');
  const hasB = lowerQ.includes('b.');
  const hasC = lowerQ.includes('c.');
  const hasD = lowerQ.includes('d.');
  const choiceCount = [hasA, hasB, hasC, hasD].filter(Boolean).length;
  if (choiceCount >= 3) return 'multiple-choice';
  if (choiceCount >= 1) return 'partial-choices';
  return 'identification';
}

function formatQuestion(rawQuestion) {
  const qType = getQuestionType(rawQuestion);
  if (qType === 'identification') {
    // FIXED: was \\s+, now \s+
    return String(rawQuestion || '').trim().replace(/\s+/g, ' ');
  }

  let stem = rawQuestion;
  let options = [];
  // Try to find options like "a. option b. option ..." - FIXED REGEX
  let stemEnd = rawQuestion.search(/\?\s*[a-d]\./i);
  if (stemEnd === -1) stemEnd = rawQuestion.search(/\.\s*[a-d]\./i);
  if (stemEnd > 0) {
    stem = rawQuestion.substring(0, stemEnd).trim();
    const optionsPart = rawQuestion.substring(stemEnd).trim();
    // FIXED: was \\s*, now \s*
    const optionRegex = /([a-d])\.\s*(.*?)(?=\s*[a-d]\.|$)/gis;
    let match;
    while ((match = optionRegex.exec(optionsPart)) !== null) {
      options.push({ letter: match[1].toLowerCase(), text: match[2].trim() });
    }
  } else {
    // fallback split
    const parts = rawQuestion.split(/([a-d]\.)/i);
    stem = parts[0].trim();
    for (let j = 1; j < parts.length; j += 2) {
      if (parts[j] && parts[j + 1]) {
        const letter = parts[j].trim().toLowerCase();
        const text = parts[j + 1].trim();
        if (letter.length === 1 && 'abcd'.includes(letter)) {
          options.push({ letter: letter, text: text });
        }
      }
    }
  }

  let formatted = stem;
  options.forEach(opt => {
    formatted += `<br> ${opt.letter}. ${opt.text}`;
  });

  if (qType === 'partial-choices') {
    formatted += `<br><small>(Partial choices - enter letter or full answer)</small>`;
  }

  // FIXED: was \\s*, now \s*
  formatted = formatted.replace(/\s*\n\s*/g, ' ').replace(/<br>\s*<br>/g, '<br>');
  return formatted;
}

