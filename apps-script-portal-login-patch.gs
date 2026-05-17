/**
 * HCT VARK Apps Script patch: robust email + access-code portal login.
 *
 * Paste this near the bottom of Code.gs after removing older duplicate copies of:
 * - verifyStudentCode
 * - getPortalDataByCode
 *
 * Why this patch exists:
 * The older verifyStudentCode() used parseInt() on the access code. That can
 * cause correct-looking codes to fail when Sheets or pasted input changes the
 * value shape. These helpers normalize both the sheet value and the typed value
 * consistently before comparing.
 */

function normalizePortalEmail(value) {
  return String(value || "").trim().toLowerCase();
}

function normalizeAccessCode(value) {
  if (value === null || value === undefined) return "";

  if (typeof value === "number" && isFinite(value)) {
    return String(Math.round(value)).padStart(6, "0");
  }

  var text = String(value).trim();
  var numericText = text.match(/^(\d+)(?:\.0+)?$/);
  if (numericText) {
    return numericText[1].padStart(6, "0");
  }

  return text.replace(/\D/g, "").padStart(6, "0").slice(-6);
}

function verifyStudentCode(email, code) {
  var sh = getOrCreatePortalSheet();
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return null;

  var rows = sh.getRange(2, 1, lastRow - 1, CODE_COLS.length).getValues();
  var normalEmail = normalizePortalEmail(email);
  var normalCode = normalizeAccessCode(code);

  for (var i = 0; i < rows.length; i++) {
    var storedEmail = normalizePortalEmail(rows[i][0]);
    var storedCode = normalizeAccessCode(rows[i][1]);

    Logger.log(
      "portal login compare email=[" + storedEmail + "] vs [" + normalEmail + "] " +
      "code=[" + storedCode + "] vs [" + normalCode + "]"
    );

    if (storedEmail === normalEmail && storedCode === normalCode) {
      var r = rows[i];
      return {
        email: r[0],
        code: storedCode,
        name: r[2],
        topModality: r[3],
        resultType: r[4],
        V: Number(r[5]) || 0,
        A: Number(r[6]) || 0,
        R: Number(r[7]) || 0,
        K: Number(r[8]) || 0,
        exam: r[9],
        surveyDate: r[10]
      };
    }
  }

  return null;
}

function getPortalDataByCode(email, code) {
  var student = verifyStudentCode(email, code);
  if (!student) {
    return {
      status: "error",
      message: "Email or access code is incorrect. Please use the same email and the 6-digit code from your VARK results email."
    };
  }

  var lV = Number(student.V) || 0;
  var lA = Number(student.A) || 0;
  var lR = Number(student.R) || 0;
  var lK = Number(student.K) || 0;

  var scores = [
    { code: "V", n: lV },
    { code: "A", n: lA },
    { code: "R", n: lR },
    { code: "K", n: lK }
  ].sort(function(a, b) {
    return b.n - a.n;
  });

  var diff12 = scores[0].n - scores[1].n;
  var diff13 = scores[0].n - scores[2].n;
  var topCodes;
  if (diff12 >= 5) {
    topCodes = [scores[0].code];
  } else if (diff13 <= 4) {
    topCodes = [scores[0].code, scores[1].code, scores[2].code];
  } else {
    topCodes = [scores[0].code, scores[1].code];
  }

  var studentOut = {
    name: student.name,
    school: "",
    exam: student.exam,
    email: normalizePortalEmail(student.email),
    submittedAt: student.surveyDate || "",
    resultType: student.resultType,
    topModality: student.topModality,
    topCode: scores[0].code,
    topCodes: topCodes,
    prefStrength: prefLabel(scores[0].n),
    scores: { V: lV, A: lA, R: lR, K: lK }
  };

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var responses = ss.getSheetByName(SHEET_NAME);
    if (responses && responses.getLastRow() >= 2) {
      var lastRow = responses.getLastRow();
      var emailCol = responses.getRange(2, 6, lastRow - 1, 1).getValues();
      var schoolCol = responses.getRange(2, 4, lastRow - 1, 1).getValues();
      var normEmail = normalizePortalEmail(email);
      for (var i = emailCol.length - 1; i >= 0; i--) {
        if (normalizePortalEmail(emailCol[i][0]) === normEmail) {
          studentOut.school = String(schoolCol[i][0] || "");
          break;
        }
      }
    }
  } catch (err) {
    Logger.log("Could not enrich school: " + err.message);
  }

  var tips = {};
  ["V", "A", "R", "K"].forEach(function(c) {
    var m = VARK_META[c];
    if (m) tips[c] = { tagline: m.tagline, items: m.tips };
  });

  return {
    status: "success",
    student: studentOut,
    resources: readLibrary(),
    bodySystems: BODY_SYSTEMS,
    resourceTypes: RESOURCE_TYPES,
    tips: tips
  };
}
