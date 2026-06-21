/**
 * PENTING: PENYELESAIAN RALAT PERMISSION DRIVEAPP (GET FOLDER & CREATE FILE)
 * Untuk mengelakkan ralat "You do not have permission to call DriveApp.Folder.createFile",
 * Anda DIWAJIBKAN menjalankan (Run) fungsi `forceAuthorization()` ini SEKALI SAHAJA secara
 * manual dalam editor Apps Script sebelum melakukan Deployment sebagai Web App.
 */
function forceAuthorization() {
  var dummyUrl = "https://www.google.com";
  var response = UrlFetchApp.fetch(dummyUrl);
  SpreadsheetApp.getActiveSpreadsheet();
  
  // Trik untuk memaksa Google Apps Script meminta skop (scope) penuh ke atas Google Drive
  // (membenarkan skrip mencipta dan mengedit fail, bukan sekadar membaca).
  var tempFolder = DriveApp.getRootFolder(); 
  var tempFile = tempFolder.createFile("LatihAI_Temp_Auth.txt", "Fail sementara untuk memaksa kebenaran (permission) Drive.");
  tempFile.setTrashed(true); // Padam terus fail sementara tersebut
  
  Logger.log("Kebenaran (Permission) penuh Drive dan Spreadsheet telah berjaya disahkan. Sila teruskan dengan 'New Deployment' sebagai Web App.");
}

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var action = data.action;
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // GLOBAL SETTINGS
    var settingsSheet = ss.getSheetByName("SETTINGS");
    if (!settingsSheet) {
      settingsSheet = ss.insertSheet("SETTINGS");
      settingsSheet.appendRow(["GLOBAL_MODE", "ulangkaji"]);
      settingsSheet.appendRow(["INACTIVE_SETS", "[]"]);
      settingsSheet.appendRow(["SET_DURATIONS", "{}"]);
      settingsSheet.appendRow(["GLOBAL_DIFFICULTY", "sederhana"]);
    }
    
    var globalMode = settingsSheet.getRange("B1").getValue() || "ulangkaji";
    
    // Check for INACTIVE_SETS row
    var a2Val = settingsSheet.getRange("A2").getValue();
    if (a2Val !== "INACTIVE_SETS") {
      settingsSheet.getRange("A2").setValue("INACTIVE_SETS");
      var currentB2 = settingsSheet.getRange("B2").getValue();
      if (!currentB2) settingsSheet.getRange("B2").setValue("[]");
    }

    // Check for SET_DURATIONS row
    var a3Val = settingsSheet.getRange("A3").getValue();
    if (a3Val !== "SET_DURATIONS") {
      settingsSheet.getRange("A3").setValue("SET_DURATIONS");
      var currentB3 = settingsSheet.getRange("B3").getValue();
      if (!currentB3) settingsSheet.getRange("B3").setValue("{}");
    }

    // Check for GLOBAL_DIFFICULTY row
    var a4Val = settingsSheet.getRange("A4").getValue();
    if (a4Val !== "GLOBAL_DIFFICULTY") {
      settingsSheet.getRange("A4").setValue("GLOBAL_DIFFICULTY");
      var currentB4 = settingsSheet.getRange("B4").getValue();
      if (!currentB4) settingsSheet.getRange("B4").setValue("sederhana");
    }
    var globalDifficulty = settingsSheet.getRange("B4").getValue() || "sederhana";
    
    var inactiveSetsStr = settingsSheet.getRange("B2").getValue();
    var inactiveSets = [];
    try { inactiveSets = JSON.parse(inactiveSetsStr) || []; } catch(err) { inactiveSets = []; }

    var setDurationsStr = settingsSheet.getRange("B3").getValue();
    var setDurations = {};
    try { setDurations = JSON.parse(setDurationsStr) || {}; } catch(err) { setDurations = {}; }

    var pSheet = ss.getSheetByName("PELAJAR");
    if (!pSheet) {
      pSheet = ss.insertSheet("PELAJAR");
      pSheet.appendRow(["Nama", "Markah", "Tarikh Terakhir", "Sejarah Topik", "Level Pet"]);
    }

    if (action === "init") {
      var keySheet = ss.getSheetByName("API_KEY");
      if (!keySheet) return createJsonResponse({ error: "Tab 'API_KEY' tiada." });
      var apiKey = keySheet.getRange("A1").getValue();
      
      var qSheet = ss.getSheetByName("SOALAN") || ss.getSheetByName("soalan");
      if (!qSheet) return createJsonResponse({ error: "Tab 'SOALAN' tiada dalam Google Sheet." });
      
      var lastRow = qSheet.getLastRow();
      var lastCol = qSheet.getLastColumn();
      var sets = [];
      
      if (lastRow > 0 && lastCol > 0) {
        var headers = qSheet.getRange(1, 1, 1, lastCol).getValues()[0];
        var qData = lastRow > 1 ? qSheet.getRange(2, 1, lastRow - 1, lastCol).getValues() : [];
        
        for (var c = 0; c < lastCol; c++) {
          var setName = headers[c];
          if (setName && setName.toString().trim() !== "") {
            var topics = [];
            for (var r = 0; r < qData.length; r++) {
              if (qData[r][c] && qData[r][c].toString().trim() !== "") {
                topics.push(qData[r][c].toString().trim());
              }
            }
            sets.push({ 
              setName: setName.toString().trim(), 
              topics: topics,
              isActive: !inactiveSets.includes(setName.toString().trim())
            });
          }
        }
      }
      
      var userPoints = 0;
      var answeredTopics = {};
      var petLevel = 0;
      if (data.name) {
         var pData = pSheet.getDataRange().getValues();
         for(var i=1; i<pData.length; i++) {
            if(pData[i][0] === data.name) {
               userPoints = pData[i][1] || 0;
               try {
                 var strTopics = pData[i][3];
                 if(strTopics) answeredTopics = JSON.parse(strTopics);
               } catch(err) { answeredTopics = {}; }
               petLevel = pData[i][4] || 0;
               break;
            }
         }
      }

      return createJsonResponse({
        success: true,
        apiKey: apiKey,
        questionSets: sets,
        setDurations: setDurations,
        userPoints: userPoints,
        answeredTopics: answeredTopics,
        petLevel: petLevel,
        globalMode: globalMode,
        globalDifficulty: globalDifficulty
      });
    }
    
    if (action === "save_score") {
      var pData = pSheet.getDataRange().getValues();
      var found = false;
      var topicJSON = JSON.stringify(data.answeredTopics || {});
      var incomingPetLevel = data.petLevel !== undefined ? data.petLevel : 0;
      for (var i = 1; i < pData.length; i++) {
        if (pData[i][0] === data.name) {
          pSheet.getRange(i + 1, 2).setValue(data.points);
          pSheet.getRange(i + 1, 3).setValue(new Date());
          pSheet.getRange(i + 1, 4).setValue(topicJSON);
          pSheet.getRange(i + 1, 5).setValue(incomingPetLevel);
          found = true;
          break;
        }
      }
      if (!found) {
        pSheet.appendRow([data.name, data.points, new Date(), topicJSON, incomingPetLevel]);
      }
      return createJsonResponse({ success: true });
    }
    
    if (action === "submit_exam") {
      try {
        var videoUrl = "Tiada Video";
        if (data.videoBase64 && data.videoBase64.includes("base64,")) {
          var b64String = data.videoBase64.split("base64,")[1];
          var blob = Utilities.newBlob(Utilities.base64Decode(b64String), "video/webm", "Exam_" + data.name + "_" + new Date().getTime() + ".webm");
          var folder;
          try {
            // Cuba simpan dalam folder spesifik
            folder = DriveApp.getFolderById("11g-UNAQtjOFYYarUM2nigGb7PgyJa321");
          } catch (folderError) {
            // FALLBACK PENTING: Jika ID folder di atas tiada kebenaran atau tidak wujud,
            // sistem akan menyimpan video tersebut di dalam Root Folder pembuat sistem.
            folder = DriveApp.getRootFolder();
          }
          var file = folder.createFile(blob);
          videoUrl = file.getUrl();
        }

        // Menyelaras penyimpanan log markah peperiksaan
        var examSS;
        try {
          examSS = SpreadsheetApp.openById("1VUdZrFieEMOzik_bxMwdFrNfKwLpd3mVwy7cdYRo2mU");
        } catch(ssErr) {
          examSS = SpreadsheetApp.getActiveSpreadsheet();
        }
        
        var rekodSheet = examSS.getSheetByName("REKOD_PEPERIKSAAN");
        if (!rekodSheet) {
          rekodSheet = examSS.insertSheet("REKOD_PEPERIKSAAN");
          rekodSheet.appendRow(["Tarikh/Masa", "Nama Pelajar", "Markah", "Jumlah Soalan", "Peratus", "Masa Diambil (s)", "Lokasi GPS", "Pautan Video", "Butiran Terperinci"]);
        }
        
        rekodSheet.appendRow([new Date(), data.name, data.score, data.total, data.percentage + "%", data.timeTaken, data.gpsLocation || "Tiada Lokasi", videoUrl, data.details]);

        return createJsonResponse({ success: true });
      } catch (err) {
        return createJsonResponse({ error: "Gagal menyimpan rekod peperiksaan: " + err.message });
      }
    }

    if (action === "get_leaderboard") {
      var scores = [];
      var pData = pSheet.getDataRange().getValues();
      for (var i = 1; i < pData.length; i++) {
        if (pData[i][0]) {
          scores.push({ name: pData[i][0], points: pData[i][1] || 0 });
        }
      }
      scores.sort(function(a, b) { return b.points - a.points; });
      scores = scores.slice(0, 10);
      return createJsonResponse({ success: true, leaderboard: scores });
    }

    if (action === "get_all_students") {
      if (data.adminPwd !== "101010") return createJsonResponse({ error: "Katalaluan Admin Tidak Sah." });
      var students = [];
      var pData = pSheet.getDataRange().getValues();
      for (var i = 1; i < pData.length; i++) {
        if (pData[i][0]) {
          students.push({ name: pData[i][0], points: pData[i][1] || 0 });
        }
      }
      return createJsonResponse({ success: true, students: students });
    }

    if (action === "admin_action") {
      if (data.adminPwd !== "101010") return createJsonResponse({ error: "Akses ditolak" });
      
      if (data.task === "set_settings") {
        settingsSheet.getRange("B1").setValue(data.mode);
        settingsSheet.getRange("B4").setValue(data.difficulty);
        return createJsonResponse({ success: true });
      }
      
      if (data.task === "toggle_set") {
        var idx = inactiveSets.indexOf(data.targetName);
        if (idx > -1) {
            inactiveSets.splice(idx, 1);
        } else {
            inactiveSets.push(data.targetName);
        }
        settingsSheet.getRange("B2").setValue(JSON.stringify(inactiveSets));
        return createJsonResponse({ success: true, inactiveSets: inactiveSets });
      }

      if (data.task === "set_duration") {
        setDurations[data.targetName] = parseInt(data.duration) || 0;
        settingsSheet.getRange("B3").setValue(JSON.stringify(setDurations));
        return createJsonResponse({ success: true, setDurations: setDurations });
      }

      if (data.task === "reset_all") {
        var lastRow = pSheet.getLastRow();
        if (lastRow > 1) {
          pSheet.getRange(2, 2, lastRow - 1, 1).setValue(0);
          pSheet.getRange(2, 4, lastRow - 1, 1).setValue("{}");
          pSheet.getRange(2, 5, lastRow - 1, 1).setValue(0);
        }
        return createJsonResponse({ success: true });
      }
      
      if (data.task === "reset_student" || data.task === "delete_student") {
        var pData = pSheet.getDataRange().getValues();
        for (var i = 1; i < pData.length; i++) {
          if (pData[i][0] === data.targetName) {
            if (data.task === "reset_student") {
              pSheet.getRange(i + 1, 2).setValue(0);
              pSheet.getRange(i + 1, 4).setValue("{}");
              pSheet.getRange(i + 1, 5).setValue(0);
            } else if (data.task === "delete_student") {
              pSheet.deleteRow(i + 1);
            }
            break;
          }
        }
        return createJsonResponse({ success: true });
      }
    }
    
    return createJsonResponse({ error: "Aksi tidak dikenali." });

  } catch (err) {
    return createJsonResponse({ error: err.message });
  }
}

function doGet(e) {
  return ContentService.createTextOutput("Sistem Backend AI berfungsi.").setMimeType(ContentService.MimeType.TEXT);
}

function createJsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
