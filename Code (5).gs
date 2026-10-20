/**
 * ============================================================================
 * ISED UMROH ASSISTANT - GOOGLE APPS SCRIPT BACKEND (PRODUCTION READY)
 * Database: 10 Sheets Resmi Sesuai RPD I-S-E-D + Kwitansi & Manifes + WA API
 * ============================================================================
 */

const SHEETS_SPEC = {
  CONFIG: ["config_key", "config_value", "description"],
  LEADS: [
    "lead_id", "organization_id", "created_at", "updated_at", "name", "phone",
    "city", "source", "campaign", "package_id", "departure_period", "participant_count",
    "budget_range", "need_summary", "temperature", "lead_score", "score_confidence",
    "funnel_stage", "lead_status", "owner_user_id", "next_action", "next_follow_up_at",
    "last_contact_at", "lost_reason", "notes", "passport_no", "room_type", "visa_status", "vaccine_status"
  ],
  CONVERSATIONS: [
    "conversation_id", "organization_id", "lead_id", "timestamp",
    "direction", "message_text", "summary", "created_by"
  ],
  FOLLOW_UPS: [
    "follow_up_id", "organization_id", "lead_id", "due_at", "follow_up_type",
    "recommended_message", "task_status", "completed_at", "result", "next_follow_up_at", "assigned_to"
  ],
  ACTIVITIES: [
    "activity_id", "organization_id", "lead_id", "activity_type",
    "old_value", "new_value", "activity_note", "actor_user_id", "created_at"
  ],
  PACKAGES: [
    "package_id", "organization_id", "package_name", "price", "departure_date",
    "hotel_information", "facilities", "payment_terms", "quota_information",
    "terms_and_conditions", "is_active", "updated_at"
  ],
  TEMPLATES: [
    "template_id", "organization_id", "template_name", "funnel_stage",
    "objection_category", "template_text", "is_active", "updated_at"
  ],
  PAYMENTS: [
    "receipt_id", "organization_id", "lead_id", "jamaah_name", "package_id",
    "payment_type", "amount", "payment_method", "payment_date", "notes", "created_by", "created_at"
  ],
  AI_ANALYSES: [
    "analysis_id", "organization_id", "lead_id", "created_at", "model_version",
    "input_summary", "lead_score", "temperature", "confidence", "objection_category",
    "recommended_action", "suggested_follow_up", "output_json", "reviewed_by"
  ],
  USERS: [
    "user_id", "organization_id", "username", "password", "email", "display_name", "role", "is_active", "created_at"
  ]
};

/**
 * 1. AUTO-SETUP & STYLING SEMUA SHEET
 */
function setupAwal() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetNames = Object.keys(SHEETS_SPEC);

  sheetNames.forEach(name => {
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
    }
    const headers = SHEETS_SPEC[name];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight("bold")
      .setBackground("#0F2942")
      .setFontColor("#FFFFFF");
    sheet.setFrozenRows(1);

    const maxCols = sheet.getMaxColumns();
    if (maxCols > headers.length) {
      sheet.deleteColumns(headers.length + 1, maxCols - headers.length);
    }
    const maxRows = sheet.getMaxRows();
    if (maxRows > 200) {
      sheet.deleteRows(100, maxRows - 100);
    }
  });

  const userSheet = ss.getSheetByName("USERS");
  if (userSheet.getLastRow() === 1) {
    userSheet.appendRow([
      "USR-01", "ORG-001", "admin", "123", "admin@umroh.id",
      "Admin Travel", "ADMIN", true, new Date().toISOString()
    ]);
  }

  const cfgSheet = ss.getSheetByName("CONFIG");
  if (cfgSheet.getLastRow() === 1) {
    cfgSheet.getRange(2, 1, 7, 3).setValues([
      ["org_name", "Al-Haramain Perdana Umroh", "Nama Biro Travel"],
      ["hot_threshold", "70", "Ambang minimal skor HOT"],
      ["warm_threshold", "40", "Ambang minimal skor WARM"],
      ["gemini_model", "gemini-2.5-flash", "Model AI Gemini"],
      ["wa_api_endpoint", "https://api.fonnte.com/send", "Endpoint WA Business Gateway"],
      ["wa_api_token", "", "Token WA Business API"],
      ["wa_api_sender", "", "Nomor Pengirim WA Business"]
    ]);
  }
}

/**
 * 2. ROUTER UTAMA (doGet & doPost)
 */
function doGet(e) {
  return jsonResponse({ status: "success", message: "ISED Umroh Assistant API Active" });
}

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents || "{}");
    const action = payload.action;

    if (action === "login") return handleLogin(payload);
    if (action === "getInitialData") return getInitialData();
    if (action === "saveLead") return saveLead(payload.lead);
    if (action === "savePayment") return savePayment(payload.payment);
    if (action === "addConversation") return addConversation(payload.conversation);
    if (action === "addFollowUp") return addFollowUp(payload.followUp);
    if (action === "updateFollowUp") return updateFollowUp(payload.followUp);
    if (action === "addPackage") return addPackage(payload.pkg);
    if (action === "addTemplate") return addTemplate(payload.template);
    if (action === "saveConfig") return saveConfig(payload.config);
    if (action === "deleteDataRow") return deleteDataRow(payload); // FUNGSI HAPUS DATA
    if (action === "sendWhatsAppApi") return sendWhatsAppApi(payload);
    if (action === "analyzeWithGemini") return analyzeWithGemini(payload);
    if (action === "uploadFile") return uploadFileToDrive(payload);

    return jsonResponse({ status: "error", message: "Action tidak dikenali: " + action });
  } catch (err) {
    return jsonResponse({ status: "error", message: err.message });
  }
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * 3. HELPER BACA TABEL KE ARRAY OBJECT (BATCH READ)
 */
function readSheetAsObjects(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  return values.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i]; });
    return obj;
  });
}

/**
 * 4. FUNGSI LOGIN (DENGAN KONVERSI STRING PASSWORD ANTI-BUG ANGKA)
 */
function handleLogin(payload) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const users = readSheetAsObjects("USERS");
    const found = users.find(u =>
      String(u.username).trim() === String(payload.username).trim() &&
      String(u.password).trim() === String(payload.password).trim() &&
      String(u.is_active) !== "false"
    );
    if (!found) {
      return jsonResponse({ status: "error", message: "Username atau password tidak sesuai." });
    }
    return jsonResponse({
      status: "success",
      user: {
        user_id: found.user_id,
        organization_id: found.organization_id,
        username: found.username,
        display_name: found.display_name,
        role: found.role
      }
    });
  } catch (e) {
    return jsonResponse({ status: "error", message: e.message });
  } finally {
    lock.releaseLock();
  }
}

/**
 * 5. LOAD INITIAL DATA
 */
function getInitialData() {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const cache = CacheService.getScriptCache();
    let packages = JSON.parse(cache.get("CACHE_PACKAGES") || "null");
    if (!packages) {
      packages = readSheetAsObjects("PACKAGES");
      cache.put("CACHE_PACKAGES", JSON.stringify(packages), 300);
    }

    const data = {
      packages: packages,
      leads: readSheetAsObjects("LEADS"),
      conversations: readSheetAsObjects("CONVERSATIONS"),
      followUps: readSheetAsObjects("FOLLOW_UPS"),
      activities: readSheetAsObjects("ACTIVITIES"),
      templates: readSheetAsObjects("TEMPLATES"),
      payments: readSheetAsObjects("PAYMENTS")
    };
    return jsonResponse({ status: "success", data: data });
  } catch (e) {
    return jsonResponse({ status: "error", message: e.message });
  } finally {
    lock.releaseLock();
  }
}

/**
 * 6. SIMPAN / UPDATE LEADS, MANIFES & VALIDASI ANTI-DUPLIKASI WA
 */
function saveLead(lead) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("LEADS");
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const now = new Date().toISOString();
    const cleanPhone = String(lead.phone || "").replace(/[^0-9]/g, "");

    let existingRowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      const rowId = String(data[i][0]);
      const rowPhone = String(data[i][5]).replace(/[^0-9]/g, "");
      if (rowPhone === cleanPhone && rowId !== String(lead.lead_id)) {
        return jsonResponse({ status: "error", message: "Nomor WhatsApp sudah terdaftar di database!" });
      }
      if (rowId === String(lead.lead_id)) {
        existingRowIndex = i + 1;
      }
    }

    const rowValues = headers.map(h => {
      if (h === "updated_at") return now;
      if (h === "created_at") return lead.created_at || now;
      return lead[h] !== undefined ? lead[h] : "";
    });

    if (existingRowIndex > -1) {
      sheet.getRange(existingRowIndex, 1, 1, headers.length).setValues([rowValues]);
    } else {
      sheet.appendRow(rowValues);
    }

    return jsonResponse({ status: "success", lead_id: lead.lead_id });
  } catch (e) {
    return jsonResponse({ status: "error", message: e.message });
  } finally {
    lock.releaseLock();
  }
}

/**
 * 7. SIMPAN KWITANSI PEMBAYARAN (DP & PELUNASAN)
 */
function savePayment(payment) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("PAYMENTS");
    const now = new Date().toISOString();
    const row = SHEETS_SPEC.PAYMENTS.map(h => {
      if (h === "organization_id") return payment.organization_id || "ORG-001";
      if (h === "created_at") return now;
      return payment[h] !== undefined ? payment[h] : "";
    });
    sheet.appendRow(row);
    return jsonResponse({ status: "success", receipt_id: payment.receipt_id });
  } catch (e) {
    return jsonResponse({ status: "error", message: e.message });
  } finally {
    lock.releaseLock();
  }
}

/**
 * 8. CRUD CONVERSATIONS, FOLLOW_UPS, PACKAGES, TEMPLATES, CONFIG
 */
function addConversation(conv) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CONVERSATIONS");
    const row = SHEETS_SPEC.CONVERSATIONS.map(h => conv[h] || "");
    sheet.appendRow(row);
    return jsonResponse({ status: "success" });
  } catch (e) {
    return jsonResponse({ status: "error", message: e.message });
  } finally {
    lock.releaseLock();
  }
}

function addFollowUp(fu) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("FOLLOW_UPS");
    const row = SHEETS_SPEC.FOLLOW_UPS.map(h => fu[h] || "");
    sheet.appendRow(row);
    return jsonResponse({ status: "success" });
  } catch (e) {
    return jsonResponse({ status: "error", message: e.message });
  } finally {
    lock.releaseLock();
  }
}

function updateFollowUp(fu) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("FOLLOW_UPS");
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(fu.follow_up_id)) {
        const updatedRow = headers.map((h, idx) => fu[h] !== undefined ? fu[h] : data[i][idx]);
        sheet.getRange(i + 1, 1, 1, headers.length).setValues([updatedRow]);
        break;
      }
    }
    return jsonResponse({ status: "success" });
  } catch (e) {
    return jsonResponse({ status: "error", message: e.message });
  } finally {
    lock.releaseLock();
  }
}

function addPackage(pkg) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PACKAGES");
    const row = SHEETS_SPEC.PACKAGES.map(h => pkg[h] !== undefined ? pkg[h] : new Date().toISOString());
    sheet.appendRow(row);
    CacheService.getScriptCache().remove("CACHE_PACKAGES");
    return jsonResponse({ status: "success" });
  } catch (e) {
    return jsonResponse({ status: "error", message: e.message });
  } finally {
    lock.releaseLock();
  }
}

function addTemplate(tpl) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("TEMPLATES");
    const row = SHEETS_SPEC.TEMPLATES.map(h => tpl[h] !== undefined ? tpl[h] : new Date().toISOString());
    sheet.appendRow(row);
    return jsonResponse({ status: "success" });
  } catch (e) {
    return jsonResponse({ status: "error", message: e.message });
  } finally {
    lock.releaseLock();
  }
}

function saveConfig(cfg) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CONFIG");
    const rows = [
      ["org_name", cfg.org_name || "", "Nama Biro Travel"],
      ["license_no", cfg.license_no || "", "Izin PPIU"],
      ["hot_threshold", String(cfg.hot_threshold || 70), "Ambang minimal skor HOT"],
      ["warm_threshold", String(cfg.warm_threshold || 40), "Ambang minimal skor WARM"],
      ["wa_api_endpoint", cfg.wa_api_endpoint || "", "Endpoint WA Business Gateway"],
      ["wa_api_token", cfg.wa_api_token || "", "Token WA Business API"],
      ["wa_api_sender", cfg.wa_api_sender || "", "Nomor Pengirim WA Business"]
    ];
    sheet.getRange(2, 1, rows.length, 3).setValues(rows);
    return jsonResponse({ status: "success" });
  } catch (e) {
    return jsonResponse({ status: "error", message: e.message });
  } finally {
    lock.releaseLock();
  }
}

/**
 * 9. FUNGSI HAPUS DATA GLOBAL
 */
function deleteDataRow(payload) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(payload.sheetName);
    
    if (!sheet) throw new Error("Sheet tidak ditemukan");

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idColIndex = headers.indexOf(payload.idColumn);

    if (idColIndex === -1) throw new Error("Kolom ID tidak ditemukan");

    let rowIndexToDelete = -1;
    // Mulai dari indeks 1 untuk melewati baris header
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idColIndex]) === String(payload.idValue)) {
        rowIndexToDelete = i + 1; // +1 karena array index mulai dari 0 dan row sheet dari 1
        break;
      }
    }

    if (rowIndexToDelete > -1) {
      sheet.deleteRow(rowIndexToDelete);
      return jsonResponse({ status: "success", message: "Data berhasil dihapus dari sistem." });
    } else {
      return jsonResponse({ status: "error", message: "Data tidak ditemukan di database." });
    }
  } catch (e) {
    return jsonResponse({ status: "error", message: e.message });
  } finally {
    lock.releaseLock();
  }
}

/**
 * 10. INTEGRASI PENGIRIMAN PESAN WA BUSINESS API GATEWAY
 */
function sendWhatsAppApi(payload) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const cfgList = readSheetAsObjects("CONFIG");
    const endpointObj = cfgList.find(c => c.config_key === "wa_api_endpoint");
    const tokenObj = cfgList.find(c => c.config_key === "wa_api_token");
    const endpoint = endpointObj ? endpointObj.config_value : "";
    const token = tokenObj ? tokenObj.config_value : "";

    const convSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CONVERSATIONS");
    const now = new Date().toISOString();
    const recipients = payload.recipients || [];

    recipients.forEach(r => {
      const msg = String(payload.message || "").replace(/\{nama\}/g, r.name || "Bapak/Ibu");
      if (endpoint && token && token !== "TOKEN_WA_BUSINESS_API") {
        UrlFetchApp.fetch(endpoint, {
          method: "post",
          headers: { Authorization: token },
          payload: { target: r.phone, message: msg },
          muteHttpExceptions: true
        });
      }
      convSheet.appendRow([
        "CV-" + new Date().getTime(), "ORG-001", r.lead_id || "-", now,
        "OUTBOUND", msg, "WA Business API", "System API"
      ]);
    });

    return jsonResponse({ status: "success", sent_count: recipients.length });
  } catch (e) {
    return jsonResponse({ status: "error", message: e.message });
  } finally {
    lock.releaseLock();
  }
}

/**
 * 11. LAYANAN AI TERPUSAT (GEMINI API SERVER-SIDE)
 */
function analyzeWithGemini(payload) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const apiKey = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");
    if (!apiKey) {
      return jsonResponse({
        status: "error",
        message: "GEMINI_API_KEY belum diatur di Script Properties Google Apps Script."
      });
    }

    const prompt = `Kamu adalah AI Sales Assistant Umroh dengan metode Formula I-S-E-D (Iklan, Silaturahmi, Edukasi, Daftar).
DILARANG mengarang harga, hotel, atau fasilitas di luar data paket berikut:
Data Paket Fakta: ${JSON.stringify(payload.packageData || {})}
Profil & Kebutuhan Jamaah: ${payload.needSummary || ""}
Riwayat Percakapan: ${payload.conversationText || ""}
Berikan respons dalam format JSON terstruktur.`;

    const url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=" + apiKey;
    const response = UrlFetchApp.fetch(url, {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify({
        contents: [{ parts: [{ text: prompt }] }]
      }),
      muteHttpExceptions: true
    });

    const result = JSON.parse(response.getContentText());
    const aiText = result.candidates?.[0]?.content?.parts?.[0]?.text || "{}";

    const logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("AI_ANALYSES");
    logSheet.appendRow([
      "AI-" + new Date().getTime(), "ORG-001", payload.leadId || "-", new Date().toISOString(),
      "gemini-2.5-flash", payload.needSummary || "", "", "", "", "", "", "", aiText, "Admin"
    ]);

    return jsonResponse({ status: "success", aiOutput: aiText });
  } catch (e) {
    return jsonResponse({ status: "error", message: e.message });
  } finally {
    lock.releaseLock();
  }
}

/**
 * 12. PROTOKOL FILE UPLOAD & AUTO-FOLDER
 */
function uploadFileToDrive(payload) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const folderName = "MicroApp_Uploads";
    const folders = DriveApp.getFoldersByName(folderName);
    const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);

    const base64Data = payload.base64.split(",")[1] || payload.base64;
    const decoded = Utilities.base64Decode(base64Data);
    const blob = Utilities.newBlob(decoded, payload.mimeType || "image/jpeg", payload.fileName || "upload.jpg");
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    return jsonResponse({ status: "success", url: file.getUrl() });
  } catch (e) {
    return jsonResponse({ status: "error", message: e.message });
  } finally {
    lock.releaseLock();
  }
}

/* PLACEHOLDER: FUNGSI CRUD BARU */