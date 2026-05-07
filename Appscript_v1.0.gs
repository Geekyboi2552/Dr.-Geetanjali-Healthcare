// ============================================================
// DR. GEETANJALI HEALTHCARE — LAB INVENTORY SYSTEM
// Google Apps Script — Complete Automation Engine
// Version: 1.0 | Author: System Architect
// ============================================================

// ===== CONFIGURATION — UPDATE THESE BEFORE DEPLOYMENT =====
const CONFIG = {
  ADMIN_EMAIL: "divyanshu.luthra@geetanjalidiagnostics.com",      // Hemant Sir's email
  PROCUREMENT_EMAIL: "shivamluthra755@gmail.com",
  ADMIN_NAME: "Divyanshu",
  ORG_NAME: "Dr. Geetanjali Healthcare",
  
  // Sheet Names (must match exactly)
  SHEETS: {
    REQUEST_LOG: "Request Log",
    STOCK: "Current Stock",
    PROCUREMENT_TRACKER: "Active Procurement Tracker",
    DEMAND_POOL: "Demand Pool",
    DASHBOARD: "Admin Dashboard",
    TECHNICIANS: "Technician List",
    AUDIT: "Audit Trail",
  },

  // Thresholds
  PENDING_REMINDER_DAYS: 3,
  PENDING_ESCALATE_DAYS: 7,
  EXPIRY_ALERT_DAYS: 30,

  // Item name normalization map (add synonyms here)
  ITEM_ALIASES: {
    "calculators": "Calculator",
    "calc": "Calculator",
    "blue tube": "Blue Tubes",
    "reagent kit": "Reagent Kits",
    "glove": "Gloves",
    "gloves latex": "Latex Gloves",
  }
};

// ============================================================
// MODULE 1: FORM SUBMISSION TRIGGER
// Fires when a technician submits the Google Form
// ============================================================
function onFormSubmit(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Read from Form Responses 2 sheet directly — most reliable method
    const formSheet = ss.getSheetByName("Form Responses");
    const lastRow = formSheet.getLastRow();
    const headers = formSheet.getRange(1, 1, 1, formSheet.getLastColumn()).getValues()[0];
    const rowData = formSheet.getRange(lastRow, 1, 1, formSheet.getLastColumn()).getValues()[0];
    
    // Helper — finds value by header name (case-insensitive, trims spaces)
    function getVal(headerName) {
      const idx = headers.findIndex(h => 
        h.toString().toLowerCase().trim() === headerName.toLowerCase().trim()
      );
      return idx >= 0 && rowData[idx] ? rowData[idx].toString().trim() : "";
    }
    
    // Extract all values
    const timestamp    = new Date();
    const techName     = getVal("Technician Name");
    const department   = getVal("Department / Section");
    const category     = getVal("Item Category");
    const rawItemName  = getVal("Item Name");
    const brand        = getVal("Brand Preferred (optional)");
    const qtyRequired  = parseInt(getVal("Quantity Required")) || 0;
    const unit         = getVal("Unit");
    const urgency      = getVal("Urgency Level") || "Low";
    const stockLeft    = getVal("Current Stock Left (approx.)");
    const reason       = getVal("Reason for Requirement");
    const itemName     = normalizeItemName(rawItemName);
    
    // Write to Request Log
    const requestLog = ss.getSheetByName(CONFIG.SHEETS.REQUEST_LOG);
    const reqLastRow = requestLog.getLastRow();
    const writeRow   = reqLastRow + 1;
    const requestId  = generateRequestId(writeRow);
    
    requestLog.getRange(writeRow, 1, 1, 15).setValues([[
      requestId,
      timestamp,
      techName,
      department,
      category,
      itemName,
      brand,
      qtyRequired,
      unit,
      urgency,
      stockLeft,
      reason,
      "Pending",
      "Awaiting Approval",
      ""
    ]]);
    
    // Formatting, duplicate check, audit, email
    applyUrgencyFormatting(requestLog, writeRow, urgency);
    const duplicateInfo = checkDuplicate(itemName, qtyRequired, requestId, ss);
    logAudit(ss, requestId, "SUBMITTED", techName, itemName, qtyRequired, urgency, "New request submitted");
    sendAdminAlert(requestId, techName, department, itemName, brand, qtyRequired, unit, urgency, stockLeft, reason, duplicateInfo);
    refreshDashboard(ss);

  } catch(err) {
    Logger.log("onFormSubmit Error: " + err.toString());
    MailApp.sendEmail(CONFIG.ADMIN_EMAIL, "⚠️ System Error - Lab Inventory", 
      "Error in form submission: " + err.toString());
  }
}
// ============================================================
// MODULE 2: DUPLICATE DETECTION ENGINE
// ============================================================
function checkDuplicate(itemName, newQty, newRequestId, ss) {
  const requestLog = ss.getSheetByName(CONFIG.SHEETS.REQUEST_LOG);
  const procTracker = ss.getSheetByName(CONFIG.SHEETS.PROCUREMENT_TRACKER);
  
  const activeStatuses = ["Pending", "Approved", "Ordered", "In Transit"];
  const data = requestLog.getDataRange().getValues();
  
  let duplicates = [];
  let totalExistingQty = 0;
  let relatedRequestIds = [];
  
  // Check all existing requests for same item with active status
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const existingItem = row[5];   // Column F: Item Name
    const existingStatus = row[12]; // Column M: Admin Status
    const existingQty = row[7];    // Column H: Quantity
    const existingId = row[0];     // Column A: Request ID
    
    if (existingId === newRequestId) continue; // Skip current row
    
    if (normalizeItemName(existingItem) === normalizeItemName(itemName) && 
        activeStatuses.includes(existingStatus)) {
      duplicates.push({
        requestId: existingId,
        status: existingStatus,
        qty: existingQty,
        tech: row[2]
      });
      totalExistingQty += parseInt(existingQty) || 0;
      relatedRequestIds.push(existingId);
    }
  }
  
  if (duplicates.length > 0) {
    // Mark the new request as potential duplicate in Request Log
    const newRow = findRowById(requestLog, newRequestId);
    if (newRow > 0) {
      requestLog.getRange(newRow, 15).setValue(
        "⚠️ DUPLICATE ALERT: " + duplicates.length + " existing request(s) found"
      );
      // Highlight in orange
      requestLog.getRange(newRow, 1, 1, 15).setBackground("#FF9800").setFontColor("#000000");
    }
    
    // Update Demand Pool
    updateDemandPool(ss, itemName, newRequestId, newQty, duplicates);
    
    return {
      isDuplicate: true,
      count: duplicates.length,
      totalExistingQty: totalExistingQty,
      details: duplicates,
      suggestedTotal: totalExistingQty + newQty,
      relatedIds: relatedRequestIds.join(", ")
    };
  }
  
  return { isDuplicate: false };
}

// ============================================================
// MODULE 3: ADMIN APPROVAL TRIGGER
// Run this via button/dropdown in Admin Sheet
// ============================================================
function onAdminApproval(requestId, decision, priority, notes) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const requestLog = ss.getSheetByName(CONFIG.SHEETS.REQUEST_LOG);
  const procTracker = ss.getSheetByName(CONFIG.SHEETS.PROCUREMENT_TRACKER);
  
  const rowNum = findRowById(requestLog, requestId);
  if (rowNum === -1) {
    SpreadsheetApp.getUi().alert("Request ID not found: " + requestId);
    return;
  }
  
  const rowData = requestLog.getRange(rowNum, 1, 1, 15).getValues()[0];
  const techName = rowData[2];
  const techEmail = getTechnicianEmail(techName, ss);
  const itemName = rowData[5];
  const qtyRequired = rowData[7];
  const urgency = rowData[9];
  
  const timestamp = new Date();
  
  if (decision === "Approved") {
    // Update status
    requestLog.getRange(rowNum, 13).setValue("Approved");
    requestLog.getRange(rowNum, 15).setValue(
      "Approved by Hemant Sir on " + formatDate(timestamp) + ". Priority: " + priority + ". " + notes
    );
    requestLog.getRange(rowNum, 1, 1, 15).setBackground("#E8F5E9");
    
    // Add to Procurement Tracker
    addToProcurementTracker(procTracker, requestId, itemName, qtyRequired, rowData, priority);
    
    // Email Procurement Team
    sendProcurementAlert(requestId, itemName, qtyRequired, rowData[8], priority, techName, urgency);
    
    // Notify Technician
    if (techEmail) {
      sendTechNotification(techEmail, techName, requestId, itemName, qtyRequired, "Approved", 
        "Your request has been approved and forwarded to Procurement.");
    }
    
  } else if (decision === "Rejected") {
    requestLog.getRange(rowNum, 13).setValue("Rejected");
    requestLog.getRange(rowNum, 15).setValue(
      "Rejected by Hemant Sir on " + formatDate(timestamp) + ". Reason: " + notes
    );
    requestLog.getRange(rowNum, 1, 1, 15).setBackground("#FFEBEE");
    
    // Notify Technician
    if (techEmail) {
      sendTechNotification(techEmail, techName, requestId, itemName, qtyRequired, "Rejected", 
        "Reason: " + notes);
    }
  }
  
  // Log to audit trail
  logAudit(ss, requestId, "ADMIN_" + decision.toUpperCase(), "Hemant Sir", itemName, qtyRequired, urgency, notes);
  refreshDashboard(ss);
}

// ============================================================
// MODULE 4: PROCUREMENT DELIVERY UPDATE
// Run when procurement marks item as Delivered
// ============================================================
function onDeliveryUpdate(requestId, deliveredQty, deliveryDate, vendorName, poNumber) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const requestLog = ss.getSheetByName(CONFIG.SHEETS.REQUEST_LOG);
  const stockSheet = ss.getSheetByName(CONFIG.SHEETS.STOCK);
  const procTracker = ss.getSheetByName(CONFIG.SHEETS.PROCUREMENT_TRACKER);
  
  const rowNum = findRowById(requestLog, requestId);
  if (rowNum === -1) return;
  
  const rowData = requestLog.getRange(rowNum, 1, 1, 15).getValues()[0];
  const techName = rowData[2];
  const techEmail = getTechnicianEmail(techName, ss);
  const itemName = rowData[5];
  
  // Update Request Log
  requestLog.getRange(rowNum, 14).setValue("Delivered");
  requestLog.getRange(rowNum, 15).setValue(
    rowData[14] + " | Delivered: " + deliveredQty + " units on " + formatDate(new Date(deliveryDate)) + 
    " | Vendor: " + vendorName + " | PO: " + poNumber
  );
  requestLog.getRange(rowNum, 1, 1, 15).setBackground("#E3F2FD");
  
  // Auto-update stock quantity
  updateStockQuantity(stockSheet, itemName, parseInt(deliveredQty), vendorName, deliveryDate);
  
  // Update Procurement Tracker
  updateProcurementTrackerDelivery(procTracker, requestId, deliveredQty, deliveryDate);
  
  // Notify Technician
  if (techEmail) {
    sendTechNotification(techEmail, techName, requestId, itemName, deliveredQty, "Delivered",
      "Your requested item has been delivered and stock has been updated.");
  }
  
  // Log audit
  logAudit(ss, requestId, "DELIVERED", vendorName, itemName, deliveredQty, "", 
    "PO: " + poNumber + " | Delivered by: " + vendorName);
  
  refreshDashboard(ss);
}

// ============================================================
// MODULE 5: SCHEDULED TRIGGERS
// Set up time-based triggers in Apps Script
// ============================================================

// Run daily — check pending request escalations
function dailyEscalationCheck() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const requestLog = ss.getSheetByName(CONFIG.SHEETS.REQUEST_LOG);
  const data = requestLog.getDataRange().getValues();
  const today = new Date();
  
  let reminders = [];
  let escalations = [];
  let lowStockAlerts = [];
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const status = row[12]; // Admin Status
    const timestamp = new Date(row[1]);
    const daysOld = Math.floor((today - timestamp) / (1000 * 60 * 60 * 24));
    const requestId = row[0];
    const itemName = row[5];
    const techName = row[2];
    const urgency = row[9];
    
    if (status === "Pending") {
      if (daysOld >= CONFIG.PENDING_ESCALATE_DAYS) {
        escalations.push({ requestId, itemName, techName, daysOld, urgency });
        // Mark as escalated
        requestLog.getRange(i + 1, 13).setValue("Pending - ESCALATED");
        requestLog.getRange(i + 1, 1, 1, 15).setBackground("#FF5722").setFontColor("#FFFFFF");
      } else if (daysOld >= CONFIG.PENDING_REMINDER_DAYS) {
        reminders.push({ requestId, itemName, techName, daysOld, urgency });
      }
    }
  }
  
  // Check low stock
  lowStockAlerts = getLowStockItems(ss);
  
  // Send daily digest to Admin
  if (reminders.length > 0 || escalations.length > 0 || lowStockAlerts.length > 0) {
    sendDailyDigest(reminders, escalations, lowStockAlerts);
  }
}

// Run weekly — stock summary
function weeklyStockReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const lowStockItems = getLowStockItems(ss);
  const expiryItems = getExpiringItems(ss);
  sendWeeklyStockSummary(lowStockItems, expiryItems);
}

// ============================================================
// EMAIL FUNCTIONS
// ============================================================
function sendAdminAlert(requestId, techName, dept, itemName, brand, qty, unit, urgency, stockLeft, reason, duplicateInfo) {
  const urgencyEmoji = urgency === "Critical" ? "🔴" : urgency === "High" ? "🟠" : urgency === "Medium" ? "🟡" : "🟢";
  const isDuplicate = duplicateInfo && duplicateInfo.isDuplicate;
  
  const subject = `${urgencyEmoji} New Request [${requestId}] — ${itemName} — ${techName}${isDuplicate ? " ⚠️ DUPLICATE" : ""}`;
  
  let body = `
<html><body style="font-family: Arial, sans-serif; max-width: 600px;">
<div style="background: #1565C0; color: white; padding: 16px 20px; border-radius: 8px 8px 0 0;">
  <h2 style="margin:0">New Inventory Request</h2>
  <p style="margin:4px 0 0; opacity:0.8">${CONFIG.ORG_NAME}</p>
</div>
<div style="border: 1px solid #e0e0e0; border-top: none; padding: 20px; border-radius: 0 0 8px 8px;">
  <table style="width:100%; border-collapse: collapse;">
    <tr><td style="padding:6px; color:#666; width:40%">Request ID</td><td style="padding:6px; font-weight:bold">${requestId}</td></tr>
    <tr style="background:#f5f5f5"><td style="padding:6px; color:#666">Technician</td><td style="padding:6px">${techName} — ${dept}</td></tr>
    <tr><td style="padding:6px; color:#666">Item Name</td><td style="padding:6px; font-weight:bold; font-size:16px">${itemName}</td></tr>
    <tr style="background:#f5f5f5"><td style="padding:6px; color:#666">Brand</td><td style="padding:6px">${brand || "Any"}</td></tr>
    <tr><td style="padding:6px; color:#666">Quantity Required</td><td style="padding:6px; font-size:16px; font-weight:bold">${qty} ${unit}</td></tr>
    <tr style="background:#f5f5f5"><td style="padding:6px; color:#666">Urgency</td><td style="padding:6px">${urgencyEmoji} <strong>${urgency}</strong></td></tr>
    <tr><td style="padding:6px; color:#666">Current Stock Left</td><td style="padding:6px">${stockLeft}</td></tr>
    <tr style="background:#f5f5f5"><td style="padding:6px; color:#666">Reason</td><td style="padding:6px">${reason}</td></tr>
  </table>`;

  if (isDuplicate) {
    body += `
  <div style="background:#FFF3E0; border-left:4px solid #FF9800; padding:12px; margin:16px 0; border-radius:4px;">
    <h3 style="margin:0 0 8px; color:#E65100">⚠️ Duplicate / Existing Request Detected</h3>
    <p><strong>${duplicateInfo.totalExistingQty} ${unit}</strong> of "${itemName}" already in pipeline across <strong>${duplicateInfo.count}</strong> existing request(s).</p>
    <p>Related Request IDs: ${duplicateInfo.relatedIds}</p>
    <p>Suggested consolidated total: <strong>${duplicateInfo.suggestedTotal} ${unit}</strong></p>
    <p style="margin:0; color:#666">Please review before approving. Use "Merge Requests" to consolidate.</p>
  </div>`;
  }

  body += `
  <div style="margin:20px 0; text-align:center;">
    <a href="${getApprovalLink()}" style="background:#2E7D32; color:white; padding:12px 24px; border-radius:6px; text-decoration:none; margin-right:8px">✅ Open Admin Panel</a>
  </div>
  <p style="font-size:12px; color:#999; text-align:center">This is an automated alert from ${CONFIG.ORG_NAME} Inventory System</p>
</div></body></html>`;

  MailApp.sendEmail({
    to: CONFIG.ADMIN_EMAIL,
    subject: subject,
    htmlBody: body
  });
}

function sendProcurementAlert(requestId, itemName, qty, unit, priority, techName, urgency) {
  const subject = `📦 New Procurement Request [${requestId}] — ${itemName} — Priority: ${priority}`;
  const body = `
<html><body style="font-family: Arial, sans-serif; max-width: 600px;">
<div style="background:#4A148C; color:white; padding:16px 20px; border-radius:8px 8px 0 0;">
  <h2 style="margin:0">Procurement Action Required</h2>
  <p style="margin:4px 0 0; opacity:0.8">${CONFIG.ORG_NAME}</p>
</div>
<div style="border:1px solid #e0e0e0; border-top:none; padding:20px; border-radius:0 0 8px 8px;">
  <p>This request has been <strong>APPROVED</strong> by ${CONFIG.ADMIN_NAME} and requires procurement action.</p>
  <table style="width:100%; border-collapse:collapse;">
    <tr style="background:#f5f5f5"><td style="padding:8px; color:#666">Request ID</td><td style="padding:8px; font-weight:bold">${requestId}</td></tr>
    <tr><td style="padding:8px; color:#666">Item</td><td style="padding:8px; font-size:16px; font-weight:bold">${itemName}</td></tr>
    <tr style="background:#f5f5f5"><td style="padding:8px; color:#666">Quantity</td><td style="padding:8px">${qty} ${unit}</td></tr>
    <tr><td style="padding:8px; color:#666">Priority</td><td style="padding:8px; font-weight:bold">${priority}</td></tr>
    <tr style="background:#f5f5f5"><td style="padding:8px; color:#666">Requested By</td><td style="padding:8px">${techName}</td></tr>
    <tr><td style="padding:8px; color:#666">Urgency</td><td style="padding:8px">${urgency}</td></tr>
  </table>
  <div style="background:#E8F5E9; border-left:4px solid #4CAF50; padding:12px; margin:16px 0; border-radius:4px;">
    <p style="margin:0"><strong>Action Required:</strong> Please update the Procurement Queue Sheet with vendor details, PO number, and delivery date. Mark status when ordered and delivered.</p>
  </div>
  <p style="font-size:12px; color:#999">Automated alert — ${CONFIG.ORG_NAME}</p>
</div></body></html>`;

  MailApp.sendEmail({ to: CONFIG.PROCUREMENT_EMAIL, subject: subject, htmlBody: body });
}

function sendTechNotification(email, techName, requestId, itemName, qty, status, message) {
  const statusColors = { "Approved": "#2E7D32", "Rejected": "#C62828", "Delivered": "#1565C0" };
  const color = statusColors[status] || "#555";
  const subject = `[Update] Your Request [${requestId}] — ${itemName} — ${status}`;
  const body = `
<html><body style="font-family:Arial,sans-serif; max-width:500px;">
<div style="background:${color}; color:white; padding:14px 18px; border-radius:8px 8px 0 0;">
  <h2 style="margin:0">Request Update: ${status}</h2>
</div>
<div style="border:1px solid #e0e0e0; border-top:none; padding:16px; border-radius:0 0 8px 8px;">
  <p>Dear ${techName},</p>
  <p>Your inventory request has been updated:</p>
  <table style="width:100%; border-collapse:collapse;">
    <tr style="background:#f5f5f5"><td style="padding:6px; color:#666">Request ID</td><td style="padding:6px">${requestId}</td></tr>
    <tr><td style="padding:6px; color:#666">Item</td><td style="padding:6px; font-weight:bold">${itemName}</td></tr>
    <tr style="background:#f5f5f5"><td style="padding:6px; color:#666">Quantity</td><td style="padding:6px">${qty}</td></tr>
    <tr><td style="padding:6px; color:#666">Status</td><td style="padding:6px; font-weight:bold; color:${color}">${status}</td></tr>
  </table>
  <p>${message}</p>
  <p style="font-size:12px; color:#999">For questions, contact ${CONFIG.ADMIN_NAME}. — ${CONFIG.ORG_NAME}</p>
</div></body></html>`;

  MailApp.sendEmail({ to: email, subject: subject, htmlBody: body });
}

function sendDailyDigest(reminders, escalations, lowStockItems) {
  const subject = `📋 Daily Digest — ${escalations.length} Escalations, ${reminders.length} Reminders, ${lowStockItems.length} Low Stock`;
  
  let body = `<html><body style="font-family:Arial,sans-serif; max-width:600px;">
<div style="background:#37474F; color:white; padding:16px 20px; border-radius:8px 8px 0 0;">
  <h2 style="margin:0">Daily Inventory Digest</h2>
  <p style="margin:4px 0 0; opacity:0.8">${formatDate(new Date())} — ${CONFIG.ORG_NAME}</p>
</div>
<div style="border:1px solid #e0e0e0; border-top:none; padding:20px; border-radius:0 0 8px 8px;">`;

  if (escalations.length > 0) {
    body += `<div style="background:#FFEBEE; border-left:4px solid #C62828; padding:12px; margin:12px 0; border-radius:4px;">
    <h3 style="margin:0 0 8px; color:#C62828">🔴 ${escalations.length} Escalated Requests (7+ Days Pending)</h3>
    <ul style="margin:0; padding-left:20px;">`;
    escalations.forEach(e => {
      body += `<li>${e.requestId} — ${e.itemName} (${e.techName}) — <strong>${e.daysOld} days old</strong></li>`;
    });
    body += `</ul></div>`;
  }

  if (reminders.length > 0) {
    body += `<div style="background:#FFF3E0; border-left:4px solid #FF9800; padding:12px; margin:12px 0; border-radius:4px;">
    <h3 style="margin:0 0 8px; color:#E65100">🟠 ${reminders.length} Pending Requests (3+ Days)</h3>
    <ul style="margin:0; padding-left:20px;">`;
    reminders.forEach(r => {
      body += `<li>${r.requestId} — ${r.itemName} (${r.techName}) — ${r.daysOld} days pending</li>`;
    });
    body += `</ul></div>`;
  }

  if (lowStockItems.length > 0) {
    body += `<div style="background:#FFF8E1; border-left:4px solid #FFC107; padding:12px; margin:12px 0; border-radius:4px;">
    <h3 style="margin:0 0 8px; color:#F57F17">⚠️ ${lowStockItems.length} Low / Out of Stock Items</h3>
    <ul style="margin:0; padding-left:20px;">`;
    lowStockItems.forEach(s => {
      body += `<li>${s.item} — <strong>${s.current} ${s.unit}</strong> remaining (threshold: ${s.threshold})</li>`;
    });
    body += `</ul></div>`;
  }

  body += `<p style="font-size:12px; color:#999; text-align:center">Automated digest — ${CONFIG.ORG_NAME} Inventory System</p>
</div></body></html>`;

  MailApp.sendEmail({ to: CONFIG.ADMIN_EMAIL, subject: subject, htmlBody: body });
}

// ============================================================
// HELPER FUNCTIONS
// ============================================================
function generateRequestId(lastRow) {
  const today = new Date();
  const yy = String(today.getFullYear()).slice(-2);
  const mm = String(today.getMonth() + 1).padStart(2, '0');
  const dd = String(today.getDate()).padStart(2, '0');
  const seq = String(new Date().getTime()).slice(-4);
  return `REQ-${yy}${mm}${dd}-${seq}`;
}

function normalizeItemName(rawName) {
  if (!rawName || typeof rawName !== 'string') return "";
  const cleaned = rawName.trim().toLowerCase();
  
  if (CONFIG.ITEM_ALIASES[cleaned]) return CONFIG.ITEM_ALIASES[cleaned];
  
  const singular = cleaned.endsWith('s') && cleaned.length > 3 ? cleaned.slice(0, -1) : cleaned;
  if (CONFIG.ITEM_ALIASES[singular]) return CONFIG.ITEM_ALIASES[singular];
  
  return rawName.trim().replace(/\w\S*/g, txt => txt.charAt(0).toUpperCase() + txt.substr(1).toLowerCase());
}

function findRowById(sheet, requestId) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === requestId) return i + 1;
  }
  return -1;
}

function applyUrgencyFormatting(sheet, rowNum, urgency) {
  const colors = {
    "Critical": { bg: "#FFCDD2", font: "#B71C1C" },
    "High":     { bg: "#FFE0B2", font: "#BF360C" },
    "Medium":   { bg: "#FFF9C4", font: "#F57F17" },
    "Low":      { bg: "#E8F5E9", font: "#1B5E20" }
  };
  const style = colors[urgency] || colors["Low"];
  sheet.getRange(rowNum, 10).setBackground(style.bg).setFontColor(style.font).setFontWeight("bold");
}

function updateStockQuantity(stockSheet, itemName, addQty, vendorName, deliveryDate) {
  const data = stockSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (normalizeItemName(data[i][1]) === normalizeItemName(itemName)) {
      const currentQty = parseInt(data[i][3]) || 0;
      const newQty = currentQty + addQty;
      const threshold = parseInt(data[i][4]) || 0;
      const status = newQty === 0 ? "Out of Stock" : newQty <= threshold ? "Low Stock" : "In Stock";
      
      stockSheet.getRange(i + 1, 4).setValue(newQty);  // Update quantity
      stockSheet.getRange(i + 1, 7).setValue(vendorName); // Update vendor
      stockSheet.getRange(i + 1, 8).setValue(new Date(deliveryDate)); // Last purchase
      stockSheet.getRange(i + 1, 11).setValue(status); // Update status
      
      // Traffic-light formatting
      const statusColors = {
        "In Stock":    "#E8F5E9",
        "Low Stock":   "#FFF9C4",
        "Out of Stock":"#FFCDD2"
      };
      stockSheet.getRange(i + 1, 11).setBackground(statusColors[status]);
      return;
    }
  }
  // Item not found — add new row
  stockSheet.appendRow([
    "ITEM-" + (stockSheet.getLastRow()),
    itemName, "", addQty, 0, "", vendorName, new Date(deliveryDate), "", "", "In Stock"
  ]);
}

function getLowStockItems(ss) {
  const stockSheet = ss.getSheetByName(CONFIG.SHEETS.STOCK);
  const data = stockSheet.getDataRange().getValues();
  const lowItems = [];
  for (let i = 1; i < data.length; i++) {
    const current = parseInt(data[i][3]) || 0;
    const threshold = parseInt(data[i][4]) || 0;
    if (current <= threshold) {
      lowItems.push({
        item: data[i][1],
        current: current,
        threshold: threshold,
        unit: data[i][5],
        status: data[i][10]
      });
    }
  }
  return lowItems;
}

function getExpiringItems(ss) {
  const stockSheet = ss.getSheetByName(CONFIG.SHEETS.STOCK);
  const data = stockSheet.getDataRange().getValues();
  const today = new Date();
  const alertDate = new Date(today.getTime() + CONFIG.EXPIRY_ALERT_DAYS * 24 * 60 * 60 * 1000);
  const expiringItems = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][8]) {
      const expiry = new Date(data[i][8]);
      if (expiry <= alertDate && expiry >= today) {
        const daysLeft = Math.floor((expiry - today) / (1000 * 60 * 60 * 24));
        expiringItems.push({ item: data[i][1], expiry: formatDate(expiry), daysLeft: daysLeft });
      }
    }
  }
  return expiringItems;
}

function updateDemandPool(ss, itemName, newRequestId, newQty, existingDuplicates) {
  const demandPool = ss.getSheetByName(CONFIG.SHEETS.DEMAND_POOL);
  existingDuplicates.forEach(dup => {
    demandPool.appendRow([
      itemName,
      newRequestId,
      newQty,
      dup.requestId,
      dup.qty,
      dup.status,
      "Pending Merge Decision",
      new Date()
    ]);
  });
}

function addToProcurementTracker(procTracker, requestId, itemName, qty, rowData, priority) {
  procTracker.appendRow([
    itemName, qty, 0, qty, 0, 0, "", requestId, "Clear", "Single", priority, new Date()
  ]);
}

function updateProcurementTrackerDelivery(procTracker, requestId, deliveredQty, deliveryDate) {
  const data = procTracker.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][7] && data[i][7].toString().includes(requestId)) {
      procTracker.getRange(i + 1, 6).setValue(parseInt(deliveredQty));
      procTracker.getRange(i + 1, 4).setValue(Math.max(0, parseInt(data[i][3]) - parseInt(deliveredQty)));
      break;
    }
  }
}

function refreshDashboard(ss) {
  const dashboard = ss.getSheetByName(CONFIG.SHEETS.DASHBOARD);
  const requestLog = ss.getSheetByName(CONFIG.SHEETS.REQUEST_LOG);
  if (!dashboard || !requestLog) return;
  
  const data = requestLog.getDataRange().getValues();
  let pending = 0, approved = 0, rejected = 0, ordered = 0, delivered = 0, duplicates = 0;
  
  for (let i = 1; i < data.length; i++) {
    const status = data[i][12];
    if (status.includes("Pending")) pending++;
    else if (status === "Approved") approved++;
    else if (status === "Rejected") rejected++;
    else if (status === "Ordered") ordered++;
    else if (status === "Delivered") delivered++;
    if (data[i][14] && data[i][14].toString().includes("DUPLICATE")) duplicates++;
  }
  
  // Write KPI values to dashboard (these cells are referenced by dashboard formulas)
  dashboard.getRange("B2").setValue(pending);
  dashboard.getRange("B3").setValue(approved);
  dashboard.getRange("B4").setValue(rejected);
  dashboard.getRange("B5").setValue(ordered);
  dashboard.getRange("B6").setValue(delivered);
  dashboard.getRange("B7").setValue(duplicates);
  dashboard.getRange("B8").setValue(new Date());
}

function logAudit(ss, requestId, action, actor, itemName, qty, urgency, notes) {
  const audit = ss.getSheetByName(CONFIG.SHEETS.AUDIT);
  if (!audit) return;
  audit.appendRow([new Date(), requestId, action, actor, itemName, qty, urgency, notes]);
}

function getTechnicianEmail(techName, ss) {
  const techSheet = ss.getSheetByName(CONFIG.SHEETS.TECHNICIANS);
  if (!techSheet) return null;
  const data = techSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === techName) return data[i][1];
  }
  return null;
}

function getApprovalLink() {
  return SpreadsheetApp.getActiveSpreadsheet().getUrl() + "#gid=0";
}

function formatDate(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "dd-MMM-yyyy");
}

function sendWeeklyStockSummary(lowStockItems, expiryItems) {
  const subject = `📊 Weekly Stock Summary — ${CONFIG.ORG_NAME}`;
  let body = `<html><body style="font-family:Arial,sans-serif;"><h2>Weekly Stock Summary</h2>`;
  body += `<h3>Low / Out of Stock Items (${lowStockItems.length})</h3><ul>`;
  lowStockItems.forEach(i => body += `<li>${i.item}: ${i.current}/${i.threshold} ${i.unit}</li>`);
  body += `</ul><h3>Expiring within 30 Days (${expiryItems.length})</h3><ul>`;
  expiryItems.forEach(i => body += `<li>${i.item}: Expires ${i.expiry} (${i.daysLeft} days left)</li>`);
  body += `</ul></body></html>`;
  MailApp.sendEmail({ to: CONFIG.ADMIN_EMAIL, subject: subject, htmlBody: body });
}

// ============================================================
// MODULE 6: SETUP FUNCTION — RUN ONCE ON DEPLOYMENT
// ============================================================
function setupSystem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Create all sheets if they don't exist
  const sheetsToCreate = [
    { name: CONFIG.SHEETS.REQUEST_LOG, headers: [
      "Request ID","Timestamp","Technician","Department","Category","Item Name","Brand",
      "Qty Required","Unit","Urgency","Stock Left","Reason","Admin Status","Procurement Status","Notes"
    ]},
    { name: CONFIG.SHEETS.STOCK, headers: [
      "Item ID","Item Name","Category","Current Qty","Min Threshold","Unit","Vendor",
      "Last Purchase Date","Expiry Date","Storage Location","Status"
    ]},
    { name: CONFIG.SHEETS.PROCUREMENT_TRACKER, headers: [
      "Item Name","Total Requested Qty","Pending Qty","Approved Qty","Ordered Qty",
      "In Transit Qty","Expected Delivery","Linked Request IDs","Duplicate Alert","Consolidation Status","Priority","Created"
    ]},
    { name: CONFIG.SHEETS.DEMAND_POOL, headers: [
      "Item Name","New Request ID","New Qty","Existing Request ID","Existing Qty",
      "Existing Status","Merge Decision","Timestamp"
    ]},
    { name: CONFIG.SHEETS.AUDIT, headers: [
      "Timestamp","Request ID","Action","Actor","Item Name","Qty","Urgency","Notes"
    ]},
    { name: CONFIG.SHEETS.TECHNICIANS, headers: ["Technician Name","Email","Department","Phone"]},
    { name: CONFIG.SHEETS.DASHBOARD, headers: ["Metric","Value"]},
  ];
  
  sheetsToCreate.forEach(sheetConfig => {
    let sheet = ss.getSheetByName(sheetConfig.name);
    if (!sheet) {
      sheet = ss.insertSheet(sheetConfig.name);
    }
    // Write headers
    const headerRange = sheet.getRange(1, 1, 1, sheetConfig.headers.length);
    headerRange.setValues([sheetConfig.headers]);
    headerRange.setFontWeight("bold").setBackground("#37474F").setFontColor("#FFFFFF");
    sheet.setFrozenRows(1);
    sheet.setColumnWidths(1, sheetConfig.headers.length, 150);
  });
  
  // Set up Dashboard seed data
  const dashboard = ss.getSheetByName(CONFIG.SHEETS.DASHBOARD);
  const dashboardData = [
    ["Metric", "Value"],
    ["Total Pending", 0],
    ["Total Approved", 0],
    ["Total Rejected", 0],
    ["Total Ordered", 0],
    ["Total Delivered", 0],
    ["Duplicate Requests Prevented", 0],
    ["Last Refreshed", new Date()],
  ];
  dashboard.getRange(1, 1, dashboardData.length, 2).setValues(dashboardData);
  
  // Set up triggers
  setupTriggers();
  
  SpreadsheetApp.getUi().alert("✅ Dr. Geetanjali Healthcare Inventory System setup complete!\n\nNext steps:\n1. Fill in Technician List sheet\n2. Link your Google Form to this spreadsheet\n3. Configure CONFIG variables at top of Code.gs");
}

function setupTriggers() {
  // Remove existing triggers
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));

  // FORM SUBMISSION TRIGGER
  ScriptApp.newTrigger("onFormSubmit")
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onFormSubmit()
    .create();

  // Daily escalation
  ScriptApp.newTrigger("dailyEscalationCheck")
    .timeBased()
    .atHour(8)
    .everyDays(1)
    .create();

  // Weekly stock report
  ScriptApp.newTrigger("weeklyStockReport")
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(9)
    .create();

  Logger.log("All triggers configured successfully");
}

// ============================================================
// MODULE 7: ADMIN UI MENU
// Adds custom menu to the spreadsheet
// ============================================================
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🏥 Lab Inventory')
  .addItem('📋 Setup System (First Time)', 'setupSystem')
  .addSeparator()
  .addItem('✅ Approve Selected Request', 'approveSelectedRequest')
  .addItem('❌ Reject Selected Request', 'rejectSelectedRequest')
  .addItem('🔀 Merge Duplicate Requests', 'mergeDuplicateRequests')
  .addSeparator()
  .addItem('📊 Refresh Dashboard', 'refreshDashboardMenu')
  .addItem('📧 Send Daily Digest Now', 'sendDailyDigestNow')
  .addItem('📦 Check Low Stock', 'checkLowStockNow')
  .addSeparator()
  .addItem('🚀 Open Command Center Dashboard', 'openDashboard')
  .addToUi();
    
}

function approveSelectedRequest() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSheet();
  const activeRow = sheet.getActiveRange().getRow();
  if (activeRow <= 1) { ui.alert("Please select a data row"); return; }
  
  const requestId = sheet.getRange(activeRow, 1).getValue();
  const itemName = sheet.getRange(activeRow, 6).getValue();
  
  const priorityResult = ui.prompt(`Approving: ${requestId} — ${itemName}`, "Priority (Normal / High / Critical):", ui.ButtonSet.OK_CANCEL);
  if (priorityResult.getSelectedButton() !== ui.Button.OK) return;
  
  const notesResult = ui.prompt("Approval Notes (optional):", ui.ButtonSet.OK_CANCEL);
  const notes = notesResult.getSelectedButton() === ui.Button.OK ? notesResult.getResponseText() : "";
  
  onAdminApproval(requestId, "Approved", priorityResult.getResponseText(), notes);
  ui.alert(`✅ Request ${requestId} approved and forwarded to Procurement!`);
}

function rejectSelectedRequest() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSheet();
  const activeRow = sheet.getActiveRange().getRow();
  if (activeRow <= 1) { ui.alert("Please select a data row"); return; }
  
  const requestId = sheet.getRange(activeRow, 1).getValue();
  const reasonResult = ui.prompt(`Rejecting: ${requestId}`, "Reason for rejection:", ui.ButtonSet.OK_CANCEL);
  if (reasonResult.getSelectedButton() !== ui.Button.OK) return;
  
  onAdminApproval(requestId, "Rejected", "", reasonResult.getResponseText());
  ui.alert(`Request ${requestId} rejected.`);
}

function refreshDashboardMenu() {
  refreshDashboard(SpreadsheetApp.getActiveSpreadsheet());
  SpreadsheetApp.getUi().alert("Dashboard refreshed!");
}

function sendDailyDigestNow() {
  dailyEscalationCheck();
  SpreadsheetApp.getUi().alert("Daily digest sent to " + CONFIG.ADMIN_EMAIL);
}

function checkLowStockNow() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const items = getLowStockItems(ss);
  if (items.length === 0) {
    SpreadsheetApp.getUi().alert("✅ All stock levels are healthy!");
  } else {
    const msg = items.map(i => `• ${i.item}: ${i.current} ${i.unit} (threshold: ${i.threshold})`).join("\n");
    SpreadsheetApp.getUi().alert(`⚠️ ${items.length} Low Stock Items:\n\n${msg}`);
  }
}

function mergeDuplicateRequests() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.prompt(
    "Merge Requests",
    "Enter comma-separated Request IDs to merge (e.g. REQ-2501-0002, REQ-2501-0005):",
    ui.ButtonSet.OK_CANCEL
  );
  if (result.getSelectedButton() !== ui.Button.OK) return;
  
  const ids = result.getResponseText().split(",").map(id => id.trim());
  ui.alert(`Merge initiated for ${ids.length} requests. Please manually update the Master Request with consolidated quantity and mark others as "Merged - See ${ids[0]}".`);
}
function openDashboard() {
  const html = HtmlService.createHtmlOutputFromFile("Dashboard")
    .setWidth(1600)
    .setHeight(900);
    
  SpreadsheetApp.getUi().showModalDialog(
    html,
    "🏥 Dr. Geetanjali Inventory Command Center"
  );
}
function getDashboardData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const requestSheet = ss.getSheetByName("Request Log");
  const data = requestSheet.getDataRange().getValues();

  let results = [];

  for (let i = 1; i < data.length; i++) {
    results.push({
      department: data[i][3],
      urgency: data[i][9],
      category: data[i][4],
      requests: Number(data[i][7]) || 0
    });
  }

  return results;
}


function getDashboardData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const requestSheet = ss.getSheetByName("Request Log");
  const stockSheet = ss.getSheetByName("Current Stock");

  const requestData = requestSheet.getDataRange().getValues();
  const stockData = stockSheet.getDataRange().getValues();

  let requests = [];
  let stockSummary = {
    inStock: 0,
    lowStock: 0,
    outStock: 0
  };

  // REQUEST LOG DATA
  for (let i = 1; i < requestData.length; i++) {
    requests.push({
      requestId: requestData[i][0],
      technician: requestData[i][2],
      department: requestData[i][3],
      category: requestData[i][4],
      itemName: requestData[i][5],
      qty: Number(requestData[i][7]) || 0,
      urgency: requestData[i][9],
      adminStatus: requestData[i][12]
    });
  }

  // STOCK DATA
  for (let i = 1; i < stockData.length; i++) {
    const status = stockData[i][10];

    if (status === "In Stock") stockSummary.inStock++;
    else if (status === "Low Stock") stockSummary.lowStock++;
    else if (status === "Out of Stock") stockSummary.outStock++;
  }

  return {
    requests: requests,
    stock: stockSummary
  };
}

