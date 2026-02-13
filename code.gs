// PROPERTY OF USER - DEPLOY AS WEB APP (v2)
const SHEET_NAME = "WarrantyData";
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.tryLock(10000);
  try {
    const doc = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = doc.getSheetByName(SHEET_NAME);
    
    if (!sheet) {
      sheet = doc.insertSheet(SHEET_NAME);
      // Added Technician Columns + Note + Recripte info
      const headers = [
        "KEY", "Work Order", "Spare Part Code", "Spare Part Name", "Old Material Code",
        "Qty", "Serial Number", "Store Code", "Store Name", "วันที่รับซาก",
        "ผู้รับซาก", "Plant", "Keep", "CI Name", "Problem", "Product Type", "Product",
        "Warranty Action", "Recorder", "Timestamp", 
        "Booking Slip", "Booking Date", "Claim Receiver", "Claim Date", "ClaimSup",
        "รหัสช่าง", "ชื่อช่าง", "Mobile", "Plantcenter",
        "Note", "Recripte", "RecripteDate" 
      ];
      sheet.appendRow(headers);
    }
    const data = JSON.parse(e.postData.contents);
    
    // START: Sanitize Booking Date (Force "DD/MM/YYYY" by split)
    if (data['Booking Date']) {
      let bd = String(data['Booking Date']);
      // Handle "T" (ISO) or " " (DateTime)
      bd = bd.split('T')[0].split(' ')[0];
      data['Booking Date'] = bd;
    }
    // END: Sanitize
    
    const operation = data.operation || 'save'; // 'save' or 'delete'
    
    // Construct KEY
    const wo = data['work order'] || data['Work Order'] || "";
    const key = (wo) + (data['Spare Part Code'] || "");
    data.KEY = key;

    // Check & Add Missing Columns dynamically
    let headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const newColumns = ["Plantcenter", "Note", "Recripte", "RecripteDate", "Claim Date", "ClaimSup", "Datefinish"];
    
    newColumns.forEach(newCol => {
        const exists = headers.some(h => String(h).toLowerCase().trim() === newCol.toLowerCase());
        if (!exists) {
            sheet.getRange(1, headers.length + 1).setValue(newCol);
            // Refresh headers
            headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
        }
    });

    // Find Row FIRST to determine if we need to preserve Timestamp
    const keyIndex = headers.indexOf("KEY");
    const timestampIndex = headers.indexOf("Timestamp");
    const allValues = sheet.getDataRange().getValues();
    let rowIndex = -1;
    for (let i = 1; i < allValues.length; i++) {
      if (String(allValues[i][keyIndex]) === String(key)) {
        rowIndex = i + 1;
        break;
      }
    }
    
    // TIMESTAMP LOGIC:
    if (rowIndex > 0) {
        if (timestampIndex !== -1) {
            data.Timestamp = allValues[rowIndex - 1][timestampIndex];
        } else {
             data.Timestamp = new Date();
        }
    } else {
        data.Timestamp = new Date();
    }
    
    const rowData = headers.map(header => {

      // 1. Try exact match
      let val = data[header];

      // 2. Try mapped key
      const mappedKey = mapHeaderToKey(header);
      if (val === undefined) {
        val = data[mappedKey];
      }

      // 3. Try robust lookup (ignore case/spaces) for Header & Mapped Key
      if (val === undefined) {
        val = getRobustValue(data, header);
      }
      if (val === undefined) {
        val = getRobustValue(data, mappedKey);
      }
      
      // Special Fallbacks
      if (val === undefined) {
         if (mappedKey === 'Plantcenter') val = data['Plantcenter'];
         if (mappedKey === 'Recripte') val = data['Recripte'];
         if (mappedKey === 'RecripteDate') val = data['RecripteDate'];
         if (mappedKey === 'Note') val = data['Note']; 
      }
      return val || "";
    });
    
    // --- DELETE OPERATION ---
    if (operation === 'delete') {
      if (rowIndex > 0) {
        sheet.deleteRow(rowIndex);
        return response({ status: "success", action: "deleted" });
      } else {
        return response({ status: "error", message: "Record not found to delete" });
      }
    }
    
    // --- SAVE (UPDATE/INSERT) OPERATION ---
    if (rowIndex > 0) {
      // Update
      sheet.getRange(rowIndex, 1, 1, rowData.length).setValues([rowData]);
      return response({ status: "success", action: "updated", row: rowIndex });
    } else {
      // Append
      sheet.appendRow(rowData);
      return response({ status: "success", action: "created" });
    }
  } catch (e) {
    return response({ status: "error", message: e.toString() });
  } finally {
    lock.releaseLock();
  }
}

function response(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}

function mapHeaderToKey(header) {
  const map = {
    "Work Order": "work order",
    "Old Material Code": "old material code",
    "Date Received": "วันที่รับซาก", // Keep for backward compatibility if header is English
    "Receiver": "ผู้รับซาก",         // Keep for backward compatibility if header is English
    "Plant": "plant",
    "Recorder": "user",
    "Warranty Action": "ActionStatus",
    "Mobile": "Phone", 
    "Plantcenter": "Plantcenter",
    "Note": "Note", // Explicit map
    "Recripte": "Recripte",
    "RecripteDate": "RecripteDate",
    "Datefinish": "Datefinish"
  };
  
  const normalized = String(header).trim();
  if (map[normalized]) return map[normalized];
  
  const lower = normalized.toLowerCase();
  if (lower === 'plantcenter' || lower === 'plant center') return 'Plantcenter';
  
  return map[header] || header;
}

function getRobustValue(data, key) {
    if (!key) return undefined;
    const cleanKey = String(key).replace(/\s+/g, '').toLowerCase();
    const foundKey = Object.keys(data).find(k => String(k).replace(/\s+/g, '').toLowerCase() === cleanKey);
    return foundKey ? data[foundKey] : undefined;
}