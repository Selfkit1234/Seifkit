const ITEMS_SHEET_NAME = 'Items';
const BORROW_HISTORY_SHEET_NAME = 'BorrowHistory';
const USERS_SHEET_NAME = 'Users';

// แคชตัวแปร Spreadsheet ไว้ใช้ร่วมกัน
let _activeSS = null;
function getActiveSS() { 
  if (!_activeSS) _activeSS = SpreadsheetApp.getActiveSpreadsheet();
  return _activeSS; 
}

// ==========================================
// ⚡ ฟังก์ชันดักจับการแก้ไขอัตโนมัติ (ไม่ต้องกดซิงค์เอง)
// ==========================================
function onEdit(e) {
  try {
    const range = e.range;
    const sheet = range.getSheet();
    const sheetName = sheet.getName();
    
    // ระบบจะทำงานอัตโนมัติ เฉพาะเมื่อมีการแก้ข้อมูลในชีต BorrowHistory เท่านั้น
    if (sheetName === BORROW_HISTORY_SHEET_NAME) {
      autoSyncOldData(true); // สั่งซิงค์ข้อมูลแบบเงียบๆ (Silent Mode)
    }
  } catch(err) {
    console.log("onEdit Error: " + err.toString());
  }
}

// ==========================================
// 1. จัดการข้อมูลชีตเบื้องต้น
// ==========================================
function initializeSheets() {
  const ss = getActiveSS();
  
  let itemsSheet = ss.getSheetByName(ITEMS_SHEET_NAME);
  if (!itemsSheet) {
    itemsSheet = ss.insertSheet(ITEMS_SHEET_NAME);
    itemsSheet.getRange(1, 1, 1, 11).setValues([[
      'ID', 'Name', 'Description', 'Image', 'InitialStock', 'AvailableQuantity', 'QR code', 'รูปภาพ', 'Consumed', 'CurrentTotal', 'Borrowed'
    ]]).setFontWeight('bold');
  }
  
  let historySheet = ss.getSheetByName(BORROW_HISTORY_SHEET_NAME);
  if (!historySheet) {
    historySheet = ss.insertSheet(BORROW_HISTORY_SHEET_NAME);
    // แก้ไขข้อผิดพลาดของวงเล็บ ] เกินตรงนี้เรียบร้อยแล้ว
    historySheet.getRange(1, 1, 1, 13).setValues([
      ['HistoryID', 'BorrowerName', 'ItemID', 'ItemName', 'BorrowQty', 'BorrowDate', 'ExpectedReturnDate', 'Status', 'BorrowingGroupID', 'ReturnedQty', 'ConsumedQty', 'ActualReturnDate', 'ReturnerName']
    ]).setFontWeight('bold');
    
    historySheet.getRange("F:F").setNumberFormat("yyyy-MM-dd HH:mm:ss");
    historySheet.getRange("G:G").setNumberFormat("yyyy-MM-dd");
    historySheet.getRange("L:L").setNumberFormat("yyyy-MM-dd HH:mm:ss");
  }
  
  let usersSheet = ss.getSheetByName(USERS_SHEET_NAME);
  if (!usersSheet) {
    usersSheet = ss.insertSheet(USERS_SHEET_NAME);
    usersSheet.getRange(1, 1, 1, 4).setValues([['ID', 'Name', 'Email', 'Role']]).setFontWeight('bold');
  }
}

// ==========================================
// 🌟 2. API ROUTER FOR GITHUB PAGES
// ==========================================
function doGet(e) {
  return ContentService.createTextOutput("Backend System is running smoothly. Please use POST method from GitHub Pages.").setMimeType(ContentService.MimeType.TEXT);
}

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return ContentService.createTextOutput(JSON.stringify({ success: false, error: "No data provided" })).setMimeType(ContentService.MimeType.JSON);
    }

    const requestData = JSON.parse(e.postData.contents);
    const action = requestData.action;
    let result;

    switch (action) {
      case 'getItems': result = getItems(); break;
      case 'processBorrow': result = processBorrowLogic(requestData.payload); break;
      case 'getBorrowedItems': result = searchBorrowedItems(requestData.query || ""); break;
      case 'returnItems': result = returnMultipleItems(requestData.payload); break;
      case 'getStats': result = getStatistics(); break;
      case 'getAllData':
        const fastItems = getItems();
        const fastHistory = getBorrowHistory();
        const fastStats = getStatistics(fastItems, fastHistory);
        result = {
          items: fastItems, history: fastHistory, stats: fastStats,
          borrowedItems: Array.isArray(fastHistory) ? fastHistory.filter(item => item.status === 'borrowed') : []
        };
        break;
      case 'addItem': result = addItem(requestData.payload); break;
      case 'updateItem': result = updateItem(requestData.payload); break;
      case 'deleteItem': result = deleteItem(requestData.id); break;
      default: result = { success: false, error: 'ไม่พบ Action (' + action + ') ที่ระบุ' };
    }

    return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: error.toString() })).setMimeType(ContentService.MimeType.JSON);
  }
}

// ==========================================
// 3. ระบบจัดการอุปกรณ์ (Items)
// ==========================================
function getItems() {
  try {
    const ss = getActiveSS();
    const sheet = ss.getSheetByName(ITEMS_SHEET_NAME);
    if (!sheet) return { error: "Sheet not found" };

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];

    const defaultImg = getDefaultImage();
    const items = [];
    const len = data.length;

    for (let i = 1; i < len; i++) {
      const row = data[i];
      if (row[0] === '') continue;
      
      const initialStock = Number(row[4]) || 0; 
      const availableQty = Number(row[5]) || 0; 
      const qrCode = row[6] || "";              
      const extraImage = row[7] || "";          
      const consumed = Number(row[8]) || 0;     
      const currentTotal = Number(row[9]) || 0; 
      const borrowed = Number(row[10]) || 0;    
      
      items.push({
        id: row[0],
        name: row[1],
        description: row[2],
        image: row[3] || defaultImg,
        initialQuantity: initialStock, 
        availableQuantity: availableQty, 
        available: availableQty > 0,
        qrCode: qrCode,                
        extraImage: extraImage,        
        consumed: consumed,            
        totalQuantity: currentTotal,   
        borrowed: borrowed             
      });
    }
    return items;
  } catch (error) { return { error: `Error in getItems: ${error.message}` }; }
}

function addItem(itemData) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = getActiveSS();
    const sheet = ss.getSheetByName(ITEMS_SHEET_NAME);
    const newId = getNextId(sheet);
    const qty = Number(itemData.quantity) || 1;
    
    sheet.appendRow([ 
      newId, itemData.name, itemData.description, itemData.image || getDefaultImage(), 
      qty, qty, itemData.qrCode || "", itemData.extraImage || "", 0, qty, 0 
    ]);
    return { success: true };
  } catch (error) { return { success: false, error: error.toString() }; } finally { lock.releaseLock(); }
}

function updateItem(itemData) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    if (!itemData || !itemData.id) return { success: false, error: 'ข้อมูลไม่ถูกต้อง' };
    
    const ss = getActiveSS();
    const sheet = ss.getSheetByName(ITEMS_SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(itemData.id)) {
        const currentConsumed = Number(data[i][8]) || 0; 
        const currentBorrowed = Number(data[i][10]) || 0; 
        
        const newInitial = Number(itemData.quantity) || 1; 
        const newCurrentTotal = newInitial - currentConsumed; 
        const newAvailable = newCurrentTotal - currentBorrowed; 
        
        if (newAvailable < 0) return { success: false, error: 'ไม่สามารถลดจำนวนลงได้ เพราะทำให้ยอดพร้อมให้ยืมติดลบ' };
        
        sheet.getRange(i + 1, 2, 1, 10).setValues([[
          itemData.name, 
          itemData.description, 
          itemData.image || getDefaultImage(),
          newInitial, 
          newAvailable,
          data[i][6], 
          data[i][7], 
          currentConsumed,
          newCurrentTotal,
          currentBorrowed
        ]]);
        return { success: true };
      }
    }
    return { success: false, error: 'ไม่พบอุปกรณ์ที่ต้องการแก้ไข' };
  } catch (error) { return { success: false, error: error.toString() }; } finally { lock.releaseLock(); }
}

function deleteItem(itemId) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = getActiveSS();
    const sheet = ss.getSheetByName(ITEMS_SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(itemId)) {
        if (Number(data[i][10]) > 0) { 
            return { success: false, error: 'ไม่สามารถลบได้เนื่องจากอุปกรณ์นี้มีคนกำลังยืมอยู่ กรุณารอให้คืนครบก่อน' };
        }
        sheet.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: false, error: 'ไม่พบอุปกรณ์ที่ต้องการลบ' };
  } catch (error) { return { success: false, error: error.toString() }; } finally { lock.releaseLock(); }
}

// ==========================================
// 4. ระบบประวัติ (History) - [เชื่อมโยงชื่ออุปกรณ์จาก Items]
// ==========================================
function formatDateNative(d) {
  if (!d || !(d instanceof Date) || isNaN(d.getTime())) return String(d || '');
  return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
}

function formatDateTimeNative(d) {
  if (!d || !( d instanceof Date) || isNaN(d.getTime())) return String(d || '');
  return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0') + ' ' +
         String(d.getHours()).padStart(2, '0') + ':' + String(d.getMinutes()).padStart(2, '0') + ':' + String(d.getSeconds()).padStart(2, '0');
}

function getBorrowHistory() {
  try {
    const ss = getActiveSS();
    const sheet = ss.getSheetByName(BORROW_HISTORY_SHEET_NAME);
    const itemsSheet = ss.getSheetByName(ITEMS_SHEET_NAME);
    if (!sheet || !itemsSheet) return [];
    
    // สร้างดิกชันนารีเก็บชื่ออุปกรณ์ล่าสุดโดยใช้ ID เป็น Key เพื่อแก้ปัญหารายการไม่ตรงกัน
    const itemsData = itemsSheet.getDataRange().getValues();
    const itemDict = {};
    for (let i = 1; i < itemsData.length; i++) {
      const id = String(itemsData[i][0]).trim();
      if (id) itemDict[id] = itemsData[i][1];
    }
    
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    
    const history = [];
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const historyId = row[0];
      if (historyId === null || historyId === '' || isNaN(Number(historyId))) continue;
      
      let bQty = row[4];
      if (bQty instanceof Date) { bQty = 1; } else { bQty = Number(bQty); if (isNaN(bQty) || bQty <= 0) bQty = 1; }
      
      const itemId = String(row[2]).trim();
      const liveItemName = itemDict[itemId] || row[3]; // ใช้ชื่ออัปเดตล่าสุดจากหน้า Items เสมอ
      const statusRaw = String(row[7] || '').trim().toLowerCase(); 

      history.push({
        historyId: Number(historyId), 
        borrowerName: row[1] ? String(row[1]).trim() : '', 
        itemId: Number(itemId), 
        itemName: liveItemName, 
        borrowQty: bQty, 
        borrowDate: formatDateTimeNative(row[5]), 
        returnDate: formatDateNative(row[6]), 
        status: statusRaw,
        borrowingGroupId: row[8] || '', 
        returnedQty: Number(row[9] || 0), 
        consumedQty: Number(row[10] || 0),
        actualReturnDate: formatDateTimeNative(row[11]), 
        returnerName: row[12] || '' 
      });
    }
    return history.reverse();
  } catch (error) { return { error: `Error: ${error.message}` }; }
}

function searchBorrowedItems(query) {
  try {
    const history = getBorrowHistory();
    if (history.error) return history;
    
    let borrowedItems = history.filter(item => item.status === 'borrowed');
    
    if (query) {
      const lowerQuery = String(query).toLowerCase().trim();
      borrowedItems = borrowedItems.filter(item => 
        (item.itemName && String(item.itemName).toLowerCase().includes(lowerQuery)) || 
        (item.borrowerName && String(item.borrowerName).toLowerCase().includes(lowerQuery)) ||
        (String(item.itemId) === lowerQuery)
      );
    }
    return borrowedItems;
  } catch (error) { return []; }
}

// ==========================================
// 5. ระบบบันทึกการ ยืม - คืน 
// ==========================================
function processBorrowLogic(borrowData) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = getActiveSS();
    const itemsSheet = ss.getSheetByName(ITEMS_SHEET_NAME);
    const historySheet = ss.getSheetByName(BORROW_HISTORY_SHEET_NAME);
    const borrowingGroupId = 'GRP-' + new Date().getTime();
    
    const itemsToBorrow = borrowData.items; 
    if (!itemsToBorrow || itemsToBorrow.length === 0) return { success: false, error: "ไม่พบข้อมูลรายการ" };

    const borrowerName = borrowData.borrowerName;
    const borrowDateObj = new Date(); 
    let returnDateObj = borrowData.returnDate ? new Date(borrowData.returnDate) : new Date(); 
    
    const itemsData = itemsSheet.getDataRange().getValues();
    const itemsToBorrowInfo = [];

    for (let j = 0; j < itemsToBorrow.length; j++) {
      const reqId = Number(itemsToBorrow[j].id);
      const reqQty = Number(itemsToBorrow[j].qty || 1);
      
      let found = false;
      for (let i = 1; i < itemsData.length; i++) {
        if (Number(itemsData[i][0]) === reqId) {
          found = true;
          const currentAvailable = Number(itemsData[i][5]); 
          const currentBorrowed = Number(itemsData[i][10]); 
          
          if (currentAvailable < reqQty) return { success: false, error: `ของ "${itemsData[i][1]}" มีไม่พอ` };
          
          itemsData[i][5] = currentAvailable - reqQty;
          itemsData[i][10] = currentBorrowed + reqQty;
          
          itemsToBorrowInfo.push({ 
            rowIndex: i + 1, id: reqId, name: itemsData[i][1], borrowQty: reqQty, 
            newAvailableQty: currentAvailable - reqQty, newBorrowed: currentBorrowed + reqQty 
          });
          break;
        }
      }
      if (!found) return { success: false, error: `ไม่พบสินค้ารหัส ${reqId}` };
    }

    itemsToBorrowInfo.forEach(item => { 
      itemsSheet.getRange(item.rowIndex, 6).setValue(item.newAvailableQty);
      itemsSheet.getRange(item.rowIndex, 11).setValue(item.newBorrowed);
    });

    const newHistoryRows = [];
    const nextHistoryId = getNextId(historySheet);
    for (let i = 0; i < itemsToBorrowInfo.length; i++) {
      const item = itemsToBorrowInfo[i];
      newHistoryRows.push([
        nextHistoryId + i, borrowerName, item.id, item.name, item.borrowQty, 
        borrowDateObj, returnDateObj, 'borrowed', borrowingGroupId, 0, 0, '', '' 
      ]);
    }

    if (newHistoryRows.length > 0) {
        historySheet.getRange(historySheet.getLastRow() + 1, 1, newHistoryRows.length, newHistoryRows[0].length).setValues(newHistoryRows);
        SpreadsheetApp.flush(); 
        
        const emails = getAdminEmails(ss);
        if (emails.length > 0) {
            let toEmails = emails.join(',');
            let emailSubject = `แจ้งเตือน: มีการยืมอุปกรณ์โดย ${borrowerName}`;
            let emailBody = `ระบบยืม-คืนอุปกรณ์\n\nมีการทำรายการยืมใหม่:\nชื่อผู้ยืม: ${borrowerName}\nวันที่ยืม: ${borrowData.borrowDate || formatDateTimeNative(new Date())}\nกำหนดคืน: ${borrowData.returnDate || '-'}\n\nรายการที่ยืม:\n`;
            itemsToBorrowInfo.forEach((item, index) => { emailBody += `${index + 1}. ${item.name} (จำนวน: ${item.borrowQty})\n`; });
            try { MailApp.sendEmail(toEmails, emailSubject, emailBody); } catch(e) {}
        }
    }
    return { success: true };
  } catch (error) { return { success: false, error: error.toString() }; } finally { lock.releaseLock(); }
}

function returnMultipleItems(returnPayload) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const returnItemsData = returnPayload.items;
    const returnerName = returnPayload.returnerName || '';

    if (!returnItemsData || returnItemsData.length === 0) return { success: false, error: 'ไม่พบรายการที่ต้องการคืน' };

    const ss = getActiveSS();
    const itemsSheet = ss.getSheetByName(ITEMS_SHEET_NAME);
    const historySheet = ss.getSheetByName(BORROW_HISTORY_SHEET_NAME);
    
    const historyData = historySheet.getDataRange().getValues();
    const itemsData = itemsSheet.getDataRange().getValues();
    
    const actualReturnDateObj = new Date(); 
    let successCount = 0, borrowerNameForEmail = "", returnedItemsInfo = [];

    for (let t = 0; t < returnItemsData.length; t++) {
      const returnData = returnItemsData[t];
      const targetHistoryId = Number(returnData.historyId);
      const returnedQty = Number(returnData.returnedQty || 0);
      const consumedQty = Number(returnData.consumedQty || 0);
      
      let historyRowIndex = -1, itemIdToReturn = -1, originalBorrowQty = 0, itemNameForEmail = "";
      
      for (let i = 1; i < historyData.length; i++) {
        const currentStatus = String(historyData[i][7] || '').trim().toLowerCase();
        if (Number(historyData[i][0]) === targetHistoryId && currentStatus === 'borrowed') {
          historyRowIndex = i; itemIdToReturn = Number(historyData[i][2]); itemNameForEmail = historyData[i][3]; originalBorrowQty = Number(historyData[i][4]);
          if(!borrowerNameForEmail) borrowerNameForEmail = historyData[i][1]; 
          historyData[i][7] = 'returned';
          break;
        }
      }
      
      if (historyRowIndex === -1) continue;
      
      for (let i = 1; i < itemsData.length; i++) {
        if (Number(itemsData[i][0]) === itemIdToReturn) {
          const currentInitial = Number(itemsData[i][4]); 
          const currentConsumed = Number(itemsData[i][8]); 
          const currentBorrowed = Number(itemsData[i][10]); 
          
          const newConsumed = currentConsumed + consumedQty; 
          const newCurrentTotal = currentInitial - newConsumed; 
          const newBorrowed = currentBorrowed - originalBorrowQty; 
          const newAvailable = newCurrentTotal - newBorrowed; 
          
          itemsSheet.getRange(i + 1, 6).setValue(newAvailable);
          itemsSheet.getRange(i + 1, 9, 1, 3).setValues([[newConsumed, newCurrentTotal, newBorrowed]]);
          break;
        }
      }
      
      const groupId = historyData[historyRowIndex][8]; 
      historySheet.getRange(historyRowIndex + 1, 8, 1, 6).setValues([['returned', groupId, returnedQty, consumedQty, actualReturnDateObj, returnerName]]); 
      returnedItemsInfo.push({ name: itemNameForEmail, returned: returnedQty, consumed: consumedQty });
      successCount++;
    }

    SpreadsheetApp.flush(); 

    if (successCount > 0) {
        const emails = getAdminEmails(ss);
        if (emails.length > 0) {
            let toEmails = emails.join(',');
            let emailSubject = `แจ้งเตือน: มีการคืนอุปกรณ์โดย ${returnerName || 'ไม่ระบุ'}`;
            let emailBody = `ระบบยืม-คืนอุปกรณ์\n\nผู้ทำรายการคืน: ${returnerName || 'ไม่ระบุ'}\nวันที่คืน: ${returnPayload.returnDate || formatDateTimeNative(new Date())}\n\nรายการที่คืน:\n`;
            returnedItemsInfo.forEach((item, index) => { emailBody += `${index + 1}. ${item.name} (ดี: ${item.returned}, ชำรุด/ใช้ไป: ${item.consumed})\n`; });
            try { MailApp.sendEmail(toEmails, emailSubject, emailBody); } catch(e) {}
        }
    }
    return { success: true, message: `ดำเนินการคืนสำเร็จ ${successCount} รายการ` };
  } catch (error) { return { success: false, error: error.toString() }; } finally { lock.releaseLock(); }
}

// ==========================================
// 6. สถิติ และ ฟังก์ชันช่วยเหลือ
// ==========================================
function getStatistics(preFetchedItems, preFetchedHistory) {
  try {
    const itemsResult = preFetchedItems || getItems();
    const historyResult = preFetchedHistory || getBorrowHistory();
    
    let availableCount = 0, borrowedCount = 0, totalConsumedOverall = 0, totalInitialItems = 0;
    const borrowCountByItem = {}, returnCountByItem = {}; 

    if (!historyResult.error) {
        for(let i=0; i<historyResult.length; i++) {
          const record = historyResult[i];
          const itemName = record.itemName ? record.itemName.trim() : "";
          if (itemName) { 
              borrowCountByItem[itemName] = (borrowCountByItem[itemName] || 0) + 1; 
              if (record.status === 'returned') returnCountByItem[itemName] = (returnCountByItem[itemName] || 0) + 1;
          }
        }
    }

    const itemUsageStats = [];
    if (!itemsResult.error) {
      for(let i=0; i<itemsResult.length; i++) {
          const item = itemsResult[i];
          availableCount += item.availableQuantity;
          borrowedCount += item.borrowed;
          totalConsumedOverall += item.consumed;
          totalInitialItems += item.initialQuantity;
          
          itemUsageStats.push({
              name: item.name, initialTotal: item.initialQuantity, consumed: item.consumed, 
              currentTotal: item.totalQuantity, borrowed: item.borrowed, available: item.availableQuantity,
              borrowTimes: borrowCountByItem[item.name] || 0, returnTimes: returnCountByItem[item.name] || 0  
          });
      }
    }

    const mostBorrowedItems = Object.entries(borrowCountByItem).map(([name, count]) => ({ name, count })).sort((a, b) => b.count - a.count);

    return {
      totalItems: totalInitialItems, availableItems: availableCount, borrowedItems: borrowedCount,
      totalConsumed: totalConsumedOverall, totalBorrows: historyResult.length || 0,
      mostBorrowedItems: mostBorrowedItems, itemUsageStats: itemUsageStats, rawHistory: historyResult 
    };
  } catch (error) { return { error: "ไม่สามารถดึงข้อมูลสถิติได้: " + error.message }; }
}

function getNextId(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return 1;
  const range = sheet.getRange(1, 1, lastRow, 1).getValues();
  let lastId = 0;
  for (let i = range.length - 1; i >= 0; i--) {
    if (range[i][0] !== "" && !isNaN(Number(range[i][0]))) { lastId = Number(range[i][0]); break; }
  }
  return lastId + 1;
}

function getDefaultImage() { 
  return 'data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMjAwIiBoZWlnaHQ9IjE1MCIgdmlld0JveD0iMCAwIDIwMCAxNTAiIGZpbGw9Im5vbmUiIHhtbG5zPSJodHRwOi8vd3d3LnczLm9yZy8yMDAwL3N2ZyI+PHJlY3Qgd2lkdGg9IjIwMCIgaGVpZ2h0PSIxNTAiIGZpbGw9IiNmM2Y0ZjYiLz48dGV4dCB4PSIxMDAiIHk9Ijc1IiB0ZXh0LWFuY2hvcj0ibWlkZGxlIiBmaWxsPSIjNjc3NDhkIiBmb250LXNpemU9IjE0Ij7guYTguKHguYjguKHguLXguKPguLnguJo8L3RleHQ+PC9zdmc+'; 
}

function getAdminEmails(ss) {
    let emailList = [];
    try {
        const usersSheet = ss.getSheetByName(USERS_SHEET_NAME);
        if (usersSheet) {
            const usersData = usersSheet.getDataRange().getValues();
            for (let r = 1; r < usersData.length; r++) {
                let email = usersData[r][2]; 
                let role = usersData[r][3] ? usersData[r][3].toString().toLowerCase().trim() : ''; 
                if (email && email.toString().includes('@') && role === 'admin') emailList.push(email.toString().trim());
            }
        }
    } catch(e) {}
    return emailList;
}

// ==========================================
// 🛠️ 7. ฟังก์ชันพิเศษ: ซิงค์ข้อมูลอัตโนมัติ (Bulk Write Optimization)
// ==========================================
function autoSyncOldData(isSilent) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const itemsSheet = ss.getSheetByName(ITEMS_SHEET_NAME);
  const historySheet = ss.getSheetByName(BORROW_HISTORY_SHEET_NAME);
  
  if (!itemsSheet || !historySheet) return;
  
  if (isSilent) {
    ss.toast("🔄 กำลังปรับปรุงยอดสต็อกอัตโนมัติ...", "ระบบยืม-คืน", 2);
  }
  
  const historyData = historySheet.getDataRange().getValues();
  const itemsData = itemsSheet.getDataRange().getValues();
  
  const statsTracker = {}; 
  
  for (let i = 1; i < historyData.length; i++) {
    const itemId = String(historyData[i][2]).trim();
    const status = String(historyData[i][7] || '').trim().toLowerCase(); 
    const borrowQty = Number(historyData[i][4]) || 0;
    const consumedQty = Number(historyData[i][10]) || 0; 
    
    if (!itemId) continue;
    if (!statsTracker[itemId]) {
      statsTracker[itemId] = { borrowed: 0, consumed: 0 };
    }
    
    if (status === 'borrowed') {
      statsTracker[itemId].borrowed += borrowQty; 
    } else if (status === 'returned') {
      statsTracker[itemId].consumed += consumedQty; 
    }
  }
  
  const availableCol = [];
  const statsCols = []; 
  
  for (let r = 1; r < itemsData.length; r++) {
    const itemId = String(itemsData[r][0]).trim();
    
    if (!itemId || itemId === '') {
      availableCol.push([itemsData[r][5]]);
      statsCols.push([itemsData[r][8], itemsData[r][9], itemsData[r][10]]);
      continue;
    }
    
    const initialStock = Number(itemsData[r][4]) || 0; 
    const stats = statsTracker[itemId] || { borrowed: 0, consumed: 0 };
    
    const consumed = stats.consumed;                           
    const currentTotal = initialStock - consumed;              
    const borrowed = stats.borrowed;                           
    const available = currentTotal - borrowed;                 
    
    availableCol.push([available]);
    statsCols.push([consumed, currentTotal, borrowed]);
  }
  
  if (availableCol.length > 0) {
    itemsSheet.getRange(2, 6, availableCol.length, 1).setValues(availableCol);
    itemsSheet.getRange(2, 9, statsCols.length, 3).setValues(statsCols);
  }
  
  if (isSilent) {
    ss.toast("✅ อัปเดตสต็อกเรียบร้อยแล้ว!", "ระบบยืม-คืน", 2);
  } else {
    SpreadsheetApp.getUi().alert('✅ ซิงค์ข้อมูลสถิติเก่า เรียบร้อยแล้ว!');
  }
}

// สร้างเมนูสำหรับกดแมนนวลบน Google Sheets
function onOpen() {
  SpreadsheetApp.getUi().createMenu('🛠️ จัดการระบบ')
    .addItem('🔄 ซิงค์ข้อมูลสต็อกแมนนวล', 'autoSyncOldData')
    .addToUi();
}
