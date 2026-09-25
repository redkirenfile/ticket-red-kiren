// 🔑 ฟังก์ชันสำหรับกดปุ่ม "เรียกใช้ (Run)" 1 ครั้งเพื่อกดยืนยันสิทธิ์สร้างไฟล์ใน Google Drive
function setupPermission() {
  // สร้างและลบไฟล์ทดสอบทันที เพื่อบังคับให้ Google ขอสิทธิ์ Write เข้า Google Drive
  const testFile = DriveApp.createFile("test_drive_permission.txt", "test");
  testFile.setTrashed(true);
  const ss = getSS();
  Logger.log("✅ ยืนยันสิทธิ์ Google Drive และ Google Sheets สำเร็จ 100% แล้ว!");
}

function getSS() {
  if (typeof SPREADSHEET_ID !== 'undefined' && SPREADSHEET_ID && SPREADSHEET_ID !== 'YOUR_SPREADSHEET_ID_HERE') {
    try {
      return SpreadsheetApp.openById(SPREADSHEET_ID);
    } catch(e) {}
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}
/**
 * ============================================================
 * Google Apps Script — แฟ้มคดีกิเลนแดง : สังหารหมู่เขาศูนย์ (Red Kiren File) Theater Ticket Backend (Updated & Optimized)
 * ============================================================
 */

// ใส่ ID ของ Google Sheet ของคุณโดยตรง
const SPREADSHEET_ID = '15RKxq6rzJXtN5R77cxdPZNp_JU6rDOE12ijuysTM76w';

// Sheet names
const SHEET_ORDERS   = 'Orders';
const SHEET_TICKETS  = 'Tickets';
const SHEET_CHECKINS = 'CheckIns';
const SHEET_SETTINGS = 'Settings';

// ─── MAIN HANDLER ─────────────────────────────────────────────────────────
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    
    // ตั้งค่า CORS ตอบกลับแบบ JSON
    const output = (resObj) => ContentService
      .createTextOutput(JSON.stringify(resObj))
      .setMimeType(ContentService.MimeType.JSON);

    // ป้องกันปัญหาเบอร์โทรเลข 0 หายเมื่อบันทึก: เพิ่มเครื่องหมาย ' นำหน้า
    if (data.phone) {
      data.phone = formatPhoneToWrite(data.phone);
    }

    // 1. จัดการอัปเดตการตั้งค่าส่วนกลาง (เช่น เปิด/ปิด Early Bird)
    if (data.action === 'updateSetting') {
      const ss = getSS();
      const settingsSheet = getOrCreateSettings(ss);
      const key = data.key;
      const value = String(data.value);
      
      const sData = settingsSheet.getDataRange().getValues();
      let foundRow = -1;
      for (let i = 1; i < sData.length; i++) {
        if (sData[i][0] === key) {
          foundRow = i + 1;
          break;
        }
      }
      if (foundRow >= 0) {
        settingsSheet.getRange(foundRow, 2).setValue(value);
      } else {
        settingsSheet.appendRow([key, value]);
      }
      return output({ success: true });
    }

    // 2. จัดการบันทึกการเช็คอิน / ยกเลิกเช็คอิน / กู้คืนตั๋ว รายใบ
    if (data.action === 'checkin' || data.action === 'undoCheckin' || data.action === 'restoreTicket') {
      const ss = getSS();
      const ticketsSheet = getOrCreateSheet(ss, SHEET_TICKETS, []);
      const ticketId = data.ticketId;
      
      const tData = ticketsSheet.getDataRange().getValues();
      const tHeaders = tData[0] || [];
      const idxTId = findColIndex(tHeaders, ['รหัสบัตร', 'ticketid', 'ticket id']) + 1;
      const idxTStatus = findColIndex(tHeaders, ['สถานะเช็คอิน', 'สถานะ', 'status']) + 1;
      const idxTTime = findColIndex(tHeaders, ['เวลาเช็คอิน', 'time']) + 1;

      let foundRow = -1;
      for (let i = 1; i < tData.length; i++) {
        if (String(tData[i][idxTId - 1]).trim() === String(ticketId).trim()) {
          foundRow = i + 1;
          break;
        }
      }

      if (foundRow >= 0) {
        let statusVal = 'ยังไม่เช็คอิน';
        let timeVal = '';
        if (data.action === 'checkin') {
          statusVal = 'เช็คอินแล้ว';
          timeVal = data.checkInTime || new Date().toLocaleString('th-TH', { timeZone: 'Asia/Bangkok' });
          
          // บันทึกลง Log CheckIns
          const checkinsSheet = getOrCreateSheet(ss, SHEET_CHECKINS, ['รหัสบัตร', 'ชื่อ-นามสกุล', 'เบอร์โทร', 'ประเภทบัตร', 'เวลาเช็คอิน']);
          checkinsSheet.appendRow([
            ticketId,
            data.name || '',
            formatPhoneToWrite(data.phone),
            data.type || '',
            timeVal
          ]);
        }

        ticketsSheet.getRange(foundRow, idxTStatus).setValue(statusVal);
        ticketsSheet.getRange(foundRow, idxTTime).setValue(timeVal);
        
        // ใส่สีสถานะเพื่อให้ดูในชีทง่ายขึ้น
        const statusCell = ticketsSheet.getRange(foundRow, idxTStatus);
        if (statusVal === 'เช็คอินแล้ว') statusCell.setBackground('#d4edda');
        else statusCell.setBackground('#ffffff');

        return output({ success: true });
      }
      return output({ success: false, error: 'ไม่พบรหัสบัตร ' + ticketId });
    }

    // 2.1 จัดการยกเลิกตั๋ว / ลบตั๋วถาวร (ให้แถวหายไปจาก Google Sheets ทันทีตามต้องการ)
    if (data.action === 'cancelTicket' || data.action === 'deleteTicket') {
      const ss = getSS();
      const ticketsSheet = getOrCreateSheet(ss, SHEET_TICKETS, []);
      const ticketId = data.ticketId;
      let orderId = data.orderId;

      // 1. ลบแถวใน Tickets sheet
      const tData = ticketsSheet.getDataRange().getValues();
      const tHeaders = tData[0] || [];
      const idxTId = findColIndex(tHeaders, ['รหัสบัตร', 'ticketid', 'ticket id']);
      const idxTOId = findColIndex(tHeaders, ['เลขที่คำสั่งซื้อ', 'orderid', 'order id']);

      if (idxTId >= 0) {
        for (let i = tData.length - 1; i >= 1; i--) {
          if (String(tData[i][idxTId]).trim() === String(ticketId).trim()) {
            if (!orderId && idxTOId >= 0) {
              orderId = tData[i][idxTOId];
            }
            ticketsSheet.deleteRow(i + 1);
            break;
          }
        }
      }

      // 2. ปรับลดจำนวนตั๋วใน Orders sheet หรือลบออร์เดอร์หากไม่เหลือตั๋ว
      if (orderId) {
        const ordersSheet = getOrCreateSheet(ss, SHEET_ORDERS, []);
        const oData = ordersSheet.getDataRange().getValues();
        const oHeaders = oData[0] || [];
        const idxOId = findColIndex(oHeaders, ['เลขที่คำสั่งซื้อ', 'orderid', 'order id']);
        const idxOQty = findColIndex(oHeaders, ['จำนวนใบ', 'จำนวนบัตร (ใบ)', 'จำนวนบัตร', 'จำนวน', 'qty']);
        const idxOTotal = findColIndex(oHeaders, ['ราคารวม', 'ยอดเงินรวม (บาท)', 'ยอดรวม', 'total', 'amount']);
        const idxOPrice = findColIndex(oHeaders, ['ราคาต่อใบ', 'ราคาต่อใบ (บาท)', 'ราคา', 'priceperticket', 'price']);
        const idxOTkts = findColIndex(oHeaders, ['รหัสบัตรทั้งหมด', 'รหัสตั๋วทั้งหมด', 'ticketids', 'tickets']);

        if (idxOId >= 0) {
          for (let i = oData.length - 1; i >= 1; i--) {
            if (String(oData[i][idxOId]).trim() === String(orderId).trim()) {
              const rowNum = i + 1;
              let currentTkts = [];
              if (idxOTkts >= 0) {
                currentTkts = String(oData[i][idxOTkts] || '').split(',').map(s => s.trim()).filter(Boolean);
                currentTkts = currentTkts.filter(id => id !== ticketId);
              }
              const oldQty = idxOQty >= 0 ? Number(oData[i][idxOQty]) || 0 : 0;
              const newQty = currentTkts.length > 0 ? currentTkts.length : Math.max(0, oldQty - 1);

              if (newQty <= 0) {
                ordersSheet.deleteRow(rowNum);
              } else {
                const pricePer = idxOPrice >= 0 ? Number(oData[i][idxOPrice]) || 0 : 0;
                if (idxOQty >= 0) ordersSheet.getRange(rowNum, idxOQty + 1).setValue(newQty);
                if (idxOTotal >= 0 && pricePer > 0) ordersSheet.getRange(rowNum, idxOTotal + 1).setValue(newQty * pricePer);
                if (idxOTkts >= 0) ordersSheet.getRange(rowNum, idxOTkts + 1).setValue(currentTkts.join(', '));
              }
              break;
            }
          }
        }
      }
      return output({ success: true, message: 'ลบตั๋วออกจาก Google Sheets เรียบร้อยแล้ว' });
    }

    // 3. จัดการยกเลิกคำสั่งซื้อ / ลบคำสั่งซื้อถาวร (ลบตั๋วทุกใบและออร์เดอร์ออกจาก Google Sheets ทันที)
    if (data.action === 'cancelOrder' || data.action === 'deleteOrder') {
      const ss = getSS();
      const ordersSheet = getOrCreateSheet(ss, SHEET_ORDERS, []);
      const ticketsSheet = getOrCreateSheet(ss, SHEET_TICKETS, []);
      const orderId = data.orderId;

      // 1. ลบตั๋วทั้งหมดของออร์เดอร์นี้ใน Tickets sheet
      const tData = ticketsSheet.getDataRange().getValues();
      const tHeaders = tData[0] || [];
      const idxTOId = findColIndex(tHeaders, ['เลขที่คำสั่งซื้อ', 'orderid', 'order id']);

      if (idxTOId >= 0) {
        for (let i = tData.length - 1; i >= 1; i--) {
          if (String(tData[i][idxTOId]).trim() === String(orderId).trim()) {
            ticketsSheet.deleteRow(i + 1);
          }
        }
      }

      // 2. ลบแถวใน Orders sheet
      const oData = ordersSheet.getDataRange().getValues();
      const oHeaders = oData[0] || [];
      const idxOId = findColIndex(oHeaders, ['เลขที่คำสั่งซื้อ', 'orderid', 'order id']);

      if (idxOId >= 0) {
        for (let i = oData.length - 1; i >= 1; i--) {
          if (String(oData[i][idxOId]).trim() === String(orderId).trim()) {
            ordersSheet.deleteRow(i + 1);
            break;
          }
        }
      }
      return output({ success: true, message: 'ลบออร์เดอร์ออกจาก Google Sheets เรียบร้อยแล้ว' });
    }

    // 4. จัดการแก้ไขข้อมูลลูกค้า (จาก Staff Edit Modal)
    if (data.action === 'saveEdit') {
      const ss = getSS();
      const ordersSheet = getOrCreateSheet(ss, SHEET_ORDERS, []);
      const ticketsSheet = getOrCreateSheet(ss, SHEET_TICKETS, []);
      const orderId = data.orderId;
      const cleanPhoneVal = formatPhoneToWrite(data.phone);
      
      // อัปเดตใน Orders sheet
      const oData = ordersSheet.getDataRange().getValues();
      const oHeaders = oData[0] || [];
      const idxOId = oHeaders.indexOf('เลขที่คำสั่งซื้อ') + 1;
      const idxOName = oHeaders.indexOf('ชื่อ-นามสกุล') + 1;
      const idxOPhone = oHeaders.indexOf('เบอร์โทร') + 1;
      const idxOEmail = oHeaders.indexOf('อีเมล') + 1;
      const idxOType = oHeaders.indexOf('ประเภทบัตร') + 1;
      const idxOPrice = oHeaders.indexOf('ราคาต่อใบ') + 1;
      const idxOTotal = oHeaders.indexOf('ราคารวม') + 1;
      const idxONote = oHeaders.indexOf('หมายเหตุ') + 1;
      const idxOShowDate = getColumnIndex(ordersSheet, 'รอบการแสดง');

      for (let i = 1; i < oData.length; i++) {
        if (oData[i][idxOId - 1] === orderId) {
          const rowNum = i + 1;
          ordersSheet.getRange(rowNum, idxOName).setValue(data.name);
          ordersSheet.getRange(rowNum, idxOPhone).setValue(cleanPhoneVal);
          if (idxOEmail) ordersSheet.getRange(rowNum, idxOEmail).setValue(data.email || '');
          ordersSheet.getRange(rowNum, idxOType).setValue(data.type);
          if (idxOPrice) ordersSheet.getRange(rowNum, idxOPrice).setValue(data.pricePerTicket);
          if (idxOTotal) ordersSheet.getRange(rowNum, idxOTotal).setValue(data.total);
          if (idxONote) ordersSheet.getRange(rowNum, idxONote).setValue(data.note || '');
          if (idxOShowDate) ordersSheet.getRange(rowNum, idxOShowDate).setValue(data.showDate || '');
          break;
        }
      }

      // อัปเดตใน Tickets sheet ทุกใบของออร์เดอร์นี้
      const tData = ticketsSheet.getDataRange().getValues();
      const tHeaders = tData[0] || [];
      const idxTOId = tHeaders.indexOf('เลขที่คำสั่งซื้อ') + 1;
      const idxTName = tHeaders.indexOf('ชื่อ-นามสกุล') + 1;
      const idxTPhone = tHeaders.indexOf('เบอร์โทร') + 1;
      const idxTType = tHeaders.indexOf('ประเภทบัตร') + 1;
      const idxTShowDate = getColumnIndex(ticketsSheet, 'รอบการแสดง');

      for (let i = 1; i < tData.length; i++) {
        if (tData[i][idxTOId - 1] === orderId) {
          const rowNum = i + 1;
          ticketsSheet.getRange(rowNum, idxTName).setValue(data.name);
          ticketsSheet.getRange(rowNum, idxTPhone).setValue(cleanPhoneVal);
          ticketsSheet.getRange(rowNum, idxTType).setValue(data.type);
          if (idxTShowDate) ticketsSheet.getRange(rowNum, idxTShowDate).setValue(data.showDate || '');
        }
      }

      return output({ success: true });
    }

    // 5. บันทึกคำสั่งซื้อใหม่ (จองตั๋วปกติ หรือ โควต้าสตาฟ)
    return handleNewOrder(data);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Handler สำหรับดึงข้อมูลแบบ GET
function doGet(e) {
  const output = (resObj) => ContentService
    .createTextOutput(JSON.stringify(resObj))
    .setMimeType(ContentService.MimeType.JSON);

  try {
    const action = e.parameter.action;
    const ss = getSS();

    // 1. ดึงเฉพาะการตั้งค่าส่วนกลาง (เช่น เช็คสถานะ Early Bird ในหน้าจองลูกค้า + ยอดจองกลางเพื่อคำนวณที่นั่งเหลือ)
    if (action === 'getSettings') {
      const settingsSheet = getOrCreateSettings(ss);
      const data = settingsSheet.getDataRange().getValues();
      const settings = {};
      for (let i = 1; i < data.length; i++) {
        const key = data[i][0];
        let val = data[i][1];
        if (val === 'true') val = true;
        if (val === 'false') val = false;
        settings[key] = val;
      }

      // ดึงยอดขายสะสมเพื่อคืนกลับไปให้ฝั่งผู้ซื้อคำนวณที่นั่งคงเหลือตรงกัน (Central Stock)
      const ticketsSheet = getOrCreateSheet(ss, SHEET_TICKETS, []);
      const tData = ticketsSheet.getDataRange().getValues();
      const tHeaders = tData[0] || [];
      const idxTStatus = tHeaders.indexOf('สถานะเช็คอิน') + 1;
      const idxTShowDate = tHeaders.indexOf('รอบการแสดง') + 1;
      const idxTType = tHeaders.indexOf('ประเภทบัตร') + 1;

      const soldCounts = {};
      const soldCountsByType = {};
      for (let i = 1; i < tData.length; i++) {
        const row = tData[i];
        const status = row[idxTStatus - 1];
        const showDate = row[idxTShowDate - 1];
        const rawType = idxTType > 0 ? String(row[idxTType - 1] || '').toUpperCase() : '';
        if (showDate && status !== 'ยกเลิกแล้ว') {
          soldCounts[showDate] = (soldCounts[showDate] || 0) + 1;
          const typeKey = rawType.includes('EARLY') ? 'earlybird' :
                          rawType.includes('STUDENT') ? 'student' :
                          rawType.includes('REGULAR') ? 'regular' : 'other';
          const slotTypeKey = `${showDate}|${typeKey}`;
          soldCountsByType[slotTypeKey] = (soldCountsByType[slotTypeKey] || 0) + 1;
        }
      }
      settings["soldCounts"] = soldCounts;
      settings["soldCountsByType"] = soldCountsByType;

      return output(settings);
    }
    
    // 2. ดึงตั๋วและประวัติทั้งหมดเพื่อทำ Cloud Sync ในหลังบ้าน (Staff)
    if (action === 'getAll') {
      const settingsSheet = getOrCreateSettings(ss);
      const sData = settingsSheet.getDataRange().getValues();
      const settings = {};
      for (let i = 1; i < sData.length; i++) {
        const key = sData[i][0];
        let val = sData[i][1];
        if (val === 'true') val = true;
        if (val === 'false') val = false;
        settings[key] = val;
      }

      const ticketsSheet = getOrCreateSheet(ss, SHEET_TICKETS, []);
      const tData = ticketsSheet.getDataRange().getValues();
      const tickets = {};
      
      const ordersSheet = getOrCreateSheet(ss, SHEET_ORDERS, []);
      const oData = ordersSheet.getDataRange().getValues();
      const orders = [];

      // ค้นหาตำแหน่งคอลัมน์ของชีทคำสั่งซื้อ (Orders)
      const oHeaders = oData[0] || [];
      const idxOId = findColIndex(oHeaders, ['เลขที่คำสั่งซื้อ', 'เลขที่ออเดอร์', 'เลขออเดอร์', 'orderid', 'order id']) + 1;
      const idxOTs = findColIndex(oHeaders, ['วันเวลา', 'เวลา', 'วันที่', 'timestamp', 'date']) + 1;
      const idxOName = findColIndex(oHeaders, ['ชื่อ-นามสกุล', 'ชื่อนามสกุล', 'ชื่อ', 'name', 'fullname']) + 1;
      const idxONick = findColIndex(oHeaders, ['ชื่อเล่น', 'nickname', 'nick']) + 1;
      const idxOPhone = findColIndex(oHeaders, ['เบอร์โทร', 'เบอร์โทรศัพท์', 'เบอร์', 'phone', 'tel']) + 1;
      const idxOEmail = findColIndex(oHeaders, ['อีเมล', 'email']) + 1;
      const idxOType = findColIndex(oHeaders, ['ประเภทบัตร', 'ประเภท', 'tickettype', 'type']) + 1;
      const idxOQty = findColIndex(oHeaders, ['จำนวนใบ', 'จำนวน', 'qty', 'quantity']) + 1;
      const idxOPrice = findColIndex(oHeaders, ['ราคาต่อใบ', 'ราคา', 'priceperticket', 'price']) + 1;
      const idxOTotal = findColIndex(oHeaders, ['ราคารวม', 'ยอดรวม', 'total', 'amount']) + 1;
      const idxONote = findColIndex(oHeaders, ['หมายเหตุ', 'note']) + 1;
      const idxOTickets = findColIndex(oHeaders, ['รหัสบัตรทั้งหมด', 'รหัสบัตร', 'ticketids', 'tickets']) + 1;
      const idxOSlip = findColIndex(oHeaders, ['สลิปการโอนเงิน', 'สลิปโอนเงิน', 'สลิป', 'หลักฐานการโอน', 'slipurl', 'slip url', 'slip_url', 'slip', 'sliplink']) + 1;
      const idxOShowDate = findColIndex(oHeaders, ['รอบการแสดง', 'รอบ', 'showdate', 'round']) + 1;

      // ค้นหาตำแหน่งคอลัมน์ของชีทตั๋วรายใบ (Tickets)
      const tHeaders = tData[0] || [];
      const idxTId = findColIndex(tHeaders, ['รหัสบัตร', 'ticketid', 'ticket id']) + 1;
      const idxTOId = findColIndex(tHeaders, ['เลขที่คำสั่งซื้อ', 'orderid', 'order id']) + 1;
      const idxTName = findColIndex(tHeaders, ['ชื่อ-นามสกุล', 'ชื่อ', 'name']) + 1;
      const idxTNick = findColIndex(tHeaders, ['ชื่อเล่น', 'nickname', 'nick']) + 1;
      const idxTPhone = findColIndex(tHeaders, ['เบอร์โทร', 'phone', 'tel']) + 1;
      const idxTType = findColIndex(tHeaders, ['ประเภทบัตร', 'type']) + 1;
      const idxTStatus = findColIndex(tHeaders, ['สถานะเช็คอิน', 'สถานะ', 'status']) + 1;
      const idxTTime = findColIndex(tHeaders, ['เวลาเช็คอิน', 'time']) + 1;
      const idxTShowDate = findColIndex(tHeaders, ['รอบการแสดง', 'showdate']) + 1;
      const idxTSlip = findColIndex(tHeaders, ['สลิปการโอนเงิน', 'สลิป', 'slipurl', 'slip url', 'slip']) + 1;

      // สร้าง Map แหล่งรวมคำสั่งซื้อ
      const ordersMap = {};
      for (let i = 1; i < oData.length; i++) {
        const row = oData[i];
        if (!row[idxOId - 1]) continue;
        const slipVal = (idxOSlip ? row[idxOSlip - 1] : '') || '';
        const order = {
          orderId: row[idxOId - 1],
          timestamp: row[idxOTs - 1],
          name: row[idxOName - 1],
          nickname: idxONick ? (row[idxONick - 1] || '') : '',
          phone: readCleanPhone(row[idxOPhone - 1]),
          email: idxOEmail ? row[idxOEmail - 1] : '',
          typeName: row[idxOType - 1],
          qty: Number(row[idxOQty - 1] || 0),
          pricePerTicket: Number(row[idxOPrice - 1] || 0),
          total: Number(row[idxOTotal - 1] || 0),
          note: idxONote ? row[idxONote - 1] : '',
          ticketIds: idxOTickets ? (row[idxOTickets - 1] || '').split(', ') : [],
          slipUrl: slipVal,
          slipImage: slipVal,
          showDate: idxOShowDate ? row[idxOShowDate - 1] : '',
          cancelled: false
        };
        orders.push(order);
        ordersMap[order.orderId] = order;
      }

      // แมปตั๋วรายใบ
      for (let i = 1; i < tData.length; i++) {
        const row = tData[i];
        const tId = row[idxTId - 1];
        if (!tId) continue;
        const oId = row[idxTOId - 1];
        const parentOrder = ordersMap[oId] || {};
        
        const status = row[idxTStatus - 1];
        const isCancelled = status === 'ยกเลิกแล้ว';
        const isCheckedIn = status === 'เช็คอินแล้ว';

        let ticketNum = 1;
        const match = tId.match(/-T(\d+)$/);
        if (match) ticketNum = Number(match[1]);

        const tSlip = (idxTSlip ? row[idxTSlip - 1] : '') || parentOrder.slipUrl || parentOrder.slipImage || '';

        tickets[tId] = {
          ticketId: tId,
          ticketNum: ticketNum,
          orderId: oId,
          name: row[idxTName - 1] || parentOrder.name || '',
          nickname: (idxTNick ? row[idxTNick - 1] : '') || parentOrder.nickname || '',
          phone: readCleanPhone(row[idxTPhone - 1]) || parentOrder.phone || '',
          email: parentOrder.email || '',
          note: parentOrder.note || '',
          type: row[idxTType - 1] || parentOrder.typeName || '',
          showDate: (idxTShowDate ? row[idxTShowDate - 1] : '') || parentOrder.showDate || '',
          pricePerTicket: parentOrder.pricePerTicket || 0,
          total: parentOrder.total || 0,
          qty: parentOrder.qty || 1,
          checkedIn: isCheckedIn,
          checkInTime: row[idxTTime - 1] || null,
          cancelled: isCancelled,
          cancelledAt: isCancelled ? (row[idxTTime - 1] || new Date().toISOString()) : null,
          slipUrl: tSlip,
          slipImage: tSlip
        };
      }

      // สแกนสถานะยกเลิกยกยวงของออร์เดอร์
      orders.forEach(o => {
        const oTkts = Object.values(tickets).filter(t => t.orderId === o.orderId);
        if (oTkts.length > 0 && oTkts.every(t => t.cancelled)) {
          o.cancelled = true;
        }
      });

      return output({
        success: true,
        earlybird_enabled: settings.earlybird_enabled !== false,
        settings: settings,
        tickets: tickets,
        orders: orders
      });
    }

    // 3. ดึงสถานะคำสั่งซื้อเจาะจง (สำหรับสืบค้นในหน้า status.html ข้ามอุปกรณ์)
    if (action === 'searchOrder') {
      const query = String(e.parameter.query || '').trim();
      const type = e.parameter.type; // 'phone' หรือ 'order'
      
      const ordersSheet = getOrCreateSheet(ss, SHEET_ORDERS, []);
      const oData = ordersSheet.getDataRange().getValues();
      const oHeaders = oData[0] || [];
      const idxOId = findColIndex(oHeaders, ['เลขที่คำสั่งซื้อ', 'เลขที่ออเดอร์', 'เลขออเดอร์', 'orderid', 'order id']);
      const idxOPhone = findColIndex(oHeaders, ['เบอร์โทร', 'เบอร์โทรศัพท์', 'เบอร์', 'phone', 'tel', 'telephone', 'mobile']);

      const ticketsSheet = getOrCreateSheet(ss, SHEET_TICKETS, []);
      const tData = ticketsSheet.getDataRange().getValues();
      const tHeaders = tData[0] || [];
      const idxTId = findColIndex(tHeaders, ['รหัสบัตร', 'ticketid', 'ticket id']);
      const idxTOId = findColIndex(tHeaders, ['เลขที่คำสั่งซื้อ', 'orderid', 'order id']);
      const idxTPhone = findColIndex(tHeaders, ['เบอร์โทร', 'เบอร์โทรศัพท์', 'เบอร์', 'phone', 'tel', 'telephone', 'mobile']);

      const matchedOrderIds = [];
      const qClean = readCleanPhone(query);

      if (type === 'phone') {
        // 1. ค้นหาใน Orders sheet
        if (idxOPhone >= 0 && idxOId >= 0) {
          for (let i = 1; i < oData.length; i++) {
            const oId = oData[i][idxOId];
            if (!oId) continue;
            const pClean = readCleanPhone(oData[i][idxOPhone]);
            if (pClean && (pClean === qClean || (qClean.length >= 9 && pClean.endsWith(qClean.slice(-9))))) {
              if (matchedOrderIds.indexOf(oId) < 0) matchedOrderIds.push(oId);
            }
          }
        }
        // 2. ค้นหาใน Tickets sheet เพิ่มเติม
        if (idxTPhone >= 0 && idxTOId >= 0) {
          for (let i = 1; i < tData.length; i++) {
            const oId = tData[i][idxTOId];
            if (!oId) continue;
            const pClean = readCleanPhone(tData[i][idxTPhone]);
            if (pClean && (pClean === qClean || (qClean.length >= 9 && pClean.endsWith(qClean.slice(-9))))) {
              if (matchedOrderIds.indexOf(oId) < 0) matchedOrderIds.push(oId);
            }
          }
        }
      } else {
        // ค้นหาด้วยรหัสคำสั่งซื้อ หรือรหัสบัตร
        if (idxOId >= 0) {
          for (let i = 1; i < oData.length; i++) {
            const oId = oData[i][idxOId];
            if (!oId) continue;
            if (String(oId).toUpperCase().indexOf(query.toUpperCase()) >= 0) {
              if (matchedOrderIds.indexOf(oId) < 0) matchedOrderIds.push(oId);
            }
          }
        }
        if (idxTOId >= 0) {
          for (let i = 1; i < tData.length; i++) {
            const oId = tData[i][idxTOId];
            if (!oId) continue;
            const tId = idxTId >= 0 ? tData[i][idxTId] : '';
            if (
              (tId && String(tId).toUpperCase().indexOf(query.toUpperCase()) >= 0) ||
              String(oId).toUpperCase().indexOf(query.toUpperCase()) >= 0
            ) {
              if (matchedOrderIds.indexOf(oId) < 0) matchedOrderIds.push(oId);
            }
          }
        }
      }

      const idxTName = findColIndex(tHeaders, ['ชื่อ-นามสกุล', 'ชื่อ', 'name']);
      const idxTType = findColIndex(tHeaders, ['ประเภทบัตร', 'type']);
      const idxTStatus = findColIndex(tHeaders, ['สถานะเช็คอิน', 'สถานะ', 'status']);
      const idxTTime = findColIndex(tHeaders, ['เวลาเช็คอิน', 'time']);
      const idxTShowDate = findColIndex(tHeaders, ['รอบการแสดง', 'showdate']);
      const idxTSlip = findColIndex(tHeaders, ['สลิปการโอนเงิน', 'สลิป', 'slipurl', 'slip url', 'slip']);

      const tickets = {};
      const orders = [];

      const idxOTs = findColIndex(oHeaders, ['วันเวลา', 'เวลา', 'วันที่', 'timestamp', 'date']);
      const idxOName = findColIndex(oHeaders, ['ชื่อ-นามสกุล', 'ชื่อนามสกุล', 'ชื่อ', 'name', 'fullname']);
      const idxOEmail = findColIndex(oHeaders, ['อีเมล', 'email']);
      const idxOType = findColIndex(oHeaders, ['ประเภทบัตร', 'ประเภท', 'tickettype', 'type']);
      const idxOQty = findColIndex(oHeaders, ['จำนวนใบ', 'จำนวนบัตร (ใบ)', 'จำนวนบัตร', 'จำนวน', 'qty']);
      const idxOPrice = findColIndex(oHeaders, ['ราคาต่อใบ (บาท)', 'ราคาต่อใบ', 'ราคา', 'priceperticket', 'price']);
      const idxOTotal = findColIndex(oHeaders, ['ราคารวม', 'ยอดเงินรวม (บาท)', 'ยอดรวม', 'total', 'amount']);
      const idxONote = findColIndex(oHeaders, ['หมายเหตุ', 'note']);
      const idxOTickets = findColIndex(oHeaders, ['รหัสบัตรทั้งหมด', 'รหัสตั๋วทั้งหมด', 'ticketids', 'tickets']);
      const idxOSlip = findColIndex(oHeaders, ['สลิปการโอนเงิน', 'สลิปโอนเงิน', 'สลิป', 'หลักฐานการโอน', 'slipurl', 'slip url', 'slip_url', 'slip', 'sliplink']);
      const idxOShowDate = findColIndex(oHeaders, ['รอบการแสดง', 'รอบ', 'showdate', 'round']);

      const ordersMap = {};
      for (let i = 1; i < oData.length; i++) {
        const row = oData[i];
        const oId = idxOId >= 0 ? row[idxOId] : '';
        if (oId && matchedOrderIds.indexOf(oId) >= 0) {
          const slipVal = (idxOSlip >= 0 ? row[idxOSlip] : '') || '';
          const order = {
            orderId: oId,
            timestamp: idxOTs >= 0 ? row[idxOTs] : '',
            name: idxOName >= 0 ? row[idxOName] : '',
            phone: idxOPhone >= 0 ? readCleanPhone(row[idxOPhone]) : '',
            email: idxOEmail >= 0 ? row[idxOEmail] : '',
            typeName: idxOType >= 0 ? row[idxOType] : '',
            qty: idxOQty >= 0 ? Number(row[idxOQty] || 0) : 0,
            pricePerTicket: idxOPrice >= 0 ? Number(row[idxOPrice] || 0) : 0,
            total: idxOTotal >= 0 ? Number(row[idxOTotal] || 0) : 0,
            note: idxONote >= 0 ? row[idxONote] : '',
            ticketIds: idxOTickets >= 0 && row[idxOTickets] ? String(row[idxOTickets]).split(',').map(s => s.trim()).filter(Boolean) : [],
            slipUrl: slipVal,
            slipImage: slipVal,
            showDate: idxOShowDate >= 0 ? row[idxOShowDate] : '',
            cancelled: false
          };
          orders.push(order);
          ordersMap[oId] = order;
        }
      }

      const matchedOrderIdsSet = new Set(matchedOrderIds);

      for (let i = 1; i < tData.length; i++) {
        const row = tData[i];
        const oId = idxTOId >= 0 ? row[idxTOId] : '';
        if (oId && matchedOrderIdsSet.has(oId)) {
          const tId = idxTId >= 0 ? row[idxTId] : '';
          if (!tId) continue;
          const parentOrder = ordersMap[oId] || {};
          const status = idxTStatus >= 0 ? row[idxTStatus] : '';
          const isCancelled = status === 'ยกเลิกแล้ว';
          const isCheckedIn = status === 'เช็คอินแล้ว';

          let ticketNum = 1;
          const match = String(tId).match(/-T(\d+)$/);
          if (match) ticketNum = Number(match[1]);

          const tSlip = (idxTSlip >= 0 ? row[idxTSlip] : '') || parentOrder.slipUrl || parentOrder.slipImage || '';

          tickets[tId] = {
            ticketId: tId,
            ticketNum: ticketNum,
            orderId: oId,
            name: (idxTName >= 0 ? row[idxTName] : '') || parentOrder.name || '',
            phone: (idxTPhone >= 0 ? readCleanPhone(row[idxTPhone]) : '') || parentOrder.phone || '',
            email: parentOrder.email || '',
            note: parentOrder.note || '',
            type: (idxTType >= 0 ? row[idxTType] : '') || parentOrder.typeName || '',
            showDate: (idxTShowDate >= 0 ? row[idxTShowDate] : '') || parentOrder.showDate || '',
            pricePerTicket: parentOrder.pricePerTicket || 0,
            total: parentOrder.total || 0,
            qty: parentOrder.qty || 1,
            checkedIn: isCheckedIn,
            checkInTime: (idxTTime >= 0 ? row[idxTTime] : null) || null,
            cancelled: isCancelled,
            cancelledAt: isCancelled ? (idxTTime >= 0 ? row[idxTTime] : null) || new Date().toISOString() : null,
            slipUrl: tSlip,
            slipImage: tSlip
          };
        }
      }

      // ตรวจสอบความสมบูรณ์ หากใน Tickets ยังไม่มีข้อมูล ให้สร้าง ticket เสมือนขึ้นมาทันทีเพื่อแสดงในหน้า Status
      orders.forEach(o => {
        const oTkts = Object.values(tickets).filter(t => t.orderId === o.orderId);
        if (oTkts.length === 0 && o.qty > 0) {
          const tIds = (o.ticketIds && o.ticketIds.length) ? o.ticketIds : Array.from({ length: o.qty }, (_, k) => `${o.orderId}-T${k + 1}`);
          tIds.forEach((tid, idx) => {
            tickets[tid] = {
              ticketId: tid,
              ticketNum: idx + 1,
              orderId: o.orderId,
              name: o.name,
              phone: o.phone,
              email: o.email || '',
              note: o.note || '',
              type: o.typeName || '',
              showDate: o.showDate || '',
              pricePerTicket: o.pricePerTicket || 0,
              total: o.total || 0,
              qty: o.qty || 1,
              checkedIn: false,
              checkInTime: null,
              cancelled: false,
              cancelledAt: null,
              slipUrl: o.slipUrl || '',
              slipImage: o.slipImage || ''
            };
          });
        }
      });

      return output({ success: true, tickets: tickets, orders: orders });
    }

    return output({ status: 'ok', service: 'Theater Ticket API' });

  } catch (err) {
    return output({ success: false, error: err.message });
  }
}

// ─── HANDLE NEW ORDER ─────────────────────────────────────────────────────
function handleNewOrder(data) {
  const ss = getSS();
  
  // ── บันทึกไฟล์รูปสลิปลง Google Drive (ถ้ามีแนบมา) ──
  let slipUrl = '—';
  if (data.slipImage && data.slipImage.startsWith('data:image')) {
    try {
      const base64Data = data.slipImage.split(',')[1];
      const mimeType = data.slipImage.split(';')[0].split(':')[1];
      const cleanName = (data.name || '').trim().replace(/[/\\?%*:|"<>]/g, '_');
      const ext = mimeType.indexOf('png') >= 0 ? 'png' : 'jpg';
      const fileName = `slip-${data.orderId}${cleanName ? '-' + cleanName : ''}.${ext}`;
      const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType, fileName);
      
      let file;
      try {
        let folder;
        const folderName = 'Theater Slips - แฟ้มคดีกิเลนแดง';
        const folders = DriveApp.getFoldersByName(folderName);
        if (folders.hasNext()) {
          folder = folders.next();
        } else {
          folder = DriveApp.createFolder(folderName);
        }
        file = folder.createFile(blob);
      } catch(folderErr) {
        file = DriveApp.createFile(blob);
      }

      try {
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      } catch(shareErr) {}
      
      file.setDescription(`เลขที่คำสั่งซื้อ: ${data.orderId}\nผู้จอง: ${data.name || '—'}\nเบอร์โทร: ${data.phone || '—'}\nรอบการแสดง: ${data.showDate || '—'}\nเวลาทำรายการ: ${data.timestamp || ''}`);
      const fileId = file.getId();
      slipUrl = `https://drive.google.com/file/d/${fileId}/view?usp=sharing`;
    } catch (e) {
      slipUrl = 'บันทึกรูปไม่สำเร็จ: ' + e.message;
    }
  } else if (data.slipImage) {
    // โควต้าทีมงาน หรือข้อมูลอื่นๆ
    slipUrl = data.slipImage;
  }

  // ป้องกันเบอร์โทรเลข 0 หายเมื่อสั่งจองใหม่
  const cleanPhoneVal = formatPhoneToWrite(data.phone);

  // ── บันทึกแถวข้อมูลลง Orders ──
  const ordersSheet = getOrCreateSheet(ss, SHEET_ORDERS, [
    'เลขที่คำสั่งซื้อ',
    'วันเวลา',
    'ชื่อ-นามสกุล',
    'ชื่อเล่น',
    'เบอร์โทร',
    'อีเมล',
    'ประเภทบัตร',
    'จำนวนใบ',
    'ราคาต่อใบ',
    'ราคารวม',
    'รอบการแสดง',
    'หมายเหตุ',
    'รหัสบัตรทั้งหมด',
    'สลิปการโอนเงิน',
  ]);
  
  const oHeaders = ordersSheet.getDataRange().getValues()[0] || [];
  const rowData = [];
  
  function assignOrderCol(names, val) {
    const idx = findColIndex(oHeaders, names);
    if (idx >= 0) rowData[idx] = val;
  }

  assignOrderCol(['เลขที่คำสั่งซื้อ', 'orderid', 'order id'], data.orderId);
  assignOrderCol(['วันเวลา', 'timestamp', 'date'], data.timestamp);
  assignOrderCol(['ชื่อ-นามสกุล', 'name', 'fullname'], data.name);
  
  let nickIdx = findColIndex(oHeaders, ['ชื่อเล่น', 'nickname', 'nick']);
  if (nickIdx < 0 && data.nickname) {
    const newCol = ordersSheet.getLastColumn() + 1;
    ordersSheet.getRange(1, newCol).setValue('ชื่อเล่น');
    ordersSheet.getRange(1, newCol).setBackground('#4a2080').setFontColor('#ffffff').setFontWeight('bold');
    oHeaders.push('ชื่อเล่น');
    nickIdx = newCol - 1;
  }
  if (nickIdx >= 0) rowData[nickIdx] = data.nickname || '—';

  assignOrderCol(['เบอร์โทร', 'phone', 'tel'], cleanPhoneVal);
  assignOrderCol(['อีเมล', 'email'], data.email || '—');
  assignOrderCol(['ประเภทบัตร', 'tickettype', 'type'], data.ticketType);
  assignOrderCol(['จำนวนใบ', 'qty'], data.qty);
  assignOrderCol(['ราคาต่อใบ', 'price'], data.pricePerTicket);

  let subtotalIdx = findColIndex(oHeaders, ['ราคาก่อนลด', 'ยอดรวมก่อนลด', 'subtotal']);
  if (subtotalIdx >= 0) rowData[subtotalIdx] = data.subtotal || data.total;

  let discIdx = findColIndex(oHeaders, ['ส่วนลด', 'discount']);
  if (discIdx >= 0) rowData[discIdx] = data.discount || 0;

  let promoIdx = findColIndex(oHeaders, ['โค้ดส่วนลด', 'promocode', 'promo code', 'promo']);
  if (promoIdx >= 0) rowData[promoIdx] = data.promoCode || '—';

  assignOrderCol(['ราคารวม', 'total'], data.total);
  assignOrderCol(['รอบการแสดง', 'showdate'], data.showDate || '—');
  assignOrderCol(['หมายเหตุ', 'note'], data.note || '—');
  assignOrderCol(['รหัสบัตรทั้งหมด', 'tickets'], data.tickets);
  
  let slipIdx = findColIndex(oHeaders, ['สลิปการโอนเงิน', 'สลิปโอนเงิน', 'สลิป', 'หลักฐานการโอน', 'หลักฐานการโอนเงิน', 'ลิงก์สลิป', 'ลิงค์สลิป', 'slipurl', 'slip url', 'slip_url', 'slip', 'sliplink']);
  if (slipIdx < 0) {
    const newCol = ordersSheet.getLastColumn() + 1;
    ordersSheet.getRange(1, newCol).setValue('สลิปการโอนเงิน');
    ordersSheet.getRange(1, newCol).setBackground('#4a2080').setFontColor('#ffffff').setFontWeight('bold');
    oHeaders.push('สลิปการโอนเงิน');
    slipIdx = newCol - 1;
  }
  rowData[slipIdx] = slipUrl;

  for (let i = 0; i < oHeaders.length; i++) {
    if (rowData[i] === undefined) rowData[i] = '';
  }
  ordersSheet.appendRow(rowData);

  // ── บันทึกแยกรายใบลง Tickets ──
  const ticketsSheet = getOrCreateSheet(ss, SHEET_TICKETS, [
    'รหัสบัตร',
    'เลขที่คำสั่งซื้อ',
    'ชื่อ-นามสกุล',
    'ชื่อเล่น',
    'เบอร์โทร',
    'ประเภทบัตร',
    'รอบการแสดง',
    'สถานะเช็คอิน',
    'เวลาเช็คอิน',
    'สลิปการโอนเงิน'
  ]);
  
  const tHeaders = ticketsSheet.getDataRange().getValues()[0] || [];
  const ticketIds = data.tickets.split(', ');
  let tSlipIdx = findColIndex(tHeaders, ['สลิปการโอนเงิน', 'สลิปโอนเงิน', 'สลิป', 'หลักฐานการโอน', 'slipurl', 'slip url', 'slip']);
  if (tSlipIdx < 0) {
    const newCol = ticketsSheet.getLastColumn() + 1;
    ticketsSheet.getRange(1, newCol).setValue('สลิปการโอนเงิน');
    ticketsSheet.getRange(1, newCol).setBackground('#4a2080').setFontColor('#ffffff').setFontWeight('bold');
    tHeaders.push('สลิปการโอนเงิน');
    tSlipIdx = newCol - 1;
  }
  
  let tNickIdx = findColIndex(tHeaders, ['ชื่อเล่น', 'nickname', 'nick']);
  if (tNickIdx < 0 && data.nickname) {
    const newCol = ticketsSheet.getLastColumn() + 1;
    ticketsSheet.getRange(1, newCol).setValue('ชื่อเล่น');
    ticketsSheet.getRange(1, newCol).setBackground('#4a2080').setFontColor('#ffffff').setFontWeight('bold');
    tHeaders.push('ชื่อเล่น');
    tNickIdx = newCol - 1;
  }

  ticketIds.forEach(tid => {
    const tRowData = [];
    const assignTCol = (names, val) => {
      const idx = findColIndex(tHeaders, names);
      if (idx >= 0) tRowData[idx] = val;
    };
    assignTCol(['รหัสบัตร', 'ticketid'], tid);
    assignTCol(['เลขที่คำสั่งซื้อ', 'orderid'], data.orderId);
    assignTCol(['ชื่อ-นามสกุล', 'name'], data.name);
    if (tNickIdx >= 0) tRowData[tNickIdx] = data.nickname || '—';
    assignTCol(['เบอร์โทร', 'phone'], cleanPhoneVal);
    assignTCol(['ประเภทบัตร', 'type'], data.ticketType);
    assignTCol(['รอบการแสดง', 'showdate'], data.showDate || '—');
    assignTCol(['สถานะเช็คอิน', 'status'], 'ยังไม่เช็คอิน');
    assignTCol(['เวลาเช็คอิน', 'checkintime'], '');
    tRowData[tSlipIdx] = slipUrl;

    for (let i = 0; i < tHeaders.length; i++) {
      if (tRowData[i] === undefined) tRowData[i] = '';
    }
    ticketsSheet.appendRow(tRowData);
  });
  
  return ContentService
    .createTextOutput(JSON.stringify({ success: true, orderId: data.orderId }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ─── HELPERS ──────────────────────────────────────────────────────────────
function getOrCreateSheet(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (headers.length > 0) {
      sheet.appendRow(headers);
      const headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setBackground('#4a2080');
      headerRange.setFontColor('#ffffff');
      headerRange.setFontWeight('bold');
      headerRange.setFontSize(11);
      sheet.setFrozenRows(1);
      
      headers.forEach((_, i) => sheet.autoResizeColumn(i + 1));
    }
  }
  return sheet;
}

function getOrCreateSettings(ss) {
  let sheet = ss.getSheetByName(SHEET_SETTINGS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_SETTINGS);
    sheet.appendRow(['Key', 'Value']);
    sheet.appendRow(['earlybird_enabled', 'true']);
    
    const range = sheet.getRange(1, 1, 1, 2);
    range.setBackground('#4a2080').setFontColor('#ffffff').setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.autoResizeColumn(1);
    sheet.autoResizeColumn(2);
  }
  return sheet;
}

// ค้นหาดัชนีคอลัมน์แบบยืดหยุ่น (รองรับหลายชื่อ ทั้งไทย/อังกฤษ และ case-insensitive)
function findColIndex(headers, possibleNames) {
  if (!headers || !headers.length) return -1;
  for (let i = 0; i < headers.length; i++) {
    const h = String(headers[i] || '').trim().toLowerCase().replace(/[\s_\-\.\(\)]/g, '');
    for (let j = 0; j < possibleNames.length; j++) {
      const target = String(possibleNames[j] || '').trim().toLowerCase().replace(/[\s_\-\.\(\)]/g, '');
      if (h === target) return i;
    }
  }
  return -1;
}

function getColumnIndex(sheet, headerName) {
  const lastCol = sheet.getLastColumn();
  if (lastCol > 0) {
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const idx = headers.indexOf(headerName);
    if (idx >= 0) return idx + 1;
  }
  // ถ้าไม่พบคอลัมน์นี้ ให้ทำการเพิ่มคอลัมน์ใหม่ที่ท้ายหัวตารางแบบอัตโนมัติ
  const newCol = lastCol + 1;
  sheet.getRange(1, newCol).setValue(headerName);
  sheet.getRange(1, newCol).setBackground('#1c1c1e').setFontColor('#ffffff').setFontWeight('bold');
  return newCol;
}

// บังคับแปลงเป็น Text โดยการเพิ่มสัญลักษณ์ ' นำหน้า ป้องกันเลข 0 ด้านหน้าหายบน Sheets
function formatPhoneToWrite(p) {
  let s = String(p || '').trim().replace(/\D/g, '');
  if (s.startsWith('66')) {
    s = '0' + s.substring(2);
  }
  if (s.length === 9 && !s.startsWith('0')) {
    s = '0' + s;
  }
  return "'" + s;
}

// ล้างรูปแบบเบอร์โทรและเติม 0 ด้านหน้าในกรณีที่ดึงมาจากชีทเลข 0 หายไป
function readCleanPhone(p) {
  let s = String(p || '').trim().replace(/\D/g, '');
  if (s.startsWith('66')) {
    s = '0' + s.substring(2);
  }
  if (s.length === 9 && !s.startsWith('0')) {
    s = '0' + s;
  }
  return s;
}
