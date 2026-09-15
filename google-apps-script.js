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
 *
 * วิธีติดตั้ง:
 * 1. ไปที่ https://script.google.com → สร้างโปรเจกต์ใหม่
 * 2. วางโค้ดนี้ทั้งหมดในไฟล์ Code.gs
 * 3. แก้ SPREADSHEET_ID ให้ตรงกับ Google Sheet ของคุณ
 * 4. กด Deploy → New deployment → Web app
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 5. Copy URL ที่ได้ ไปวางใน CONFIG.APPS_SCRIPT_URL ในไฟล์ app.js
 *    และ APPS_SCRIPT_URL ในไฟล์ staff.html
 *
 * โครงสร้าง Google Sheet:
 * - Sheet1 "Orders" — รายการคำสั่งซื้อ
 * - Sheet2 "Tickets" — รายการบัตรแต่ละใบ
 * - Sheet3 "CheckIns" — log การเช็คอิน
 * - Sheet4 "Settings" — เก็บการตั้งค่าส่วนกลาง (เช่น เปิด/ปิด Early Bird)
 * ============================================================
 */

// ⚠️ แก้ค่านี้: ใส่ ID ของ Google Sheet ของคุณ
// (เอาจาก URL: https://docs.google.com/spreadsheets/d/[SPREADSHEET_ID]/edit)
const SPREADSHEET_ID = 'YOUR_SPREADSHEET_ID_HERE';

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

    // 2. จัดการบันทึกการเช็คอิน / ยกเลิกเช็คอิน / ยกเลิกตั๋ว / กู้คืนตั๋ว รายใบ
    if (data.action === 'checkin' || data.action === 'undoCheckin' || data.action === 'cancelTicket' || data.action === 'restoreTicket') {
      const ss = getSS();
      const ticketsSheet = getOrCreateSheet(ss, SHEET_TICKETS, []);
      const ticketId = data.ticketId;
      
      const tData = ticketsSheet.getDataRange().getValues();
      const tHeaders = tData[0] || [];
      const idxTId = tHeaders.indexOf('รหัสบัตร') + 1;
      const idxTStatus = tHeaders.indexOf('สถานะเช็คอิน') + 1;
      const idxTTime = tHeaders.indexOf('เวลาเช็คอิน') + 1;

      let foundRow = -1;
      for (let i = 1; i < tData.length; i++) {
        if (tData[i][idxTId - 1] === ticketId) {
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
        } else if (data.action === 'cancelTicket') {
          statusVal = 'ยกเลิกแล้ว';
        }

        ticketsSheet.getRange(foundRow, idxTStatus).setValue(statusVal);
        ticketsSheet.getRange(foundRow, idxTTime).setValue(timeVal);
        
        // ใส่สีสถานะเพื่อให้ดูในชีทง่ายขึ้น
        const statusCell = ticketsSheet.getRange(foundRow, idxTStatus);
        if (statusVal === 'เช็คอินแล้ว') statusCell.setBackground('#d4edda');
        else if (statusVal === 'ยกเลิกแล้ว') statusCell.setBackground('#f8d7da');
        else statusCell.setBackground('#ffffff');

        return output({ success: true });
      }
      return output({ success: false, error: 'ไม่พบรหัสบัตร ' + ticketId });
    }

    // 3. จัดการยกเลิกคำสั่งซื้อ / กู้คืนคำสั่งซื้อ (ทั้งออร์เดอร์)
    if (data.action === 'cancelOrder' || data.action === 'restoreOrder') {
      const ss = getSS();
      const ticketsSheet = getOrCreateSheet(ss, SHEET_TICKETS, []);
      const orderId = data.orderId;
      
      const tData = ticketsSheet.getDataRange().getValues();
      const tHeaders = tData[0] || [];
      const idxTOId = tHeaders.indexOf('เลขที่คำสั่งซื้อ') + 1;
      const idxTStatus = tHeaders.indexOf('สถานะเช็คอิน') + 1;
      const idxTTime = tHeaders.indexOf('เวลาเช็คอิน') + 1;

      const statusVal = data.action === 'cancelOrder' ? 'ยกเลิกแล้ว' : 'ยังไม่เช็คอิน';

      for (let i = 1; i < tData.length; i++) {
        if (tData[i][idxTOId - 1] === orderId) {
          const rowNum = i + 1;
          ticketsSheet.getRange(rowNum, idxTStatus).setValue(statusVal);
          ticketsSheet.getRange(rowNum, idxTTime).setValue('');
          
          const statusCell = ticketsSheet.getRange(rowNum, idxTStatus);
          if (statusVal === 'ยกเลิกแล้ว') statusCell.setBackground('#f8d7da');
          else statusCell.setBackground('#ffffff');
        }
      }
      return output({ success: true });
    }

    // 3.1 จัดการลบตั๋วถาวร (Permanent Delete Ticket)
    if (data.action === 'deleteTicket') {
      const ss = getSS();
      const ticketsSheet = getOrCreateSheet(ss, SHEET_TICKETS, []);
      const ticketId = data.ticketId;
      const orderId = data.orderId;

      // 1. ลบแถวใน Tickets sheet
      const tData = ticketsSheet.getDataRange().getValues();
      const tHeaders = tData[0] || [];
      const idxTId = tHeaders.indexOf('รหัสบัตร');

      for (let i = tData.length - 1; i >= 1; i--) {
        if (tData[i][idxTId] === ticketId) {
          ticketsSheet.deleteRow(i + 1);
          break;
        }
      }

      // 2. ปรับลดจำนวนตั๋วใน Orders sheet หรือลบออร์เดอร์หากไม่เหลือตั๋ว
      if (orderId) {
        const ordersSheet = getOrCreateSheet(ss, SHEET_ORDERS, []);
        const oData = ordersSheet.getDataRange().getValues();
        const oHeaders = oData[0] || [];
        const idxOId = oHeaders.indexOf('เลขที่คำสั่งซื้อ');
        const idxOQty = oHeaders.indexOf('จำนวนบัตร (ใบ)');
        const idxOTotal = oHeaders.indexOf('ยอดเงินรวม (บาท)');
        const idxOPrice = oHeaders.indexOf('ราคาต่อใบ (บาท)');
        const idxOTkts = oHeaders.indexOf('รหัสตั๋วทั้งหมด');

        for (let i = 1; i < oData.length; i++) {
          if (oData[i][idxOId] === orderId) {
            const rowNum = i + 1;
            let currentTkts = String(oData[i][idxOTkts] || '').split(',').map(s => s.trim()).filter(Boolean);
            currentTkts = currentTkts.filter(id => id !== ticketId);
            const newQty = currentTkts.length;
            if (newQty <= 0) {
              ordersSheet.deleteRow(rowNum);
            } else {
              const pricePer = Number(oData[i][idxOPrice]) || 0;
              if (idxOQty >= 0) ordersSheet.getRange(rowNum, idxOQty + 1).setValue(newQty);
              if (idxOTotal >= 0) ordersSheet.getRange(rowNum, idxOTotal + 1).setValue(newQty * pricePer);
              if (idxOTkts >= 0) ordersSheet.getRange(rowNum, idxOTkts + 1).setValue(currentTkts.join(', '));
            }
            break;
          }
        }
      }
      return output({ success: true });
    }

    // 3.2 จัดการลบคำสั่งซื้อถาวรทั้งออร์เดอร์ (Permanent Delete Order)
    if (data.action === 'deleteOrder') {
      const ss = getSS();
      const ordersSheet = getOrCreateSheet(ss, SHEET_ORDERS, []);
      const ticketsSheet = getOrCreateSheet(ss, SHEET_TICKETS, []);
      const orderId = data.orderId;

      // 1. ลบตั๋วทั้งหมดของออร์เดอร์นี้ใน Tickets sheet
      const tData = ticketsSheet.getDataRange().getValues();
      const tHeaders = tData[0] || [];
      const idxTOId = tHeaders.indexOf('เลขที่คำสั่งซื้อ');

      for (let i = tData.length - 1; i >= 1; i--) {
        if (tData[i][idxTOId] === orderId) {
          ticketsSheet.deleteRow(i + 1);
        }
      }

      // 2. ลบแถวใน Orders sheet
      const oData = ordersSheet.getDataRange().getValues();
      const oHeaders = oData[0] || [];
      const idxOId = oHeaders.indexOf('เลขที่คำสั่งซื้อ');

      for (let i = oData.length - 1; i >= 1; i--) {
        if (oData[i][idxOId] === orderId) {
          ordersSheet.deleteRow(i + 1);
          break;
        }
      }
      return output({ success: true });
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

      const soldCounts = {};
      for (let i = 1; i < tData.length; i++) {
        const row = tData[i];
        const status = row[idxTStatus - 1];
        const showDate = row[idxTShowDate - 1];
        if (showDate && status !== 'ยกเลิกแล้ว') {
          soldCounts[showDate] = (soldCounts[showDate] || 0) + 1;
        }
      }
      settings["soldCounts"] = soldCounts;

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
      const idxOId = oHeaders.indexOf('เลขที่คำสั่งซื้อ') + 1;
      const idxOTs = oHeaders.indexOf('วันเวลา') + 1;
      const idxOName = oHeaders.indexOf('ชื่อ-นามสกุล') + 1;
      const idxOPhone = oHeaders.indexOf('เบอร์โทร') + 1;
      const idxOEmail = oHeaders.indexOf('อีเมล') + 1;
      const idxOType = oHeaders.indexOf('ประเภทบัตร') + 1;
      const idxOQty = oHeaders.indexOf('จำนวนใบ') + 1;
      const idxOPrice = oHeaders.indexOf('ราคาต่อใบ') + 1;
      const idxOTotal = oHeaders.indexOf('ราคารวม') + 1;
      const idxONote = oHeaders.indexOf('หมายเหตุ') + 1;
      const idxOTickets = oHeaders.indexOf('รหัสบัตรทั้งหมด') + 1;
      const idxOSlip = oHeaders.indexOf('สลิปการโอนเงิน') + 1;
      const idxOShowDate = oHeaders.indexOf('รอบการแสดง') + 1;

      // ค้นหาตำแหน่งคอลัมน์ของชีทตั๋วรายใบ (Tickets)
      const tHeaders = tData[0] || [];
      const idxTId = tHeaders.indexOf('รหัสบัตร') + 1;
      const idxTOId = tHeaders.indexOf('เลขที่คำสั่งซื้อ') + 1;
      const idxTName = tHeaders.indexOf('ชื่อ-นามสกุล') + 1;
      const idxTPhone = tHeaders.indexOf('เบอร์โทร') + 1;
      const idxTType = tHeaders.indexOf('ประเภทบัตร') + 1;
      const idxTStatus = tHeaders.indexOf('สถานะเช็คอิน') + 1;
      const idxTTime = tHeaders.indexOf('เวลาเช็คอิน') + 1;
      const idxTShowDate = tHeaders.indexOf('รอบการแสดง') + 1;

      // สร้าง Map แหล่งรวมคำสั่งซื้อ
      const ordersMap = {};
      for (let i = 1; i < oData.length; i++) {
        const row = oData[i];
        if (!row[idxOId - 1]) continue;
        const order = {
          orderId: row[idxOId - 1],
          timestamp: row[idxOTs - 1],
          name: row[idxOName - 1],
          phone: readCleanPhone(row[idxOPhone - 1]),
          email: idxOEmail ? row[idxOEmail - 1] : '',
          typeName: row[idxOType - 1],
          qty: Number(row[idxOQty - 1] || 0),
          pricePerTicket: Number(row[idxOPrice - 1] || 0),
          total: Number(row[idxOTotal - 1] || 0),
          note: idxONote ? row[idxONote - 1] : '',
          ticketIds: idxOTickets ? (row[idxOTickets - 1] || '').split(', ') : [],
          slipImage: idxOSlip ? row[idxOSlip - 1] : '',
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

        tickets[tId] = {
          ticketId: tId,
          ticketNum: ticketNum,
          orderId: oId,
          name: row[idxTName - 1] || parentOrder.name || '',
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
          cancelledAt: isCancelled ? (row[idxTTime - 1] || new Date().toISOString()) : null
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
      const query = e.parameter.query;
      const type = e.parameter.type; // 'phone' หรือ 'order'
      
      const ordersSheet = getOrCreateSheet(ss, SHEET_ORDERS, []);
      const oData = ordersSheet.getDataRange().getValues();
      const oHeaders = oData[0] || [];
      const idxOId = oHeaders.indexOf('เลขที่คำสั่งซื้อ') + 1;
      const idxOPhone = oHeaders.indexOf('เบอร์โทร') + 1;

      const matchedOrderIds = [];
      const qClean = readCleanPhone(query);

      for (let i = 1; i < oData.length; i++) {
        const oId = oData[i][idxOId - 1];
        if (!oId) continue;
        const phone = oData[i][idxOPhone - 1];
        if (type === 'phone') {
          const pClean = readCleanPhone(phone);
          if (pClean === qClean) matchedOrderIds.push(oId);
        } else {
          if (String(oId || '').toUpperCase().indexOf(query.toUpperCase()) >= 0) {
            matchedOrderIds.push(oId);
          }
        }
      }

      const ticketsSheet = getOrCreateSheet(ss, SHEET_TICKETS, []);
      const tData = ticketsSheet.getDataRange().getValues();
      const tHeaders = tData[0] || [];
      const idxTId = tHeaders.indexOf('รหัสบัตร') + 1;
      const idxTOId = tHeaders.indexOf('เลขที่คำสั่งซื้อ') + 1;
      const idxTName = tHeaders.indexOf('ชื่อ-นามสกุล') + 1;
      const idxTPhone = tHeaders.indexOf('เบอร์โทร') + 1;
      const idxTType = tHeaders.indexOf('ประเภทบัตร') + 1;
      const idxTStatus = tHeaders.indexOf('สถานะเช็คอิน') + 1;
      const idxTTime = tHeaders.indexOf('เวลาเช็คอิน') + 1;
      const idxTShowDate = tHeaders.indexOf('รอบการแสดง') + 1;

      const tickets = {};
      const orders = [];

      const oHeadersFull = oHeaders;
      const idxOTs = oHeadersFull.indexOf('วันเวลา') + 1;
      const idxOName = oHeadersFull.indexOf('ชื่อ-นามสกุล') + 1;
      const idxOEmail = oHeadersFull.indexOf('อีเมล') + 1;
      const idxOType = oHeadersFull.indexOf('ประเภทบัตร') + 1;
      const idxOQty = oHeadersFull.indexOf('จำนวนใบ') + 1;
      const idxOPrice = oHeadersFull.indexOf('ราคาต่อใบ') + 1;
      const idxOTotal = oHeadersFull.indexOf('ราคารวม') + 1;
      const idxONote = oHeadersFull.indexOf('หมายเหตุ') + 1;
      const idxOTickets = oHeadersFull.indexOf('รหัสบัตรทั้งหมด') + 1;
      const idxOSlip = oHeadersFull.indexOf('สลิปการโอนเงิน') + 1;
      const idxOShowDate = oHeadersFull.indexOf('รอบการแสดง') + 1;

      const ordersMap = {};
      for (let i = 1; i < oData.length; i++) {
        const row = oData[i];
        const oId = row[idxOId - 1];
        if (matchedOrderIds.indexOf(oId) >= 0) {
          const order = {
            orderId: oId,
            timestamp: row[idxOTs - 1],
            name: row[idxOName - 1],
            phone: readCleanPhone(row[idxOPhone - 1]),
            email: idxOEmail ? row[idxOEmail - 1] : '',
            typeName: row[idxOType - 1],
            qty: Number(row[idxOQty - 1] || 0),
            pricePerTicket: Number(row[idxOPrice - 1] || 0),
            total: Number(row[idxOTotal - 1] || 0),
            note: idxONote ? row[idxONote - 1] : '',
            ticketIds: idxOTickets ? (row[idxOTickets - 1] || '').split(', ') : [],
            slipImage: idxOSlip ? row[idxOSlip - 1] : '',
            showDate: idxOShowDate ? row[idxOShowDate - 1] : '',
            cancelled: false
          };
          orders.push(order);
          ordersMap[oId] = order;
        }
      }

      const matchedOrderIdsSet = new Set(matchedOrderIds);

      for (let i = 1; i < tData.length; i++) {
        const row = tData[i];
        const oId = row[idxTOId - 1];
        if (oId && matchedOrderIdsSet.has(oId)) {
          const tId = row[idxTId - 1];
          if (!tId) continue;
          const parentOrder = ordersMap[oId] || {};
          const status = row[idxTStatus - 1];
          const isCancelled = status === 'ยกเลิกแล้ว';
          const isCheckedIn = status === 'เช็คอินแล้ว';

          let ticketNum = 1;
          const match = tId.match(/-T(\d+)$/);
          if (match) ticketNum = Number(match[1]);

          tickets[tId] = {
            ticketId: tId,
            ticketNum: ticketNum,
            orderId: oId,
            name: row[idxTName - 1] || parentOrder.name || '',
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
            cancelledAt: isCancelled ? (row[idxTTime - 1] || new Date().toISOString()) : null
          };
        }
      }

      orders.forEach(o => {
        const oTkts = Object.values(tickets).filter(t => t.orderId === o.orderId);
        if (oTkts.length > 0 && oTkts.every(t => t.cancelled)) {
          o.cancelled = true;
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
      const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType, `slip-${data.orderId}.jpg`);
      
      let folder;
      const folders = DriveApp.getFoldersByName('Theater Slips - แฟ้มคดีกิเลนแดง');
      if (folders.hasNext()) {
        folder = folders.next();
      } else {
        folder = DriveApp.createFolder('Theater Slips - แฟ้มคดีกิเลนแดง');
      }
      const file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      slipUrl = file.getUrl();
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
  rowData[oHeaders.indexOf('เลขที่คำสั่งซื้อ')] = data.orderId;
  rowData[oHeaders.indexOf('วันเวลา')] = data.timestamp;
  rowData[oHeaders.indexOf('ชื่อ-นามสกุล')] = data.name;
  rowData[oHeaders.indexOf('เบอร์โทร')] = cleanPhoneVal;
  rowData[oHeaders.indexOf('อีเมล')] = data.email || '—';
  rowData[oHeaders.indexOf('ประเภทบัตร')] = data.ticketType;
  rowData[oHeaders.indexOf('จำนวนใบ')] = data.qty;
  rowData[oHeaders.indexOf('ราคาต่อใบ')] = data.pricePerTicket;
  rowData[oHeaders.indexOf('ราคารวม')] = data.total;
  rowData[oHeaders.indexOf('รอบการแสดง')] = data.showDate || '—';
  rowData[oHeaders.indexOf('หมายเหตุ')] = data.note || '—';
  rowData[oHeaders.indexOf('รหัสบัตรทั้งหมด')] = data.tickets;
  rowData[oHeaders.indexOf('สลิปการโอนเงิน')] = slipUrl;

  for (let i = 0; i < oHeaders.length; i++) {
    if (rowData[i] === undefined) rowData[i] = '';
  }
  ordersSheet.appendRow(rowData);

  // ── บันทึกแยกรายใบลง Tickets ──
  const ticketsSheet = getOrCreateSheet(ss, SHEET_TICKETS, [
    'รหัสบัตร',
    'เลขที่คำสั่งซื้อ',
    'ชื่อ-นามสกุล',
    'เบอร์โทร',
    'ประเภทบัตร',
    'รอบการแสดง',
    'สถานะเช็คอิน',
    'เวลาเช็คอิน',
  ]);
  
  const tHeaders = ticketsSheet.getDataRange().getValues()[0] || [];
  const ticketIds = data.tickets.split(', ');
  
  ticketIds.forEach(tid => {
    const tRowData = [];
    tRowData[tHeaders.indexOf('รหัสบัตร')] = tid;
    tRowData[tHeaders.indexOf('เลขที่คำสั่งซื้อ')] = data.orderId;
    tRowData[tHeaders.indexOf('ชื่อ-นามสกุล')] = data.name;
    tRowData[tHeaders.indexOf('เบอร์โทร')] = cleanPhoneVal;
    tRowData[tHeaders.indexOf('ประเภทบัตร')] = data.ticketType;
    tRowData[tHeaders.indexOf('รอบการแสดง')] = data.showDate || '—';
    tRowData[tHeaders.indexOf('สถานะเช็คอิน')] = 'ยังไม่เช็คอิน';
    tRowData[tHeaders.indexOf('เวลาเช็คอิน')] = '';

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
  sheet.getRange(1, newCol).setBackground('#4a2080').setFontColor('#ffffff').setFontWeight('bold');
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
