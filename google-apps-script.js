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

    // 2.1 จัดการยกเลิกตั๋วรายใบ (Soft Cancel: เปลี่ยนสถานะเป็น "ยกเลิกแล้ว" ไม่ลบแถว เพื่อให้อยู่ในแท็บยกเลิกแล้ว)
    if (data.action === 'cancelTicket') {
      const ss = getSS();
      const ticketsSheet = getOrCreateSheet(ss, SHEET_TICKETS, []);
      const ticketId = data.ticketId;
      const orderId = data.orderId;

      const tData = ticketsSheet.getDataRange().getValues();
      const tHeaders = tData[0] || [];
      const idxTId = findColIndex(tHeaders, ['รหัสบัตร', 'ticketid', 'ticket id']) + 1;
      const idxTOId = findColIndex(tHeaders, ['เลขที่คำสั่งซื้อ', 'orderid', 'order id']) + 1;
      const idxTStatus = findColIndex(tHeaders, ['สถานะเช็คอิน', 'สถานะ', 'status']) + 1;
      const idxTTime = findColIndex(tHeaders, ['เวลาเช็คอิน', 'time']) + 1;

      let foundRow = -1;
      let matchedOId = orderId;
      for (let i = 1; i < tData.length; i++) {
        if (String(tData[i][idxTId - 1]).trim() === String(ticketId).trim()) {
          foundRow = i + 1;
          if (!matchedOId && idxTOId > 0) matchedOId = String(tData[i][idxTOId - 1]).trim();
          break;
        }
      }

      if (foundRow >= 0) {
        if (idxTStatus > 0) {
          const statusCell = ticketsSheet.getRange(foundRow, idxTStatus);
          statusCell.setValue('ยกเลิกแล้ว');
          statusCell.setBackground('#f8d7da');
        }
        if (idxTTime > 0) {
          ticketsSheet.getRange(foundRow, idxTTime).setValue('');
        }

        // ตรวจสอบว่าตั๋วทุกใบของออร์เดอร์นี้ถูกยกเลิกแล้วหรือไม่
        if (matchedOId) {
          const updatedTData = ticketsSheet.getDataRange().getValues();
          let allCancelled = true;
          let hasTicket = false;
          for (let i = 1; i < updatedTData.length; i++) {
            if (String(updatedTData[i][idxTOId - 1]).trim() === matchedOId) {
              hasTicket = true;
              if (String(updatedTData[i][idxTStatus - 1]).trim() !== 'ยกเลิกแล้ว') {
                allCancelled = false;
                break;
              }
            }
          }
          if (hasTicket && allCancelled) {
            const ordersSheet = getOrCreateSheet(ss, SHEET_ORDERS, []);
            const oData = ordersSheet.getDataRange().getValues();
            const oHeaders = oData[0] || [];
            const idxOId = findColIndex(oHeaders, ['เลขที่คำสั่งซื้อ', 'orderid', 'order id']);
            const idxOStatus = findColIndex(oHeaders, ['สถานะ', 'status']);
            const idxONote = findColIndex(oHeaders, ['หมายเหตุ', 'note', 'remark']);
            if (idxOId >= 0) {
              for (let i = 1; i < oData.length; i++) {
                if (String(oData[i][idxOId]).trim() === matchedOId) {
                  const rowNum = i + 1;
                  if (idxOStatus >= 0) {
                    ordersSheet.getRange(rowNum, idxOStatus + 1).setValue('ยกเลิกแล้ว').setBackground('#f8d7da');
                  } else if (idxONote >= 0) {
                    const noteVal = String(oData[i][idxONote] || '');
                    if (!noteVal.includes('[ยกเลิกแล้ว]')) {
                      ordersSheet.getRange(rowNum, idxONote + 1).setValue(noteVal ? `[ยกเลิกแล้ว] ${noteVal}` : '[ยกเลิกแล้ว]');
                    }
                  }
                }
              }
            }
          }
        }

        return output({ success: true, message: 'ยกเลิกบัตรเรียบร้อยแล้ว' });
      }
      return output({ success: false, error: 'ไม่พบรหัสบัตร ' + ticketId });
    }

    // 2.2 จัดการยกเลิกคำสั่งซื้อ (Soft Cancel: เปลี่ยนสถานะตั๋วทุกใบของออร์เดอร์เป็น "ยกเลิกแล้ว")
    if (data.action === 'cancelOrder') {
      const ss = getSS();
      const ticketsSheet = getOrCreateSheet(ss, SHEET_TICKETS, []);
      const ordersSheet = getOrCreateSheet(ss, SHEET_ORDERS, []);
      const orderId = String(data.orderId || '').trim();

      const tData = ticketsSheet.getDataRange().getValues();
      const tHeaders = tData[0] || [];
      const idxTOId = findColIndex(tHeaders, ['เลขที่คำสั่งซื้อ', 'orderid', 'order id']) + 1;
      const idxTStatus = findColIndex(tHeaders, ['สถานะเช็คอิน', 'สถานะ', 'status']) + 1;
      const idxTTime = findColIndex(tHeaders, ['เวลาเช็คอิน', 'time']) + 1;

      let count = 0;
      if (idxTOId > 0 && idxTStatus > 0) {
        for (let i = 1; i < tData.length; i++) {
          if (String(tData[i][idxTOId - 1]).trim() === orderId) {
            const rowNum = i + 1;
            const statusCell = ticketsSheet.getRange(rowNum, idxTStatus);
            statusCell.setValue('ยกเลิกแล้ว');
            statusCell.setBackground('#f8d7da');
            if (idxTTime > 0) {
              ticketsSheet.getRange(rowNum, idxTTime).setValue('');
            }
            count++;
          }
        }
      }

      // อัปเดตสถานะใน Orders sheet
      const oData = ordersSheet.getDataRange().getValues();
      const oHeaders = oData[0] || [];
      const idxOId = findColIndex(oHeaders, ['เลขที่คำสั่งซื้อ', 'orderid', 'order id']);
      const idxOStatus = findColIndex(oHeaders, ['สถานะ', 'status']);
      const idxONote = findColIndex(oHeaders, ['หมายเหตุ', 'note', 'remark']);

      if (idxOId >= 0) {
        for (let i = 1; i < oData.length; i++) {
          if (String(oData[i][idxOId]).trim() === orderId) {
            const rowNum = i + 1;
            if (idxOStatus >= 0) {
              ordersSheet.getRange(rowNum, idxOStatus + 1).setValue('ยกเลิกแล้ว').setBackground('#f8d7da');
            } else if (idxONote >= 0) {
              const noteVal = String(oData[i][idxONote] || '');
              if (!noteVal.includes('[ยกเลิกแล้ว]')) {
                ordersSheet.getRange(rowNum, idxONote + 1).setValue(noteVal ? `[ยกเลิกแล้ว] ${noteVal}` : '[ยกเลิกแล้ว]');
              }
            }
          }
        }
      }

      return output({ success: true, message: `ยกเลิกออร์เดอร์ ${count} ใบเรียบร้อยแล้ว` });
    }

    // 2.3 จัดการกู้คืนคำสั่งซื้อ (เปลี่ยนสถานะตั๋วทุกใบของออร์เดอร์กลับเป็น "ยังไม่เช็คอิน")
    if (data.action === 'restoreOrder') {
      const ss = getSS();
      const ticketsSheet = getOrCreateSheet(ss, SHEET_TICKETS, []);
      const ordersSheet = getOrCreateSheet(ss, SHEET_ORDERS, []);
      const orderId = String(data.orderId || '').trim();

      const tData = ticketsSheet.getDataRange().getValues();
      const tHeaders = tData[0] || [];
      const idxTOId = findColIndex(tHeaders, ['เลขที่คำสั่งซื้อ', 'orderid', 'order id']) + 1;
      const idxTStatus = findColIndex(tHeaders, ['สถานะเช็คอิน', 'สถานะ', 'status']) + 1;
      const idxTTime = findColIndex(tHeaders, ['เวลาเช็คอิน', 'time']) + 1;

      let count = 0;
      if (idxTOId > 0 && idxTStatus > 0) {
        for (let i = 1; i < tData.length; i++) {
          if (String(tData[i][idxTOId - 1]).trim() === orderId) {
            const rowNum = i + 1;
            const statusCell = ticketsSheet.getRange(rowNum, idxTStatus);
            statusCell.setValue('ยังไม่เช็คอิน');
            statusCell.setBackground('#ffffff');
            if (idxTTime > 0) {
              ticketsSheet.getRange(rowNum, idxTTime).setValue('');
            }
            count++;
          }
        }
      }

      // ปลดสถานะยกเลิกใน Orders sheet
      const oData = ordersSheet.getDataRange().getValues();
      const oHeaders = oData[0] || [];
      const idxOId = findColIndex(oHeaders, ['เลขที่คำสั่งซื้อ', 'orderid', 'order id']);
      const idxOStatus = findColIndex(oHeaders, ['สถานะ', 'status']);
      const idxONote = findColIndex(oHeaders, ['หมายเหตุ', 'note', 'remark']);

      if (idxOId >= 0) {
        for (let i = 1; i < oData.length; i++) {
          if (String(oData[i][idxOId]).trim() === orderId) {
            const rowNum = i + 1;
            if (idxOStatus >= 0) {
              ordersSheet.getRange(rowNum, idxOStatus + 1).setValue('ปกติ').setBackground('#ffffff');
            } else if (idxONote >= 0) {
              const noteVal = String(oData[i][idxONote] || '').replace('[ยกเลิกแล้ว]', '').trim();
              ordersSheet.getRange(rowNum, idxONote + 1).setValue(noteVal);
            }
          }
        }
      }

      return output({ success: true, message: `กู้คืนออร์เดอร์ ${count} ใบเรียบร้อยแล้ว` });
    }

    // 2.4 จัดการลบตั๋วถาวร (ลบแถวออกจาก Google Sheets เมื่อกด ลบถาวร)
    if (data.action === 'deleteTicket') {
      const ss = getSS();
      const ticketsSheet = getOrCreateSheet(ss, SHEET_TICKETS, []);
      const ticketId = String(data.ticketId || '').trim();
      let orderId = String(data.orderId || '').trim();

      // 1. ลบแถวใน Tickets sheet (ลบทุกแถวที่ตรงกับ ticketId)
      const tData = ticketsSheet.getDataRange().getValues();
      const tHeaders = tData[0] || [];
      const idxTId = findColIndex(tHeaders, ['รหัสบัตร', 'ticketid', 'ticket id']);
      const idxTOId = findColIndex(tHeaders, ['เลขที่คำสั่งซื้อ', 'orderid', 'order id']);

      if (idxTId >= 0) {
        for (let i = tData.length - 1; i >= 1; i--) {
          if (String(tData[i][idxTId]).trim() === ticketId) {
            if (!orderId && idxTOId >= 0) {
              orderId = String(tData[i][idxTOId]).trim();
            }
            ticketsSheet.deleteRow(i + 1);
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
        const idxOTotal = findColIndex(oHeaders, ['ราคารวม', 'ยอดเงินรวม (บาท)', 'ยอดรวม', 'ยอดบัตรรวม', 'ยอดเงินรวม', 'ยอดชำระ', 'ราคาสุทธิ', 'ยอดสุทธิ', 'total', 'totalprice', 'amount']);
        const idxOPrice = findColIndex(oHeaders, ['ราคาต่อใบ', 'ราคาต่อใบ (บาท)', 'ราคา', 'priceperticket', 'price']);
        const idxOTkts = findColIndex(oHeaders, ['รหัสบัตรทั้งหมด', 'รหัสตั๋วทั้งหมด', 'ticketids', 'tickets']);

        if (idxOId >= 0) {
          for (let i = oData.length - 1; i >= 1; i--) {
            if (String(oData[i][idxOId]).trim() === orderId) {
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
            }
          }
        }
      }
      cleanGhostOrders(ss);
      return output({ success: true, message: 'ลบตั๋วออกจาก Google Sheets เรียบร้อยแล้ว' });
    }

    // 2.5 จัดการลบคำสั่งซื้อถาวร (ลบตั๋วทุกใบและออร์เดอร์ออกจาก Google Sheets เมื่อกด ลบถาวร)
    if (data.action === 'deleteOrder') {
      const ss = getSS();
      const ordersSheet = getOrCreateSheet(ss, SHEET_ORDERS, []);
      const ticketsSheet = getOrCreateSheet(ss, SHEET_TICKETS, []);
      const orderId = String(data.orderId || '').trim();

      if (orderId) {
        // 1. ลบตั๋วทั้งหมดของออร์เดอร์นี้ใน Tickets sheet
        const tData = ticketsSheet.getDataRange().getValues();
        const tHeaders = tData[0] || [];
        const idxTOId = findColIndex(tHeaders, ['เลขที่คำสั่งซื้อ', 'orderid', 'order id']);
        const idxTId = findColIndex(tHeaders, ['รหัสบัตร', 'ticketid', 'ticket id']);

        for (let i = tData.length - 1; i >= 1; i--) {
          const rowOId = idxTOId >= 0 ? String(tData[i][idxTOId] || '').trim() : '';
          const rowTId = idxTId >= 0 ? String(tData[i][idxTId] || '').trim() : '';
          if (rowOId === orderId || rowTId.startsWith(orderId)) {
            ticketsSheet.deleteRow(i + 1);
          }
        }

        // 2. ลบแถวใน Orders sheet (ลบทุกแถวที่ตรงกับ orderId ป้องกันแถวซ้ำ)
        const oData = ordersSheet.getDataRange().getValues();
        const oHeaders = oData[0] || [];
        const idxOId = findColIndex(oHeaders, ['เลขที่คำสั่งซื้อ', 'orderid', 'order id']);

        if (idxOId >= 0) {
          for (let i = oData.length - 1; i >= 1; i--) {
            if (String(oData[i][idxOId] || '').trim() === orderId) {
              ordersSheet.deleteRow(i + 1);
            }
          }
        }
      }
      cleanGhostOrders(ss);
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
      const idxOId = findColIndex(oHeaders, ['เลขที่คำสั่งซื้อ', 'orderid', 'order id']) + 1;
      const idxOName = findColIndex(oHeaders, ['ชื่อ-นามสกุล', 'name', 'fullname']) + 1;
      const idxOPhone = findColIndex(oHeaders, ['เบอร์โทร', 'phone', 'tel']) + 1;
      const idxOEmail = findColIndex(oHeaders, ['อีเมล', 'email']) + 1;
      const idxOType = findColIndex(oHeaders, ['ประเภทบัตร', 'tickettype', 'type']) + 1;
      const idxOPrice = findColIndex(oHeaders, ['ราคาต่อใบ', 'price', 'priceperticket']) + 1;
      const idxOTotal = findColIndex(oHeaders, ['ราคารวม', 'ยอดเงินรวม (บาท)', 'ยอดรวม', 'ยอดบัตรรวม', 'ยอดเงินรวม', 'ยอดชำระ', 'ราคาสุทธิ', 'ยอดสุทธิ', 'total', 'totalprice', 'amount']) + 1;
      const idxONote = findColIndex(oHeaders, ['หมายเหตุ', 'note', 'remark']) + 1;
      const idxOShowDate = findColIndex(oHeaders, ['รอบการแสดง', 'showdate', 'round']) + 1;

      for (let i = 1; i < oData.length; i++) {
        if (idxOId > 0 && oData[i][idxOId - 1] === orderId) {
          const rowNum = i + 1;
          if (idxOName > 0) ordersSheet.getRange(rowNum, idxOName).setValue(data.name);
          if (idxOPhone > 0) ordersSheet.getRange(rowNum, idxOPhone).setValue(cleanPhoneVal);
          if (idxOEmail > 0) ordersSheet.getRange(rowNum, idxOEmail).setValue(data.email || '');
          if (idxOType > 0) ordersSheet.getRange(rowNum, idxOType).setValue(data.type);
          if (idxOPrice > 0) ordersSheet.getRange(rowNum, idxOPrice).setValue(data.pricePerTicket);
          if (idxOTotal > 0) ordersSheet.getRange(rowNum, idxOTotal).setValue(data.total);
          if (idxONote > 0) ordersSheet.getRange(rowNum, idxONote).setValue(data.note || '');
          if (idxOShowDate > 0) ordersSheet.getRange(rowNum, idxOShowDate).setValue(data.showDate || '');
          break;
        }
      }

      // อัปเดตใน Tickets sheet ทุกใบของออร์เดอร์นี้
      const tData = ticketsSheet.getDataRange().getValues();
      const tHeaders = tData[0] || [];
      const idxTOId = findColIndex(tHeaders, ['เลขที่คำสั่งซื้อ', 'orderid', 'order id']) + 1;
      const idxTName = findColIndex(tHeaders, ['ชื่อ-นามสกุล', 'name']) + 1;
      const idxTPhone = findColIndex(tHeaders, ['เบอร์โทร', 'phone', 'tel']) + 1;
      const idxTType = findColIndex(tHeaders, ['ประเภทบัตร', 'type']) + 1;
      const idxTShowDate = findColIndex(tHeaders, ['รอบการแสดง', 'showdate']) + 1;

      for (let i = 1; i < tData.length; i++) {
        if (idxTOId > 0 && tData[i][idxTOId - 1] === orderId) {
          const rowNum = i + 1;
          if (idxTName > 0) ticketsSheet.getRange(rowNum, idxTName).setValue(data.name);
          if (idxTPhone > 0) ticketsSheet.getRange(rowNum, idxTPhone).setValue(cleanPhoneVal);
          if (idxTType > 0) ticketsSheet.getRange(rowNum, idxTType).setValue(data.type);
          if (idxTShowDate > 0) ticketsSheet.getRange(rowNum, idxTShowDate).setValue(data.showDate || '');
        }
      }

      return output({ success: true });
    }

    // 5. บันทึกคำสั่งซื้อใหม่ (จองตั๋วปกติ หรือ โควต้าสตาฟ)
    if (data.action === 'newOrder' || (!data.action && (data.name || data.tickets))) {
      return handleNewOrder(data);
    }

    return output({ success: false, error: 'Unrecognized action: ' + (data.action || 'empty') });

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

    // 0. ตรวจสอบและทำความสะอาดข้อมูลอัตโนมัติ
    fixMojibakeStatuses(ss);
    cleanGhostOrders(ss);
    backfillOrderTotals(ss);
    cleanDuplicateTickets(ss);
    syncTicketsFromOrders(ss);

    if (action === 'cleanGhost' || action === 'fixTotals' || action === 'backfillTotals') {
      const fixedMojibake = fixMojibakeStatuses(ss);
      const ghostCount = cleanGhostOrders(ss);
      const fixedCount = backfillOrderTotals(ss);
      const cleanedCount = cleanDuplicateTickets(ss);
      const syncedCount = syncTicketsFromOrders(ss);
      const slipsCount = formatAllSlipLinks(ss);
      return output({
        success: true,
        message: `แก้คำเพี้ยน/ต่างดาว (${fixedMojibake} ช่อง), ลบแถวผีใน Orders (${ghostCount} แถว), อัปเดตยอดบัตรรวม (${fixedCount} รายการ), ลบตั๋วซ้ำ (${cleanedCount} รายการ), ซิงค์ตั๋ว (${syncedCount} ใบ) และจัดฟอร์แมตสลิป (${slipsCount} ช่อง) เรียบร้อยแล้ว`
      });
    }

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

      // ยกเลิกการจำกัดสิทธิ์เบอร์ซ้ำ (คืนค่าว่างเพื่อให้เครื่องที่แคชไว้ไม่บล็อกเบอร์)
      settings["usedPromoPhones"] = {
        'NMC600': [],
        'FB750': []
      };

      return output(settings);
    }

    // กรณีมีเครื่องลูกค้าที่ยังแคชโค้ดเก่าแล้วยิง checkPromo มา ให้ตอบว่าไม่เคยใช้เสมอ
    if (action === 'checkPromo') {
      return output({
        success: true,
        alreadyUsed: false
      });
    }
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
      const tickets = {};
      
      const ordersSheet = getOrCreateSheet(ss, SHEET_ORDERS, []);
      const oRange = ordersSheet.getDataRange();
      const oData = oRange.getValues();
      const oFormulas = oRange.getFormulas();
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
      const idxOQty = findColIndex(oHeaders, ['จำนวนใบ', 'จำนวน', 'qty', 'quantity', 'ยอดบัตรรวม (ใบ)', 'ยอดบัตรรวม']) + 1;
      const idxOPrice = findColIndex(oHeaders, ['ราคาต่อใบ', 'ราคา', 'priceperticket', 'price']) + 1;
      const idxOTotal = findColIndex(oHeaders, ['ราคารวม', 'ยอดรวม', 'total', 'amount', 'totalprice', 'ยอดเงินรวม', 'ยอดบัตรรวม', 'ยอดชำระ', 'ยอดสุทธิ', 'ราคาสุทธิ']) + 1;
      const idxONote = findColIndex(oHeaders, ['หมายเหตุ', 'note']) + 1;
      const idxOTickets = findColIndex(oHeaders, ['รหัสบัตรทั้งหมด', 'รหัสบัตร', 'ticketids', 'tickets']) + 1;
      const idxOSlip = findColIndex(oHeaders, ['สลิปการโอนเงิน', 'สลิปโอนเงิน', 'สลิป', 'หลักฐานการโอน', 'slipurl', 'slip url', 'slip_url', 'slip', 'sliplink']) + 1;
      const idxOShowDate = findColIndex(oHeaders, ['รอบการแสดง', 'รอบ', 'showdate', 'round']) + 1;

      // ค้นหาตำแหน่งคอลัมน์ของชีทตั๋วรายใบ (Tickets)
      const tRange = ticketsSheet.getDataRange();
      const tData = tRange.getValues();
      const tFormulas = tRange.getFormulas();
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
        const oName = idxOName ? String(row[idxOName - 1] || '').trim() : '';
        if (!oName || oName === '—') continue; // ข้ามแถวผีที่ไม่มีชื่อลูกค้า

        const idxOStatus = findColIndex(oHeaders, ['สถานะ', 'status']) + 1;
        const isOrderCancelled = (idxOStatus > 0 && String(row[idxOStatus - 1] || '').includes('ยกเลิก')) ||
                                 (idxONote > 0 && String(row[idxONote - 1] || '').includes('[ยกเลิกแล้ว]'));

        let slipVal = (idxOSlip ? row[idxOSlip - 1] : '') || '';
        const fVal = (idxOSlip && oFormulas[i]) ? oFormulas[i][idxOSlip - 1] : '';
        if (fVal && String(fVal).indexOf('http') >= 0) {
          const m = String(fVal).match(/https?:\/\/[^"',)]+/);
          if (m) slipVal = m[0];
        }
        const order = {
          orderId: row[idxOId - 1],
          timestamp: row[idxOTs - 1],
          name: oName,
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
          cancelled: isOrderCancelled
        };
        orders.push(order);
        ordersMap[order.orderId] = order;
      }

      // แมปตั๋วรายใบ พร้อมตัดรายการซ้ำออก
      const seenTicketKeys = new Set();
      for (let i = 1; i < tData.length; i++) {
        const row = tData[i];
        let tId = String(row[idxTId - 1] || '').trim();
        if (!tId) continue;
        const oId = String(row[idxTOId - 1] || '').trim();
        const parentOrder = ordersMap[oId] || {};
        
        const status = row[idxTStatus - 1];
        const isCancelled = status === 'ยกเลิกแล้ว' || String(status || '').indexOf('‡') >= 0;
        const isCheckedIn = status === 'เช็คอินแล้ว';

        let ticketNum = 1;
        const match = tId.match(/-T0*(\d+)$/);
        if (match) ticketNum = Number(match[1]);

        const stdTId = `${oId}-T${String(ticketNum).padStart(2, '0')}`;
        const dedupKey = `${oId}|${ticketNum}`;

        if (seenTicketKeys.has(dedupKey)) {
          // หากแถวที่เจอทีหลังมีการเช็คอิน ให้อัปเดตสถานะเช็คอินให้ตั๋วตัวจริง
          if (isCheckedIn && tickets[stdTId] && !tickets[stdTId].checkedIn) {
            tickets[stdTId].checkedIn = true;
            tickets[stdTId].checkInTime = row[idxTTime - 1] || null;
          }
          continue;
        }
        seenTicketKeys.add(dedupKey);

        let tSlip = (idxTSlip ? row[idxTSlip - 1] : '') || parentOrder.slipUrl || parentOrder.slipImage || '';
        const tfVal = (idxTSlip && tFormulas[i]) ? tFormulas[i][idxTSlip - 1] : '';
        if (tfVal && String(tfVal).indexOf('http') >= 0) {
          const m = String(tfVal).match(/https?:\/\/[^"',)]+/);
          if (m) tSlip = m[0];
        }

        tickets[stdTId] = {
          ticketId: stdTId,
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

      // ตรวจสอบความสมบูรณ์ หากใน Tickets ยังไม่มีข้อมูล ให้สร้าง ticket เสมือนขึ้นมาทันทีเพื่อแสดงในหน้า Staff
      orders.forEach(o => {
        const oTkts = Object.values(tickets).filter(t => t.orderId === o.orderId);
        const targetQty = o.qty > 0 ? o.qty : 1;
        if (oTkts.length < targetQty) {
          for (let k = 1; k <= targetQty; k++) {
            const tid = `${o.orderId}-T${String(k).padStart(2, '0')}`;
            if (!tickets[tid]) {
              tickets[tid] = {
                ticketId: tid,
                ticketNum: k,
                orderId: o.orderId,
                name: o.name,
                nickname: o.nickname || '',
                phone: o.phone,
                email: o.email,
                note: o.note,
                type: o.typeName,
                showDate: o.showDate,
                pricePerTicket: o.pricePerTicket,
                total: o.total,
                qty: targetQty,
                checkedIn: false,
                checkInTime: null,
                cancelled: o.cancelled || false,
                cancelledAt: null,
                slipUrl: o.slipUrl,
                slipImage: o.slipImage
              };
            }
          }
        }
      });

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
      const oRange = ordersSheet.getDataRange();
      const oData = oRange.getValues();
      const oFormulas = oRange.getFormulas();
      const oHeaders = oData[0] || [];
      const idxOId = findColIndex(oHeaders, ['เลขที่คำสั่งซื้อ', 'เลขที่ออเดอร์', 'เลขออเดอร์', 'orderid', 'order id']);
      const idxOPhone = findColIndex(oHeaders, ['เบอร์โทร', 'เบอร์โทรศัพท์', 'เบอร์', 'phone', 'tel', 'telephone', 'mobile']);

      const ticketsSheet = getOrCreateSheet(ss, SHEET_TICKETS, []);
      const tRange = ticketsSheet.getDataRange();
      const tData = tRange.getValues();
      const tFormulas = tRange.getFormulas();
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
      const idxOQty = findColIndex(oHeaders, ['จำนวนใบ', 'จำนวนบัตร (ใบ)', 'จำนวนบัตร', 'จำนวน', 'qty', 'ยอดบัตรรวม (ใบ)', 'ยอดบัตรรวม']);
      const idxOPrice = findColIndex(oHeaders, ['ราคาต่อใบ (บาท)', 'ราคาต่อใบ', 'ราคา', 'priceperticket', 'price']);
      const idxOTotal = findColIndex(oHeaders, ['ราคารวม', 'ยอดเงินรวม (บาท)', 'ยอดรวม', 'total', 'amount', 'totalprice', 'ยอดเงินรวม', 'ยอดบัตรรวม', 'ยอดชำระ', 'ยอดสุทธิ', 'ราคาสุทธิ']);
      const idxONote = findColIndex(oHeaders, ['หมายเหตุ', 'note']);
      const idxOTickets = findColIndex(oHeaders, ['รหัสบัตรทั้งหมด', 'รหัสตั๋วทั้งหมด', 'ticketids', 'tickets']);
      const idxOSlip = findColIndex(oHeaders, ['สลิปการโอนเงิน', 'สลิปโอนเงิน', 'สลิป', 'หลักฐานการโอน', 'slipurl', 'slip url', 'slip_url', 'slip', 'sliplink']);
      const idxOShowDate = findColIndex(oHeaders, ['รอบการแสดง', 'รอบ', 'showdate', 'round']);

      const ordersMap = {};
      for (let i = 1; i < oData.length; i++) {
        const row = oData[i];
        const oId = idxOId >= 0 ? row[idxOId] : '';
        if (oId && matchedOrderIds.indexOf(oId) >= 0) {
          let slipVal = (idxOSlip >= 0 ? row[idxOSlip] : '') || '';
          const fVal = (idxOSlip >= 0 && oFormulas[i]) ? oFormulas[i][idxOSlip] : '';
          if (fVal && String(fVal).indexOf('http') >= 0) {
            const m = String(fVal).match(/https?:\/\/[^"',)]+/);
            if (m) slipVal = m[0];
          }
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

          let tSlip = (idxTSlip >= 0 ? row[idxTSlip] : '') || parentOrder.slipUrl || parentOrder.slipImage || '';
          const tfVal = (idxTSlip >= 0 && tFormulas[i]) ? tFormulas[i][idxTSlip] : '';
          if (tfVal && String(tfVal).indexOf('http') >= 0) {
            const m = String(tfVal).match(/https?:\/\/[^"',)]+/);
            if (m) tSlip = m[0];
          }

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

  assignOrderCol(['เบอร์โทร', 'phone', 'tel', 'mobile'], cleanPhoneVal);
  assignOrderCol(['อีเมล', 'email'], data.email || '—');
  assignOrderCol(['ประเภทบัตร', 'tickettype', 'type'], data.ticketType);
  assignOrderCol(['จำนวนใบ', 'จำนวนบัตร (ใบ)', 'จำนวนบัตร', 'จำนวนตั๋ว', 'จำนวน', 'ยอดบัตรรวม (ใบ)', 'ยอดบัตรรวม', 'qty', 'quantity', 'ticketcount'], data.qty);
  assignOrderCol(['ราคาต่อใบ', 'ราคาต่อใบ (บาท)', 'ราคา', 'price', 'priceperticket', 'unitprice'], data.pricePerTicket);

  let subtotalIdx = findColIndex(oHeaders, ['ราคาก่อนลด', 'ยอดรวมก่อนลด', 'subtotal']);
  if (subtotalIdx >= 0) rowData[subtotalIdx] = data.subtotal || data.total;

  let discIdx = findColIndex(oHeaders, ['ส่วนลด', 'discount']);
  if (discIdx >= 0) rowData[discIdx] = data.discount || 0;

  let promoIdx = findColIndex(oHeaders, ['โค้ดส่วนลด', 'promocode', 'promo code', 'promo']);
  if (promoIdx >= 0) rowData[promoIdx] = data.promoCode || '—';

  const finalTotal = data.total !== undefined ? Number(data.total) : (Number(data.qty || 1) * Number(data.pricePerTicket || 0));
  let totalIdx = findColIndex(oHeaders, ['ราคารวม', 'ยอดเงินรวม (บาท)', 'ยอดรวม', 'ยอดบัตรรวม', 'ยอดเงินรวม', 'ยอดชำระ', 'ราคาสุทธิ', 'ยอดสุทธิ', 'total', 'totalprice', 'amount', 'totalamount', 'grandtotal']);
  if (totalIdx < 0) {
    const newCol = ordersSheet.getLastColumn() + 1;
    ordersSheet.getRange(1, newCol).setValue('TotalPrice');
    ordersSheet.getRange(1, newCol).setBackground('#4a2080').setFontColor('#ffffff').setFontWeight('bold');
    oHeaders.push('TotalPrice');
    totalIdx = newCol - 1;
  }
  rowData[totalIdx] = finalTotal;

  assignOrderCol(['รอบการแสดง', 'showdate', 'รอบ', 'round'], data.showDate || '—');
  assignOrderCol(['หมายเหตุ', 'note', 'remark'], data.note || '—');
  assignOrderCol(['รหัสบัตรทั้งหมด', 'tickets', 'ticketids', 'รหัสตั๋ว'], data.tickets);
  
  let slipIdx = findColIndex(oHeaders, ['สลิปการโอนเงิน', 'สลิปโอนเงิน', 'สลิป', 'หลักฐานการโอน', 'หลักฐานการโอนเงิน', 'ลิงก์สลิป', 'ลิงค์สลิป', 'slipurl', 'slip url', 'slip_url', 'slip', 'sliplink']);
  if (slipIdx < 0) {
    const newCol = ordersSheet.getLastColumn() + 1;
    ordersSheet.getRange(1, newCol).setValue('สลิปการโอนเงิน');
    ordersSheet.getRange(1, newCol).setBackground('#4a2080').setFontColor('#ffffff').setFontWeight('bold');
    oHeaders.push('สลิปการโอนเงิน');
    slipIdx = newCol - 1;
  }
  const slipDisplay = (slipUrl && String(slipUrl).startsWith('http'))
    ? `=HYPERLINK("${slipUrl}", "📄 ดูรูปสลิป")`
    : slipUrl;
  rowData[slipIdx] = slipDisplay;

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
    assignTCol(['จำนวนบัตรรวม', 'จำนวนใบ', 'จำนวน', 'qty'], data.qty);
    assignTCol(['ยอดบัตรรวม', 'ยอดรวม', 'ยอดเงินรวม', 'ราคารวม', 'total', 'totalprice', 'amount'], finalTotal);
    tRowData[tSlipIdx] = slipDisplay;

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

// ─── AUTO-BACKFILL / FIX MISSING TOTALS ────────────────────────────────────
// ฟังก์ชันตรวจสอบและเติมค่า TotalPrice/ยอดรวม ให้ทุกออเดอร์ใน Orders sheet โดยอัตโนมัติ
function backfillOrderTotals(ss) {
  let count = 0;
  try {
    const ordersSheet = ss.getSheetByName(SHEET_ORDERS);
    if (!ordersSheet) return 0;
    const data = ordersSheet.getDataRange().getValues();
    if (data.length <= 1) return 0;
    const headers = data[0] || [];
    
    let idxTotal = findColIndex(headers, ['totalprice', 'total', 'ราคารวม', 'ยอดรวม', 'ยอดบัตรรวม', 'ยอดเงินรวม', 'ยอดชำระ', 'ราคาสุทธิ', 'ยอดสุทธิ', 'amount']);
    const idxQty = findColIndex(headers, ['qty', 'จำนวนใบ', 'จำนวนบัตร', 'จำนวน', 'ยอดบัตรรวม']);
    const idxPrice = findColIndex(headers, ['price', 'priceperticket', 'ราคาต่อใบ', 'ราคา']);
    const idxType = findColIndex(headers, ['tickettype', 'type', 'ประเภทบัตร', 'ประเภท']);
    const idxNote = findColIndex(headers, ['note', 'หมายเหตุ', 'remark']);

    if (idxTotal < 0) {
      // หากยังไม่มีคอลัมน์ TotalPrice ให้สร้างขึ้นมาอัตโนมัติ
      const newCol = ordersSheet.getLastColumn() + 1;
      ordersSheet.getRange(1, newCol).setValue('TotalPrice');
      ordersSheet.getRange(1, newCol).setBackground('#4a2080').setFontColor('#ffffff').setFontWeight('bold');
      idxTotal = newCol - 1;
    }

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const currentVal = row[idxTotal];
      // เติมเฉพาะแถวที่ยังว่างอยู่
      if (currentVal === '' || currentVal === null || currentVal === undefined) {
        const qty = idxQty >= 0 ? Number(row[idxQty]) || 1 : 1;
        let price = idxPrice >= 0 ? Number(row[idxPrice]) || 0 : 0;
        
        const typeStr = idxType >= 0 ? String(row[idxType] || '').toUpperCase() : '';
        const isQuota = typeStr.includes('โควต้า') || typeStr.includes('STAFF') || typeStr.includes('QUOTA');
        
        if (price <= 0 && !isQuota) {
          if (typeStr.includes('EARLY')) price = 500;
          else if (typeStr.includes('STUDENT')) price = 550;
          else if (typeStr.includes('REGULAR')) price = 900;
        }

        let calculatedTotal = isQuota ? 0 : qty * price;
        // ตรวจสอบส่วนลดเพิ่มเติมจากหมายเหตุ (เช่น NMC600, FB750, NMC300)
        if (idxNote >= 0) {
          const noteStr = String(row[idxNote] || '');
          if (noteStr.includes('NMC600')) {
            calculatedTotal = Math.max(0, calculatedTotal - 300);
          } else if (noteStr.includes('FB750')) {
            calculatedTotal = Math.max(0, calculatedTotal - 150);
          } else if (noteStr.includes('ลด 300') || noteStr.includes('NMC300')) {
            calculatedTotal = Math.max(0, calculatedTotal - 300);
          }
        }

        ordersSheet.getRange(i + 1, idxTotal + 1).setValue(calculatedTotal);
        count++;
      }
    }
  } catch (err) {
    console.warn('backfillOrderTotals error:', err);
  }
  return count;
}

// ฟังก์ชันแปลงลิงก์สลิปทั้งหมดในชีท Orders และ Tickets ให้แสดงผลเป็นปุ่มข้อความ "📄 ดูรูปสลิป"
function formatAllSlipLinks(ss) {
  let count = 0;
  try {
    [SHEET_ORDERS, SHEET_TICKETS].forEach(sheetName => {
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet) return;
      const data = sheet.getDataRange().getValues();
      if (data.length <= 1) return;
      const headers = data[0] || [];
      const slipIdx = findColIndex(headers, ['สลิปการโอนเงิน', 'สลิปโอนเงิน', 'สลิป', 'หลักฐานการโอน', 'slipurl', 'slip url', 'slip']);
      if (slipIdx < 0) return;

      for (let i = 1; i < data.length; i++) {
        const val = String(data[i][slipIdx] || '').trim();
        if (val.startsWith('http://') || val.startsWith('https://')) {
          sheet.getRange(i + 1, slipIdx + 1).setFormula(`=HYPERLINK("${val}", "📄 ดูรูปสลิป")`);
          count++;
        }
      }
    });
  } catch (err) {
    console.warn('formatAllSlipLinks error:', err);
  }
  return count;
}

// ฟังก์ชันตรวจสอบและเติมตั๋วที่ขาดหายไปลงชีท Tickets จาก Orders sheet อัตโนมัติ
function syncTicketsFromOrders(ss) {
  let addedCount = 0;
  try {
    const ordersSheet = ss.getSheetByName(SHEET_ORDERS);
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
    if (!ordersSheet) return 0;

    const oData = ordersSheet.getDataRange().getValues();
    if (oData.length <= 1) return 0;
    const oHeaders = oData[0] || [];

    const idxOId = findColIndex(oHeaders, ['เลขที่คำสั่งซื้อ', 'เลขที่ออเดอร์', 'เลขออเดอร์', 'orderid', 'order id']);
    if (idxOId < 0) return 0;

    const idxOName = findColIndex(oHeaders, ['ชื่อ-นามสกุล', 'ชื่อนามสกุล', 'ชื่อ', 'name', 'fullname']);
    const idxONick = findColIndex(oHeaders, ['ชื่อเล่น', 'nickname', 'nick']);
    const idxOPhone = findColIndex(oHeaders, ['เบอร์โทร', 'เบอร์โทรศัพท์', 'เบอร์', 'phone', 'tel']);
    const idxOType = findColIndex(oHeaders, ['ประเภทบัตร', 'ประเภท', 'tickettype', 'type']);
    const idxOQty = findColIndex(oHeaders, ['จำนวนใบ', 'จำนวน', 'qty', 'quantity', 'ยอดบัตรรวม (ใบ)', 'ยอดบัตรรวม']);
    const idxOShowDate = findColIndex(oHeaders, ['รอบการแสดง', 'รอบ', 'showdate', 'round']);
    const idxOTickets = findColIndex(oHeaders, ['รหัสบัตรทั้งหมด', 'รหัสบัตร', 'ticketids', 'tickets']);
    const idxOSlip = findColIndex(oHeaders, ['สลิปการโอนเงิน', 'สลิปโอนเงิน', 'สลิป', 'หลักฐานการโอน', 'slipurl', 'slip url', 'slip']);
    const idxOTotal = findColIndex(oHeaders, ['totalprice', 'total', 'ราคารวม', 'ยอดรวม', 'ยอดบัตรรวม', 'ยอดเงินรวม', 'amount']);
    const idxOStatus = findColIndex(oHeaders, ['สถานะ', 'status']);

    const oFormulas = ordersSheet.getDataRange().getFormulas();

    // รวบรวมตั๋วที่มีอยู่แล้วใน Tickets sheet ตามคีย์ `${orderId}|${ticketNum}`
    const tData = ticketsSheet.getDataRange().getValues();
    const tHeaders = tData[0] || [];
    const idxTId = findColIndex(tHeaders, ['รหัสบัตร', 'ticketid', 'ticket id']);
    const idxTOId = findColIndex(tHeaders, ['เลขที่คำสั่งซื้อ', 'orderid', 'order id']);
    const existingTicketKeys = new Set();
    if (idxTId >= 0) {
      for (let i = 1; i < tData.length; i++) {
        const idVal = String(tData[i][idxTId] || '').trim();
        if (!idVal) continue;
        const oId = idxTOId >= 0 ? String(tData[i][idxTOId] || '').trim() : '';
        const m = idVal.match(/-T0*(\d+)$/);
        const num = m ? Number(m[1]) : 1;
        const realOId = oId || idVal.replace(/-T\d+$/, '');
        if (realOId) existingTicketKeys.add(`${realOId}|${num}`);
      }
    }

    const rowsToAdd = [];
    for (let i = 1; i < oData.length; i++) {
      const oRow = oData[i];
      const orderId = String(oRow[idxOId] || '').trim();
      if (!orderId) continue;

      const qty = idxOQty >= 0 ? Math.max(1, Number(oRow[idxOQty]) || 1) : 1;
      const name = idxOName >= 0 ? String(oRow[idxOName] || '').trim() : '';
      if (!name || name === '—') continue; // ข้ามแถวผี ไม่ต้องเติมตั๋ว
      const nick = idxONick >= 0 ? String(oRow[idxONick] || '').trim() : '';
      const phone = idxOPhone >= 0 ? formatPhoneToWrite(oRow[idxOPhone]) : '';
      const type = idxOType >= 0 ? String(oRow[idxOType] || '').trim() : 'REGULAR';
      const showDate = idxOShowDate >= 0 ? String(oRow[idxOShowDate] || '').trim() : '';
      const total = idxOTotal >= 0 ? Number(oRow[idxOTotal]) || 0 : 0;
      const status = idxOStatus >= 0 && String(oRow[idxOStatus] || '').includes('ยกเลิก') ? 'ยกเลิกแล้ว' : 'ยังไม่เช็คอิน';

      // สลิป
      let slipVal = idxOSlip >= 0 ? (oRow[idxOSlip] || '') : '';
      const fVal = (idxOSlip >= 0 && oFormulas[i]) ? oFormulas[i][idxOSlip] : '';
      if (fVal && String(fVal).indexOf('http') >= 0) {
        slipVal = fVal;
      } else if (String(slipVal).startsWith('http')) {
        slipVal = `=HYPERLINK("${slipVal}", "📄 ดูรูปสลิป")`;
      }

      for (let k = 1; k <= qty; k++) {
        const key = `${orderId}|${k}`;
        if (!existingTicketKeys.has(key)) {
          const tid = `${orderId}-T${String(k).padStart(2, '0')}`;
          const newTRow = [];
          for (let c = 0; c < tHeaders.length; c++) newTRow.push('');

          const assign = (names, val) => {
            const cIdx = findColIndex(tHeaders, names);
            if (cIdx >= 0) newTRow[cIdx] = val;
          };

          assign(['รหัสบัตร', 'ticketid'], tid);
          assign(['เลขที่คำสั่งซื้อ', 'orderid'], orderId);
          assign(['ชื่อ-นามสกุล', 'name'], name);
          assign(['ชื่อเล่น', 'nickname', 'nick'], nick);
          assign(['เบอร์โทร', 'phone'], phone);
          assign(['ประเภทบัตร', 'type'], type);
          assign(['รอบการแสดง', 'showdate'], showDate);
          assign(['สถานะเช็คอิน', 'status'], status);
          assign(['เวลาเช็คอิน', 'time'], '');
          assign(['สลิปการโอนเงิน', 'สลิป', 'slipurl', 'slip url', 'slip'], slipVal);
          assign(['จำนวนบัตรรวม', 'จำนวนใบ', 'จำนวน', 'qty'], qty);
          assign(['ยอดบัตรรวม', 'ยอดรวม', 'ยอดเงินรวม', 'ราคารวม', 'total', 'totalprice'], total);

          rowsToAdd.push(newTRow);
          existingTicketKeys.add(key);
          addedCount++;
        }
      }
    }

    if (rowsToAdd.length > 0) {
      for (let r = 0; r < rowsToAdd.length; r++) {
        ticketsSheet.appendRow(rowsToAdd[r]);
      }
    }
  } catch (err) {
    console.warn('syncTicketsFromOrders error:', err);
  }
  return addedCount;
}

// ฟังก์ชันลบตั๋วที่ซ้ำกันในชีท Tickets (เช่น ตั๋วที่มีทั้ง -T01 และ -T1 หรือรายการที่เบิ้ลขึ้นมา)
function cleanDuplicateTickets(ss) {
  let deletedCount = 0;
  try {
    const ticketsSheet = ss.getSheetByName(SHEET_TICKETS);
    if (!ticketsSheet) return 0;
    const tData = ticketsSheet.getDataRange().getValues();
    if (tData.length <= 1) return 0;
    const tHeaders = tData[0] || [];
    const idxTId = findColIndex(tHeaders, ['รหัสบัตร', 'ticketid', 'ticket id']);
    const idxTOId = findColIndex(tHeaders, ['เลขที่คำสั่งซื้อ', 'orderid', 'order id']);
    const idxTStatus = findColIndex(tHeaders, ['สถานะเช็คอิน', 'สถานะ', 'status']);
    if (idxTId < 0) return 0;

    const bestRows = new Map(); // key -> { rowNum, rawTId, isCheckedIn, hasPad }

    for (let i = 1; i < tData.length; i++) {
      const rowNum = i + 1;
      const rawTId = String(tData[i][idxTId] || '').trim();
      if (!rawTId) continue;
      const oId = idxTOId >= 0 ? String(tData[i][idxTOId] || '').trim() : '';
      
      const m = rawTId.match(/-T0*(\d+)$/);
      const ticketNum = m ? Number(m[1]) : 1;
      const realOId = oId || rawTId.replace(/-T\d+$/, '');
      const key = `${realOId}|${ticketNum}`;
      const isCheckedIn = idxTStatus >= 0 && String(tData[i][idxTStatus] || '') === 'เช็คอินแล้ว';
      const hasPad = !!rawTId.match(/-T\d{2,}$/); // เช่น -T01

      if (!bestRows.has(key)) {
        bestRows.set(key, { rowNum, rawTId, isCheckedIn, hasPad });
      } else {
        const existing = bestRows.get(key);
        if (isCheckedIn && !existing.isCheckedIn) {
          bestRows.set(key, { rowNum, rawTId, isCheckedIn, hasPad });
        } else if (!existing.isCheckedIn && hasPad && !existing.hasPad) {
          bestRows.set(key, { rowNum, rawTId, isCheckedIn, hasPad });
        }
      }
    }

    const rowsToKeep = new Set();
    for (const item of bestRows.values()) {
      rowsToKeep.add(item.rowNum);
    }

    // ลบแถวซ้ำจากล่างขึ้นบน เพื่อไม่ให้กระทบตำแหน่ง index แถว
    for (let i = tData.length - 1; i >= 1; i--) {
      const rowNum = i + 1;
      const rawTId = String(tData[i][idxTId] || '').trim();
      if (!rawTId) {
        ticketsSheet.deleteRow(rowNum);
        deletedCount++;
        continue;
      }
      if (!rowsToKeep.has(rowNum)) {
        ticketsSheet.deleteRow(rowNum);
        deletedCount++;
      }
    }
  } catch (err) {
    console.warn('cleanDuplicateTickets error:', err);
  }
  return deletedCount;
}

// ฟังก์ชันแก้คำเพี้ยน/ภาษาต่างดาวในช่องสถานะ (เช่น ‡∏¢... ให้เป็น 'ยกเลิกแล้ว')
function fixMojibakeStatuses(ss) {
  let fixedCount = 0;
  try {
    const sheetsToCheck = [SHEET_TICKETS, SHEET_ORDERS];
    sheetsToCheck.forEach(sheetName => {
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet) return;
      const data = sheet.getDataRange().getValues();
      if (data.length <= 1) return;
      const headers = data[0] || [];
      const idxStatus = findColIndex(headers, ['สถานะเช็คอิน', 'สถานะ', 'status']);
      if (idxStatus < 0) return;

      for (let i = 1; i < data.length; i++) {
        const val = String(data[i][idxStatus] || '').trim();
        if (val.indexOf('‡') >= 0 || val.indexOf('â') >= 0 || val.indexOf('∏') >= 0) {
          const cell = sheet.getRange(i + 1, idxStatus + 1);
          cell.setValue('ยกเลิกแล้ว');
          cell.setBackground('#f8d7da');
          fixedCount++;
        }
      }
    });
  } catch (err) {
    console.warn('fixMojibakeStatuses error:', err);
  }
  return fixedCount;
}

// ฟังก์ชันลบแถวผี/แถวว่าง/แถวซ้ำที่ตกค้างใน Orders sheet (เช่น แถวที่มีแต่เลขออร์เดอร์แต่ไม่มีชื่อ หรือยอดเป็น 0)
function cleanGhostOrders(ss) {
  let deletedCount = 0;
  try {
    const ordersSheet = ss.getSheetByName(SHEET_ORDERS);
    if (!ordersSheet) return 0;
    const oData = ordersSheet.getDataRange().getValues();
    if (oData.length <= 1) return 0;
    const oHeaders = oData[0] || [];

    const idxOId = findColIndex(oHeaders, ['เลขที่คำสั่งซื้อ', 'orderid', 'order id']);
    const idxOName = findColIndex(oHeaders, ['ชื่อ-นามสกุล', 'name', 'fullname']);
    const idxOPhone = findColIndex(oHeaders, ['เบอร์โทร', 'phone', 'tel']);
    const idxOQty = findColIndex(oHeaders, ['จำนวนใบ', 'qty']);
    const idxOTotal = findColIndex(oHeaders, ['ราคารวม', 'ยอดเงินรวม (บาท)', 'ยอดรวม', 'total', 'totalprice', 'amount']);

    if (idxOId < 0) return 0;

    // รวบรวมออร์เดอร์ที่ถูกต้องสมบูรณ์ (มีชื่อ และไม่ใช่แถวว่าง)
    const validOrders = new Set();
    for (let i = 1; i < oData.length; i++) {
      const oId = String(oData[i][idxOId] || '').trim();
      const name = idxOName >= 0 ? String(oData[i][idxOName] || '').trim() : '';
      if (oId && name && name !== '—') {
        validOrders.add(oId);
      }
    }

    // ลบแถวที่ผิดปกติจากล่างขึ้นบน
    for (let i = oData.length - 1; i >= 1; i--) {
      const rowNum = i + 1;
      const oId = String(oData[i][idxOId] || '').trim();
      const name = idxOName >= 0 ? String(oData[i][idxOName] || '').trim() : '';
      const phone = idxOPhone >= 0 ? String(oData[i][idxOPhone] || '').trim() : '';
      const qty = idxOQty >= 0 ? Number(oData[i][idxOQty]) || 0 : 0;
      const total = idxOTotal >= 0 ? Number(oData[i][idxOTotal]) || 0 : 0;

      // เงื่อนไขแถวผี:
      // 1. ไม่มีเลขออร์เดอร์
      // 2. ไม่มีชื่อ (หรือชื่อเป็น '—') และ (ยอดรวมเป็น 0 หรือไม่มีเบอร์ หรือจำนวนเป็น 0)
      // 3. เป็นแถวซ้ำของ orderId เดียวกันที่ไม่มีชื่อ ในขณะที่มีแถวหลักที่มีชื่ออยู่แล้ว
      const isNoId = !oId;
      const isEmptyNameGhost = (!name || name === '—') && (total === 0 || !phone || phone === '—' || qty <= 0);
      const isDuplicateGhost = (!name || name === '—') && validOrders.has(oId);

      if (isNoId || isEmptyNameGhost || isDuplicateGhost) {
        ordersSheet.deleteRow(rowNum);
        deletedCount++;
      }
    }
  } catch (err) {
    console.warn('cleanGhostOrders error:', err);
  }
  return deletedCount;
}

// ฟังก์ชันล้างข้อมูลผีและจัดระเบียบชีททั้งหมด (สามารถกดรันใน Apps Script Editor ได้ทันที)
function cleanGhostData() {
  const ss = getSS();
  const fixedMojibake = fixMojibakeStatuses(ss);
  const ghostCount = cleanGhostOrders(ss);
  const dupCount = cleanDuplicateTickets(ss);
  const totalsCount = backfillOrderTotals(ss);
  const slipsCount = formatAllSlipLinks(ss);
  const syncCount = syncTicketsFromOrders(ss);
  Logger.log(`แก้คำเพี้ยน/ต่างดาว: ${fixedMojibake} ช่อง, ลบแถวผีใน Orders: ${ghostCount} แถว, ลบตั๋วซ้ำ: ${dupCount} ใบ, เติมยอด: ${totalsCount}, แปลงสลิป: ${slipsCount}, ซิงค์ตั๋ว: ${syncCount}`);
}

// ฟังก์ชันสำหรับกดรันใน Google Apps Script Editor โดยตรง เพื่อเติมยอดรวม ลบตั๋วซ้ำ และแปลงสลิปทั้งหมด
function fixOrderTotals() {
  const ss = getSS();
  const fixedMojibake = fixMojibakeStatuses(ss);
  const countGhosts = cleanGhostOrders(ss);
  const countTotals = backfillOrderTotals(ss);
  const countSlips = formatAllSlipLinks(ss);
  const countCleaned = cleanDuplicateTickets(ss);
  const countTickets = syncTicketsFromOrders(ss);
  Logger.log(`แก้คำเพี้ยน ${fixedMojibake} ช่อง, ลบแถวผี ${countGhosts} แถว, อัปเดตยอดรวม ${countTotals} แถว, ลบตั๋วซ้ำ ${countCleaned} ใบ, เติมตั๋ว ${countTickets} ใบ, แปลงสลิป ${countSlips} ช่อง`);
}

function fixSlipLinks() {
  const ss = getSS();
  const count = formatAllSlipLinks(ss);
  Logger.log(`แปลงลิงก์สลิปเป็น [📄 ดูรูปสลิป] เรียบร้อยแล้วทั้งหมด ${count} ช่อง`);
}

function fixTickets() {
  const ss = getSS();
  const countGhosts = cleanGhostOrders(ss);
  const countCleaned = cleanDuplicateTickets(ss);
  const count = syncTicketsFromOrders(ss);
  Logger.log(`ลบแถวผี ${countGhosts} แถว, ลบตั๋วที่ซ้ำกัน ${countCleaned} ใบ และซิงค์ตั๋วที่ขาด ${count} ใบ เรียบร้อยแล้ว`);
}

function fixDuplicates() {
  const ss = getSS();
  const countGhosts = cleanGhostOrders(ss);
  const count = cleanDuplicateTickets(ss);
  Logger.log(`ลบแถวผี ${countGhosts} แถว และลบตั๋วที่ซ้ำกันในชีท Tickets เรียบร้อยแล้วทั้งหมด ${count} ใบ`);
}


