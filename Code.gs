const SPREADSHEET_ID = "1KQD-F-LuOCRR5ONjWWLCAUZUU1JeZI0qUgF8s2Wvgdo";

// LINE Login Channel ID
const LINE_LOGIN_CHANNEL_ID = "2012195926";

const CUSTOMER_SHEET = "CUSTOMERS";
const ORDER_SHEET = "Form_Responses";

// ===== API (รับคำขอจากหน้าเว็บที่โฮสต์บน GitHub Pages) =====

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    let data;

    if (body.action === "getCustomer") {
      data = getCustomer(body.idToken);
    } else if (body.action === "saveSubmission") {
      data = saveSubmission(
        body.idToken,
        body.mode,
        body.order,
        body.name,
        body.address,
        body.zipcode,
        body.phone
      );
    } else {
      throw new Error("ไม่พบ action ที่ร้องขอ");
    }

    return jsonOut({ ok: true, data: data });
  } catch (err) {
    return jsonOut({ ok: false, error: err.message });
  }
}

// ไว้เช็กว่า API ทำงานอยู่ (เปิด URL /exec ในเบราว์เซอร์จะเห็นข้อความนี้)
function doGet() {
  return jsonOut({ ok: true, message: "API พร้อมใช้งาน" });
}

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ===== สร้างชีตและหัวตารางที่จำเป็น โดยไม่ลบข้อมูลเดิม =====

function setupSheets() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  let customers = ss.getSheetByName(CUSTOMER_SHEET);
  if (!customers) {
    customers = ss.insertSheet(CUSTOMER_SHEET);
  }

  if (customers.getLastRow() === 0) {
    customers.appendRow([
      "LINE User ID",
      "ชื่อ",
      "ที่อยู่",
      "รหัสไปรษณีย์",
      "เบอร์โทร",
      "อัปเดตล่าสุด"
    ]);
  }

  let orders = ss.getSheetByName(ORDER_SHEET);
  if (!orders) {
    orders = ss.insertSheet(ORDER_SHEET);
  }

  if (orders.getLastRow() === 0) {
    orders.appendRow([
      "วันที่บันทึก",
      "รอบพรี/สินค้า",
      "ชื่อ",
      "ที่อยู่",
      "รหัสไปรษณีย์",
      "เบอร์โทร",
      "LINE User ID",
      "สถานะ"
    ]);
  } else {
    if (!orders.getRange("G1").getValue()) {
      orders.getRange("G1").setValue("LINE User ID");
    }
    if (!orders.getRange("H1").getValue()) {
      orders.getRange("H1").setValue("สถานะ");
    }
  }

  return "ตั้งค่าชีตเรียบร้อย";
}

// ===== ตรวจสอบ ID token กับ LINE ก่อนใช้ LINE User ID =====

function verifyLineToken(idToken) {
  if (!idToken || !LINE_LOGIN_CHANNEL_ID ||
      LINE_LOGIN_CHANNEL_ID === "ใส่_LINE_LOGIN_CHANNEL_ID") {
    throw new Error("ยังตั้งค่า LINE Login Channel ID ไม่ครบ");
  }

  const response = UrlFetchApp.fetch(
    "https://api.line.me/oauth2/v2.1/verify",
    {
      method: "post",
      payload: {
        id_token: idToken,
        client_id: LINE_LOGIN_CHANNEL_ID
      },
      muteHttpExceptions: true
    }
  );

  const result = JSON.parse(response.getContentText());

  if (response.getResponseCode() !== 200 || !result.sub) {
    throw new Error("ยืนยันบัญชี LINE ไม่สำเร็จ กรุณาเปิดหน้าเว็บผ่าน LINE ใหม่");
  }

  return result.sub;
}

// ===== ดึงข้อมูลลูกค้า =====

function getCustomer(idToken) {
  const lineUserId = verifyLineToken(idToken);
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CUSTOMER_SHEET);

  if (!sheet || sheet.getLastRow() < 2) {
    return { found: false };
  }

  const rows = sheet.getRange(
    2, 1, sheet.getLastRow() - 1, 6
  ).getDisplayValues();

  for (let i = 0; i < rows.length; i++) {
    if (rows[i][0] === lineUserId) {
      return {
        found: true,
        name: rows[i][1],
        address: rows[i][2],
        zipcode: rows[i][3],
        phone: rows[i][4]
      };
    }
  }

  return { found: false };
}

// ===== บันทึกออเดอร์ / อัปเดตที่อยู่ =====

function saveSubmission(idToken, mode, order, name, address, zipcode, phone) {
  const lineUserId = verifyLineToken(idToken);

  name = String(name || "").trim();
  address = String(address || "").trim();
  zipcode = String(zipcode || "").trim();
  phone = String(phone || "").trim();
  order = String(order || "").trim();

  if (!name || !address || !zipcode || !phone) {
    throw new Error("กรุณากรอกชื่อ ที่อยู่ รหัสไปรษณีย์ และเบอร์โทรให้ครบ");
  }

  if (mode === "newOrder" && !order) {
    throw new Error("กรุณากรอกรอบพรี/สินค้า");
  }

  if (mode !== "newOrder" && mode !== "updateAddress") {
    throw new Error("ไม่พบรูปแบบการบันทึกที่ถูกต้อง");
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);

  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const customers = ss.getSheetByName(CUSTOMER_SHEET);
    const orders = ss.getSheetByName(ORDER_SHEET);

    if (!customers || !orders) {
      throw new Error("ยังไม่มีชีตที่จำเป็น กรุณารัน setupSheets ก่อน");
    }

    // อัปเดตข้อมูลลูกค้าคนเดิม หรือเพิ่มลูกค้าใหม่
    const customerRow = findCustomerRow(customers, lineUserId);
    const customerData = [
      lineUserId, name, address, zipcode, phone, new Date()
    ];

    if (customerRow) {
      customers.getRange(customerRow, 1, 1, 6).setValues([customerData]);
    } else {
      customers.appendRow(customerData);
    }

    if (mode === "updateAddress") {
      // เปลี่ยนที่อยู่ให้เฉพาะออเดอร์ที่ยังรอจัดส่ง
      const lastRow = orders.getLastRow();

      if (lastRow >= 2) {
        const rows = orders.getRange(2, 1, lastRow - 1, 8).getValues();

        rows.forEach((row, index) => {
          if (
            String(row[6]) === lineUserId &&
            String(row[7]).trim() === "รอจัดส่ง"
          ) {
            orders.getRange(index + 2, 3, 1, 4).setValues([[
              name, address, zipcode, phone
            ]]);
          }
        });
      }

      return { status: "success", message: "อัปเดตที่อยู่เรียบร้อย" };
    }

    // ทุกออเดอร์ใหม่จะมีแถวของตัวเอง
    orders.appendRow([
      new Date(),
      order,
      name,
      address,
      zipcode,
      phone,
      lineUserId,
      "รอจัดส่ง"
    ]);

    return { status: "success", message: "บันทึกออเดอร์เรียบร้อย" };

  } finally {
    lock.releaseLock();
  }
}

function findCustomerRow(sheet, lineUserId) {
  if (sheet.getLastRow() < 2) return null;

  const ids = sheet.getRange(
    2, 1, sheet.getLastRow() - 1, 1
  ).getDisplayValues();

  for (let i = 0; i < ids.length; i++) {
    if (ids[i][0] === lineUserId) return i + 2;
  }

  return null;
}
