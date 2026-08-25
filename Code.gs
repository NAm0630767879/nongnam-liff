const SPREADSHEET_ID = '1Jd5_7CzCBz1Oto0I3h99tDEdOeYDl_QyqWlOEkrvll4';
const CHANNEL_ACCESS_TOKEN = 'PASTE_YOUR_NEW_CHANNEL_ACCESS_TOKEN_HERE'; // ⚠️ token เดิมหลุดในแชทแล้ว ต้องไปออกใหม่ที่ LINE Developers Console ก่อนใช้งาน แล้วนำมาใส่แทนที่นี่
const SLIP_FOLDER_ID = '1TMeFgsHgCBewRAjFw_84kCX7pbnNiZT9'; // โฟลเดอร์ Google Drive สำหรับเก็บสลิป
const SELLER_LINE_USER_ID = 'U2bb7e570edbcae8dbb5d2e86a42b6388'; // LINE User ID ของแม่ค้า สำหรับรับแจ้งเตือนออเดอร์ใหม่

function getSheet(name) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  return ss.getSheetByName(name);
}

// ===================== ทางเข้าเว็บ (GET) =====================
function doGet(e) {
  const params = (e && e.parameter) ? e.parameter : {};
  const action = params.action;

  if (!action) {
    return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('สั่งเครื่องดื่ม - ร้านน้ำชงน้องน้ำ')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  try {
    if (action === 'getMenu') {
      return jsonResponse({ ok: true, data: getMenuData() });
    }
    if (action === 'getOrdersByUser') {
      return jsonResponse({ ok: true, data: getOrdersByUser(params.userId) });
    }
    if (action === 'getAllOrders') {
      return jsonResponse({ ok: true, data: getAllOrders() });
    }
    if (action === 'getRevenue') {
      return jsonResponse({ ok: true, data: getRevenue(params.range) });
    }
    if (action === 'getCostSummary') {
      return jsonResponse({ ok: true, data: getCostSummary(params.month) });
    }
    return jsonResponse({ ok: false, error: 'unknown action: ' + action });
  } catch (err) {
    return jsonResponse({ ok: false, error: String(err) });
  }
}

// ===================== ทางเข้าเว็บ (POST) =====================
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action;

    if (action === 'createOrder') {
      return createOrder(data);
    }
    if (action === 'updateOrderStatus') {
      return updateOrderStatus(data);
    }
    if (action === 'addExpense') {
      return addExpense(data);
    }
    if (action === 'deleteExpense') {
      return deleteExpense(data);
    }
    return jsonResponse({ ok: false, error: 'unknown action: ' + action });
  } catch (err) {
    return jsonResponse({ ok: false, error: String(err) });
  }
}

// ===================== เมนู =====================
// โครงสร้างชีต Menu จริง: A=code, B=menu, C=price, D=img
function getMenuData() {
  const menuSheet = getSheet('Menu');
  if (!menuSheet) throw new Error('ไม่พบชีตชื่อ "Menu"');

  const menuRows = menuSheet.getDataRange().getValues();
  const menu = menuRows.slice(1)
    .filter(r => r[0])
    .map(r => ({
      id: r[0],
      name: r[1],
      price: r[2],
      img: r[3] || '',
      status: 'available'
    }));

  let toppings = [];
  const toppingSheet = getSheet('Topping');
  if (toppingSheet) {
    const toppingRows = toppingSheet.getDataRange().getValues();
    toppings = toppingRows.slice(1)
      .filter(r => r[0])
      .map(r => ({ id: r[0], name: r[1], price: r[2] }));
  }

  return { menu, toppings };
}

// ===================== คำสั่งซื้อ =====================
// โครงสร้างชีต Orders (A ถึง K):
// OrderID | LineUserID | CustomerName | Telephone | Timestamp | Items | TotalPrice | PaymentMethod | SlipURL | PaymentStatus | OrderStatus
function createOrder(data) {
  const sheet = getSheet('Orders');
  if (!sheet) throw new Error('ไม่พบชีตชื่อ "Orders"');

  const orderId = 'ORD' + new Date().getTime();
  const timestamp = new Date();

  // ถ้ามีการแนบสลิปโอนเงินมาด้วย (data.image เป็น base64) ให้บันทึกลง Drive
  let slipUrl = '';
  if (data.image) {
    slipUrl = saveImageToDrive(data.image, orderId);
  }

  // ถ้ามีสลิปแนบมาด้วย = รอตรวจสอบสลิป, ถ้าเลือกชำระที่ร้าน = รอชำระที่ร้าน
  const paymentStatus = slipUrl ? 'รอตรวจสอบสลิป' : 'รอชำระที่ร้าน';

  sheet.appendRow([
    orderId,
    data.userId || '',
    data.name || '',
    data.telephone || '',
    timestamp,
    JSON.stringify(data.items),
    data.totalPrice,
    data.paymentMethod || '',
    slipUrl,
    paymentStatus,
    'รอชง'
  ]);

  if (data.userId) {
    pushMessage(data.userId,
      `รับคำสั่งซื้อแล้วค่ะ 🧋\nหมายเลขคำสั่งซื้อ: ${orderId}\nยอดรวม: ${data.totalPrice} บาท\nสถานะการชำระเงิน: ${paymentStatus}\nสถานะออเดอร์: รอชง`);
  }

  // แจ้งเตือนแม่ค้าเมื่อมีออเดอร์ใหม่เข้ามา
  notifySellerNewOrder(orderId, data);

  return jsonResponse({ ok: true, orderId: orderId, slipUrl: slipUrl });
}

// แจ้งเตือนแม่ค้าผ่าน LINE เมื่อมีออเดอร์ใหม่
function notifySellerNewOrder(orderId, data) {
  if (!SELLER_LINE_USER_ID) return;

  const itemsText = (data.items || [])
    .map(i => `- ${i.name} × ${i.qty}`)
    .join('\n');

  const custText = [data.name, data.telephone].filter(Boolean).join(' · ');

  const text =
    `🔔 มีออเดอร์ใหม่เข้ามา!\n` +
    `หมายเลข: ${orderId}\n` +
    (custText ? `ลูกค้า: ${custText}\n` : '') +
    `รายการ:\n${itemsText}\n` +
    `ยอดรวม: ${data.totalPrice} บาท\n` +
    `การชำระเงิน: ${data.paymentMethod || '-'}`;

  try {
    pushMessage(SELLER_LINE_USER_ID, text);
  } catch (err) {
    // ไม่ให้การแจ้งเตือนที่ล้มเหลวไปทำให้การสร้างออเดอร์ล้มเหลวตามไปด้วย
    console.error('แจ้งเตือนแม่ค้าไม่สำเร็จ: ' + err);
  }
}

// บันทึกไฟล์สลิป (base64) ลง Google Drive แล้วคืน URL ของไฟล์
function saveImageToDrive(base64Image, orderId) {
  const blob = Utilities.newBlob(
    Utilities.base64Decode(base64Image.split(',')[1]),
    'image/jpeg',
    'slip-' + (orderId || Utilities.getUuid()) + '.jpg'
  );

  const folder = DriveApp.getFolderById(SLIP_FOLDER_ID);
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  return file.getUrl();
}

function updateOrderStatus(data) {
  const sheet = getSheet('Orders');
  if (!sheet) throw new Error('ไม่พบชีตชื่อ "Orders"');

  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === data.orderId) {
      sheet.getRange(i + 1, 11).setValue(data.newStatus); // คอลัมน์ K = OrderStatus
      const lineUserId = values[i][1];
      if (lineUserId) {
        pushMessage(lineUserId, `คำสั่งซื้อ ${data.orderId} อัปเดตสถานะเป็น: ${data.newStatus}`);
      }
      return jsonResponse({ ok: true });
    }
  }
  return jsonResponse({ ok: false, error: 'order not found' });
}

// อัปเดตสถานะการชำระเงิน (เช่น หลังแอดมินตรวจสลิปแล้ว)
function updatePaymentStatus(data) {
  const sheet = getSheet('Orders');
  if (!sheet) throw new Error('ไม่พบชีตชื่อ "Orders"');

  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === data.orderId) {
      sheet.getRange(i + 1, 10).setValue(data.newStatus); // คอลัมน์ J = PaymentStatus
      return jsonResponse({ ok: true });
    }
  }
  return jsonResponse({ ok: false, error: 'order not found' });
}

function getOrdersByUser(userId) {
  const sheet = getSheet('Orders');
  if (!sheet) throw new Error('ไม่พบชีตชื่อ "Orders"');

  const values = sheet.getDataRange().getValues();
  return values.slice(1)
    .filter(r => r[1] === userId)
    .map(rowToOrderObject)
    .reverse();
}

function getAllOrders() {
  const sheet = getSheet('Orders');
  if (!sheet) throw new Error('ไม่พบชีตชื่อ "Orders"');

  const values = sheet.getDataRange().getValues();
  return values.slice(1).filter(r => r[0]).map(rowToOrderObject).reverse();
}

function rowToOrderObject(r) {
  return {
    orderId: r[0],
    userId: r[1],
    customerName: r[2],
    telephone: r[3],
    timestamp: r[4],
    items: JSON.parse(r[5]),
    totalPrice: r[6],
    paymentMethod: r[7],
    slipUrl: r[8],
    paymentStatus: r[9],
    orderStatus: r[10]
  };
}

// ===================== รายได้ =====================
// นับรายได้จากออเดอร์ที่มีสถานะ "เสร็จสิ้น" เท่านั้น (ลูกค้ามารับของแล้ว)
function getCompletedOrders() {
  return getAllOrders().filter(o => o.orderStatus === 'เสร็จสิ้น');
}

function getRevenue(range) {
  range = range || 'today';
  const now = new Date();
  const completed = getCompletedOrders();

  if (range === 'today') {
    const todays = completed.filter(o => isSameDay(new Date(o.timestamp), now));
    const hourTotals = {};
    todays.forEach(o => {
      const h = new Date(o.timestamp).getHours();
      hourTotals[h] = (hourTotals[h] || 0) + Number(o.totalPrice || 0);
    });
    // ช่วงเวลาเปิดร้าน สมมติ 07:00 - 20:00 ปรับตามจริงได้
    const labels = [];
    const values = [];
    for (let h = 7; h <= 20; h++) {
      labels.push(h + ':00');
      values.push(hourTotals[h] || 0);
    }
    const total = todays.reduce((s, o) => s + Number(o.totalPrice || 0), 0);
    return { total, labels, values };

  } else if (range === 'week') {
    const days = lastNDays(7, now);
    const labels = days.map(d => THAI_DOW[d.getDay()] + ' ' + d.getDate() + '/' + (d.getMonth() + 1));
    const values = days.map(d => sumForDay(completed, d));
    const total = values.reduce((a, b) => a + b, 0);
    return { total, labels, values };

  } else { // month
    const start = new Date(now.getFullYear(), now.getMonth(), 1);
    const end = new Date(now.getFullYear(), now.getMonth() + 1, 0);
    const labels = [];
    const values = [];
    for (let d = start.getDate(); d <= end.getDate(); d++) {
      const day = new Date(now.getFullYear(), now.getMonth(), d);
      labels.push(String(d));
      values.push(sumForDay(completed, day));
    }
    const total = values.reduce((a, b) => a + b, 0);
    return { total, labels, values };
  }
}

function sumForDay(orders, day) {
  return orders
    .filter(o => isSameDay(new Date(o.timestamp), day))
    .reduce((s, o) => s + Number(o.totalPrice || 0), 0);
}

function isSameDay(a, b) {
  return a.getFullYear() === b.getFullYear() && a.getMonth() === b.getMonth() && a.getDate() === b.getDate();
}

function lastNDays(n, from) {
  const out = [];
  for (let i = n - 1; i >= 0; i--) {
    const d = new Date(from);
    d.setDate(d.getDate() - i);
    out.push(d);
  }
  return out;
}

const THAI_DOW = ['อา', 'จ', 'อ', 'พ', 'พฤ', 'ศ', 'ส'];
const THAI_MONTH_ABBR = ['ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.', 'ก.ค.', 'ส.ค.', 'ก.ย.', 'ต.ค.', 'พ.ย.', 'ธ.ค.'];

// ===================== ต้นทุน / รายจ่าย =====================
// โครงสร้างชีต Expenses (A ถึง D): ID | Date | Description | Amount
// ต้องสร้างชีตชื่อ "Expenses" พร้อมหัวตารางนี้ในสเปรดชีตเดียวกันก่อนใช้งาน
function getExpensesSheet() {
  const sheet = getSheet('Expenses');
  if (!sheet) throw new Error('ไม่พบชีตชื่อ "Expenses" กรุณาสร้างชีตนี้พร้อมหัวตาราง ID | Date | Description | Amount');
  return sheet;
}

function getAllExpenses() {
  const sheet = getExpensesSheet();
  const values = sheet.getDataRange().getValues();
  return values.slice(1)
    .filter(r => r[0])
    .map(r => ({ id: r[0], date: r[1], description: r[2], amount: r[3] }));
}

function addExpense(data) {
  const sheet = getExpensesSheet();
  const id = 'EXP' + new Date().getTime();
  sheet.appendRow([id, new Date(data.date), data.description || '', Number(data.amount) || 0]);
  return jsonResponse({ ok: true, data: { id, date: data.date, description: data.description, amount: data.amount } });
}

function deleteExpense(data) {
  const sheet = getExpensesSheet();
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === data.id) {
      sheet.deleteRow(i + 1);
      return jsonResponse({ ok: true });
    }
  }
  return jsonResponse({ ok: false, error: 'expense not found' });
}

// สรุปรายได้ / ต้นทุน / กำไร-ขาดทุน ของเดือนที่ระบุ (month = "YYYY-MM")
function getCostSummary(month) {
  const now = new Date();
  const [y, m] = (month || (now.getFullYear() + '-' + String(now.getMonth() + 1).padStart(2, '0')))
    .split('-').map(Number);
  const targetYear = y;
  const targetMonth = m - 1; // 0-indexed

  const completed = getCompletedOrders();
  const expenses = getAllExpenses();

  const inTargetMonth = (dateVal) => {
    const d = new Date(dateVal);
    return d.getFullYear() === targetYear && d.getMonth() === targetMonth;
  };

  const monthRevenue = completed
    .filter(o => inTargetMonth(o.timestamp))
    .reduce((s, o) => s + Number(o.totalPrice || 0), 0);

  const monthExpenses = expenses.filter(ex => inTargetMonth(ex.date));
  const monthCost = monthExpenses.reduce((s, ex) => s + Number(ex.amount || 0), 0);

  // แนวโน้มกำไร-ขาดทุนย้อนหลัง 6 เดือน (รวมเดือนที่เลือกด้วย)
  const monthlyLabels = [];
  const monthlyProfit = [];
  for (let i = 5; i >= 0; i--) {
    const d = new Date(targetYear, targetMonth - i, 1);
    const rev = completed
      .filter(o => {
        const od = new Date(o.timestamp);
        return od.getFullYear() === d.getFullYear() && od.getMonth() === d.getMonth();
      })
      .reduce((s, o) => s + Number(o.totalPrice || 0), 0);
    const cost = expenses
      .filter(ex => {
        const ed = new Date(ex.date);
        return ed.getFullYear() === d.getFullYear() && ed.getMonth() === d.getMonth();
      })
      .reduce((s, ex) => s + Number(ex.amount || 0), 0);
    monthlyLabels.push(THAI_MONTH_ABBR[d.getMonth()] + ' ' + String(d.getFullYear() + 543).slice(-2));
    monthlyProfit.push(rev - cost);
  }

  return {
    revenue: monthRevenue,
    cost: monthCost,
    profit: monthRevenue - monthCost,
    expenses: monthExpenses,
    monthlyLabels,
    monthlyProfit
  };
}

// ===================== แจ้งเตือน LINE =====================
function pushMessage(userId, text) {
  UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + CHANNEL_ACCESS_TOKEN },
    payload: JSON.stringify({
      to: userId,
      messages: [{ type: 'text', text: text }]
    })
  });
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
