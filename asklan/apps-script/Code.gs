/**
 * Asklan Pharmacy — orders endpoint (Google Apps Script)
 * ------------------------------------------------------
 * This is your existing doPost, unchanged, with an email
 * notification added: the moment a customer submits the order
 * form, the row is logged to the sheet (as before) AND an
 * Arabic order summary is emailed to NOTIFY_EMAIL.
 *
 * DEPLOY:
 *   1. Paste this over your current script at script.google.com.
 *   2. Deploy → Manage deployments → edit the existing "Web app"
 *      deployment → New version → Deploy.
 *      (Same deployment = ORDERS_ENDPOINT stays the same, so no
 *       website changes are needed.)
 *   3. First run asks for Gmail permission — approve it. Tip: run
 *      the testEmail function once from the editor to trigger that
 *      prompt and send yourself a sample before going live.
 */

// Where order notifications are emailed. Change to use another address.
var NOTIFY_EMAIL = 'tactic.digital.company@gmail.com';


function doPost(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var data = JSON.parse(e.postData.contents);

    var items = data.items.map(function(i) {
      return i.name + ' × ' + i.qty + ' = ' + i.subtotal + ' د.أ';
    }).join(' | ');

    sheet.appendRow([
      data.timestamp,
      data.customer.name,
      data.customer.phone,
      data.customer.area,
      data.customer.notes || '',
      items,
      data.total + ' د.أ'
    ]);

    // Notify by email. Wrapped so a mail failure can never
    // break the sheet logging above (the row is already saved).
    try {
      notifyByEmail(data);
    } catch (mailErr) {}

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}


function notifyByEmail(data) {
  var c = data.customer || {};
  var lines = (data.items || []).map(function(i) {
    return '• ' + i.name + ' × ' + i.qty + ' = ' + i.subtotal + ' د.أ';
  }).join('\n');

  var body =
    '🛒 طلب جديد\n\n' +
    '👤 الاسم: ' + (c.name || '') + '\n' +
    '📞 الهاتف: ' + (c.phone || '') + '\n' +
    '📍 العنوان: ' + (c.area || '') + '\n' +
    (c.notes ? '📝 ملاحظات: ' + c.notes + '\n' : '') +
    '\nالمنتجات:\n' + lines +
    '\n\n💰 المجموع: ' + (data.total || '') + ' د.أ' +
    '\n🕒 ' + (data.timestamp || '');

  MailApp.sendEmail({
    to: NOTIFY_EMAIL,
    subject: 'طلب جديد – ' + (c.name || 'عميل') + ' (' + (data.total || '') + ' د.أ)',
    body: body
  });
}


// Run once from the editor (Run → testEmail) to approve the Gmail
// permission and send yourself a sample notification.
function testEmail() {
  notifyByEmail({
    timestamp: new Date().toLocaleString('ar-JO', { timeZone: 'Asia/Amman' }),
    customer: { name: 'عميل تجريبي', phone: '0790000000', area: 'عمّان', notes: 'طلب اختبار' },
    items: [
      { name: 'Alcom Plus Lotion', qty: 2, subtotal: '19.80' },
      { name: 'INNOX Hair Mask', qty: 1, subtotal: '9.95' }
    ],
    total: '29.75'
  });
}
