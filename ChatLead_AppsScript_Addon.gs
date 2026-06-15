// =========================================================
// ADD-ON: Handles leads from the website's AI Chat Assistant
// Paste this BELOW your existing code in the same Apps Script project.
// It reuses CLINIC_WHATSAPP_TO, CLINIC_NAME, CLINIC_PHONE, sendWhatsAppMessage()
// and formatPhoneNumber() that are already defined above.
// =========================================================

// Runs when the website sends a chat lead (Name, Phone, Age)
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const name  = (data.name  || '').toString();
    const phone = (data.phone || '').toString();
    const age   = (data.age   || '').toString();

    // 1. Log this lead as a new row in the SAME response sheet
    //    Column order matches your Google Form sheet:
    //    Timestamp | Name | DOB | Address | Phone | Reason | Preferred Date | Preferred Time
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    sheet.appendRow([
      new Date(),                                  // Timestamp
      name,                                        // Name / பெயர்
      '',                                          // Date of Birth (not collected by chat)
      '',                                          // Address (not collected by chat)
      phone,                                       // Phone Number / தொலைபேஸி எண்
      'Chat Assistant enquiry (Age: ' + age + ')', // Reason for Appointment
      '',                                          // Preferred Date
      ''                                           // Preferred Time
    ]);

    // 2. Notify the clinic on WhatsApp — same channel as form bookings
    const clinicMessage =
      '💬 New CHAT lead from website\n\n' +
      'Name: ' + name + '\n' +
      'Age: ' + age + '\n' +
      'Phone: ' + phone + '\n\n' +
      'Source: AI Chat Assistant\n' +
      'Please follow up with the patient.';

    sendWhatsAppMessage(CLINIC_WHATSAPP_TO, clinicMessage);

    // 3. Optional: send a short acknowledgement to the patient too
    const patientNumber = formatPhoneNumber(phone);
    if (patientNumber) {
      const patientMessage =
        'Hello ' + name + ', thank you for reaching out to ' + CLINIC_NAME + ' via our website chat.\n\n' +
        'Our team has received your details and will contact you shortly.\n' +
        'For urgent queries, call ' + CLINIC_PHONE + '.';
      sendWhatsAppMessage('whatsapp:' + patientNumber, patientMessage);
    }

    return ContentService.createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
