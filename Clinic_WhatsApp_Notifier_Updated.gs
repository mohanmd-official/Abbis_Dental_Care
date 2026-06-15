// === CONFIGURATION — fill these in ===
const TWILIO_ACCOUNT_SID   = 'AC6768f04c65d69c549e6d01d0a7dde4c2';        // from Twilio Console
const TWILIO_AUTH_TOKEN    = 'f4c934623c2bb3b001bb47531f2facc8';        // from Twilio Console
const TWILIO_WHATSAPP_FROM = 'whatsapp:+14155238886';   // Twilio sandbox number
const CLINIC_WHATSAPP_TO   = 'whatsapp:+917395995539';
const DEFAULT_COUNTRY_CODE = '+91';
const CLINIC_NAME          = "Abbi's Dental Care";
const CLINIC_PHONE         = '+91 73959 95539';

function onFormSubmit(e) {
  const responses = e.namedValues;

  function getValue(key) {
    return responses[key] ? responses[key][0] : '';
  }

  const name      = getValue('Name / பெயர்');
  const dob       = getValue('Date of Birth / பிறந்த தேதி');
  const age       = calculateAge(dob);
  const address   = getValue('Address / முகவரி');
  const phone     = getValue('Phone Number / தொலைபேஸி எண்');
  const reason    = getValue('Reason for Appointment / சந்திப்புக்கான காரணம்');
  const prefDate  = getValue('Preferred Date / விருப்பமான தேதி');
  const prefTime  = formatTime12Hour(getValue('Preferred Time / விருப்பமான நேரம்'));

  // ---- Message to the Clinic ----
  const clinicMessage =
    '🦷 New Appointment Booking\n\n' +
    'Name: ' + name + '\n' +
    'DOB: ' + dob + '\n' +
    'Age: ' + age + '\n' +
    'Address: ' + address + '\n' +
    'Phone: ' + phone + '\n' +
    'Reason: ' + reason + '\n' +
    'Preferred Date: ' + prefDate + '\n' +
    'Preferred Time: ' + prefTime + '\n\n' +
    'Call and confirm the appointment as soon as possible!';

  sendWhatsAppMessage(CLINIC_WHATSAPP_TO, clinicMessage);

  // ---- Message to the Patient ----
  const patientMessage =
    'Hello ' + name + ', thank you for booking with ' + CLINIC_NAME + '.\n\n' +
    'Reason: ' + reason + '\n' +
    'Preferred Date: ' + prefDate + '\n' +
    'Preferred Time: ' + prefTime + '\n\n' +
    'We will confirm your appointment shortly. For urgent queries call ' + CLINIC_PHONE + '.';

  const patientNumber = formatPhoneNumber(phone);
  if (patientNumber) {
    sendWhatsAppMessage('whatsapp:' + patientNumber, patientMessage);
  }
}

function formatPhoneNumber(raw) {
  let digits = raw.replace(/[^0-9+]/g, '');
  if (digits.indexOf('+') === 0) return digits;
  if (digits.length === 10) return DEFAULT_COUNTRY_CODE + digits;
  return '+' + digits;
}

// Converts "14:30" -> "2:30 PM"
function formatTime12Hour(time24) {
  if (!time24) return '';
  const parts = time24.split(':');
  let hours = parseInt(parts[0], 10);
  const minutes = parts[1];
  const ampm = hours >= 12 ? 'PM' : 'AM';
  hours = hours % 12;
  if (hours === 0) hours = 12;
  return hours + ':' + minutes + ' ' + ampm;
}

// Calculates age in years from a DOB string (e.g. "1990-06-14")
function calculateAge(dobString) {
  if (!dobString) return '';
  const dobDate = new Date(dobString);
  if (isNaN(dobDate.getTime())) return '';

  const today = new Date();
  let age = today.getFullYear() - dobDate.getFullYear();
  const monthDiff = today.getMonth() - dobDate.getMonth();
  if (monthDiff < 0 || (monthDiff === 0 && today.getDate() < dobDate.getDate())) {
    age--;
  }
  return age;
}

function sendWhatsAppMessage(to, body) {
  const url = 'https://api.twilio.com/2010-04-01/Accounts/' + TWILIO_ACCOUNT_SID + '/Messages.json';
  const payload = { 'To': to, 'From': TWILIO_WHATSAPP_FROM, 'Body': body };
  const options = {
    method: 'post',
    payload: payload,
    headers: {
      Authorization: 'Basic ' + Utilities.base64Encode(TWILIO_ACCOUNT_SID + ':' + TWILIO_AUTH_TOKEN)
    },
    muteHttpExceptions: true
  };
  const response = UrlFetchApp.fetch(url, options);
  Logger.log(response.getContentText());
}

// =========================================================
// ADD-ON: Handles leads from the website's AI Chat Assistant
// Reuses the config and helper functions above — does not
// modify or interfere with onFormSubmit / form bookings.
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
