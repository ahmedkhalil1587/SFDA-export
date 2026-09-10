/**
 * ============================================================
 * نظام تسجيل الدخول بـ OTP - SFDA Updates
 * ============================================================
 * الشيتات المستخدمة (لازم تكون بنفس الأسماء دي بالظبط):
 *   Users      : UserID | Name | Email | Role | Status | CreatedAt | LastLogin
 *   OTP_Codes  : Email  | OTP  | GeneratedAt | ExpiresAt | Used | Attempts
 *   Sessions   : Email  | Token | CreatedAt | ExpiresAt
 *
 * طريقة النشر:
 *   1) من محرر Apps Script: Deploy > New deployment
 *   2) النوع: Web app
 *   3) Execute as: Me
 *   4) Who has access: Anyone
 *   5) انسخي الرابط (Web App URL) واستخدميه في صفحة الـ HTML بتاعتك
 * ============================================================
 */

// ---------------------- الإعدادات ----------------------
const CONFIG = {
  SHEET_USERS: "Users",
  SHEET_OTP: "OTP_Codes",
  SHEET_SESSIONS: "Sessions",
  OTP_EXPIRY_MINUTES: 10,
  SESSION_EXPIRY_DAYS: 7,
  MAX_OTP_ATTEMPTS: 5,
  EMAIL_SENDER_NAME: "integriox",
  SENDER_EMAIL: "integriox@gmail.com",
};

// ---------------------- نقطة الدخول (Web App) ----------------------
function doGet(e) {
  const action = e.parameter.action;
  let result;

  try {
    switch (action) {
      case "requestOtp":
        result = requestOtp(e.parameter.email);
        break;
      case "register":
        result = registerUser(e.parameter.name, e.parameter.email);
        break;
      case "verifyOtp":
        result = verifyOtp(e.parameter.email, e.parameter.code);
        break;
      case "checkSession":
        result = checkSession(e.parameter.token);
        break;
      case "logout":
        result = logout(e.parameter.token);
        break;
      case "listUsers":
        result = listUsers(e.parameter.token);
        break;
      case "setUserStatus":
        result = setUserStatus(e.parameter.token, e.parameter.targetEmail, e.parameter.status);
        break;
      default:
        result = { success: false, message: "إجراء غير معروف" };
    }
  } catch (err) {
    result = { success: false, message: "خطأ في السيرفر: " + err.message };
  }

  return jsonResponse(result);
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---------------------- طلب OTP ----------------------
function requestOtp(email) {
  if (!email) {
    return { success: false, message: "الإيميل مطلوب" };
  }
  email = email.trim().toLowerCase();

  const user = findUserByEmail(email);
  if (!user) {
    return { success: false, message: "الإيميل ده مش مسجل في النظام" };
  }
  if (user.status !== "Active") {
    if (user.status === "Pending") {
      return { success: false, message: "حسابك لسه بانتظار موافقة الأدمن" };
    }
    return { success: false, message: "الحساب ده موقوف، كلّم الأدمن" };
  }

  const otp = generateOtpCode();
  const now = new Date();
  const expiresAt = new Date(now.getTime() + CONFIG.OTP_EXPIRY_MINUTES * 60 * 1000);

  saveOtp(email, otp, now, expiresAt);
  sendOtpEmail(email, user.name, otp);

  return { success: true, message: "تم إرسال رمز التحقق للإيميل" };
}

function generateOtpCode() {
  return Utilities.formatString("%06d", Math.floor(Math.random() * 1000000));
}

function sendOtpEmail(email, name, otp) {
  const subject = "رمز تسجيل الدخول - SFDA Updates";
  const body =
    "مرحبًا " + name + "،\n\n" +
    "رمز تسجيل الدخول بتاعك هو: " + otp + "\n\n" +
    "الرمز صالح لمدة " + CONFIG.OTP_EXPIRY_MINUTES + " دقايق بس.\n" +
    "لو ما طلبتش تسجيل دخول، تجاهل الإيميل ده.\n\n" +
    CONFIG.EMAIL_SENDER_NAME;

  GmailApp.sendEmail(email, subject, body, {
    name: CONFIG.EMAIL_SENDER_NAME,
    from: CONFIG.SENDER_EMAIL,
  });
}

// ---------------------- حفظ OTP في الشيت ----------------------
function saveOtp(email, otp, generatedAt, expiresAt) {
  const sheet = getSheet(CONFIG.SHEET_OTP);
  const data = sheet.getDataRange().getValues();

  // لو فيه صف قديم لنفس الإيميل، نحدّثه بدل ما نضيف صف جديد
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === email) {
      sheet.getRange(i + 1, 1, 1, 6).setValues([[email, otp, generatedAt, expiresAt, false, 0]]);
      return;
    }
  }
  // مفيش صف قديم، نضيف صف جديد
  sheet.appendRow([email, otp, generatedAt, expiresAt, false, 0]);
}

// ---------------------- التحقق من OTP ----------------------
function verifyOtp(email, code) {
  if (!email || !code) {
    return { success: false, message: "الإيميل والرمز مطلوبين" };
  }
  email = email.trim().toLowerCase();

  const sheet = getSheet(CONFIG.SHEET_OTP);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    const rowEmail = String(data[i][0]).trim().toLowerCase();
    if (rowEmail !== email) continue;

    const storedOtp = String(data[i][1]).trim();
    const expiresAt = new Date(data[i][3]);
    const used = data[i][4];
    const attempts = Number(data[i][5]) || 0;

    if (used) {
      return { success: false, message: "الرمز ده اتستخدم قبل كده، اطلب رمز جديد" };
    }
    if (new Date() > expiresAt) {
      return { success: false, message: "الرمز منتهي الصلاحية، اطلب رمز جديد" };
    }
    if (attempts >= CONFIG.MAX_OTP_ATTEMPTS) {
      return { success: false, message: "تجاوزت عدد المحاولات المسموح، اطلب رمز جديد" };
    }
    if (storedOtp !== String(code).trim()) {
      sheet.getRange(i + 1, 6).setValue(attempts + 1);
      return { success: false, message: "الرمز غلط، حاول تاني" };
    }

    // الرمز صح - نعلّمه كمُستخدم
    sheet.getRange(i + 1, 5).setValue(true);

    const user = findUserByEmail(email);
    const token = createSession(email);
    updateLastLogin(email);

    return {
      success: true,
      token: token,
      user: { name: user.name, email: user.email, role: user.role },
    };
  }

  return { success: false, message: "محدش طلب رمز بالإيميل ده، اطلب رمز الأول" };
}

// ---------------------- إدارة الجلسات (Sessions) ----------------------
function createSession(email) {
  const token = Utilities.getUuid();
  const now = new Date();
  const expiresAt = new Date(now.getTime() + CONFIG.SESSION_EXPIRY_DAYS * 24 * 60 * 60 * 1000);

  const sheet = getSheet(CONFIG.SHEET_SESSIONS);
  sheet.appendRow([email, token, now, expiresAt]);

  return token;
}

function checkSession(token) {
  if (!token) {
    return { success: false, message: "التوكن مطلوب" };
  }

  const sheet = getSheet(CONFIG.SHEET_SESSIONS);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]) === token) {
      const expiresAt = new Date(data[i][3]);
      if (new Date() > expiresAt) {
        return { success: false, message: "انتهت الجلسة، سجّل دخول تاني" };
      }
      const email = String(data[i][0]).trim().toLowerCase();
      const user = findUserByEmail(email);
      if (!user || user.status !== "Active") {
        return { success: false, message: "الحساب غير متاح" };
      }
      return {
        success: true,
        user: { name: user.name, email: user.email, role: user.role },
      };
    }
  }

  return { success: false, message: "جلسة غير صالحة" };
}

function logout(token) {
  if (!token) {
    return { success: false, message: "التوكن مطلوب" };
  }
  const sheet = getSheet(CONFIG.SHEET_SESSIONS);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]) === token) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: true }; // مفيش جلسة أصلاً، يبقى الهدف اتحقق
}

// ---------------------- تسجيل حساب جديد ----------------------
function registerUser(name, email) {
  if (!name || !email) {
    return { success: false, message: "الاسم والإيميل مطلوبين" };
  }
  name = name.trim();
  email = email.trim().toLowerCase();

  if (!email.includes("@") || !email.includes(".")) {
    return { success: false, message: "اكتب إيميل صحيح" };
  }

  const existing = findUserByEmail(email);
  if (existing) {
    if (existing.status === "Pending") {
      return { success: false, message: "الحساب ده اتسجّل قبل كده ولسه بانتظار موافقة الأدمن" };
    }
    if (existing.status === "Active") {
      return { success: false, message: "الحساب ده مسجّل بالفعل، سجّل دخول عادي" };
    }
    return { success: false, message: "الحساب ده متوقف، كلّم الأدمن" };
  }

  const sheet = getSheet(CONFIG.SHEET_USERS);
  const data = sheet.getDataRange().getValues();

  let maxId = 0;
  for (let i = 1; i < data.length; i++) {
    const id = Number(data[i][0]) || 0;
    if (id > maxId) maxId = id;
  }

  sheet.appendRow([maxId + 1, name, email, "Employee", "Pending", new Date(), ""]);

  return { success: true, message: "تم التسجيل، هتقدر تسجّل دخول بعد ما الأدمن يوافق على حسابك" };
}

// ---------------------- إدارة اليوزرات (أدمن بس) ----------------------

// يتأكد إن التوكن ده بتاع أدمن فعلاً قبل ما يسمح بأي إجراء إداري
function requireAdmin(token) {
  const session = checkSession(token);
  if (!session.success) {
    return { ok: false, message: session.message };
  }
  if (session.user.role !== "Admin") {
    return { ok: false, message: "الإجراء ده متاح للأدمن بس" };
  }
  return { ok: true, admin: session.user };
}

// يرجّع كل اليوزرات - للأدمن بس، عشان يقدر يختار مين يوقف/يفعّل
function listUsers(token) {
  const check = requireAdmin(token);
  if (!check.ok) return { success: false, message: check.message };

  const sheet = getSheet(CONFIG.SHEET_USERS);
  const data = sheet.getDataRange().getValues();
  const users = [];

  for (let i = 1; i < data.length; i++) {
    users.push({
      userId: data[i][0],
      name: data[i][1],
      email: data[i][2],
      role: data[i][3],
      status: data[i][4],
      createdAt: data[i][5],
      lastLogin: data[i][6],
    });
  }

  return { success: true, users: users };
}

// يفعّل أو يوقف حساب يوزر معيّن - للأدمن بس
function setUserStatus(token, targetEmail, newStatus) {
  const check = requireAdmin(token);
  if (!check.ok) return { success: false, message: check.message };

  if (!targetEmail || !newStatus) {
    return { success: false, message: "الإيميل والحالة مطلوبين" };
  }
  if (newStatus !== "Active" && newStatus !== "Disabled") {
    return { success: false, message: "الحالة لازم تكون Active أو Disabled" };
  }

  targetEmail = targetEmail.trim().toLowerCase();

  // تحذير أمان: الأدمن مايقدرش يوقف حسابه هو نفسه (عشان مايتقفلش برّه بالغلط)
  if (targetEmail === check.admin.email.toLowerCase() && newStatus === "Disabled") {
    return { success: false, message: "مينفعش توقف حسابك انت شخصيًا" };
  }

  const targetUser = findUserByEmail(targetEmail);
  if (!targetUser) {
    return { success: false, message: "مفيش يوزر بالإيميل ده" };
  }

  const sheet = getSheet(CONFIG.SHEET_USERS);
  sheet.getRange(targetUser.rowIndex, 5).setValue(newStatus);

  // لو بنوقف الحساب، نلغي أي جلسات دخول شغالة له فورًا (يتم طرده لو داخل دلوقتي)
  if (newStatus === "Disabled") {
    revokeAllSessions(targetEmail);
  }

  return { success: true, message: "تم تحديث حالة الحساب إلى " + newStatus };
}

// يمسح كل جلسات الدخول الشغالة بتاعة إيميل معيّن
function revokeAllSessions(email) {
  const sheet = getSheet(CONFIG.SHEET_SESSIONS);
  const data = sheet.getDataRange().getValues();

  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]).trim().toLowerCase() === email) {
      sheet.deleteRow(i + 1);
    }
  }
}

// ---------------------- دوال مساعدة ----------------------
function getSheet(name) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sheet) throw new Error("الشيت مش موجود: " + name);
  return sheet;
}

function findUserByEmail(email) {
  const sheet = getSheet(CONFIG.SHEET_USERS);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][2]).trim().toLowerCase() === email) {
      return {
        rowIndex: i + 1,
        userId: data[i][0],
        name: data[i][1],
        email: data[i][2],
        role: data[i][3],
        status: data[i][4],
      };
    }
  }
  return null;
}

function updateLastLogin(email) {
  const user = findUserByEmail(email);
  if (!user) return;
  const sheet = getSheet(CONFIG.SHEET_USERS);
  sheet.getRange(user.rowIndex, 7).setValue(new Date());
}

// ---------------------- تنظيف دوري (اختياري) ----------------------
// شغّليها يدويًا مرة، أو اعملي لها Trigger زمني (يومي مثلاً) من قايمة Triggers
function cleanupExpiredRecords() {
  const now = new Date();

  const otpSheet = getSheet(CONFIG.SHEET_OTP);
  const otpData = otpSheet.getDataRange().getValues();
  for (let i = otpData.length - 1; i >= 1; i--) {
    if (new Date(otpData[i][3]) < now) {
      otpSheet.deleteRow(i + 1);
    }
  }

  const sessSheet = getSheet(CONFIG.SHEET_SESSIONS);
  const sessData = sessSheet.getDataRange().getValues();
  for (let i = sessData.length - 1; i >= 1; i--) {
    if (new Date(sessData[i][3]) < now) {
      sessSheet.deleteRow(i + 1);
    }
  }
}
