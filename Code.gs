/**
 * HE THONG THI TRUC TUYEN - GOOGLE APPS SCRIPT
 * Chay setupSheets() mot lan de tao 3 trang tinh va tieu de.
 */

// ID Google Sheets: lấy phần nằm giữa /d/ và /edit trong URL bảng tính.
// Ví dụ URL: https://docs.google.com/spreadsheets/d/ABC123xyz/edit
const SPREADSHEET_ID = "1yTSRGL-chXp7AvWdMhY5naRBOqBd86-kbX25cBrVF-E";

function getSpreadsheet_() {
  const id = String(SPREADSHEET_ID || "").trim();
  if (!id || id === "DAN_ID_GOOGLE_SHEET_CUA_BAN_VAO_DAY") {
    throw new Error("Chưa cấu hình SPREADSHEET_ID trong Code.gs. Hãy dán ID Google Sheets vào dòng cấu hình đầu file.");
  }
  try {
    return SpreadsheetApp.openById(id);
  } catch (err) {
    throw new Error("Không mở được Google Sheets. Kiểm tra SPREADSHEET_ID và quyền truy cập của tài khoản triển khai. Chi tiết: " + err.message);
  }
}

const SHEETS = {
  CONFIG: "CaiDatDeThi",
  BANK: "NganHangDe",
  RESULTS: "KetQua"
};
const HEADERS = {
  CaiDatDeThi: ["MonHoc", "ThoiGianLamBai", "NgayBatDau", "NgayKetThuc", "SoCauHoiBoc"],
  NganHangDe: ["Id", "MonHoc", "LoaiCauHoi", "NoiDungCauHoi", "HinhAnh", "DapAnA", "DapAnB", "DapAnC", "DapAnD", "DapAnDung"],
  KetQua: ["ThoiGian", "TenHocSinh", "MonHoc", "SoCauDung", "TongSoCau", "DiemSo"]
};
const SESSION_PREFIX = "EXAM_SESSION_";
const SUBMITTED_PREFIX = "EXAM_SUBMITTED_";
const ADMIN_TOKEN_PREFIX = "EXAM_ADMIN_";
const ADMIN_TOKEN_TTL = 21600;

function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) || "";
    if (action === "getSubjects") return jsonOutput_({status:"success", subjects:getSubjects_()});
    if (action === "getExam") {
      const subject = String(e.parameter.monHoc || "").trim();
      return jsonOutput_(createExam_(subject));
    }
    return jsonOutput_({status:"error", message:"Thao tác không hợp lệ."});
  } catch (err) {
    console.error(err);
    return jsonOutput_({status:"error", message:err.message || "Lỗi máy chủ."});
  }
}

function doPost(e) {
  try {
    const body = e && e.postData && e.postData.contents ? e.postData.contents : "{}";
    const params = JSON.parse(body);
    if (params.action === "submitExam") return jsonOutput_(submitExam_(params));
    if (params.action === "adminLogin") return jsonOutput_(adminLogin_(params));
    if (params.action && params.action.indexOf("admin") === 0) {
      requireAdmin_(params.token);
      return jsonOutput_(handleAdminAction_(params));
    }
    return jsonOutput_({status:"error", message:"Thao tác không hợp lệ."});
  } catch (err) {
    console.error(err);
    return jsonOutput_({status:"error", message:err.message || "Không thể xử lý dữ liệu gửi lên."});
  }
}

/** Tạo trang tính nếu thiếu; không xóa dữ liệu hiện có. */
function setupSheets() {
  const ss = getSpreadsheet_();
  Object.keys(HEADERS).forEach(name => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    const headers = HEADERS[name];
    const firstRow = sh.getRange(1, 1, 1, headers.length).getValues()[0];
    const empty = firstRow.every(v => v === "" || v === null);
    if (empty) {
      sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    } else {
      // Báo lỗi nếu dòng tiêu đề đang có nhưng sai thứ tự; tránh ghi đè dữ liệu.
      const mismatch = headers.some((h, i) => String(firstRow[i]).trim() !== h);
      if (mismatch) throw new Error("Trang tính " + name + " đã có tiêu đề khác cấu trúc. Hãy kiểm tra dòng 1 trước khi chạy lại setupSheets().");
    }
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#d9eaf7");
  });
  const config = ss.getSheetByName(SHEETS.CONFIG);
  config.getRange("C:D").setNumberFormat("dd/MM/yyyy HH:mm");
  config.getRange("B:B").setNumberFormat("0");
  config.getRange("E:E").setNumberFormat("0");
  const results = ss.getSheetByName(SHEETS.RESULTS);
  results.getRange("A:A").setNumberFormat("dd/MM/yyyy HH:mm:ss");
  results.getRange("F:F").setNumberFormat("0.00");
  return "Đã kiểm tra/tạo đủ 3 trang tính và tiêu đề cột.";
}

function getSubjects_() {
  const sh = getSheetOrThrow_(SHEETS.CONFIG);
  if (sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues()
    .filter(r => String(r[0] || "").trim())
    .map(r => String(r[0]).trim());
}

function createExam_(subject) {
  if (!subject) return {status:"error", message:"Vui lòng chọn môn thi."};
  const config = getSheetOrThrow_(SHEETS.CONFIG);
  const rows = config.getLastRow() > 1 ? config.getRange(2,1,config.getLastRow()-1,5).getValues() : [];
  const cfg = rows.find(r => String(r[0] || "").trim().toLowerCase() === subject.toLowerCase());
  if (!cfg) return {status:"error", message:"Không tìm thấy môn thi này trong CaiDatDeThi."};

  const forcedStatus = getExamStatusMap_()[String(cfg[0]).trim()];
  if (forcedStatus === "CLOSED") return {status:"error", message:"Đề thi hiện đang được giáo viên đóng."};
  const now = new Date();
  const start = parseSheetDate_(cfg[2]);
  const end = parseSheetDate_(cfg[3]);
  if (forcedStatus !== "OPEN") {
    if (start && now < start) return {status:"error", message:"Đề thi chưa đến giờ mở. Vui lòng quay lại sau."};
    if (end && now > end) return {status:"error", message:"Đề thi đã hết hạn làm bài."};
  }

  const bank = getSheetOrThrow_(SHEETS.BANK);
  if (bank.getLastRow() < 2) return {status:"error", message:"Ngân hàng đề chưa có câu hỏi."};
  const data = bank.getRange(2,1,bank.getLastRow()-1,10).getValues();
  const questions = data.filter(r =>
    String(r[1] || "").trim().toLowerCase() === subject.toLowerCase() &&
    String(r[0] || "").trim() !== "" &&
    ["MULTIPLE_CHOICE","FILL_BLANK"].includes(String(r[2] || "").trim().toUpperCase()) &&
    String(r[3] || "").trim() !== "" &&
    String(r[9] || "").trim() !== ""
  );
  if (!questions.length) return {status:"error", message:"Chưa có câu hỏi hợp lệ cho môn thi này. Kiểm tra MonHoc, LoaiCauHoi, NoiDungCauHoi và DapAnDung."};

  shuffle_(questions);
  const requested = Number(cfg[4]);
  const count = requested > 0 ? Math.min(Math.floor(requested), questions.length) : questions.length;
  const selected = questions.slice(0, count);
  const token = Utilities.getUuid();
  const answerKey = {};
  const publicQuestions = selected.map(r => {
    const id = String(r[0]).trim();
    const type = String(r[2]).trim().toUpperCase();
    answerKey[id] = {type:type, answer:String(r[9]).trim()};
    return {
      id:id, type:type, question:String(r[3]), image:safeImageUrl_(r[4]),
      a:String(r[5] || ""), b:String(r[6] || ""), c:String(r[7] || ""), d:String(r[8] || "")
    };
  });
  const minutes = Math.max(0, Number(cfg[1]) || 0);
  const session = {subject:String(cfg[0]).trim(), answerKey:answerKey, total:selected.length, createdAt:Date.now()};
  CacheService.getScriptCache().put(SESSION_PREFIX + token, JSON.stringify(session), 21600);
  return {status:"success", token:token, config:{thoiGian:minutes, monHoc:String(cfg[0])}, questions:publicQuestions};
}

function submitExam_(params) {
  const token = String(params.token || "");
  if (!token) return {status:"error", message:"Thiếu mã phiên làm bài. Vui lòng tải lại đề."};
  const cache = CacheService.getScriptCache();
  const sessionKey = SESSION_PREFIX + token;
  const submittedKey = SUBMITTED_PREFIX + token;
  if (cache.get(submittedKey)) return {status:"error", message:"Bài thi này đã được nộp trước đó."};
  const raw = cache.get(sessionKey);
  if (!raw) return {status:"error", message:"Phiên thi đã hết hạn hoặc không hợp lệ. Vui lòng tải lại đề."};
  const session = JSON.parse(raw);
  const answers = params.answers || {};
  let score = 0;
  const ids = Object.keys(session.answerKey);
  ids.forEach(id => {
    const item = session.answerKey[id];
    const userAnswer = normalizeAnswer_(answers[id]);
    const accepted = normalizeAnswer_(item.answer).split("|").map(s => s.trim()).filter(Boolean);
    if (accepted.includes(userAnswer)) score++;
  });
  const total = session.total;
  const grade = total ? Math.round((score / total * 10) * 100) / 100 : 0;
  const studentName = String(params.studentName || "").trim().slice(0,150);
  if (!studentName) return {status:"error", message:"Thiếu họ tên thí sinh."};

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    if (cache.get(submittedKey)) return {status:"error", message:"Bài thi này đã được nộp trước đó."};
    getSheetOrThrow_(SHEETS.RESULTS).appendRow([new Date(), studentName, session.subject, score, total, grade]);
    cache.put(submittedKey, "1", 21600);
    cache.remove(sessionKey);
  } finally {
    lock.releaseLock();
  }
  return {status:"success", score:score, total:total, diemSo:grade};
}

function normalizeAnswer_(value) {
  return String(value == null ? "" : value).trim().toLocaleLowerCase("vi").replace(/\s+/g, " ");
}
function parseSheetDate_(value) {
  if (!value) return null;
  if (value instanceof Date && !isNaN(value.getTime())) return value;
  const s = String(value).trim();
  const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s+(\d{1,2}):(\d{2}))?$/);
  if (!m) {
    const d = new Date(s);
    return isNaN(d.getTime()) ? null : d;
  }
  return new Date(Number(m[3]), Number(m[2])-1, Number(m[1]), Number(m[4]||0), Number(m[5]||0), 0);
}
function safeImageUrl_(value) {
  const s = String(value || "").trim();
  return /^https:\/\/\S+$/i.test(s) || /^http:\/\/\S+$/i.test(s) ? s : "";
}
function shuffle_(arr) {
  for (let i=arr.length-1; i>0; i--) {
    const j=Math.floor(Math.random()*(i+1));
    [arr[i],arr[j]]=[arr[j],arr[i]];
  }
}
function getSheetOrThrow_(name) {
  const sh = getSpreadsheet_().getSheetByName(name);
  if (!sh) throw new Error("Thiếu trang tính " + name + ". Hãy chạy setupSheets() trong Apps Script.");
  return sh;
}
function jsonOutput_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}


/* ================= QUAN TRI ADMIN ================= */
/** Chay thu cong mot lan trong Apps Script: setAdminPassword("mat-khau-manh-cua-ban") */
function setAdminPassword(password) {
  // Script độc lập không có giao diện bảng tính; truyền mật khẩu trực tiếp khi chạy hàm.
  password = String(password || "");
  if (password.length < 6) throw new Error("Mật khẩu Admin phải có ít nhất 6 ký tự.");
  const salt = Utilities.getUuid() + Utilities.getUuid();
  const hash = hashPassword_(password, salt);
  PropertiesService.getScriptProperties().setProperties({ADMIN_SALT:salt, ADMIN_HASH:hash});
  return "Đã thiết lập mật khẩu Admin. Hãy xóa mật khẩu khỏi lịch sử chạy nếu có thể.";
}
function adminLogin_(params) {
  const props = PropertiesService.getScriptProperties();
  const salt = props.getProperty("ADMIN_SALT");
  const expected = props.getProperty("ADMIN_HASH");
  if (!salt || !expected) return {status:"error", message:"Admin chưa được thiết lập mật khẩu. Chủ hệ thống hãy chạy setAdminPassword() trong Apps Script."};
  const supplied = hashPassword_(String(params.password || ""), salt);
  if (!constantTimeEquals_(supplied, expected)) return {status:"error", message:"Mật khẩu Admin không chính xác."};
  const token = Utilities.getUuid() + Utilities.getUuid();
  CacheService.getScriptCache().put(ADMIN_TOKEN_PREFIX + token, "1", ADMIN_TOKEN_TTL);
  return {status:"success", token:token, expiresIn:ADMIN_TOKEN_TTL};
}
function hashPassword_(password, salt) {
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + "|" + password, Utilities.Charset.UTF_8);
  return Utilities.base64EncodeWebSafe(digest);
}
function constantTimeEquals_(a,b) {
  a=String(a); b=String(b); if(a.length!==b.length) return false;
  let diff=0; for(let i=0;i<a.length;i++) diff |= a.charCodeAt(i)^b.charCodeAt(i);
  return diff===0;
}
function requireAdmin_(token) {
  if (!token || !CacheService.getScriptCache().get(ADMIN_TOKEN_PREFIX + String(token))) {
    throw new Error("Phiên Admin hết hạn hoặc không hợp lệ. Vui lòng đăng nhập lại.");
  }
}
function handleAdminAction_(p) {
  switch(p.action) {
    case "adminDashboard": return adminDashboard_();
    case "adminSaveSubject": return adminSaveSubject_(p);
    case "adminSetStatus": return adminSetStatus_(p);
    case "adminAddQuestion": return adminAddQuestion_(p.question || {});
    case "adminImportQuestions": return adminImportQuestions_(p.tsv || "");
    case "adminDeleteQuestion": return adminDeleteQuestion_(p.id);
    case "adminGetResults": return adminGetResults_(p);
    case "adminLogout": CacheService.getScriptCache().remove(ADMIN_TOKEN_PREFIX + String(p.token)); return {status:"success"};
    default: return {status:"error", message:"Lệnh quản trị không hợp lệ."};
  }
}
function adminDashboard_() {
  const cfgSh=getSheetOrThrow_(SHEETS.CONFIG), bankSh=getSheetOrThrow_(SHEETS.BANK);
  const cfg=cfgSh.getLastRow()>1?cfgSh.getRange(2,1,cfgSh.getLastRow()-1,5).getValues():[];
  const bank=bankSh.getLastRow()>1?bankSh.getRange(2,1,bankSh.getLastRow()-1,10).getValues():[];
  const statusMap=getExamStatusMap_(), now=new Date();
  const subjects=cfg.filter(r=>String(r[0]||"").trim()).map(r=>{
    const name=String(r[0]).trim(), start=parseSheetDate_(r[2]), end=parseSheetDate_(r[3]);
    let status=statusMap[name]||"AUTO";
    if(status==="AUTO") status=(start&&now<start)?"SCHEDULED":((end&&now>end)?"EXPIRED":"OPEN");
    return {monHoc:name,thoiGian:Number(r[1])||0,ngayBatDau:formatDateTime_(r[2]),ngayKetThuc:formatDateTime_(r[3]),soCau:Number(r[4])||0,status:status,soCauTrongNganHang:bank.filter(q=>String(q[1]).trim().toLowerCase()===name.toLowerCase()).length};
  });
  return {status:"success",subjects:subjects,totalQuestions:bank.length};
}
function adminSaveSubject_(p) {
  const name=String(p.monHoc||"").trim();
  if(!name) throw new Error("Mã/tên môn thi không được để trống.");
  const minutes=Math.max(0,Number(p.thoiGian)||0), count=Math.max(0,Math.floor(Number(p.soCau)||0));
  const start=parseAdminDate_(p.ngayBatDau), end=parseAdminDate_(p.ngayKetThuc);
  if(start&&end&&end<start) throw new Error("Ngày kết thúc phải sau ngày bắt đầu.");
  const sh=getSheetOrThrow_(SHEETS.CONFIG), last=sh.getLastRow();
  const rows=last>1?sh.getRange(2,1,last-1,5).getValues():[];
  const idx=rows.findIndex(r=>String(r[0]||"").trim().toLowerCase()===name.toLowerCase());
  const row=[name,minutes,start||"",end||"",count];
  if(idx>=0) sh.getRange(idx+2,1,1,5).setValues([row]); else sh.appendRow(row);
  sh.getRange(2,3,Math.max(1,sh.getMaxRows()-1),2).setNumberFormat("dd/MM/yyyy HH:mm");
  return {status:"success",message:idx>=0?"Đã cập nhật môn thi.":"Đã thêm môn thi."};
}
function adminSetStatus_(p) {
  const name=String(p.monHoc||"").trim(), status=String(p.status||"").toUpperCase();
  if(!name||!["OPEN","CLOSED","AUTO"].includes(status)) throw new Error("Thông tin điều khiển đề không hợp lệ.");
  const subjects=getSubjects_();
  if(!subjects.some(s=>s.toLowerCase()===name.toLowerCase())) throw new Error("Không tìm thấy môn thi.");
  const map=getExamStatusMap_();
  if(status==="AUTO") delete map[name]; else map[name]=status;
  PropertiesService.getScriptProperties().setProperty("EXAM_STATUS_MAP",JSON.stringify(map));
  return {status:"success",message:status==="OPEN"?"Đã mở đề thi.":status==="CLOSED"?"Đã đóng đề thi.":"Đã chuyển về chế độ tự động theo lịch."};
}
function getExamStatusMap_(){try{return JSON.parse(PropertiesService.getScriptProperties().getProperty("EXAM_STATUS_MAP")||"{}");}catch(e){return {};}}
function adminAddQuestion_(q) {
  const id=String(q.id||"").trim(), subject=String(q.monHoc||"").trim(), type=String(q.loaiCauHoi||"").trim().toUpperCase();
  const content=String(q.noiDungCauHoi||"").trim(), correct=String(q.dapAnDung||"").trim();
  if(!id||!subject||!content||!correct) throw new Error("Cần nhập Id, môn học, nội dung và đáp án đúng.");
  if(!["MULTIPLE_CHOICE","FILL_BLANK"].includes(type)) throw new Error("Loại câu hỏi không hợp lệ.");
  if(!getSubjects_().some(s=>s.toLowerCase()===subject.toLowerCase())) throw new Error("Môn học chưa có trong CaiDatDeThi.");
  if(type==="MULTIPLE_CHOICE") {
    if(!/^[ABCD]$/i.test(correct)) throw new Error("Đáp án trắc nghiệm phải là A, B, C hoặc D.");
    if([q.dapAnA,q.dapAnB,q.dapAnC,q.dapAnD].some(v=>!String(v||"").trim())) throw new Error("Câu trắc nghiệm phải có đủ 4 lựa chọn A, B, C, D.");
  }
  const sh=getSheetOrThrow_(SHEETS.BANK), data=sh.getLastRow()>1?sh.getRange(2,1,sh.getLastRow()-1,10).getValues():[];
  if(data.some(r=>String(r[0]).trim()===id)) throw new Error("Id câu hỏi đã tồn tại: "+id);
  sh.appendRow([id,subject,type,content,safeImageUrl_(q.hinhAnh),type==="MULTIPLE_CHOICE"?q.dapAnA||"":"",type==="MULTIPLE_CHOICE"?q.dapAnB||"":"",type==="MULTIPLE_CHOICE"?q.dapAnC||"":"",type==="MULTIPLE_CHOICE"?q.dapAnD||"":"",correct]);
  return {status:"success",message:"Đã thêm câu hỏi "+id+"."};
}
function adminImportQuestions_(tsv) {
  const lines=String(tsv||"").split(/\r?\n/).filter(line=>line.trim());
  if(!lines.length) throw new Error("Chưa có dữ liệu để nhập.");
  let start=0;
  const firstCells=lines[0].split("\t").map(v=>v.trim().toLowerCase());
  if(firstCells[0]==="id" && firstCells[1]==="monhoc") start=1;
  const sh=getSheetOrThrow_(SHEETS.BANK), existing=sh.getLastRow()>1?sh.getRange(2,1,sh.getLastRow()-1,1).getValues().flat().map(String):[];
  const seen=new Set(existing), validSubjects=getSubjects_().map(s=>s.toLowerCase()), rows=[], errors=[];
  for(let i=start;i<lines.length;i++){
    const c=lines[i].split("\t");
    if(c.length<10){errors.push("Dòng "+(i+1)+": cần đủ 10 cột.");continue;}
    const id=String(c[0]).trim(), subject=String(c[1]).trim(), type=String(c[2]).trim().toUpperCase();
    if(!id||seen.has(id)){errors.push("Dòng "+(i+1)+": Id trống hoặc bị trùng ("+id+").");continue;}
    if(!validSubjects.includes(subject.toLowerCase())){errors.push("Dòng "+(i+1)+": môn học chưa được khai báo.");continue;}
    if(!["MULTIPLE_CHOICE","FILL_BLANK"].includes(type)||!String(c[3]).trim()||!String(c[9]).trim()){errors.push("Dòng "+(i+1)+": thiếu nội dung/đáp án hoặc sai loại câu hỏi.");continue;}
    if(type==="MULTIPLE_CHOICE"&&(!/^[ABCD]$/i.test(String(c[9]).trim())||[c[5],c[6],c[7],c[8]].some(v=>!String(v||"").trim()))){errors.push("Dòng "+(i+1)+": trắc nghiệm cần đủ 4 lựa chọn và đáp án đúng A/B/C/D.");continue;}
    rows.push([id,subject,type,c[3],safeImageUrl_(c[4]),type==="MULTIPLE_CHOICE"?c[5]:"",type==="MULTIPLE_CHOICE"?c[6]:"",type==="MULTIPLE_CHOICE"?c[7]:"",type==="MULTIPLE_CHOICE"?c[8]:"",c[9]]); seen.add(id);
  }
  if(rows.length) sh.getRange(sh.getLastRow()+1,1,rows.length,10).setValues(rows);
  return {status:"success",imported:rows.length,errors:errors.slice(0,30),message:"Đã nhập "+rows.length+" câu hỏi."};
}
function adminDeleteQuestion_(id) {
  id=String(id||"").trim(); const sh=getSheetOrThrow_(SHEETS.BANK);
  if(sh.getLastRow()<2) throw new Error("Ngân hàng đề đang trống.");
  const ids=sh.getRange(2,1,sh.getLastRow()-1,1).getValues().flat();
  const idx=ids.findIndex(v=>String(v).trim()===id);
  if(idx<0) throw new Error("Không tìm thấy Id câu hỏi.");
  sh.deleteRow(idx+2); return {status:"success",message:"Đã xóa câu hỏi "+id+"."};
}
function adminGetResults_(p) {
  const sh=getSheetOrThrow_(SHEETS.RESULTS), last=sh.getLastRow();
  if(last<2) return {status:"success",results:[]};
  const values=sh.getRange(2,1,last-1,6).getValues();
  const filter=String(p.monHoc||"").trim().toLowerCase();
  const results=values.filter(r=>!filter||String(r[2]).toLowerCase()===filter).reverse().slice(0,1000).map(r=>({thoiGian:formatDateTime_(r[0]),tenHocSinh:String(r[1]),monHoc:String(r[2]),soCauDung:Number(r[3])||0,tongSoCau:Number(r[4])||0,diemSo:Number(r[5])||0}));
  return {status:"success",results:results};
}
function parseAdminDate_(value) {
  if(!value) return null;
  const s=String(value).trim();
  const m=s.match(/^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2})$/);
  if(!m) throw new Error("Ngày giờ không hợp lệ.");
  const d=new Date(Number(m[1]),Number(m[2])-1,Number(m[3]),Number(m[4]),Number(m[5]));
  if(isNaN(d.getTime())) throw new Error("Ngày giờ không hợp lệ.");
  return d;
}
function formatDateTime_(value) {
  if(!value) return "";
  const d=value instanceof Date?value:new Date(value);
  if(isNaN(d.getTime())) return String(value);
  return Utilities.formatDate(d,Session.getScriptTimeZone(),"dd/MM/yyyy HH:mm");
}
function taoMatKhauAdminMotLan() {
  setAdminPassword("Admin123");
}
