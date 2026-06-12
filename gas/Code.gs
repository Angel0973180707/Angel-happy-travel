// =====================================================
// ✈️ 天使幸福旅居活動管理系統 v2.3
// 收款帳號與管理密碼請透過試算表選單設定，不要寫入原始碼。
// =====================================================

const BANK_INFO = {
bankName: '玉山銀行',
branchName: '小港分行',
accountName: '李秀芳',
deadline: 3
};

const TR_SHEET_TABS = [
{ name: '01_活動主檔', headers: ['活動代碼','活動名稱','活動類型','出發日期','回程日期','天數','地點','團費','名額上限','已報名人數','剩餘名額','報名表連結','活動狀態','備註'] },
{ name: '02_報名資料', headers: ['報名編號','報名時間','活動代碼','活動名稱','品牌','姓名','性別','出生日期','身分證字號','手機','LINE ID','Email','地址','審核狀態','備註'] },
{ name: '03_護照簽證', headers: ['報名編號','活動代碼','姓名','護照英文姓名','護照號碼','護照到期日','護照照片連結','簽證狀態','簽證備註'] },
{ name: '04_住宿房務', headers: ['報名編號','活動代碼','姓名','房型需求','指定同房者','實際房號','室友姓名','單人房補差','房務備註'] },
{ name: '05_飲食健康', headers: ['報名編號','活動代碼','姓名','飲食需求','慢性疾病','是否可正常步行','特殊需求','緊急聯絡人','關係','緊急電話'] },
{ name: '06_繳費管理', headers: ['報名編號','活動代碼','姓名','團費','繳費類型','訂金金額','訂金日期','訂金末五碼','尾款金額','尾款日期','尾款末五碼','繳費方式','繳費狀態','對帳狀態','對帳日期','對帳人員','收款備註'] },
{ name: '07_保險資料', headers: ['報名編號','活動代碼','姓名','出生日期','身分證字號','保險公司','保單號碼','保險狀態','備註'] },
{ name: '08_LINE群組', headers: ['報名編號','活動代碼','姓名','手機','LINE ID','是否加入群組','行前說明會','護照已收','保險完成','備註'] },
{ name: '09_行前通知', headers: ['活動代碼','通知日期','通知主題','通知內容','寄送對象','寄送狀態','備註'] },
{ name: '10_成本利潤表', headers: ['活動代碼','活動名稱','收入總額','旅行社成本','機票成本','住宿成本','交通成本','保險成本','雜支成本','總成本','預估利潤','實際利潤','備註'] },
{ name: '11_同行人資料', headers: ['主報名編號','活動代碼','主報名人','同行人姓名','關係','手機','身分證／護照號碼','希望同房','備註'] }
];

function onOpen() {
SpreadsheetApp.getUi()
.createMenu('✈️ 幸福旅居系統')
.addItem('初始化系統', 'initTRSystem')
.addItem('設定管理密碼', 'setAdminPassword')
.addItem('設定收款帳號', 'setBankAccountNo')
.addItem('更新名額統計', 'updateTRQuota')
.addItem('寄送行前通知', 'sendTRNotice')
.addSeparator()
.addItem('確認同行（選取列）', 'approveSelected')
.addItem('本次無法同行（選取列）', 'rejectSelected')
.addToUi();
}

function setAdminPassword() {
const ui = SpreadsheetApp.getUi();
const result = ui.prompt('設定管理密碼', '請輸入後台管理密碼（至少 8 個字元）', ui.ButtonSet.OK_CANCEL);
if (result.getSelectedButton() !== ui.Button.OK) return;
const password = result.getResponseText().trim();
if (password.length < 8) {
ui.alert('管理密碼至少需要 8 個字元');
return;
}
PropertiesService.getScriptProperties().setProperty('ADMIN_PASSWORD', password);
ui.alert('管理密碼已儲存');
}

function getAdminPassword_() {
return PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD') || '';
}

function setBankAccountNo() {
const ui = SpreadsheetApp.getUi();
const result = ui.prompt('設定收款帳號', '請輸入收款銀行帳號（只輸入數字）', ui.ButtonSet.OK_CANCEL);
if (result.getSelectedButton() !== ui.Button.OK) return;
const accountNo = result.getResponseText().replace(/\D/g, '');
if (accountNo.length < 8) {
ui.alert('請確認收款帳號是否完整');
return;
}
PropertiesService.getScriptProperties().setProperty('BANK_ACCOUNT_NO', accountNo);
ui.alert('收款帳號已儲存');
}

function getBankAccountNo_() {
return PropertiesService.getScriptProperties().getProperty('BANK_ACCOUNT_NO') || '請洽 LINE 官方帳號';
}

function initTRSystem() {
const ss = SpreadsheetApp.getActiveSpreadsheet();
TR_SHEET_TABS.forEach(tab => {
let sheet = ss.getSheetByName(tab.name);
if (!sheet) sheet = ss.insertSheet(tab.name);
if (sheet.getLastRow() === 0) {
sheet.appendRow(tab.headers);
sheet.getRange(1, 1, 1, tab.headers.length)
.setBackground('#1a4a7a').setFontColor('#ffffff').setFontWeight('bold');
sheet.setFrozenRows(1);
}
});
const mainSheet = ss.getSheetByName('01_活動主檔');
if (mainSheet.getLastRow() === 1) {
mainSheet.appendRow([
'TR001','越南下龍灣歡樂遊','海外旅遊',
'2026/07/21','2026/07/28','8天7夜',
'香港、深圳、南寧、越南下龍灣',
'21800','20','0','20','','招生中',''
]);
}
SpreadsheetApp.getUi().alert('✅ 天使幸福旅居系統初始化完成！');
}

function normalizeActivityCode(value) {
const match = String(value || '').trim().match(/^([A-Za-z]+\d+)/);
return match ? match[1].toUpperCase() : '';
}

function isOccupiedRegistrationStatus_(value) {
const status = String(value || '').trim();
return [
'待審核', '已通過', '已繳費',
'待確認', '確認同行'
].includes(status);
}

function getRegistrationCounts_(ss) {
const sheet = ss.getSheetByName('02_報名資料');
const data = sheet ? sheet.getDataRange().getValues() : [];
const counts = {};
for (let i = 1; i < data.length; i++) {
const code = normalizeActivityCode(data[i][2]);
if (code && isOccupiedRegistrationStatus_(data[i][13])) {
counts[code] = (counts[code] || 0) + 1;
}
}
return counts;
}

function findActiveDuplicateRegistration_(ss, activityCode, idNo) {
const code = normalizeActivityCode(activityCode);
const normalizedIdNo = String(idNo || '').trim().toUpperCase();
if (!code || !normalizedIdNo) return null;
const sheet = ss.getSheetByName('02_報名資料');
const data = sheet ? sheet.getDataRange().getValues() : [];
for (let i = 1; i < data.length; i++) {
if (
normalizeActivityCode(data[i][2]) === code &&
String(data[i][8] || '').trim().toUpperCase() === normalizedIdNo &&
isOccupiedRegistrationStatus_(data[i][13])
) {
return { regNo: data[i][0], status: data[i][13] };
}
}
return null;
}

function isInternationalActivity_(activity) {
const type = String(activity && activity.type || '').toLowerCase();
return type.includes('海外') || type.includes('國外') || type.includes('international');
}

function isValidEmail_(value) {
return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(value || '').trim());
}

function getOrCreateSheet_(ss, name) {
let sheet = ss.getSheetByName(name);
if (sheet) return sheet;
const definition = TR_SHEET_TABS.find(tab => tab.name === name);
if (!definition) throw new Error('找不到工作表設定：' + name);
sheet = ss.insertSheet(name);
sheet.appendRow(definition.headers);
sheet.getRange(1, 1, 1, definition.headers.length)
.setBackground('#1a4a7a').setFontColor('#ffffff').setFontWeight('bold');
sheet.setFrozenRows(1);
return sheet;
}

function savePassportPhoto_(photo, regNo) {
if (!photo || !photo.data) return '';
const allowedTypes = ['image/jpeg', 'image/png', 'image/webp', 'application/pdf'];
const mimeType = String(photo.mimeType || '');
if (!allowedTypes.includes(mimeType)) throw new Error('護照照片格式不支援');
const bytes = Utilities.base64Decode(photo.data);
if (bytes.length > 5 * 1024 * 1024) throw new Error('護照照片不可超過 5 MB');

const folderName = '幸福旅居_護照上傳';
const folders = DriveApp.getFoldersByName(folderName);
const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
const extension = mimeType === 'application/pdf' ? '.pdf'
: mimeType === 'image/png' ? '.png'
: mimeType === 'image/webp' ? '.webp' : '.jpg';
const blob = Utilities.newBlob(bytes, mimeType, regNo + '_passport' + extension);
return folder.createFile(blob).getUrl();
}

function generateTRRegNo_(activityCode) {
const code = normalizeActivityCode(activityCode);
if (!code) throw new Error('活動代碼格式錯誤');
const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('02_報名資料');
const data = sheet.getDataRange().getValues();
let maxNo = 0;
for (let i = 1; i < data.length; i++) {
const regNo = data[i][0].toString();
if (regNo.startsWith(code + '-')) {
const num = parseInt(regNo.split('-').pop());
if (num > maxNo) maxNo = num;
}
}
return code + '-' + String(maxNo + 1).padStart(3, '0');
}

function getTRActivityInfo_(activityCode) {
const code = normalizeActivityCode(activityCode);
if (!code) return null;
const data = SpreadsheetApp.getActiveSpreadsheet()
.getSheetByName('01_活動主檔').getDataRange().getValues();
for (let i = 1; i < data.length; i++) {
if (normalizeActivityCode(data[i][0]) === code) {
return {
code: normalizeActivityCode(data[i][0]), name: data[i][1], type: data[i][2],
startDate: data[i][3], endDate: data[i][4],
fee: data[i][7], limit: parseInt(data[i][8]) || 0,
status: data[i][12], row: i + 1
};
}
}
return null;
}

function onTRFormSubmit(e) {
const lock = LockService.getScriptLock();
lock.waitLock(10000);
try {
const ss = SpreadsheetApp.getActiveSpreadsheet();
const r = e.namedValues;
const activityCode = normalizeActivityCode(
(r['活動代碼'] || r['請填入活動代碼'] || r[' 請填入活動代碼'] || [''])[0]
);
const activity = getTRActivityInfo_(activityCode);
if (!activity) throw new Error('找不到活動代碼');
if (activity.status !== '招生中') throw new Error('活動目前已停招');
const name = (r['姓名'] || [''])[0];
const email = (r['Email'] || [''])[0];
const idNo = (r['身分證字號']||[''])[0];
const duplicate = findActiveDuplicateRegistration_(ss, activityCode, idNo);
if (duplicate) {
console.warn('略過重複報名：' + duplicate.regNo);
return;
}
const counts = getRegistrationCounts_(ss);
if (activity.limit > 0 && (counts[activityCode] || 0) >= activity.limit) {
throw new Error('活動名額已滿');
}
const regNo = generateTRRegNo_(activityCode);
writeRegistration_(ss, regNo, new Date(), activity, activityCode, name, email, {
gender: (r['性別']||[''])[0], birthday: (r['出生日期']||[''])[0],
idNo: idNo, phone: (r['手機']||[''])[0],
lineId: (r['LINE ID']||[''])[0], address: (r['地址']||[''])[0],
passportName: (r['護照英文姓名']||[''])[0], passportNo: (r['護照號碼']||[''])[0],
passportExpiry: (r['護照到期日']||[''])[0], roomType: (r['房型需求']||[''])[0],
food: (r['飲食需求']||[''])[0], walking: (r['是否可正常步行']||[''])[0],
chronic: (r['慢性疾病或用藥狀況']||[''])[0], special: (r['特殊需求']||[''])[0],
emergencyName: (r['緊急聯絡人']||[''])[0], emergencyRel: (r['緊急聯絡人關係']||[''])[0],
emergencyPhone: (r['緊急聯絡電話']||[''])[0],
passportPhotoUrl: (r['護照照片上傳']||[''])[0],
companions: [],
insuranceConsent: (r['同意辦理保險']||[''])[0] === '是',
insuranceNote: (r['投保需求或既有保險說明']||[''])[0]
});
updateTRQuota();
if (email) sendPendingEmail_(email, name, activity, regNo);
} finally {
lock.releaseLock();
}
}

function writeRegistration_(ss, regNo, regTime, activity, activityCode, name, email, d) {
const code = normalizeActivityCode(activity ? activity.code : activityCode);
if (!code) throw new Error('活動代碼格式錯誤');
ss.getSheetByName('02_報名資料').appendRow([
regNo, regTime, code, activity ? activity.name : '', '天使幸福旅居',
name, d.gender, d.birthday, d.idNo, d.phone,
d.lineId, email, d.address, '待確認', ''
]);
ss.getSheetByName('03_護照簽證').appendRow([
regNo, code, name, d.passportName, d.passportNo, d.passportExpiry, d.passportPhotoUrl || '', '待確認', ''
]);
ss.getSheetByName('04_住宿房務').appendRow([
regNo, code, name, d.roomType, '', '', '', '', ''
]);
ss.getSheetByName('05_飲食健康').appendRow([
regNo, code, name, d.food, d.chronic,
d.walking, d.special, d.emergencyName, d.emergencyRel, d.emergencyPhone
]);
ss.getSheetByName('06_繳費管理').appendRow([
regNo, code, name, activity ? activity.fee : '',
'', '', '', '', '', '', '', '', '待繳費', '待對帳', '', '', ''
]);
ss.getSheetByName('07_保險資料').appendRow([
regNo, code, name, d.birthday, d.idNo, '', '', '待投保',
(d.insuranceConsent ? '已同意辦理保險' : '未勾選保險同意') +
(d.insuranceNote ? '；' + d.insuranceNote : '')
]);
ss.getSheetByName('08_LINE群組').appendRow([
regNo, code, name, d.phone, d.lineId, '否', '否', '否', '否', ''
]);
const companionSheet = getOrCreateSheet_(ss, '11_同行人資料');
(Array.isArray(d.companions) ? d.companions : []).forEach(companion => {
if (!companion || !String(companion.name || '').trim()) return;
companionSheet.appendRow([
regNo, code, name, String(companion.name).trim(),
String(companion.relation || '').trim(), String(companion.phone || '').trim(),
String(companion.idNo || '').trim().toUpperCase(), companion.sameRoom ? '是' : '否', ''
]);
});
}

function approveSelected() {
const ss = SpreadsheetApp.getActiveSpreadsheet();
const sheet = ss.getSheetByName('02_報名資料');
const row = sheet.getActiveRange().getRow();
if (row <= 1) { SpreadsheetApp.getUi().alert('請先選取要確認的資料列'); return; }
const data = sheet.getRange(row, 1, 1, 15).getValues()[0];
sheet.getRange(row, 14).setValue('確認同行');
updateTRQuota();
const activity = getTRActivityInfo_(data[2]);
sendApproveEmail_(data[11], data[5], activity, data[0]);
SpreadsheetApp.getUi().alert('✅ 已確認同行！\n\n── LINE 複製文案 ──\n\n' + buildApproveCopy_(data[5], activity ? activity.name : data[2]));
}

function rejectSelected() {
const ui = SpreadsheetApp.getUi();
const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('02_報名資料');
const row = sheet.getActiveRange().getRow();
if (row <= 1) { ui.alert('請先選取資料列'); return; }
const reasonInput = ui.prompt('無法同行原因');
if (reasonInput.getSelectedButton() !== ui.Button.OK) return;
const reason = reasonInput.getResponseText().trim() || '本次活動名額已滿';
const data = sheet.getRange(row, 1, 1, 15).getValues()[0];
sheet.getRange(row, 14).setValue('無法同行');
sheet.getRange(row, 15).setValue(reason);
updateTRQuota();
const activity = getTRActivityInfo_(data[2]);
sendRejectEmail_(data[11], data[5], activity, data[0], reason);
ui.alert('已通知！\n\n── LINE 複製文案 ──\n\n' + buildRejectCopy_(data[5], activity ? activity.name : data[2], reason));
}

function buildApproveCopy_(name, activityName) {
return `✈️ 天使幸福旅居｜確認同行通知\n親愛的 ${name}，\n您登記參與的【${activityName}】已確認您的同行名額！\n請於 ${BANK_INFO.deadline} 天內完成費用匯款，以確保您的名額。\n匯款資訊已寄至您的 Email，請確認收信。\n期待與您一起踏上這段美好旅程 🌏`;
}

function buildRejectCopy_(name, activityName, reason) {
return `✈️ 天使幸福旅居｜同行確認結果\n親愛的 ${name}，\n感謝您登記參與【${activityName}】。\n很遺憾這次無法安排您同行，${reason}。\n歡迎繼續關注我們後續活動，期待下次與您相遇 🙏`;
}

function sendPendingEmail_(email, name, activity, regNo) {
const subject = '✈️ 天使幸福旅居｜已收到您的同行登記';
const body =
`親愛的 ${name} 您好，\n\n` +
`感謝您登記參與天使幸福旅居活動！\n我們已收到您的資料，正在確認同行名額。\n\n` +
`📋 登記資訊\n─────────────────\n` +
`登記編號：${regNo}\n活動名稱：${activity ? activity.name : ''}\n` +
`出發日期：${activity ? formatDate_(activity.startDate) : ''}\n` +
`─────────────────\n\n` +
`📌 請保持手機與 LINE 暢通，確認結果將以 Email 通知。\n\n` +
`生命是一場旅居，\n最美的風景在路上，\n最好的時光與好友同行。\n\n` +
`── 天使幸福旅居 召集人 敬上`;
GmailApp.sendEmail(email, subject, body);
return true;
}

function sendApproveEmail_(email, name, activity, regNo) {
const subject = '✈️ 天使幸福旅居｜同行名額確認，請完成費用匯款';
const body =
`親愛的 ${name} 您好，\n\n很高興確認您的同行名額！期待與您一起出發 🌏\n\n` +
`📋 登記資訊\n─────────────────\n` +
`登記編號：${regNo}\n活動名稱：${activity ? activity.name : ''}\n` +
`出發日期：${activity ? formatDate_(activity.startDate) : ''}\n` +
`回程日期：${activity ? formatDate_(activity.endDate) : ''}\n` +
`活動共同費用：NT$ ${activity ? activity.fee : ''}\n` +
`─────────────────\n\n` +
`💳 匯款資訊\n─────────────────\n` +
`銀行：${BANK_INFO.bankName} ${BANK_INFO.branchName}\n` +
`戶名：${BANK_INFO.accountName}\n` +
`帳號：${getBankAccountNo_()}\n` +
`─────────────────\n\n` +
`📌 請於收到此信 ${BANK_INFO.deadline} 天內完成訂金匯款\n` +
`匯款後請至繳費回報頁填寫末五碼，以利確認。\n` +
`護照請確認有效期限距出發日 6 個月以上。\n\n` +
`── 天使幸福旅居 召集人 敬上`;
GmailApp.sendEmail(email, subject, body);
}

function sendRejectEmail_(email, name, activity, regNo, reason) {
const subject = '✈️ 天使幸福旅居｜同行確認結果通知';
const body =
`親愛的 ${name} 您好，\n\n` +
`感謝您登記參與【${activity ? activity.name : ''}】。\n` +
`很遺憾，這次無法安排您同行。\n原因：${reason}\n\n` +
`歡迎繼續關注天使幸福旅居後續活動，\n期待下次能與您同行 🙏\n\n` +
`── 天使幸福旅居 召集人 敬上`;
GmailApp.sendEmail(email, subject, body);
}

function updateTRQuota() {
const ss = SpreadsheetApp.getActiveSpreadsheet();
const mainSheet = ss.getSheetByName('01_活動主檔');
const mainData = mainSheet.getDataRange().getValues();
const counts = getRegistrationCounts_(ss);
for (let i = 1; i < mainData.length; i++) {
const code = normalizeActivityCode(mainData[i][0]);
const limit = parseInt(mainData[i][8]) || 0;
const count = counts[code] || 0;
mainSheet.getRange(i + 1, 10).setValue(count);
mainSheet.getRange(i + 1, 11).setValue(Math.max(0, limit - count));
}
}

function sendTRNotice() {
const ss = SpreadsheetApp.getActiveSpreadsheet();
const ui = SpreadsheetApp.getUi();
const codeInput = ui.prompt('請輸入活動代碼（例：TR001）');
if (codeInput.getSelectedButton() !== ui.Button.OK) return;
const code = codeInput.getResponseText().trim();
const noticeData = ss.getSheetByName('09_行前通知').getDataRange().getValues();
const regData = ss.getSheetByName('02_報名資料').getDataRange().getValues();
let notice = null;
for (let i = 1; i < noticeData.length; i++) {
if (noticeData[i][0] === code && noticeData[i][5] === '待發送') {
notice = { row: i + 1, subject: noticeData[i][2], content: noticeData[i][3] };
break;
}
}
if (!notice) { ui.alert('找不到待發送通知'); return; }
let count = 0;
for (let j = 1; j < regData.length; j++) {
if (regData[j][2] === code && regData[j][11] && regData[j][13] === '確認同行') {
GmailApp.sendEmail(regData[j][11], notice.subject, notice.content);
count++;
}
}

ss.getSheetByName('09_行前通知').getRange(notice.row, 6).setValue('已發送');
ui.alert(`✅ 已發送 ${count} 封行前通知！`);
}

function formatDate_(d) {
if (!d) return '';
if (d instanceof Date) return Utilities.formatDate(d, 'Asia/Taipei', 'yyyy/MM/dd');
return d.toString();
}

function doGet(e) {
const action = e.parameter.action;
const pwd = e.parameter.pwd;
if (action === 'getActiveEvents') return jsonOut(getActiveEvents_());
if (!getAdminPassword_() || pwd !== getAdminPassword_()) return jsonOut({ error: '密碼錯誤' });
if (action === 'getPending') return jsonOut(getPendingList_());
if (action === 'getAll') return jsonOut(getAllList_());
if (action === 'getPayments') return jsonOut(getPaymentList_());
return jsonOut({ error: '未知指令' });
}

function doPost(e) {
const params = JSON.parse(e.postData.contents);
if (params.action === 'register') return jsonOut(registerFromWeb_(params));
if (!getAdminPassword_() || params.pwd !== getAdminPassword_()) return jsonOut({ error: '密碼錯誤' });
if (params.action === 'approve') return jsonOut(approveByRegNo_(params.regNo));
if (params.action === 'reject') return jsonOut(rejectByRegNo_(params.regNo, params.reason));
if (params.action === 'submitPayment') return jsonOut(submitPayment_(params));
if (params.action === 'confirmPayment') return jsonOut(confirmPayment_(params.regNo, params.feeType));
return jsonOut({ error: '未知指令' });
}

function jsonOut(data) {
return ContentService.createTextOutput(JSON.stringify(data))
.setMimeType(ContentService.MimeType.JSON);
}

function getActiveEvents_() {
const ss = SpreadsheetApp.getActiveSpreadsheet();
const data = ss
.getSheetByName('01_活動主檔').getDataRange().getValues();
const counts = getRegistrationCounts_(ss);
const list = [];
for (let i = 1; i < data.length; i++) {
if (data[i][12] === '招生中') {
const code = normalizeActivityCode(data[i][0]);
const limit = parseInt(data[i][8]) || 0;
const registered = counts[code] || 0;
list.push({
code: code, name: data[i][1], type: data[i][2],
startDate: formatDate_(data[i][3]), endDate: formatDate_(data[i][4]),
days: data[i][5], location: data[i][6], fee: data[i][7],
limit: limit,
registered: registered,
remaining: Math.max(0, limit - registered)
});
}
}
return list;
}
function registerFromWeb_(p) {
const lock = LockService.getScriptLock();
lock.waitLock(10000);
try {
const ss = SpreadsheetApp.getActiveSpreadsheet();
const activityCode = normalizeActivityCode(p.activityCode);
const activity = getTRActivityInfo_(activityCode);
if (!activity) return { error: '找不到活動代碼' };
if (activity.status !== '招生中') return { error: '活動目前已停招' };
if (!isValidEmail_(p.email)) return { error: 'Email 格式不正確' };
if (!String(p.name || '').trim() || !String(p.idNo || '').trim()) {
return { error: '姓名與身分證字號為必填' };
}
if (!String(p.emergencyName || '').trim() || !String(p.emergencyPhone || '').trim()) {
return { error: '緊急聯絡人與電話為必填' };
}
if (!p.insuranceConsent) return { error: '請勾選保險資料使用同意' };

const international = isInternationalActivity_(activity);
if (international && (
!String(p.passportName || '').trim() ||
!String(p.passportNo || '').trim() ||
!String(p.passportExpiry || '').trim() ||
!p.passportPhoto
)) {
return { error: '國外團請完整填寫護照資料並上傳照片' };
}

const duplicate = findActiveDuplicateRegistration_(ss, activityCode, p.idNo);
if (duplicate) {
return { error: '此參加人已完成登記', regNo: duplicate.regNo };
}

const counts = getRegistrationCounts_(ss);
const registered = counts[activityCode] || 0;
if (activity.limit > 0 && registered >= activity.limit) {
return { error: '活動名額已滿' };
}

const regNo = generateTRRegNo_(activityCode);
const passportPhotoUrl = international ? savePassportPhoto_(p.passportPhoto, regNo) : '';
writeRegistration_(ss, regNo, new Date(), activity, activityCode, p.name, p.email, {
gender: p.gender||'', birthday: p.birthday||'', idNo: p.idNo||'',
phone: p.phone||'', lineId: p.lineId||'', address: p.address||'',
passportName: p.passportName||'', passportNo: p.passportNo||'',
passportExpiry: p.passportExpiry||'', passportPhotoUrl: passportPhotoUrl,
companions: Array.isArray(p.companions) ? p.companions.slice(0, 10) : [],
roomType: p.roomType||'',
food: p.food||'', walking: p.walking||'', chronic: p.chronic||'',
special: p.special||'', emergencyName: p.emergencyName||'',
emergencyRel: p.emergencyRel||'', emergencyPhone: p.emergencyPhone||'',
insuranceConsent: !!p.insuranceConsent, insuranceNote: p.insuranceNote||''
});
updateTRQuota();
let emailSent = false;
try {
emailSent = sendPendingEmail_(p.email, p.name, activity, regNo);
} catch (emailError) {
console.error('報名成功，但 Email 寄送失敗：' + emailError);
}
return { success: true, regNo, emailSent: emailSent };
} catch(err) {
return { error: err.toString() };
} finally {
lock.releaseLock();
}
}

function getPendingList_() {
const data = SpreadsheetApp.getActiveSpreadsheet()
.getSheetByName('02_報名資料').getDataRange().getValues();
const list = [];
for (let i = 1; i < data.length; i++) {
if (data[i][13] === '待確認') {
list.push({
regNo: data[i][0], time: data[i][1].toString(),
activityCode: data[i][2], activityName: data[i][3],
name: data[i][5], gender: data[i][6],
phone: data[i][9], lineId: data[i][10], email: data[i][11]
});
}
}
return list;
}

function getAllList_() {
const data = SpreadsheetApp.getActiveSpreadsheet()
.getSheetByName('02_報名資料').getDataRange().getValues();
const list = [];
for (let i = 1; i < data.length; i++) {
if (data[i][0]) {
list.push({
regNo: data[i][0], time: data[i][1].toString(),
activityCode: data[i][2], activityName: data[i][3],
name: data[i][5], phone: data[i][9],
email: data[i][11], status: data[i][13]
});
}
}
return list;
}

function getPaymentList_() {
const data = SpreadsheetApp.getActiveSpreadsheet()
.getSheetByName('06_繳費管理').getDataRange().getValues();
const list = [];
for (let i = 1; i < data.length; i++) {
if (data[i][13] === '訂金待對帳' || data[i][13] === '尾款待對帳') {
list.push({
regNo: data[i][0], activityCode: data[i][1],
name: data[i][2], feeType: data[i][4],
amount: data[i][5] || data[i][8],
last5: data[i][7] || data[i][10],
payStatus: data[i][13]
});
}
}
return list;
}

function approveByRegNo_(regNo) {
const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('02_報名資料');
const data = sheet.getDataRange().getValues();
for (let i = 1; i < data.length; i++) {
if (data[i][0] === regNo) {
sheet.getRange(i + 1, 14).setValue('確認同行');
updateTRQuota();
const activity = getTRActivityInfo_(data[i][2]);
sendApproveEmail_(data[i][11], data[i][5], activity, regNo);
return { success: true, copyText: buildApproveCopy_(data[i][5], activity ? activity.name : data[i][2]) };
}
}
return { error: '找不到登記編號' };
}

function rejectByRegNo_(regNo, reason) {
const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('02_報名資料');
const data = sheet.getDataRange().getValues();
const r = reason || '本次活動名額已滿';
for (let i = 1; i < data.length; i++) {
if (data[i][0] === regNo) {
sheet.getRange(i + 1, 14).setValue('無法同行');
sheet.getRange(i + 1, 15).setValue(r);
updateTRQuota();
const activity = getTRActivityInfo_(data[i][2]);
sendRejectEmail_(data[i][11], data[i][5], activity, regNo, r);
return { success: true, copyText: buildRejectCopy_(data[i][5], activity ? activity.name : data[i][2], r) };
}
}
return { error: '找不到登記編號' };
}

function submitPayment_(params) {
const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('06_繳費管理');
const data = sheet.getDataRange().getValues();
for (let i = 1; i < data.length; i++) {
if (data[i][0] === params.regNo) {
const today = new Date().toLocaleDateString('zh-TW');
if (params.type === '訂金' || params.type === '全額') {
sheet.getRange(i+1,5).setValue(params.type);
sheet.getRange(i+1,6).setValue(params.amount);
sheet.getRange(i+1,7).setValue(today);
sheet.getRange(i+1,8).setValue(params.last5);
sheet.getRange(i+1,13).setValue('訂金待對帳');
} else {
sheet.getRange(i+1,9).setValue(params.amount);
sheet.getRange(i+1,10).setValue(today);
sheet.getRange(i+1,11).setValue(params.last5);
sheet.getRange(i+1,13).setValue('尾款待對帳');
}
return { success: true };
}
}
return { error: '找不到登記編號' };
}

function confirmPayment_(regNo, feeType) {
const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('06_繳費管理');
const data = sheet.getDataRange().getValues();
for (let i = 1; i < data.length; i++) {
if (data[i][0] === regNo) {
sheet.getRange(i+1,13).setValue('已收款');
sheet.getRange(i+1,14).setValue('已確認');
sheet.getRange(i+1,15).setValue(new Date().toLocaleDateString('zh-TW'));
return { success: true };
}
}
return { error: 'Registration number not found' };
}
