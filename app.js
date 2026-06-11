const config = {
googleClientId: "",
spreadsheetId: "",
cellRange: "A1:H107",
className: "1年A組",
...(window.TIMETABLE_CONFIG || {}),
};
const weekdays = ["月", "火", "水", "木", "金"];
const dayIndexes = [1, 2, 3, 4, 5];
const cacheKey = `timetable:v5:${config.spreadsheetId}`;
const authorizationKey = `timetable:google-authorized:${config.googleClientId}`;
const tokenExpiryKey = "googleAccessTokenExpiresAt";
const tokenScope = "https://www.googleapis.com/auth/spreadsheets.readonly";
const firstWeekRows = [
{ name: 6, content: 7 },
{ name: 10, content: 11 },
{ name: 15, content: 16 },
{ name: 19, content: 20 },
];
const firstWeekdayColumn = 1;
const rowsPerWeek = 21;
const weekCount = 5;
const firstDateRow = 3;
const excludedMonths = new Set([8]);
const lessonAbbreviations = new Map([
["オウンドメディア演習", "OM演習"],
["ファッションプロモーション", "FP"],
["ムービーワーク", "MOV"],
["グラフィックデザインし", "Gデザ"],
["グラフィックデザイン", "Gデザ"],
["メディア活用論", "M活用"],
["カラープランニング", "カラー"],
["プロダクトデザイン", "Pデザ"],
["特別講義", "特講"],
["自由選択", "選択"],
["プロモーションフォト", "Pフォト"],
["ファッショーマーケティング論", "FMK論"],
["ファッションマーケティング論", "FMK論"],
["エディトリアルワーク", "エディ"],
["ファッション商品知識", "知識"],
["デジタルマーケティング", "デジタル"],
]);
const sampleLessons = [
[
{ name: "国語", content: "漢字テスト" },
{ name: "数学", content: "一次方程式" },
{ name: "英語", content: "Unit 3" },
{ name: "理科", content: "植物の分類" },
{ name: "社会", content: "鎌倉時代" },
],
[
{ name: "数学", content: "一次方程式" },
{ name: "国語", content: "漢字テスト" },
{ name: "理科", content: "植物の分類" },
{ name: "英語", content: "Unit 3" },
{ name: "数学", content: "一次方程式" },
],
[
{ name: "英語", content: "Unit 3" },
{ name: "理科", content: "物質の性質" },
{ name: "国語", content: "古典の読解" },
{ name: "社会", content: "鎌倉時代" },
{ name: "体育", content: "バレーボール" },
],
[
{ name: "社会", content: "鎌倉時代" },
{ name: "英語", content: "Unit 3" },
{ name: "数学", content: "一次方程式" },
{ name: "国語", content: "古典の読解" },
{ name: "理科", content: "物質の性質" },
],
];
const elements = {
accountButton: document.querySelector("#account-button"),
accountImage: document.querySelector("#account-image"),
accountInitial: document.querySelector("#account-initial"),
appTitle: document.querySelector("#app-title"),
className: document.querySelector("#class-name"),
dataSource: document.querySelector("#data-source"),
detailClose: document.querySelector("#detail-close"),
detailContent: document.querySelector("#detail-content"),
detailDialog: document.querySelector("#lesson-detail"),
detailTitle: document.querySelector("#detail-title"),
lastUpdated: document.querySelector("#last-updated"),
loadingOverlay: document.querySelector("#loading-overlay"),
notice: document.querySelector("#notice"),
nextWeek: document.querySelector("#next-week"),
previousWeek: document.querySelector("#previous-week"),
refreshButton: document.querySelector("#refresh-button"),
signInButton: document.querySelector("#sign-in-button"),
timetable: document.querySelector("#timetable"),
weekLabel: document.querySelector("#week-label"),
weekRange: document.querySelector("#week-range"),
};
let accessToken = sessionStorage.getItem("googleAccessToken") || "";
const storedTokenExpiry = Number(sessionStorage.getItem(tokenExpiryKey) || 0);
if (storedTokenExpiry && storedTokenExpiry <= Date.now()) {
accessToken = "";
sessionStorage.removeItem("googleAccessToken");
sessionStorage.removeItem(tokenExpiryKey);
}
let tokenClient;
let automaticAuthorizationPending = false;
let authorizationRestoreAttempts = 0;
let pendingAuthorization = { interactive: true, manual: false };
let currentWeeks = [{ dates: [], lessons: sampleLessons, month: null, sheetTitle: "" }];
let selectedWeekIndex = 0;
function startOfWeek(date = new Date()) {
const copy = new Date(date);
const day = copy.getDay();
const distance = day === 0 ? -6 : 1 - day;
copy.setDate(copy.getDate() + distance);
copy.setHours(0, 0, 0, 0);
return copy;
}
function todayWeekdayIndex() {
const day = new Date().getDay();
return dayIndexes.indexOf(day);
}
function cell(className, text) {
const node = document.createElement("div");
node.className = `grid-cell ${className}`;
if (text) node.textContent = text;
return node;
}
function abbreviateLessonName(name) {
const normalized = String(name || "").trim();
return lessonAbbreviations.get(normalized) || normalized;
}
function openLessonDetail(lesson) {
elements.detailTitle.textContent = lesson.name;
elements.detailContent.textContent = lesson.content || "内容はありません";
elements.detailDialog.showModal();
}
function closeLessonDetail() {
elements.detailDialog.close();
}
function renderTimetable(lessons, highlightToday = true) {
const today = highlightToday ? todayWeekdayIndex() : -1;
const fragment = document.createDocumentFragment();
fragment.append(cell("corner-cell", ""));
weekdays.forEach((weekday, column) => {
const header = cell("weekday-cell", weekday);
if (column === today) header.classList.add("today-column");
fragment.append(header);
});
lessons.forEach((row, rowIndex) => {
fragment.append(cell("period-cell", `${rowIndex + 1}限`));
weekdays.forEach((_, column) => {
const lesson = row[column] || { name: "", content: "" };
const lessonCell = cell("lesson-cell", "");
if (column === today) lessonCell.classList.add("today-column");
if (lesson.name) {
const trigger = document.createElement("button");
trigger.className = "lesson-trigger";
trigger.type = "button";
trigger.setAttribute(
"aria-label",
`${lesson.name}${lesson.content ? `、${lesson.content}` : ""}`,
);
trigger.addEventListener("click", () => openLessonDetail(lesson));
const name = document.createElement("span");
name.className = "lesson-name";
name.textContent = abbreviateLessonName(lesson.name);
trigger.append(name);
lessonCell.append(trigger);
}
fragment.append(lessonCell);
});
});
elements.timetable.replaceChildren(fragment);
}
function parseDayNumber(value) {
const match = String(value ?? "").match(/\d+/);
return match ? Number(match[0]) : null;
}
function dateFromIso(value) {
if (!value) return null;
const date = new Date(`${value}T00:00:00`);
return Number.isNaN(date.getTime()) ? null : date;
}
function findCurrentWeekIndex(weeks) {
const today = new Date();
today.setHours(0, 0, 0, 0);
let nearestIndex = 0;
let nearestDistance = Number.POSITIVE_INFINITY;
weeks.forEach((week, index) => {
const dates = week.dates.map(dateFromIso).filter(Boolean);
if (!dates.length) return;
const start = dates[0];
const end = dates.at(-1);
if (today >= start && today <= end) {
nearestIndex = index;
nearestDistance = 0;
return;
}
const distance = Math.min(
Math.abs(today.getTime() - start.getTime()),
Math.abs(today.getTime() - end.getTime()),
);
if (distance < nearestDistance) {
nearestIndex = index;
nearestDistance = distance;
}
});
return nearestIndex;
}
function formatDateRange(week) {
const dates = week.dates.map(dateFromIso).filter(Boolean);
if (!dates.length) return "週間時間割";
const start = dates[0];
const end = dates.at(-1);
const startLabel = `${start.getMonth() + 1}月${start.getDate()}日`;
const endLabel =
start.getMonth() === end.getMonth()
? `${end.getDate()}日`
: `${end.getMonth() + 1}月${end.getDate()}日`;
return `${startLabel}〜${endLabel}`;
}
function weekDisplayLabel(week) {
return week.sheetTitle ? `${week.sheetTitle}の時間割` : "時間割";
}
function selectWeek(index) {
if (!currentWeeks.length) return;
selectedWeekIndex = Math.max(0, Math.min(index, currentWeeks.length - 1));
const week = currentWeeks[selectedWeekIndex];
const currentWeekIndex = findCurrentWeekIndex(currentWeeks);
renderTimetable(week.lessons, selectedWeekIndex === currentWeekIndex);
const dateRange = formatDateRange(week);
elements.appTitle.textContent = dateRange;
document.title = `${dateRange} | 時間割`;
elements.weekLabel.textContent = weekDisplayLabel(week);
elements.weekRange.textContent = "";
elements.previousWeek.disabled = selectedWeekIndex === 0;
elements.nextWeek.disabled = selectedWeekIndex === currentWeeks.length - 1;
}
function findSheetStartDate(values, sheetMonth, year) {
const firstDateValues = values[firstDateRow - 1] || [];
const days = weekdays.map((_, dayIndex) =>
parseDayNumber(firstDateValues[firstWeekdayColumn + dayIndex]),
);
const firstValidIndex = days.findIndex(Number.isFinite);
if (firstValidIndex === -1) {
const firstOfMonth = new Date(year, sheetMonth - 1, 1);
const weekday = firstOfMonth.getDay();
const daysFromMonday = weekday === 0 ? 6 : weekday - 1;
firstOfMonth.setDate(firstOfMonth.getDate() - daysFromMonday);
return firstOfMonth;
}
const day = days[firstValidIndex];
const monthOffset = day > 20 ? -1 : 0;
const date = new Date(year, sheetMonth - 1 + monthOffset, day);
date.setDate(date.getDate() - firstValidIndex);
return date;
}
function buildWeekDates(sheetStartDate, weekIndex) {
return weekdays.map((_, dayIndex) => {
const date = new Date(sheetStartDate);
date.setDate(date.getDate() + weekIndex * 7 + dayIndex);
const monthText = String(date.getMonth() + 1).padStart(2, "0");
const dayText = String(date.getDate()).padStart(2, "0");
return `${date.getFullYear()}-${monthText}-${dayText}`;
});
}
function shouldHideLesson(dateIso, periodIndex) {
const date = dateFromIso(dateIso);
return (
date?.getMonth() === 5 &&
date.getDate() === 17 &&
(periodIndex === 2 || periodIndex === 3)
);
}
function parseSheet(values, sheetMonth, sheetTitle, year) {
if (!Array.isArray(values) || !values.length) return [];
const sheetStartDate = findSheetStartDate(values, sheetMonth, year);
return Array.from({ length: weekCount }, (_, weekIndex) => {
const rowOffset = rowsPerWeek * weekIndex;
const dates = buildWeekDates(sheetStartDate, weekIndex);
const lessons = firstWeekRows.map(
({ name: nameRowNumber, content: contentRowNumber }, periodIndex) => {
const nameRow = values[nameRowNumber + rowOffset - 1] || [];
const contentRow = values[contentRowNumber + rowOffset - 1] || [];
return weekdays.map((_, dayIndex) => {
const columnIndex = firstWeekdayColumn + dayIndex;
if (shouldHideLesson(dates[dayIndex], periodIndex)) {
return { name: "", content: "" };
}
const name = String(nameRow[columnIndex] || "").trim();
return {
name,
content: name ? String(contentRow[columnIndex] || "").trim() : "",
};
});
},
);
return { dates, lessons, month: sheetMonth, sheetTitle };
}).filter((week) => week.dates.some(Boolean));
}
function setLoading(loading) {
elements.loadingOverlay.hidden = !loading;
elements.refreshButton.disabled = loading;
}
function showNotice(message = "") {
elements.notice.textContent = message;
elements.notice.hidden = !message;
}
function formatUpdatedAt(timestamp) {
if (!timestamp) return "未取得";
const date = new Date(timestamp);
const today = new Date();
const isToday = date.toDateString() === today.toDateString();
const time = new Intl.DateTimeFormat("ja-JP", {
hour: "2-digit",
minute: "2-digit",
}).format(date);
return `${isToday ? "今日" : `${date.getMonth() + 1}月${date.getDate()}日`} ${time}`;
}
function readCache() {
try {
return JSON.parse(localStorage.getItem(cacheKey) || "null");
} catch {
return null;
}
}
function cacheIsFromToday(cache) {
if (!cache?.updatedAt) return false;
return new Date(cache.updatedAt).toDateString() === new Date().toDateString();
}
function hasPreviousAuthorization() {
return Boolean(
localStorage.getItem(authorizationKey) || localStorage.getItem(cacheKey),
);
}
function applyCache(cache) {
if (!Array.isArray(cache?.weeks) || !cache.weeks.length) return false;
currentWeeks = cache.weeks;
selectWeek(findCurrentWeekIndex(currentWeeks));
elements.lastUpdated.textContent = formatUpdatedAt(cache.updatedAt);
elements.dataSource.textContent = "Google スプレッドシートから取得しています";
return true;
}
async function fetchProfile() {
if (!accessToken) return;
const response = await fetch("https://www.googleapis.com/oauth2/v3/userinfo", {
headers: { Authorization: `Bearer ${accessToken}` },
});
if (!response.ok) return;
const profile = await response.json();
if (profile.picture) {
elements.accountImage.src = profile.picture;
elements.accountImage.hidden = false;
elements.accountInitial.hidden = true;
} else if (profile.name) {
elements.accountInitial.textContent = profile.name.slice(0, 1);
}
elements.accountButton.title = profile.name || "Googleアカウント";
}
async function googleApiFetch(url) {
const response = await fetch(url, {
headers: { Authorization: `Bearer ${accessToken}` },
});
if (response.status === 401) {
accessToken = "";
sessionStorage.removeItem("googleAccessToken");
sessionStorage.removeItem(tokenExpiryKey);
requestAccessToken({ interactive: false });
return null;
}
if (!response.ok) {
const error = await response.json().catch(() => ({}));
throw new Error(error.error?.message || "時間割を取得できませんでした。");
}
return response.json();
}
async function fetchMonthSheetTitles() {
const endpoint = new URL(
`https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(config.spreadsheetId)}`,
);
endpoint.searchParams.set("fields", "sheets.properties.title");
const data = await googleApiFetch(endpoint);
if (!data) return null;
return data.sheets
.map((sheet) => sheet.properties?.title || "")
.map((title) => {
const match = title.match(/^(\d{1,2})月$/);
return match ? { title, month: Number(match[1]) } : null;
})
.filter(
(sheet) =>
sheet &&
sheet.month >= 1 &&
sheet.month <= 12 &&
!excludedMonths.has(sheet.month),
)
.sort((a, b) => a.month - b.month);
}
async function fetchAllMonthValues(monthSheets) {
const endpoint = new URL(
`https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(config.spreadsheetId)}/values:batchGet`,
);
endpoint.searchParams.set("majorDimension", "ROWS");
monthSheets.forEach(({ title }) => {
endpoint.searchParams.append("ranges", `'${title}'!${config.cellRange}`);
});
const data = await googleApiFetch(endpoint);
return data?.valueRanges || null;
}
function mergeMonthWeeks(monthSheets, valueRanges) {
const year = new Date().getFullYear();
const seen = new Set();
const weeks = [];
monthSheets.forEach((sheet, index) => {
const parsedWeeks = parseSheet(
valueRanges[index]?.values || [],
sheet.month,
sheet.title,
year,
);
parsedWeeks.forEach((week) => {
const key = week.dates.filter(Boolean).join(",");
if (key && !seen.has(key)) {
seen.add(key);
weeks.push(week);
}
});
});
return weeks.sort((a, b) => {
const firstA = a.dates.find(Boolean) || "";
const firstB = b.dates.find(Boolean) || "";
return firstA.localeCompare(firstB);
});
}
async function fetchTimetable({ manual = false } = {}) {
if (!config.spreadsheetId || !config.googleClientId) {
showNotice("現在はサンプル表示です。config.js にGoogle OAuthとスプレッドシートを設定すると実データへ切り替わります。");
return;
}
if (!accessToken) {
requestAccessToken({
interactive: !hasPreviousAuthorization(),
manual,
});
return;
}
setLoading(true);
showNotice("");
try {
const monthSheets = await fetchMonthSheetTitles();
if (!monthSheets) return;
if (!monthSheets.length) {
throw new Error("「○月」という名前の時間割シートが見つかりません。");
}
const valueRanges = await fetchAllMonthValues(monthSheets);
if (!valueRanges) return;
const weeks = mergeMonthWeeks(monthSheets, valueRanges);
if (!weeks.length) {
throw new Error("月別シートに時間割データがありません。");
}
const cache = { weeks, updatedAt: Date.now() };
localStorage.setItem(cacheKey, JSON.stringify(cache));
applyCache(cache);
elements.signInButton.hidden = true;
if (manual) showNotice("最新の時間割に更新しました。");
} catch (error) {
showNotice(error.message || "時間割を取得できませんでした。");
} finally {
setLoading(false);
}
}
function requestAccessToken({ interactive = true, manual = false } = {}) {
if (!config.googleClientId) {
showNotice("Googleログインを利用するには config.js の googleClientId を設定してください。");
return;
}
if (!window.google?.accounts?.oauth2) {
if (interactive) {
showNotice("Googleログインの準備中です。数秒後にもう一度お試しください。");
}
return;
}
pendingAuthorization = { interactive, manual };
tokenClient ||= window.google.accounts.oauth2.initTokenClient({
client_id: config.googleClientId,
scope: `${tokenScope} openid profile email`,
callback: async (response) => {
const request = pendingAuthorization;
automaticAuthorizationPending = false;
if (response.error) {
elements.signInButton.hidden = false;
if (request.interactive) {
showNotice("Googleログインを完了できませんでした。");
}
return;
}
accessToken = response.access_token;
sessionStorage.setItem("googleAccessToken", accessToken);
sessionStorage.setItem(
tokenExpiryKey,
String(Date.now() + Math.max(0, Number(response.expires_in) - 60) * 1000),
);
localStorage.setItem(authorizationKey, "true");
elements.signInButton.hidden = true;
fetchProfile().catch(() => {});
await fetchTimetable({ manual: request.manual });
},
});
tokenClient.requestAccessToken({ prompt: interactive ? "consent" : "" });
}
function restoreGoogleAuthorization() {
if (
accessToken ||
automaticAuthorizationPending ||
!hasPreviousAuthorization()
) {
return;
}
if (!window.google?.accounts?.oauth2) {
authorizationRestoreAttempts += 1;
if (authorizationRestoreAttempts >= 40) {
elements.signInButton.hidden = false;
return;
}
window.setTimeout(restoreGoogleAuthorization, 250);
return;
}
authorizationRestoreAttempts = 0;
automaticAuthorizationPending = true;
requestAccessToken({ interactive: false });
}
function initialize() {
elements.className.textContent = config.className;
selectWeek(0);
const cache = readCache();
if (cache) applyCache(cache);
const configured = Boolean(config.googleClientId && config.spreadsheetId);
elements.signInButton.hidden =
Boolean(accessToken) || hasPreviousAuthorization() || !configured;
if (!configured) {
showNotice("現在はサンプル表示です。config.js にGoogle OAuthとスプレッドシートを設定すると実データへ切り替わります。");
} else if (accessToken) {
fetchProfile();
if (!cacheIsFromToday(cache)) fetchTimetable();
} else {
restoreGoogleAuthorization();
}
}
elements.refreshButton.addEventListener("click", () => fetchTimetable({ manual: true }));
elements.signInButton.addEventListener("click", () =>
requestAccessToken({ interactive: true }),
);
elements.accountButton.addEventListener("click", () =>
requestAccessToken({ interactive: true }),
);
elements.previousWeek.addEventListener("click", () => selectWeek(selectedWeekIndex - 1));
elements.nextWeek.addEventListener("click", () => selectWeek(selectedWeekIndex + 1));
elements.detailClose.addEventListener("click", closeLessonDetail);
elements.detailDialog.addEventListener("click", (event) => {
if (event.target === elements.detailDialog) closeLessonDetail();
});
initialize();
