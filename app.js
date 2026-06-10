const config = {
  googleClientId: "",
  spreadsheetId: "",
  range: "時間割!A1:K7",
  title: "週間時間割",
  className: "1年A組",
  ...(window.TIMETABLE_CONFIG || {}),
};

const weekdays = ["月", "火", "水", "木", "金"];
const dayIndexes = [1, 2, 3, 4, 5];
const cacheKey = `timetable:${config.spreadsheetId}:${config.range}`;
const tokenScope = "https://www.googleapis.com/auth/spreadsheets.readonly";

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
  [
    { name: "体育", content: "バレーボール" },
    { name: "社会", content: "地理の学習" },
    { name: "体育", content: "バレーボール" },
    { name: "数学", content: "一次方程式" },
    { name: "英語", content: "Unit 3" },
  ],
  [
    { name: "LHR", content: "学級活動" },
    { name: "総合", content: "探究活動" },
    { name: "社会", content: "地理の学習" },
    { name: "体育", content: "バレーボール" },
    { name: "国語", content: "漢字テスト" },
  ],
];

const elements = {
  accountButton: document.querySelector("#account-button"),
  accountImage: document.querySelector("#account-image"),
  accountInitial: document.querySelector("#account-initial"),
  appTitle: document.querySelector("#app-title"),
  className: document.querySelector("#class-name"),
  dataSource: document.querySelector("#data-source"),
  lastUpdated: document.querySelector("#last-updated"),
  loadingOverlay: document.querySelector("#loading-overlay"),
  notice: document.querySelector("#notice"),
  refreshButton: document.querySelector("#refresh-button"),
  signInButton: document.querySelector("#sign-in-button"),
  timetable: document.querySelector("#timetable"),
  weekRange: document.querySelector("#week-range"),
};

let accessToken = sessionStorage.getItem("googleAccessToken") || "";
let tokenClient;
let currentLessons = sampleLessons;

function startOfWeek(date = new Date()) {
  const copy = new Date(date);
  const day = copy.getDay();
  const distance = day === 0 ? -6 : 1 - day;
  copy.setDate(copy.getDate() + distance);
  copy.setHours(0, 0, 0, 0);
  return copy;
}

function formatWeekRange() {
  const monday = startOfWeek();
  const friday = new Date(monday);
  friday.setDate(monday.getDate() + 4);
  const sameMonth = monday.getMonth() === friday.getMonth();
  const first = `${monday.getMonth() + 1}月${monday.getDate()}日`;
  const second = sameMonth
    ? `${friday.getDate()}日`
    : `${friday.getMonth() + 1}月${friday.getDate()}日`;
  return `${first}–${second}`;
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

function renderTimetable(lessons) {
  const today = todayWeekdayIndex();
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

      if (!lesson.name && !lesson.content) {
        const empty = document.createElement("span");
        empty.className = "empty-lesson";
        empty.textContent = "—";
        lessonCell.append(empty);
      } else {
        const name = document.createElement("span");
        name.className = "lesson-name";
        name.textContent = lesson.name;
        const content = document.createElement("span");
        content.className = "lesson-content";
        content.textContent = lesson.content;
        lessonCell.append(name, content);
      }
      fragment.append(lessonCell);
    });
  });

  elements.timetable.replaceChildren(fragment);
}

function parseSheet(values) {
  if (!Array.isArray(values) || values.length < 2) {
    throw new Error("シートに時間割データがありません。");
  }

  const rows = values.slice(1, 7);
  const pairedColumns = (values[0]?.length || 0) >= 11;

  return Array.from({ length: 6 }, (_, rowIndex) => {
    const row = rows[rowIndex] || [];
    return weekdays.map((_, dayIndex) => {
      if (pairedColumns) {
        return {
          name: String(row[1 + dayIndex * 2] || "").trim(),
          content: String(row[2 + dayIndex * 2] || "").trim(),
        };
      }

      const raw = String(row[1 + dayIndex] || "").trim();
      const [name = "", ...content] = raw.split(/\r?\n/);
      return { name: name.trim(), content: content.join(" ").trim() };
    });
  });
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

function applyCache(cache) {
  if (!cache?.lessons) return false;
  currentLessons = cache.lessons;
  renderTimetable(currentLessons);
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

async function fetchTimetable({ manual = false } = {}) {
  if (!config.spreadsheetId || !config.googleClientId) {
    showNotice("現在はサンプル表示です。config.js にGoogle OAuthとスプレッドシートを設定すると実データへ切り替わります。");
    return;
  }

  if (!accessToken) {
    requestAccessToken();
    return;
  }

  setLoading(true);
  showNotice("");
  try {
    const endpoint = new URL(
      `https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(config.spreadsheetId)}/values/${encodeURIComponent(config.range)}`,
    );
    endpoint.searchParams.set("majorDimension", "ROWS");

    const response = await fetch(endpoint, {
      headers: { Authorization: `Bearer ${accessToken}` },
    });

    if (response.status === 401) {
      accessToken = "";
      sessionStorage.removeItem("googleAccessToken");
      requestAccessToken();
      return;
    }

    if (!response.ok) {
      const error = await response.json().catch(() => ({}));
      throw new Error(error.error?.message || "時間割を取得できませんでした。");
    }

    const data = await response.json();
    const lessons = parseSheet(data.values);
    const cache = { lessons, updatedAt: Date.now() };
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

function requestAccessToken() {
  if (!config.googleClientId) {
    showNotice("Googleログインを利用するには config.js の googleClientId を設定してください。");
    return;
  }

  if (!window.google?.accounts?.oauth2) {
    showNotice("Googleログインの準備中です。数秒後にもう一度お試しください。");
    return;
  }

  tokenClient ||= window.google.accounts.oauth2.initTokenClient({
    client_id: config.googleClientId,
    scope: `${tokenScope} openid profile email`,
    callback: async (response) => {
      if (response.error) {
        showNotice("Googleログインを完了できませんでした。");
        return;
      }
      accessToken = response.access_token;
      sessionStorage.setItem("googleAccessToken", accessToken);
      elements.signInButton.hidden = true;
      fetchProfile().catch(() => {});
      await fetchTimetable();
    },
  });
  tokenClient.requestAccessToken({ prompt: accessToken ? "" : "consent" });
}

function initialize() {
  document.title = config.title;
  elements.appTitle.textContent = config.title;
  elements.className.textContent = config.className;
  elements.weekRange.textContent = formatWeekRange();
  renderTimetable(sampleLessons);

  const cache = readCache();
  if (cache) applyCache(cache);

  const configured = Boolean(config.googleClientId && config.spreadsheetId);
  elements.signInButton.hidden = Boolean(accessToken) || !configured;

  if (!configured) {
    showNotice("現在はサンプル表示です。config.js にGoogle OAuthとスプレッドシートを設定すると実データへ切り替わります。");
  } else if (accessToken) {
    fetchProfile();
    if (!cacheIsFromToday(cache)) fetchTimetable();
  }
}

elements.refreshButton.addEventListener("click", () => fetchTimetable({ manual: true }));
elements.signInButton.addEventListener("click", requestAccessToken);
elements.accountButton.addEventListener("click", requestAccessToken);

initialize();
