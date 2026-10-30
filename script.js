/**
 * Doodh Hisab Core Application Logic
 * Centralized Calculation Engine, IndexedDB Persistence,
 * Google Sheets read-only data, Date-Range Calculator, and Excel Export.
 */

// --- 1. Global State & Default Settings ---
const APP_STATE = {
  entries: [], // Array of daily milk records
  settings: {
    pricePerLitre: 33,
    pricePerPav: 8,
    pavPerLitre: 4,
    openingBalance: 0,
    currency: "₹",
  },
  currentSessionMode: "morning",
  activeTab: "today",
  dashboardFilter: "all",
  chartInstance: null,
  lastCalculatedRange: null, // Cached range calculation result for export
};

const HINDI_WEEKDAYS = [
  { hi: "रविवार", en: "Sunday" },
  { hi: "सोमवार", en: "Monday" },
  { hi: "मंगलवार", en: "Tuesday" },
  { hi: "बुधवार", en: "Wednesday" },
  { hi: "गुरुवार", en: "Thursday" },
  { hi: "शुक्रवार", en: "Friday" },
  { hi: "शनिवार", en: "Saturday" },
];
const SHEET_ENDPOINT_URL =
  "https://script.google.com/macros/s/AKfycbziy3EydOvDhkR4BLdhJVBIWMnRkz3wSrhM08qSjNJl1er0yg7js16X_FH44we5MilpcQ/exec";

// --- 2. IndexedDB Persistent Storage Helper ---
const DB_NAME = "DoodhHisabDB";
const DB_VERSION = 1;
let dbInstance = null;
let milkSummaryWeekOffset = 0;

function openDatabase() {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(DB_NAME, DB_VERSION);

    request.onupgradeneeded = (e) => {
      const db = e.target.result;
      if (!db.objectStoreNames.contains("entries")) {
        db.createObjectStore("entries", { keyPath: "date" });
      }
      if (!db.objectStoreNames.contains("settings")) {
        db.createObjectStore("settings", { keyPath: "key" });
      }
    };

    request.onsuccess = (e) => {
      dbInstance = e.target.result;
      resolve(dbInstance);
    };

    request.onerror = (e) => {
      console.error("IndexedDB open error:", e);
      reject(e);
    };
  });
}

async function replaceEntriesInDB(entries) {
  if (!dbInstance) await openDatabase();
  return new Promise((resolve, reject) => {
    const tx = dbInstance.transaction("entries", "readwrite");
    const store = tx.objectStore("entries");
    store.clear();
    entries.forEach((entry) => store.put(entry));
    tx.oncomplete = () => resolve(true);
    tx.onerror = (e) => reject(e);
  });
}

async function getAllEntriesFromDB() {
  if (!dbInstance) await openDatabase();
  return new Promise((resolve, reject) => {
    const tx = dbInstance.transaction("entries", "readonly");
    const store = tx.objectStore("entries");
    const req = store.getAll();
    req.onsuccess = () => resolve(req.result || []);
    req.onerror = (e) => reject(e);
  });
}

async function saveSettingToDB(key, value) {
  if (!dbInstance) await openDatabase();
  return new Promise((resolve, reject) => {
    const tx = dbInstance.transaction("settings", "readwrite");
    const store = tx.objectStore("settings");
    store.put({ key, value });
    tx.oncomplete = () => resolve(true);
    tx.onerror = (e) => reject(e);
  });
}

async function getSettingsFromDB() {
  if (!dbInstance) await openDatabase();
  return new Promise((resolve) => {
    const tx = dbInstance.transaction("settings", "readonly");
    const store = tx.objectStore("settings");
    const req = store.getAll();
    req.onsuccess = () => {
      const result = req.result || [];
      const loaded = {};
      result.forEach((item) => (loaded[item.key] = item.value));
      resolve(loaded);
    };
    req.onerror = () => resolve({});
  });
}

async function clearAllDataFromDB() {
  if (!dbInstance) await openDatabase();
  return new Promise((resolve, reject) => {
    const tx = dbInstance.transaction(["entries", "settings"], "readwrite");
    tx.objectStore("entries").clear();
    tx.objectStore("settings").clear();
    tx.oncomplete = () => resolve(true);
    tx.onerror = (e) => reject(e);
  });
}

// --- 3. Centralized Calculation Engine ---
function computeDailyRecord(
  morningL,
  morningP,
  eveningL,
  eveningP,
  paid,
  deduction,
  settings = APP_STATE.settings,
) {
  const mL = Math.max(0, parseInt(morningL) || 0);
  const mP = Math.max(0, parseInt(morningP) || 0);
  const eL = Math.max(0, parseInt(eveningL) || 0);
  const eP = Math.max(0, parseInt(eveningP) || 0);
  const pAmt = Math.max(0, parseFloat(paid) || 0);
  const dAmt = Math.max(0, parseFloat(deduction) || 0);
  if (!Number.isInteger(settings.pavPerLitre) || settings.pavPerLitre <= 0) {
    throw new RangeError("पाव प्रति लीटर की संख्या 1 या उससे अधिक पूर्णांक होनी चाहिए।");
  }

  const morningAddedL = Math.floor(mP / settings.pavPerLitre);
  const morningRemP = mP % settings.pavPerLitre;
  const morningTotalL = mL + morningAddedL;
  const morningAmount =
    morningTotalL * settings.pricePerLitre + morningRemP * settings.pricePerPav;
  const morningDecimal = morningTotalL + morningRemP / settings.pavPerLitre;

  const eveningAddedL = Math.floor(eP / settings.pavPerLitre);
  const eveningRemP = eP % settings.pavPerLitre;
  const eveningTotalL = eL + eveningAddedL;
  const eveningAmount =
    eveningTotalL * settings.pricePerLitre + eveningRemP * settings.pricePerPav;
  const eveningDecimal = eveningTotalL + eveningRemP / settings.pavPerLitre;

  const totalPavRaw = mP + eP;
  const totalAddedLitresFromPav = Math.floor(
    totalPavRaw / settings.pavPerLitre,
  );
  const remainingPav = totalPavRaw % settings.pavPerLitre;
  const totalFullLitres = mL + eL + totalAddedLitresFromPav;

  const totalAmount =
    totalFullLitres * settings.pricePerLitre +
    remainingPav * settings.pricePerPav;
  const totalMilkDecimal =
    totalFullLitres + remainingPav / settings.pavPerLitre;

  let totalMilkDisplay = `${totalFullLitres} L`;
  if (remainingPav > 0) {
    totalMilkDisplay += ` ${remainingPav} पाव`;
  }

  return {
    morningLitres: mL,
    morningPav: mP,
    morningTotalLitres: morningTotalL,
    morningRemPav: morningRemP,
    morningAmount,
    morningDecimal,
    eveningLitres: eL,
    eveningPav: eP,
    eveningTotalLitres: eveningTotalL,
    eveningRemPav: eveningRemP,
    eveningAmount,
    eveningDecimal,
    totalFullLitres,
    remainingPav,
    totalMilkDisplay,
    totalMilkDecimal,
    totalAmount,
    paymentReceived: pAmt,
    deduction: dAmt,
  };
}

function recalculateLedger(
  entriesList = APP_STATE.entries,
  settings = APP_STATE.settings,
) {
  const sorted = [...entriesList].sort((a, b) => a.date.localeCompare(b.date));
  let runningBalance = settings.openingBalance || 0;

  const ledger = sorted.map((entry, index) => {
    const computed = computeDailyRecord(
      entry.morningLitres,
      entry.morningPav,
      entry.eveningLitres,
      entry.eveningPav,
      entry.paymentReceived,
      entry.deduction,
      settings,
    );

    const totalAmount = Number.isFinite(entry.sheetTotalAmount)
      ? entry.sheetTotalAmount
      : computed.totalAmount;
    runningBalance =
      runningBalance +
      totalAmount -
      computed.paymentReceived -
      computed.deduction;
    if (Number.isFinite(entry.sheetOutstandingBalance)) {
      runningBalance = entry.sheetOutstandingBalance;
    }

    return {
      ...entry,
      ...computed,
      totalAmount,
      sNo: Number.isFinite(entry.sNo) && entry.sNo > 0 ? entry.sNo : index + 1,
      outstandingBalance: runningBalance,
    };
  });

  return ledger;
}

// --- 4. Custom UI Notifications (Zero Browser Alerts) ---
function showToast(message, type = "success") {
  const toast = document.getElementById("toast");
  const toastMsg = document.getElementById("toastMessage");
  const toastIcon = document.getElementById("toastIcon");

  toastMsg.innerText = message;
  if (type === "error") {
    toastIcon.className =
      "w-6 h-6 flex items-center justify-center rounded-full text-white bg-rose-600";
    toastIcon.innerHTML =
      '<i data-lucide="x" class="w-4 h-4" aria-hidden="true"></i>';
  } else if (type === "warn") {
    toastIcon.className =
      "w-6 h-6 flex items-center justify-center rounded-full text-white bg-amber-600";
    toastIcon.innerHTML =
      '<i data-lucide="triangle-alert" class="w-4 h-4" aria-hidden="true"></i>';
  } else {
    toastIcon.className =
      "w-6 h-6 flex items-center justify-center rounded-full text-white bg-blue-600";
    toastIcon.innerHTML =
      '<i data-lucide="check" class="w-4 h-4" aria-hidden="true"></i>';
  }
  if (window.lucide) window.lucide.createIcons();

  toast.classList.remove("translate-y-[-140%]", "opacity-0");
  toast.classList.add("translate-y-0", "opacity-100");

  setTimeout(() => {
    toast.classList.remove("translate-y-0", "opacity-100");
    toast.classList.add("translate-y-[-140%]", "opacity-0");
  }, 3200);
}

function showConfirmModal(title, description, onConfirm) {
  const modal = document.getElementById("customModal");
  const modalTitle = document.getElementById("modalTitle");
  const modalDesc = document.getElementById("modalDescription");
  const confirmBtn = document.getElementById("modalConfirmBtn");
  const cancelBtn = document.getElementById("modalCancelBtn");

  modalTitle.innerText = title;
  modalDesc.innerText = description;
  modal.classList.remove("hidden");

  const cleanup = () => {
    modal.classList.add("hidden");
    confirmBtn.onclick = null;
    cancelBtn.onclick = null;
  };

  confirmBtn.onclick = () => {
    cleanup();
    if (typeof onConfirm === "function") onConfirm();
  };

  cancelBtn.onclick = () => {
    cleanup();
  };
}

// --- 5. Hindi Date & Weekday Utils ---
function getHindiWeekdayFromDateStr(dateStr) {
  if (!dateStr) return { hi: "अमान्य", en: "Invalid" };
  const [year, month, day] = dateStr.split("-").map(Number);
  const dateObj = new Date(year, month - 1, day);
  const dayIndex = dateObj.getDay();
  return HINDI_WEEKDAYS[dayIndex] || { hi: "", en: "" };
}

function formatIndianCurrency(amount) {
  return new Intl.NumberFormat("en-IN", {
    style: "currency",
    currency: "INR",
    maximumFractionDigits: 2,
  }).format(amount || 0);
}

function getIndiaDateString(date = new Date()) {
  const parts = new Intl.DateTimeFormat("en-CA", {
    timeZone: "Asia/Kolkata",
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
  }).formatToParts(date);
  const values = Object.fromEntries(
    parts.map((part) => [part.type, part.value]),
  );
  return `${values.year}-${values.month}-${values.day}`;
}

function parseSheetDate(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return getIndiaDateString(value);
  }

  const text = String(value ?? "").trim();
  if (!text) return "";

  if (/^\d{4}-\d{2}-\d{2}$/.test(text)) {
    const [year, month, day] = text.split("-").map(Number);
    const date = new Date(Date.UTC(year, month - 1, day));
    return date.getUTCFullYear() === year &&
      date.getUTCMonth() === month - 1 &&
      date.getUTCDate() === day
      ? text
      : "";
  }

  const slashDate = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (slashDate) {
    const [, day, month, year] = slashDate;
    return parseSheetDate(
      `${year}-${month.padStart(2, "0")}-${day.padStart(2, "0")}`,
    );
  }

  const parsed = new Date(text);
  return Number.isNaN(parsed.getTime()) ? "" : getIndiaDateString(parsed);
}

function getMonthPrefix(date = new Date()) {
  return getIndiaDateString(date).slice(0, 7);
}

function getPreviousMonthPrefix(date = new Date()) {
  const parts = getIndiaDateString(date).split("-").map(Number);
  const previousMonth = new Date(Date.UTC(parts[0], parts[1] - 2, 1));
  return `${previousMonth.getUTCFullYear()}-${String(previousMonth.getUTCMonth() + 1).padStart(2, "0")}`;
}

function formatDateDisplay(dateStr) {
  if (!dateStr) return "";
  const [year, month, day] = dateStr.split("-");
  const months = [
    "Jan",
    "Feb",
    "Mar",
    "Apr",
    "May",
    "Jun",
    "Jul",
    "Aug",
    "Sep",
    "Oct",
    "Nov",
    "Dec",
  ];
  return `${parseInt(day)} ${months[parseInt(month) - 1]} ${year}`;
}

// --- 6. Tab Management ---
function switchTab(tabId) {
  APP_STATE.activeTab = tabId;

  document.querySelectorAll(".tab-btn").forEach((btn) => {
    const isActive = btn.getAttribute("data-tab") === tabId;
    if (btn.getAttribute("data-tab") === tabId) {
      btn.classList.add("active", "bg-blue-600", "text-white", "shadow-sm");
      btn.classList.remove(
        "text-slate-600",
        "dark:text-slate-300",
        "hover:bg-slate-100",
        "dark:hover:bg-slate-700",
      );
    } else {
      btn.classList.remove("active", "bg-blue-600", "text-white", "shadow-sm");
      btn.classList.add(
        "text-slate-600",
        "dark:text-slate-300",
        "hover:bg-slate-100",
        "dark:hover:bg-slate-700",
      );
    }
    if (isActive)
      btn.scrollIntoView({
        behavior: "smooth",
        block: "nearest",
        inline: "center",
      });
  });

  document.querySelectorAll(".tab-content").forEach((section) => {
    if (section.id === `tab-${tabId}`) {
      section.classList.remove("hidden");
    } else {
      section.classList.add("hidden");
    }
  });

  if (tabId === "today") {
    renderTodayOverview();
  } else if (tabId === "dashboard") {
    renderDashboard();
  } else if (tabId === "records") {
    renderMasterRecords();
  } else if (tabId === "calculator") {
    runDateRangeCalculator();
  } else if (tabId === "payments") {
    renderPaymentsLedger();
  } else if (tabId === "reports") {
    renderReportsView();
  }
}

function addDaysToDateString(dateStr, days) {
  const [year, month, day] = dateStr.split("-").map(Number);
  return new Date(Date.UTC(year, month - 1, day + days))
    .toISOString()
    .slice(0, 10);
}

function getDateRangeLabel(startDate, endDate) {
  return `${formatDateDisplay(startDate)} – ${formatDateDisplay(endDate)}`;
}

function getMilkCountForPeriod(
  startDate,
  endDate,
  startSessions = ["evening"],
  endSessions = ["morning"],
) {
  let totalPav = 0;
  let totalAmount = 0;
  const pavPerLitre = APP_STATE.settings.pavPerLitre;
  const entriesByDate = new Map(
    APP_STATE.entries.map((entry) => [entry.date, entry]),
  );
  const sessions = [];
  const sessionAmount = (entry, session) => {
    const amountKey = `${session}Amount`;
    if (Number.isFinite(entry[amountKey])) return entry[amountKey];
    const litres = entry[`${session}Litres`] || 0;
    const pav = entry[`${session}Pav`] || 0;
    const convertedLitres = Math.floor(pav / pavPerLitre);
    const remainingPav = pav % pavPerLitre;
    return (
      (litres + convertedLitres) * APP_STATE.settings.pricePerLitre +
      remainingPav * APP_STATE.settings.pricePerPav
    );
  };

  for (let date = startDate; date <= endDate; date = addDaysToDateString(date, 1)) {
    const entry = entriesByDate.get(date);
    const sessionsForDate =
      date === startDate
        ? startSessions
        : date === endDate
          ? endSessions
          : ["morning", "evening"];

    sessionsForDate.forEach((session) => {
      const litres = entry?.[`${session}Litres`] || 0;
      const pav = entry?.[`${session}Pav`] || 0;
      const sessionPav = litres * pavPerLitre + pav;
      const amount = sessionAmount(entry || {}, session);
      totalPav += sessionPav;
      totalAmount += amount;
      sessions.push({
        date,
        weekday: getHindiWeekdayFromDateStr(date).hi,
        session: session === "morning" ? "सुबह" : "शाम",
        milk: `${litres} L ${pav} पाव`,
        amount,
        recorded: Boolean(entry),
      });
    });
  }

  const litres = Math.floor(totalPav / pavPerLitre);
  const remainingPav = totalPav % pavPerLitre;
  return {
    display: `${litres} L${remainingPav > 0 ? ` ${remainingPav} पाव` : ""}`,
    decimal: (totalPav / pavPerLitre).toFixed(2),
    amount: totalAmount,
    sessions,
  };
}

function renderMilkPeriodBreakdown(containerId, period) {
  const container = document.getElementById(containerId);
  container.replaceChildren();

  if (period.sessions.length === 0) {
    const emptyMessage = document.createElement("p");
    emptyMessage.className = "p-4 text-sm text-slate-500 dark:text-slate-400";
    emptyMessage.innerText = "इस अवधि के लिए कोई रिकॉर्ड उपलब्ध नहीं है।";
    container.appendChild(emptyMessage);
    return;
  }

  const table = document.createElement("table");
  table.className = "w-full min-w-[34rem] text-left text-xs sm:text-sm";
  table.innerHTML = `
    <thead class="bg-slate-100 text-[11px] tracking-wide text-slate-600 dark:bg-slate-800 dark:text-slate-300">
      <tr>
        <th class="px-3 py-2.5 font-bold">दिनांक</th>
        <th class="px-3 py-2.5 font-bold">दिन</th>
        <th class="px-3 py-2.5 font-bold">समय</th>
        <th class="px-3 py-2.5 font-bold">मात्रा</th>
        <th class="px-3 py-2.5 font-bold">स्थिति</th>
        <th class="px-3 py-2.5 text-right font-bold">रकम</th>
      </tr>
    </thead>
    <tbody class="divide-y divide-slate-100 dark:divide-slate-700"></tbody>
  `;
  const tbody = table.querySelector("tbody");

  period.sessions.forEach((item) => {
    const row = document.createElement("tr");
    row.className = "text-slate-700 transition-colors hover:bg-slate-50 dark:text-slate-200 dark:hover:bg-slate-700/60";
    const values = [
      formatDateDisplay(item.date),
      item.weekday,
      item.session,
      item.milk.replace(/ L/g, " लीटर"),
      item.recorded ? "दर्ज" : "दर्ज नहीं",
      formatIndianCurrency(item.amount),
    ];

    values.forEach((value, index) => {
      const cell = document.createElement("td");
      cell.className = `px-3 py-2.5 whitespace-nowrap ${index === 5 ? "text-right font-bold tabular-nums" : ""}`;
      cell.innerText = value;
      row.appendChild(cell);
    });
    tbody.appendChild(row);
  });

  container.appendChild(table);
}

function renderTodayOverview() {
  const todayDate = getIndiaDateString();
  const weekday = getHindiWeekdayFromDateStr(todayDate);
  const todayRecord = APP_STATE.entries.find((entry) => entry.date === todayDate);
  const dateObj = new Date(`${todayDate}T00:00:00Z`);
  const daysSinceWednesday = (dateObj.getUTCDay() - 3 + 7) % 7;
  const currentWednesdayDate = addDaysToDateString(
    todayDate,
    -daysSinceWednesday,
  );
  const wednesdayDate = addDaysToDateString(
    currentWednesdayDate,
    milkSummaryWeekOffset * 7,
  );
  const sundayDate = addDaysToDateString(wednesdayDate, 4);
  const nextWednesdayDate = addDaysToDateString(wednesdayDate, 7);
  const firstPeriod = getMilkCountForPeriod(
    wednesdayDate,
    sundayDate,
    ["evening"],
    ["morning"],
  );
  const secondPeriod = getMilkCountForPeriod(sundayDate, nextWednesdayDate);
  const pavPerLitre = APP_STATE.settings.pavPerLitre;
  const setText = (id, value) => {
    document.getElementById(id).innerText = value;
  };
  setText(
    "milkSummaryWeekLabel",
    `${milkSummaryWeekOffset === 0 ? "वर्तमान सप्ताह" : "पिछला सप्ताह"} · ${getDateRangeLabel(wednesdayDate, nextWednesdayDate)}`,
  );
  document.getElementById("milkSummaryNextWeek").disabled =
    milkSummaryWeekOffset === 0;

  setText("todayOverviewDate", formatDateDisplay(todayDate));
  setText("todayOverviewWeekday", `${weekday.hi} (${weekday.en})`);
  setText(
    "todayOverviewStatus",
    todayRecord ? "" : "आज का रिकॉर्ड अभी उपलब्ध नहीं है।",
  );

  const morningLitres = todayRecord?.morningLitres || 0;
  const morningPav = todayRecord?.morningPav || 0;
  const eveningLitres = todayRecord?.eveningLitres || 0;
  const eveningPav = todayRecord?.eveningPav || 0;
  setText(
    "todayOverviewMorning",
    `${morningLitres} L ${morningPav} पाव · ${(morningLitres + morningPav / pavPerLitre).toFixed(2)} L`,
  );
  setText(
    "todayOverviewMorningAmount",
    `रकम: ${formatIndianCurrency(todayRecord?.morningAmount || 0)}`,
  );
  setText(
    "todayOverviewEvening",
    `${eveningLitres} L ${eveningPav} पाव · ${(eveningLitres + eveningPav / pavPerLitre).toFixed(2)} L`,
  );
  setText(
    "todayOverviewEveningAmount",
    `रकम: ${formatIndianCurrency(todayRecord?.eveningAmount || 0)}`,
  );
  setText("todayOverviewTotalMilk", todayRecord?.totalMilkDisplay || "0 L");
  setText(
    "todayOverviewTotalAmount",
    formatIndianCurrency(todayRecord?.totalAmount || 0),
  );
  setText(
    "todayOverviewPayment",
    formatIndianCurrency(todayRecord?.paymentReceived || 0),
  );
  setText(
    "todayOverviewDeduction",
    formatIndianCurrency(todayRecord?.deduction || 0),
  );
  setText(
    "todayOverviewBalance",
    formatIndianCurrency(
      todayRecord?.outstandingBalance ?? APP_STATE.settings.openingBalance ?? 0,
    ),
  );

  setText(
    "milkCountFirstDates",
    getDateRangeLabel(wednesdayDate, sundayDate),
  );
  setText("milkCountFirstTotal", firstPeriod.display);
  setText("milkCountFirstDecimal", `${firstPeriod.decimal} लीटर`);
  setText("milkCountFirstAmount", formatIndianCurrency(firstPeriod.amount));
  renderMilkPeriodBreakdown("milkCountFirstBreakdown", firstPeriod);
  setText(
    "milkCountSecondDates",
    getDateRangeLabel(sundayDate, nextWednesdayDate),
  );
  setText("milkCountSecondTotal", secondPeriod.display);
  setText("milkCountSecondDecimal", `${secondPeriod.decimal} लीटर`);
  setText("milkCountSecondAmount", formatIndianCurrency(secondPeriod.amount));
  renderMilkPeriodBreakdown("milkCountSecondBreakdown", secondPeriod);
}

// --- 9. Dashboard Rendering & KPI Calculation ---
function renderDashboard() {
  let filteredEntries = [...APP_STATE.entries];

  const filter = APP_STATE.dashboardFilter;
  const customRangeEl = document.getElementById("dashCustomRange");
  const currentMonth = getMonthPrefix();
  const previousMonth = getPreviousMonthPrefix();

  if (filter === "this_month") {
    customRangeEl.classList.add("hidden");
    filteredEntries = filteredEntries.filter((e) =>
      e.date.startsWith(currentMonth),
    );
  } else if (filter === "last_month") {
    customRangeEl.classList.add("hidden");
    filteredEntries = filteredEntries.filter((e) =>
      e.date.startsWith(previousMonth),
    );
  } else if (filter === "custom") {
    customRangeEl.classList.remove("hidden");
    const start = document.getElementById("dashStartDate").value;
    const end = document.getElementById("dashEndDate").value;
    if (start && end) {
      filteredEntries = filteredEntries.filter(
        (e) => e.date >= start && e.date <= end,
      );
    }
  } else {
    customRangeEl.classList.add("hidden");
  }

  let totalL = 0;
  let totalP = 0;
  let totalAmt = 0;
  let totalPaid = 0;
  let totalDed = 0;
  let morningTotalMilkDec = 0;
  let eveningTotalMilkDec = 0;

  filteredEntries.forEach((item) => {
    totalL += (item.morningLitres || 0) + (item.eveningLitres || 0);
    totalP += (item.morningPav || 0) + (item.eveningPav || 0);
    totalAmt += item.totalAmount || 0;
    totalPaid += item.paymentReceived || 0;
    totalDed += item.deduction || 0;

    morningTotalMilkDec += item.morningDecimal || 0;
    eveningTotalMilkDec += item.eveningDecimal || 0;
  });

  const addedL = Math.floor(totalP / APP_STATE.settings.pavPerLitre);
  const remP = totalP % APP_STATE.settings.pavPerLitre;
  const grandTotalL = totalL + addedL;

  const latestOverallBalance =
    APP_STATE.entries.length > 0
      ? APP_STATE.entries[APP_STATE.entries.length - 1].outstandingBalance
      : APP_STATE.settings.openingBalance || 0;

  document.getElementById("kpiTotalMilk").innerText =
    `${grandTotalL} L ${remP > 0 ? remP + " पाव" : ""}`;
  document.getElementById("kpiTotalMilkPav").innerText =
    `कुल पाव: ${totalP} (${(grandTotalL + remP / 4).toFixed(2)} L)`;
  document.getElementById("kpiTotalAmount").innerText =
    formatIndianCurrency(totalAmt);
  document.getElementById("kpiTotalReceived").innerText =
    formatIndianCurrency(totalPaid);
  document.getElementById("kpiTotalDeductions").innerText =
    `कटौती: ${formatIndianCurrency(totalDed)}`;
  document.getElementById("kpiOutstanding").innerText =
    formatIndianCurrency(latestOverallBalance);

  document.getElementById("quickOutstanding").innerText =
    formatIndianCurrency(latestOverallBalance);
  document.getElementById("quickRate").innerText =
    `₹${APP_STATE.settings.pricePerLitre}/L`;

  const todayDateStr = getIndiaDateString();
  const todayRecord = APP_STATE.entries.find((e) => e.date === todayDateStr);
  document.getElementById("dashTodayDate").innerText =
    formatDateDisplay(todayDateStr);
  if (todayRecord) {
    document.getElementById("dashTodayMilk").innerText =
      todayRecord.totalMilkDisplay;
    document.getElementById("dashTodayPav").innerText =
      `(${todayRecord.remainingPav} पाव शेष)`;
    document.getElementById("dashTodayAmount").innerText = formatIndianCurrency(
      todayRecord.totalAmount,
    );
  } else {
    document.getElementById("dashTodayMilk").innerText = "0.0 L";
    document.getElementById("dashTodayPav").innerText = "(0 पाव)";
    document.getElementById("dashTodayAmount").innerText = "₹0.00";
  }

  const currentMonthEntries = APP_STATE.entries.filter((e) =>
    e.date.startsWith(getMonthPrefix()),
  );
  let mMonthL = 0;
  let mMonthP = 0;
  let mMonthEarnings = 0;
  currentMonthEntries.forEach((it) => {
    mMonthL += (it.morningLitres || 0) + (it.eveningLitres || 0);
    mMonthP += (it.morningPav || 0) + (it.eveningPav || 0);
    mMonthEarnings += it.totalAmount || 0;
  });
  const mMonthAdded = Math.floor(mMonthP / APP_STATE.settings.pavPerLitre);
  const mMonthRem = mMonthP % APP_STATE.settings.pavPerLitre;
  document.getElementById("dashMonthMilk").innerText =
    `${mMonthL + mMonthAdded} L ${mMonthRem > 0 ? mMonthRem + " पाव" : ""}`;
  document.getElementById("dashMonthEarnings").innerText =
    formatIndianCurrency(mMonthEarnings);

  const sumSessions = morningTotalMilkDec + eveningTotalMilkDec;
  const mornPercent =
    sumSessions > 0
      ? Math.round((morningTotalMilkDec / sumSessions) * 100)
      : 50;
  const evePercent = sumSessions > 0 ? 100 - mornPercent : 50;

  document.getElementById("dashMorningTotal").innerText =
    `${morningTotalMilkDec.toFixed(1)} L (${mornPercent}%)`;
  document.getElementById("dashEveningTotal").innerText =
    `${eveningTotalMilkDec.toFixed(1)} L (${evePercent}%)`;
  document.getElementById("dashMorningBar").style.width = `${mornPercent}%`;
  document.getElementById("dashEveningBar").style.width = `${evePercent}%`;

  const recentTbody = document.getElementById("dashRecentTableBody");
  recentTbody.innerHTML = "";
  const recentFive = [...APP_STATE.entries].slice(-6).reverse();

  if (recentFive.length === 0) {
    recentTbody.innerHTML = `<tr><td colspan="8" class="text-center py-6 text-slate-400">अभी कोई रिकॉर्ड उपलब्ध नहीं है।</td></tr>`;
  } else {
    recentFive.forEach((item) => {
      const tr = document.createElement("tr");
      tr.className = "hover:bg-slate-50 dark:hover:bg-slate-750/50 transition";
      tr.innerHTML = `
            <td class="px-4 py-3 font-semibold text-slate-900 dark:text-white">${formatDateDisplay(item.date)}</td>
            <td class="px-4 py-3 text-slate-500 font-medium">${item.weekdayHindi}</td>
            <td class="px-4 py-3 text-amber-600 font-semibold">${item.morningLitres} L, ${item.morningPav} पाव</td>
            <td class="px-4 py-3 text-indigo-600 font-semibold">${item.eveningLitres} L, ${item.eveningPav} पाव</td>
            <td class="px-4 py-3 font-bold text-blue-600 dark:text-blue-400">${item.totalMilkDisplay}</td>
            <td class="px-4 py-3 font-bold">${formatIndianCurrency(item.totalAmount)}</td>
            <td class="px-4 py-3 text-emerald-600 font-medium">${formatIndianCurrency(item.paymentReceived)}</td>
            <td class="px-4 py-3 text-right font-black text-rose-600 dark:text-rose-400">${formatIndianCurrency(item.outstandingBalance)}</td>
          `;
      recentTbody.appendChild(tr);
    });
  }

  renderTodayOverview();
}

// --- 10. Master Records Tab & Table ---
function renderMasterRecords() {
  const tbody = document.getElementById("masterRecordsTableBody");
  const emptyState = document.getElementById("noRecordsState");
  const filterDate = document.getElementById("recordsFilterDate").value;
  const filterMonth = document.getElementById("recordsFilterMonth").value;
  const filterSession = document.getElementById("recordsFilterSession").value;
  const sortOrder = document.getElementById("recordsSortOrder").value;

  populateMonthDropdown();

  let records = [...APP_STATE.entries];

  if (filterDate) {
    records = records.filter((it) => it.date === filterDate);
  }
  if (filterMonth !== "all") {
    records = records.filter((it) => it.date.startsWith(filterMonth));
  }
  if (filterSession === "morning_only") {
    records = records.filter(
      (it) =>
        (it.morningLitres > 0 || it.morningPav > 0) &&
        it.eveningLitres === 0 &&
        it.eveningPav === 0,
    );
  } else if (filterSession === "evening_only") {
    records = records.filter(
      (it) =>
        (it.eveningLitres > 0 || it.eveningPav > 0) &&
        it.morningLitres === 0 &&
        it.morningPav === 0,
    );
  } else if (filterSession === "both") {
    records = records.filter(
      (it) =>
        (it.morningLitres > 0 || it.morningPav > 0) &&
        (it.eveningLitres > 0 || it.eveningPav > 0),
    );
  }

  if (sortOrder === "asc") {
    records.sort((a, b) => a.date.localeCompare(b.date));
  } else {
    records.sort((a, b) => b.date.localeCompare(a.date));
  }

  document.getElementById("recordsCountText").innerText =
    `${records.length} रिकॉर्ड्स प्रदर्शित (कुल: ${APP_STATE.entries.length})`;

  tbody.innerHTML = "";
  if (records.length === 0) {
    emptyState.classList.remove("hidden");
  } else {
    emptyState.classList.add("hidden");
    records.forEach((row, idx) => {
      const tr = document.createElement("tr");
      tr.className =
        "hover:bg-slate-50 dark:hover:bg-slate-750 transition text-xs sm:text-sm";
      tr.innerHTML = `
            <td class="px-3 py-3 text-center text-slate-400 font-mono">${row.sNo || idx + 1}</td>
            <td class="px-4 py-3 font-bold text-slate-800 dark:text-slate-100 whitespace-nowrap">${formatDateDisplay(row.date)}</td>
            <td class="px-3 py-3 text-slate-500 whitespace-nowrap">${row.weekdayHindi}</td>
            <td class="px-3 py-3 font-semibold text-amber-700 dark:text-amber-400 bg-amber-50/20">${row.morningLitres}</td>
            <td class="px-3 py-3 text-amber-700 dark:text-amber-400 bg-amber-50/20">${row.morningPav}</td>
            <td class="px-3 py-3 font-semibold text-indigo-700 dark:text-indigo-400 bg-indigo-50/20">${row.eveningLitres}</td>
            <td class="px-3 py-3 text-indigo-700 dark:text-indigo-400 bg-indigo-50/20">${row.eveningPav}</td>
            <td class="px-4 py-3 font-black text-blue-600 dark:text-blue-400 whitespace-nowrap">${row.totalMilkDisplay}</td>
            <td class="px-4 py-3 font-extrabold">${formatIndianCurrency(row.totalAmount)}</td>
            <td class="px-3 py-3 font-semibold text-emerald-600">${row.paymentReceived > 0 ? formatIndianCurrency(row.paymentReceived) : "—"}</td>
            <td class="px-3 py-3 text-rose-500">${row.deduction > 0 ? formatIndianCurrency(row.deduction) : "—"}</td>
            <td class="px-4 py-3 font-black text-rose-600 dark:text-rose-400 whitespace-nowrap">${formatIndianCurrency(row.outstandingBalance)}</td>
          `;
      tbody.appendChild(tr);
    });
  }
}

function populateMonthDropdown() {
  const select = document.getElementById("recordsFilterMonth");
  if (select.children.length > 1) return;

  const months = new Set();
  APP_STATE.entries.forEach((it) => {
    months.add(it.date.substring(0, 7));
  });

  Array.from(months)
    .sort()
    .reverse()
    .forEach((ym) => {
      const [y, m] = ym.split("-");
      const opt = document.createElement("option");
      opt.value = ym;
      opt.innerText = `${m}/${y} (Month)`;
      select.appendChild(opt);
    });
}

// --- 11. ENHANCED DATE-RANGE CALCULATOR WITH BOUNDARY SESSION CONTROLS & EXPORTS ---
function runDateRangeCalculator() {
  const startDate =
    document.getElementById("calcStartDate").value || `${getMonthPrefix()}-01`;
  const endDate =
    document.getElementById("calcEndDate").value || getIndiaDateString();
  if (startDate > endDate) {
    showToast("प्रारंभिक तारीख समाप्ति तारीख से पहले होनी चाहिए", "warn");
    return;
  }

  if (!document.getElementById("calcStartDate").value)
    document.getElementById("calcStartDate").value = startDate;
  if (!document.getElementById("calcEndDate").value)
    document.getElementById("calcEndDate").value = endDate;

  // Update weekday badges
  const startWk = getHindiWeekdayFromDateStr(startDate);
  const endWk = getHindiWeekdayFromDateStr(endDate);
  document.getElementById("calcStartWeekdayBadge").innerText = startWk.hi;
  document.getElementById("calcEndWeekdayBadge").innerText = endWk.hi;

  // Boundary checkboxes
  const startMorning = document.getElementById("calcStartMorning").checked;
  const startEvening = document.getElementById("calcStartEvening").checked;
  const endMorning = document.getElementById("calcEndMorning").checked;
  const endEvening = document.getElementById("calcEndEvening").checked;

  // Middle sessions selector ('both', 'morning', 'evening')
  const interveningSession = document.getElementById(
    "calcInterveningSession",
  ).value;

  // Rate handling (Default or Custom Override)
  const useCustomRate = document.getElementById("calcUseCustomRate").checked;
  let rateL = APP_STATE.settings.pricePerLitre;
  let rateP = APP_STATE.settings.pricePerPav;

  if (useCustomRate) {
    document.getElementById("calcCustomRateInputs").classList.remove("hidden");
    rateL =
      parseFloat(document.getElementById("calcRatePerLitre").value) || rateL;
    rateP =
      parseFloat(document.getElementById("calcRatePerPav").value) || rateP;
  } else {
    document.getElementById("calcCustomRateInputs").classList.add("hidden");
  }

  document.getElementById("calcRateAppliedBadge").innerText =
    `लागू दर: ₹${rateL}/L, ₹${rateP}/पाव`;

  // Filter entries strictly within range
  const inRangeEntries = APP_STATE.entries.filter(
    (e) => e.date >= startDate && e.date <= endDate,
  );

  // Total calendar days
  const d1 = new Date(startDate);
  const d2 = new Date(endDate);
  const diffTime = Math.abs(d2 - d1);
  const totalCalendarDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24)) + 1;

  let suppliedDaysCount = 0;
  let sumMorningLitres = 0;
  let sumMorningPav = 0;
  let sumEveningLitres = 0;
  let sumEveningPav = 0;
  let sumGrandAmount = 0;
  let sumTotalPaid = 0;

  const breakdownRows = [];

  inRangeEntries.forEach((entry) => {
    const isStart = entry.date === startDate;
    const isEnd = entry.date === endDate;
    let incM = true;
    let incE = true;

    if (isStart && isEnd) {
      // Both dates are identical (1 day range)
      incM = startMorning && endMorning;
      incE = startEvening && endEvening;
    } else if (isStart) {
      incM = startMorning;
      incE = startEvening;
    } else if (isEnd) {
      incM = endMorning;
      incE = endEvening;
    } else {
      // Middle intervening date
      if (interveningSession === "morning") {
        incM = true;
        incE = false;
      } else if (interveningSession === "evening") {
        incM = false;
        incE = true;
      } else {
        incM = true;
        incE = true;
      }
    }

    const mL = incM ? entry.morningLitres || 0 : 0;
    const mP = incM ? entry.morningPav || 0 : 0;
    const eL = incE ? entry.eveningLitres || 0 : 0;
    const eP = incE ? entry.eveningPav || 0 : 0;

    let appliedSessionTag = "";
    if (incM && incE) appliedSessionTag = "दोनों (सुबह + शाम)";
    else if (incM) appliedSessionTag = "केवल सुबह (Morning)";
    else if (incE) appliedSessionTag = "केवल शाम (Evening)";
    else appliedSessionTag = "शून्य (कोई नहीं)";

    const dayPav = mP + eP;
    const dayAddedL = Math.floor(dayPav / APP_STATE.settings.pavPerLitre);
    const dayRemPav = dayPav % APP_STATE.settings.pavPerLitre;
    const dayTotalL = mL + eL + dayAddedL;
    const dayAmount = dayTotalL * rateL + dayRemPav * rateP;

    if (mL > 0 || mP > 0 || eL > 0 || eP > 0) {
      suppliedDaysCount++;
    }

    sumMorningLitres += mL;
    sumMorningPav += mP;
    sumEveningLitres += eL;
    sumEveningPav += eP;
    sumGrandAmount += dayAmount;
    sumTotalPaid += entry.paymentReceived || 0;

    let rowMilkDisplay = `${dayTotalL} L`;
    if (dayRemPav > 0) rowMilkDisplay += ` ${dayRemPav} पाव`;

    breakdownRows.push({
      date: entry.date,
      weekdayHindi: entry.weekdayHindi,
      appliedSessionTag,
      morningDisplay: incM
        ? `${entry.morningLitres} L ${entry.morningPav} पाव`
        : "— (शामिल नहीं)",
      eveningDisplay: incE
        ? `${entry.eveningLitres} L ${entry.eveningPav} पाव`
        : "— (शामिल नहीं)",
      rowMilkDisplay,
      dayTotalL,
      dayRemPav,
      dayAmount,
      paymentReceived: entry.paymentReceived || 0,
    });
  });

  // Summation logic across chosen boundary
  const mAdded = Math.floor(sumMorningPav / APP_STATE.settings.pavPerLitre);
  const mRem = sumMorningPav % APP_STATE.settings.pavPerLitre;
  const mFullLitres = sumMorningLitres + mAdded;

  const eAdded = Math.floor(sumEveningPav / APP_STATE.settings.pavPerLitre);
  const eRem = sumEveningPav % APP_STATE.settings.pavPerLitre;
  const eFullLitres = sumEveningLitres + eAdded;

  const combPav = sumMorningPav + sumEveningPav;
  const combAdded = Math.floor(combPav / APP_STATE.settings.pavPerLitre);
  const combRem = combPav % APP_STATE.settings.pavPerLitre;
  const grandTotalLitres = sumMorningLitres + sumEveningLitres + combAdded;
  const grandDecimalLitres =
    grandTotalLitres + combRem / APP_STATE.settings.pavPerLitre;

  // Populate UI Outputs
  document.getElementById("calcTotalDays").innerText = totalCalendarDays;
  document.getElementById("calcSuppliedDays").innerText = suppliedDaysCount;
  document.getElementById("calcMorningMilk").innerText =
    `${mFullLitres} L ${mRem > 0 ? mRem + " पाव" : ""}`;
  document.getElementById("calcMorningPavTotal").innerText =
    `${sumMorningPav} पाव (${(mFullLitres + mRem / 4).toFixed(2)} L)`;
  document.getElementById("calcEveningMilk").innerText =
    `${eFullLitres} L ${eRem > 0 ? eRem + " पाव" : ""}`;
  document.getElementById("calcEveningPavTotal").innerText =
    `${sumEveningPav} पाव (${(eFullLitres + eRem / 4).toFixed(2)} L)`;
  document.getElementById("calcGrandTotalMilk").innerText =
    `${grandTotalLitres} L ${combRem > 0 ? combRem + " पाव" : ""}`;
  document.getElementById("calcDecimalMilk").innerText =
    `(${grandDecimalLitres.toFixed(2)} Litres)`;
  document.getElementById("calcGrandTotalAmount").innerText =
    formatIndianCurrency(sumGrandAmount);
  document.getElementById("calcBreakdownNetText").innerText =
    formatIndianCurrency(sumGrandAmount);

  // Store in App State for 1-click Export
  APP_STATE.lastCalculatedRange = {
    startDate,
    endDate,
    startMorning,
    startEvening,
    endMorning,
    endEvening,
    interveningSession,
    rateL,
    rateP,
    totalCalendarDays,
    suppliedDaysCount,
    mFullLitres,
    mRem,
    eFullLitres,
    eRem,
    grandTotalLitres,
    combRem,
    grandDecimalLitres,
    sumGrandAmount,
    sumTotalPaid,
    breakdownRows,
  };

  // Populate Breakdown Table
  const breakdownTbody = document.getElementById("calcBreakdownTableBody");
  breakdownTbody.innerHTML = "";

  if (breakdownRows.length === 0) {
    breakdownTbody.innerHTML = `<tr><td colspan="8" class="text-center py-6 text-slate-400">चयनित सीमा में कोई रिकॉर्ड नहीं मिला।</td></tr>`;
  } else {
    breakdownRows.forEach((row) => {
      const tr = document.createElement("tr");
      tr.className =
        "hover:bg-slate-50 dark:hover:bg-slate-750 transition text-xs";
      tr.innerHTML = `
            <td class="px-3 py-2 font-bold text-slate-800 dark:text-slate-100 whitespace-nowrap">${formatDateDisplay(row.date)}</td>
            <td class="px-3 py-2 text-slate-500 whitespace-nowrap">${row.weekdayHindi}</td>
            <td class="px-3 py-2 whitespace-nowrap font-medium text-blue-600 dark:text-blue-400">${row.appliedSessionTag}</td>
            <td class="px-3 py-2 text-amber-700 dark:text-amber-400">${row.morningDisplay}</td>
            <td class="px-3 py-2 text-indigo-700 dark:text-indigo-400">${row.eveningDisplay}</td>
            <td class="px-3 py-2 font-black text-blue-700 dark:text-blue-300">${row.rowMilkDisplay}</td>
            <td class="px-3 py-2 text-right font-black text-slate-900 dark:text-white">${formatIndianCurrency(row.dayAmount)}</td>
            <td class="px-3 py-2 text-right text-emerald-600 font-semibold">${row.paymentReceived > 0 ? formatIndianCurrency(row.paymentReceived) : "—"}</td>
          `;
      breakdownTbody.appendChild(tr);
    });
  }
}

// --- Range Excel Export ---
function exportRangeCalculatorExcel() {
  const calc = APP_STATE.lastCalculatedRange;
  if (!calc || !calc.breakdownRows || calc.breakdownRows.length === 0) {
    showToast("निर्यात करने के लिए कोई रेंज डेटा नहीं है", "warn");
    return;
  }

  if (!window.XLSX) {
    showToast("एक्सेल लाइब्रेरी लोड नहीं हुई है", "error");
    return;
  }

  // Metadata summary rows
  const data = [
    ["दूध हिसाब – तारीख सीमा गणना रिपोर्ट (Date-Range Calculation)"],
    [
      "प्रारंभिक तारीख",
      calc.startDate,
      "सुबह शामिल",
      calc.startMorning ? "हाँ" : "नहीं",
      "शाम शामिल",
      calc.startEvening ? "हाँ" : "नहीं",
    ],
    [
      "समाप्ति तारीख",
      calc.endDate,
      "सुबह शामिल",
      calc.endMorning ? "हाँ" : "नहीं",
      "शाम शामिल",
      calc.endEvening ? "हाँ" : "नहीं",
    ],
    [
      "बीच के दिनों का सत्र",
      calc.interveningSession,
      "लागू प्रति लीटर दर",
      `₹${calc.rateL}`,
      "प्रति पाव दर",
      `₹${calc.rateP}`,
    ],
    [
      "कुल कैलेंडर दिन",
      calc.totalCalendarDays,
      "दूध आपूर्ति दिन",
      calc.suppliedDaysCount,
    ],
    [
      "कुल सुबह दूध",
      `${calc.mFullLitres} L ${calc.mRem > 0 ? calc.mRem + " पाव" : ""}`,
      "कुल शाम दूध",
      `${calc.eFullLitres} L ${calc.eRem > 0 ? calc.eRem + " पाव" : ""}`,
    ],
    [
      "रेंज का कुल दूध",
      `${calc.grandTotalLitres} L ${calc.combRem > 0 ? calc.combRem + " पाव" : ""}`,
      "कुल दशमलव दूध",
      `${calc.grandDecimalLitres.toFixed(2)} L`,
    ],
    [
      "कुल देय रकम",
      `₹${calc.sumGrandAmount}`,
      "प्राप्त भुगतान",
      `₹${calc.sumTotalPaid}`,
    ],
    [], // blank spacer row
    [
      "क्र.सं.",
      "दिनांक",
      "दिन",
      "लागू सत्र",
      "सुबह दूध",
      "शाम दूध",
      "चयनित कुल दूध",
      "रकम (₹)",
      "प्राप्त पैसे (₹)",
    ],
  ];

  calc.breakdownRows.forEach((r, idx) => {
    data.push([
      idx + 1,
      r.date,
      r.weekdayHindi,
      r.appliedSessionTag,
      r.morningDisplay,
      r.eveningDisplay,
      r.rowMilkDisplay,
      r.dayAmount,
      r.paymentReceived,
    ]);
  });

  // Total row at bottom
  data.push([
    "कुल",
    `${calc.startDate} से ${calc.endDate}`,
    "",
    "",
    `${calc.mFullLitres} L ${calc.mRem} पाव`,
    `${calc.eFullLitres} L ${calc.eRem} पाव`,
    `${calc.grandTotalLitres} L ${calc.combRem} पाव`,
    calc.sumGrandAmount,
    calc.sumTotalPaid,
  ]);

  const ws = XLSX.utils.aoa_to_sheet(data);
  ws["!cols"] = [
    { wch: 8 },
    { wch: 14 },
    { wch: 12 },
    { wch: 22 },
    { wch: 18 },
    { wch: 18 },
    { wch: 18 },
    { wch: 14 },
    { wch: 14 },
  ];

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "रेंज गणना हिसाब");
  const filename = `रेंज_दूध_हिसाब_${calc.startDate}_से_${calc.endDate}.xlsx`;
  XLSX.writeFile(wb, filename);
  showToast(`रेंज एक्सेल डाउनलोड सफल: ${filename}`, "success");
}

// --- Range CSV Export ---
function exportRangeCalculatorCsv() {
  const calc = APP_STATE.lastCalculatedRange;
  if (!calc || !calc.breakdownRows) {
    showToast("कोई डेटा उपलब्ध नहीं", "warn");
    return;
  }

  let csv = `दिनांक,दिन,लागू सत्र,सुबह दूध,शाम दूध,कुल दूध,रकम (₹),प्राप्त पैसे\n`;
  calc.breakdownRows.forEach((r) => {
    csv += `"${r.date}","${r.weekdayHindi}","${r.appliedSessionTag}","${r.morningDisplay}","${r.eveningDisplay}","${r.rowMilkDisplay}","${r.dayAmount}","${r.paymentReceived}"\n`;
  });
  csv += `"कुल","${calc.startDate} से ${calc.endDate}","","","","${calc.grandTotalLitres} L ${calc.combRem} पाव","${calc.sumGrandAmount}","${calc.sumTotalPaid}"\n`;

  const blob = new Blob(["\uFEFF" + csv], { type: "text/csv;charset=utf-8;" });
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = `दूध_हिसाब_${calc.startDate}_से_${calc.endDate}.csv`;
  link.click();
  showToast("CSV फ़ाइल डाउनलोड हो गई", "success");
}

function printSelectedRangeSlip() {
  const calc = APP_STATE.lastCalculatedRange;
  if (!calc) {
    showToast("पहले तारीख सीमा चुनें", "warn");
    return;
  }

  const slip = document.getElementById("rangePrintSlip");
  slip.replaceChildren();

  const title = document.createElement("h1");
  title.textContent = "दूध हिसाब — तारीख सीमा पर्ची";
  title.className = "text-2xl font-bold mb-1";
  slip.appendChild(title);

  const range = document.createElement("p");
  range.textContent = `${formatDateDisplay(calc.startDate)} से ${formatDateDisplay(calc.endDate)}`;
  range.className = "text-sm text-slate-600";
  slip.appendChild(range);

  const summary = document.createElement("div");
  summary.className = "grid grid-cols-2 gap-3 mt-4";
  [
    ["कुल कैलेंडर दिन", calc.totalCalendarDays],
    ["दूध आपूर्ति दिन", calc.suppliedDaysCount],
    ["कुल दूध", `${calc.grandTotalLitres} L ${calc.combRem} पाव`],
    ["कुल देय रकम", formatIndianCurrency(calc.sumGrandAmount)],
    ["प्राप्त भुगतान", formatIndianCurrency(calc.sumTotalPaid)],
    ["लागू दर", `₹${calc.rateL}/L, ₹${calc.rateP}/पाव`],
  ].forEach(([label, value]) => {
    const item = document.createElement("p");
    const strong = document.createElement("strong");
    strong.textContent = `${label}: `;
    item.append(strong, document.createTextNode(String(value)));
    summary.appendChild(item);
  });
  slip.appendChild(summary);

  const table = document.createElement("table");
  const head = table.createTHead().insertRow();
  ["दिनांक", "दिन", "सत्र", "सुबह", "शाम", "कुल दूध", "रकम", "प्राप्त"].forEach(
    (label) => {
      const cell = document.createElement("th");
      cell.textContent = label;
      head.appendChild(cell);
    },
  );

  const body = table.createTBody();
  calc.breakdownRows.forEach((row) => {
    const cells = [
      formatDateDisplay(row.date),
      row.weekdayHindi,
      row.appliedSessionTag,
      row.morningDisplay,
      row.eveningDisplay,
      row.rowMilkDisplay,
      formatIndianCurrency(row.dayAmount),
      row.paymentReceived > 0 ? formatIndianCurrency(row.paymentReceived) : "—",
    ];
    const tr = body.insertRow();
    cells.forEach((value) => {
      const cell = tr.insertCell();
      cell.textContent = String(value);
    });
  });
  slip.appendChild(table);

  document.body.classList.add("print-range-slip");
  window.onafterprint = () =>
    document.body.classList.remove("print-range-slip");
  window.print();
}

// --- 12. Payments Ledger Tab ---
function renderPaymentsLedger() {
  let totalReceived = 0;
  let totalDeductions = 0;
  APP_STATE.entries.forEach((e) => {
    totalReceived += e.paymentReceived || 0;
    totalDeductions += e.deduction || 0;
  });

  const currentBalance =
    APP_STATE.entries.length > 0
      ? APP_STATE.entries[APP_STATE.entries.length - 1].outstandingBalance
      : APP_STATE.settings.openingBalance || 0;

  document.getElementById("paymentsTotalReceived").innerText =
    formatIndianCurrency(totalReceived);
  document.getElementById("paymentsTotalDeductions").innerText =
    formatIndianCurrency(totalDeductions);
  document.getElementById("paymentsNetOutstanding").innerText =
    formatIndianCurrency(currentBalance);

  const tbody = document.getElementById("paymentsLedgerTableBody");
  tbody.innerHTML = "";

  const txRows = [...APP_STATE.entries].reverse();
  if (txRows.length === 0) {
    tbody.innerHTML = `<tr><td colspan="6" class="text-center py-6 text-slate-400">कोई लेन-देन उपलब्ध नहीं है।</td></tr>`;
    return;
  }

  txRows.forEach((row) => {
    const tr = document.createElement("tr");
    tr.className =
      "hover:bg-slate-50 dark:hover:bg-slate-750 transition text-xs sm:text-sm";
    tr.innerHTML = `
          <td class="px-4 py-3 font-semibold text-slate-800 dark:text-slate-100">${formatDateDisplay(row.date)}</td>
          <td class="px-4 py-3 text-slate-500">${row.weekdayHindi}</td>
          <td class="px-4 py-3 font-semibold">${formatIndianCurrency(row.totalAmount)}</td>
          <td class="px-4 py-3 text-emerald-600 font-bold">${row.paymentReceived > 0 ? formatIndianCurrency(row.paymentReceived) : "—"}</td>
          <td class="px-4 py-3 text-rose-500 font-medium">${row.deduction > 0 ? formatIndianCurrency(row.deduction) : "—"}</td>
          <td class="px-4 py-3 text-rose-600 dark:text-rose-400 font-black">${formatIndianCurrency(row.outstandingBalance)}</td>
        `;
    tbody.appendChild(tr);
  });
}

// --- 13. Reports & Invoices View ---
function renderReportsView() {
  const period = document.getElementById("reportPeriodSelect").value;
  document.getElementById("reportGeneratedDate").innerText =
    formatDateDisplay(getIndiaDateString());
  document.getElementById("reportRateInfo").innerText =
    `₹${APP_STATE.settings.pricePerLitre}/L, ₹${APP_STATE.settings.pricePerPav}/पाव`;

  let reportData = [...APP_STATE.entries];
  const currentMonth = getMonthPrefix();
  const previousMonth = getPreviousMonthPrefix();
  if (period === "this_month") {
    reportData = reportData.filter((e) => e.date.startsWith(currentMonth));
  } else if (period === "last_month") {
    reportData = reportData.filter((e) => e.date.startsWith(previousMonth));
  }

  let totalL = 0;
  let totalP = 0;
  let totalAmt = 0;
  let totalPaid = 0;

  const tbody = document.getElementById("reportTableBody");
  tbody.innerHTML = "";

  reportData.forEach((row) => {
    totalL += (row.morningLitres || 0) + (row.eveningLitres || 0);
    totalP += (row.morningPav || 0) + (row.eveningPav || 0);
    totalAmt += row.totalAmount || 0;
    totalPaid += row.paymentReceived || 0;

    const tr = document.createElement("tr");
    tr.className = "text-xs";
    tr.innerHTML = `
          <td class="p-2 border font-medium">${formatDateDisplay(row.date)}</td>
          <td class="p-2 border text-slate-500">${row.weekdayHindi}</td>
          <td class="p-2 border">${row.morningLitres} L ${row.morningPav} पाव</td>
          <td class="p-2 border">${row.eveningLitres} L ${row.eveningPav} पाव</td>
          <td class="p-2 border font-bold">${row.totalMilkDisplay}</td>
          <td class="p-2 border font-semibold">${formatIndianCurrency(row.totalAmount)}</td>
          <td class="p-2 border text-emerald-600">${row.paymentReceived > 0 ? formatIndianCurrency(row.paymentReceived) : "—"}</td>
          <td class="p-2 border text-rose-500">${row.deduction > 0 ? formatIndianCurrency(row.deduction) : "—"}</td>
          <td class="p-2 border text-right font-black">${formatIndianCurrency(row.outstandingBalance)}</td>
        `;
    tbody.appendChild(tr);
  });

  const addedL = Math.floor(totalP / APP_STATE.settings.pavPerLitre);
  const remP = totalP % APP_STATE.settings.pavPerLitre;
  const finalLitres = totalL + addedL;

  const latestBalance =
    reportData.length > 0
      ? reportData[reportData.length - 1].outstandingBalance
      : 0;

  document.getElementById("reportTotalMilkQty").innerText =
    `${finalLitres} L ${remP > 0 ? remP + " पाव" : ""}`;
  document.getElementById("reportTotalAmount").innerText =
    formatIndianCurrency(totalAmt);
  document.getElementById("reportTotalPaid").innerText =
    formatIndianCurrency(totalPaid);
  document.getElementById("reportFinalBalance").innerText =
    formatIndianCurrency(latestBalance);

  renderTrendChart(reportData);
}

function renderTrendChart(data) {
  const ctx = document.getElementById("milkTrendChart");
  if (!ctx) return;

  if (APP_STATE.chartInstance) {
    APP_STATE.chartInstance.destroy();
  }

  const labels = data.map((d) => d.date.substring(5));
  const morningData = data.map((d) => d.morningDecimal);
  const eveningData = data.map((d) => d.eveningDecimal);

  APP_STATE.chartInstance = new Chart(ctx, {
    type: "line",
    data: {
      labels,
      datasets: [
        {
          label: "सुबह का दूध (L)",
          data: morningData,
          borderColor: "#f59e0b",
          backgroundColor: "rgba(245, 158, 11, 0.15)",
          tension: 0.3,
          fill: true,
        },
        {
          label: "शाम का दूध (L)",
          data: eveningData,
          borderColor: "#6366f1",
          backgroundColor: "rgba(99, 102, 241, 0.15)",
          tension: 0.3,
          fill: true,
        },
      ],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: {
          position: "top",
          labels: {
            boxWidth: 12,
            font: { family: "Inter", size: 11 },
          },
        },
      },
      scales: {
        y: {
          beginAtZero: true,
          title: { display: true, text: "लीटर (Litres)" },
        },
      },
    },
  });
}

// --- 14. Google Sheets data source ---
function normalizeSheetHeader(header) {
  return String(header)
    .trim()
    .replace(/\s+/g, " ")
    .replace(/\s*([()])\s*/g, "$1")
    .toLowerCase();
}

function getSheetValue(row, aliases) {
  const values = new Map(
    Object.entries(row).map(([header, value]) => [
      normalizeSheetHeader(header),
      value,
    ]),
  );
  for (const alias of aliases) {
    const key = normalizeSheetHeader(alias);
    if (values.has(key)) return values.get(key);
  }
  return "";
}

function parseSheetNumber(value) {
  const parsed = Number.parseFloat(String(value ?? "").replace(/[₹,\s]/g, ""));
  return Number.isFinite(parsed) ? parsed : 0;
}

function parseOptionalSheetNumber(value) {
  if (value === null || value === undefined || String(value).trim() === "")
    return null;
  const parsed = Number.parseFloat(String(value).replace(/[₹,\s]/g, ""));
  return Number.isFinite(parsed) ? parsed : null;
}

function setSheetSyncStatus(status, isError = false) {
  const badge = document.getElementById("headerSyncBadge");
  if (!badge) return;

  badge.className = isError
    ? "flex items-center gap-1.5 px-2.5 py-0.5 rounded-lg bg-rose-100 dark:bg-rose-950/60 text-rose-700 dark:text-rose-300 font-bold border border-rose-300 dark:border-rose-800"
    : "flex items-center gap-1.5 px-2.5 py-0.5 rounded-lg bg-emerald-100 dark:bg-emerald-950/60 text-emerald-700 dark:text-emerald-300 font-bold border border-emerald-300 dark:border-emerald-800";
  badge.innerHTML = `<i data-lucide="${isError ? "cloud-off" : "cloud"}" class="w-3.5 h-3.5" aria-hidden="true"></i><span>${status}</span>`;
  if (window.lucide) window.lucide.createIcons();
}

async function fetchFromGoogleSheet() {
  setSheetSyncStatus("Loading..");

  try {
    const response = await fetch(SHEET_ENDPOINT_URL, { cache: "no-store" });
    if (!response.ok) throw new Error(`HTTP Error: ${response.status}`);

    const json = await response.json();
    const rawRows = Array.isArray(json) ? json : json.data || json.rows || [];
    if (!Array.isArray(rawRows) || rawRows.length === 0) {
      throw new Error("Google Sheet returned no records.");
    }

    const parsedRows = rawRows
      .map((row) => {
        const date = parseSheetDate(
          getSheetValue(row, ["दिनांक", "Date", "date"]),
        );
        if (!date) return null;

        const weekday = getHindiWeekdayFromDateStr(date);
        return {
          date,
          sNo: parseSheetNumber(getSheetValue(row, ["S.No", "S.No."])),
          sheetTotalAmount: parseOptionalSheetNumber(
            getSheetValue(row, ["कुल पैसे", "Total Amount"]),
          ),
          sheetOutstandingBalance: parseOptionalSheetNumber(
            getSheetValue(row, ["बकाया पैसे", "Outstanding Balance"]),
          ),
          weekdayHindi: String(
            getSheetValue(row, ["दिन", "Day"]) || weekday.hi,
          ).trim(),
          weekdayEnglish: weekday.en,
          morningLitres: Math.max(
            0,
            Math.trunc(
              parseSheetNumber(
                getSheetValue(row, [
                  "सुबह दूध (L)",
                  "Morning Litres",
                  "सुबह दूध",
                ]),
              ),
            ),
          ),
          morningPav: Math.max(
            0,
            Math.trunc(
              parseSheetNumber(
                getSheetValue(row, ["सुबह (पाव)", "Morning Pav", "सुबह पाव"]),
              ),
            ),
          ),
          eveningLitres: Math.max(
            0,
            Math.trunc(
              parseSheetNumber(
                getSheetValue(row, [
                  "शाम दूध (L)",
                  "Evening Litres",
                  "शाम दूध",
                ]),
              ),
            ),
          ),
          eveningPav: Math.max(
            0,
            Math.trunc(
              parseSheetNumber(
                getSheetValue(row, ["शाम (पाव)", "Evening Pav", "शाम पाव"]),
              ),
            ),
          ),
          paymentReceived: Math.max(
            0,
            parseSheetNumber(
              getSheetValue(row, [
                "प्राप्त पैसे",
                "Payment Received",
                "प्राप्त",
              ]),
            ),
          ),
          deduction: Math.max(
            0,
            parseSheetNumber(
              getSheetValue(row, ["कटौती पैसे", "Deduction", "कटौती"]),
            ),
          ),
        };
      })
      .filter(Boolean);

    if (parsedRows.length === 0) {
      throw new Error(
        "No valid rows found; check the sheet date and column headers.",
      );
    }

    const entries = recalculateLedger(parsedRows);
    await replaceEntriesInDB(entries);
    APP_STATE.entries = entries;

    setSheetSyncStatus(`Live • ${entries.length}`);
    renderTodayOverview();
    renderDashboard();
    renderMasterRecords();
    renderPaymentsLedger();
    renderReportsView();
    runDateRangeCalculator();
  } catch (error) {
    console.error("Google Sheets sync failed:", error);
    setSheetSyncStatus("सिंक विफल • स्थानीय डेटा", true);
    showToast(
      "शीट डेटा लोड नहीं हो सका। स्थानीय रिकॉर्ड दिखाए जा रहे हैं।",
      "error",
    );
  }
}

// --- 15. Default Seed Data (October 2026 Test Environment) ---
function generateSeedData() {
  return [
    {
      date: "2026-10-20",
      morningLitres: 2,
      morningPav: 1,
      eveningLitres: 2,
      eveningPav: 0,
      paymentReceived: 0,
      deduction: 0,
    },
    {
      date: "2026-10-21",
      morningLitres: 2,
      morningPav: 0,
      eveningLitres: 2,
      eveningPav: 1,
      paymentReceived: 0,
      deduction: 0,
    },
    {
      date: "2026-10-22",
      morningLitres: 3,
      morningPav: 0,
      eveningLitres: 2,
      eveningPav: 2,
      paymentReceived: 500,
      deduction: 0,
    },
    {
      date: "2026-10-23",
      morningLitres: 2,
      morningPav: 2,
      eveningLitres: 2,
      eveningPav: 0,
      paymentReceived: 0,
      deduction: 0,
    },
    {
      date: "2026-10-24",
      morningLitres: 2,
      morningPav: 0,
      eveningLitres: 3,
      eveningPav: 0,
      paymentReceived: 0,
      deduction: 50,
    },
    {
      date: "2026-10-25",
      morningLitres: 3,
      morningPav: 1,
      eveningLitres: 2,
      eveningPav: 1,
      paymentReceived: 0,
      deduction: 0,
    },
    {
      date: "2026-10-26",
      morningLitres: 2,
      morningPav: 0,
      eveningLitres: 2,
      eveningPav: 0,
      paymentReceived: 0,
      deduction: 0,
    },
    {
      date: "2026-10-27",
      morningLitres: 2,
      morningPav: 3,
      eveningLitres: 2,
      eveningPav: 2,
      paymentReceived: 0,
      deduction: 0,
    },
    {
      date: "2026-10-28",
      morningLitres: 3,
      morningPav: 0,
      eveningLitres: 3,
      eveningPav: 0,
      paymentReceived: 0,
      deduction: 0,
    },
    {
      date: "2026-10-29",
      morningLitres: 2,
      morningPav: 2,
      eveningLitres: 2,
      eveningPav: 1,
      paymentReceived: 0,
      deduction: 0,
    },
    {
      date: "2026-10-30",
      morningLitres: 3,
      morningPav: 0,
      eveningLitres: 2,
      eveningPav: 2,
      paymentReceived: 0,
      deduction: 0,
    },
  ].map((it) => {
    const weekday = getHindiWeekdayFromDateStr(it.date);
    return {
      ...it,
      weekdayHindi: weekday.hi,
      weekdayEnglish: weekday.en,
    };
  });
}

// --- 16. Event Listeners & Bootstrapping ---
document.addEventListener("DOMContentLoaded", async () => {
  try {
    await openDatabase();
    const storedSettings = await getSettingsFromDB();
    if (storedSettings && Object.keys(storedSettings).length > 0) {
      APP_STATE.settings = { ...APP_STATE.settings, ...storedSettings };
    }
    const savedPavPerLitre = Number(APP_STATE.settings.pavPerLitre);
    if (!Number.isInteger(savedPavPerLitre) || savedPavPerLitre <= 0) {
      APP_STATE.settings.pavPerLitre = 4;
      await saveSettingToDB("pavPerLitre", 4);
    }
    for (const key of ["pricePerLitre", "pricePerPav"]) {
      const value = Number(APP_STATE.settings[key]);
      if (!Number.isFinite(value) || value < 0) {
        APP_STATE.settings[key] = key === "pricePerLitre" ? 33 : 8;
        await saveSettingToDB(key, APP_STATE.settings[key]);
      }
    }
    const openingBalance = Number(APP_STATE.settings.openingBalance);
    if (!Number.isFinite(openingBalance)) {
      APP_STATE.settings.openingBalance = 0;
      await saveSettingToDB("openingBalance", 0);
    }

    const storedEntries = await getAllEntriesFromDB();
    if (storedEntries && storedEntries.length > 0) {
      APP_STATE.entries = recalculateLedger(storedEntries);
    }
  } catch (err) {
    console.error(
      "Persistent storage is unavailable; using in-memory data:",
      err,
    );
    APP_STATE.entries = [];
  }

  // Initialize Settings Values
  document.getElementById("settingRatePerLitre").value =
    APP_STATE.settings.pricePerLitre;
  document.getElementById("settingRatePerPav").value =
    APP_STATE.settings.pricePerPav;
  document.getElementById("settingPavPerLitre").value =
    APP_STATE.settings.pavPerLitre;
  document.getElementById("settingOpeningBalance").value =
    APP_STATE.settings.openingBalance;

  const today = getIndiaDateString();
  document.getElementById("calcStartDate").value = `${getMonthPrefix()}-01`;
  document.getElementById("calcEndDate").value = today;

  // Navigation Tab Clicks
  if (window.lucide) window.lucide.createIcons();
  document.querySelectorAll(".tab-btn").forEach((btn) => {
    btn.addEventListener("click", () => {
      switchTab(btn.getAttribute("data-tab"));
    });
  });
  document.querySelectorAll(".milk-summary-toggle").forEach((button) => {
    button.addEventListener("click", () => {
      const isExpanded = button.getAttribute("aria-expanded") === "true";
      const summary = document.getElementById(button.getAttribute("aria-controls"));
      button.setAttribute("aria-expanded", String(!isExpanded));
      summary.classList.toggle("hidden", isExpanded);
      button.innerText = isExpanded ? "More Details" : "Hide Details";
    });
  });

  // Defensive Safe Listener Binding Helper
  function addSafeListener(id, event, handler) {
    const el = document.getElementById(id);
    if (el) {
      el.addEventListener(event, handler);
    }
  }

  addSafeListener("milkSummaryPreviousWeek", "click", () => {
    milkSummaryWeekOffset -= 1;
    renderTodayOverview();
  });
  addSafeListener("milkSummaryNextWeek", "click", () => {
    milkSummaryWeekOffset = Math.min(0, milkSummaryWeekOffset + 1);
    renderTodayOverview();
  });
  addSafeListener("milkSummaryCurrentWeek", "click", () => {
    milkSummaryWeekOffset = 0;
    renderTodayOverview();
  });

  // Dashboard Filters
  document.querySelectorAll(".dash-filter-btn").forEach((btn) => {
    btn.addEventListener("click", () => {
      document.querySelectorAll(".dash-filter-btn").forEach((b) => {
        b.className =
          "dash-filter-btn px-3 py-1.5 rounded-lg text-xs font-bold text-slate-600 dark:text-slate-300 hover:bg-slate-100 dark:hover:bg-slate-700";
      });
      btn.className =
        "dash-filter-btn px-3 py-1.5 rounded-lg text-xs font-bold bg-blue-100 dark:bg-blue-900/60 text-blue-700 dark:text-blue-300";
      APP_STATE.dashboardFilter = btn.getAttribute("data-period");
      renderDashboard();
    });
  });

  addSafeListener("applyDashCustomFilter", "click", renderDashboard);

  // Date Range Calculator Dynamic Listeners
  const rangeInputIds = [
    "calcStartDate",
    "calcEndDate",
    "calcStartMorning",
    "calcStartEvening",
    "calcEndMorning",
    "calcEndEvening",
    "calcInterveningSession",
    "calcUseCustomRate",
    "calcRatePerLitre",
    "calcRatePerPav",
  ];

  rangeInputIds.forEach((id) => {
    const el = document.getElementById(id);
    if (el) {
      el.addEventListener("change", runDateRangeCalculator);
      el.addEventListener("input", runDateRangeCalculator);
    }
  });

  document.querySelectorAll(".calc-preset-btn").forEach((btn) => {
    btn.addEventListener("click", () => {
      const preset = btn.getAttribute("data-preset");
      const end = getIndiaDateString();
      const endEl = document.getElementById("calcEndDate");
      const startEl = document.getElementById("calcStartDate");
      if (endEl) endEl.value = end;

      if (startEl) {
        const durationByPreset = { last7: 7, last10: 10, last15: 15 };
        if (durationByPreset[preset]) {
          const [year, month, day] = end.split("-").map(Number);
          const start = new Date(
            Date.UTC(year, month - 1, day - durationByPreset[preset] + 1),
          );
          startEl.value = start.toISOString().slice(0, 10);
        } else if (preset === "thisMonth") {
          startEl.value = `${getMonthPrefix()}-01`;
        }
      }

      const sm = document.getElementById("calcStartMorning");
      const se = document.getElementById("calcStartEvening");
      const em = document.getElementById("calcEndMorning");
      const ee = document.getElementById("calcEndEvening");
      if (sm) sm.checked = true;
      if (se) se.checked = true;
      if (em) em.checked = true;
      if (ee) ee.checked = true;
      runDateRangeCalculator();
    });
  });

  // Range Calculator Export Actions
  addSafeListener("calcExportExcelBtn", "click", exportRangeCalculatorExcel);
  addSafeListener("calcExportCsvBtn", "click", exportRangeCalculatorCsv);
  addSafeListener("calcPrintSummaryBtn", "click", printSelectedRangeSlip);

  // Master Records Filter Listeners
  [
    "recordsFilterDate",
    "recordsFilterMonth",
    "recordsFilterSession",
    "recordsSortOrder",
  ].forEach((id) => {
    const el = document.getElementById(id);
    if (el) el.addEventListener("change", renderMasterRecords);
  });

  addSafeListener("clearRecordsFilterBtn", "click", () => {
    const fd = document.getElementById("recordsFilterDate");
    const fm = document.getElementById("recordsFilterMonth");
    const fs = document.getElementById("recordsFilterSession");
    const so = document.getElementById("recordsSortOrder");
    if (fd) fd.value = "";
    if (fm) fm.value = "all";
    if (fs) fs.value = "all";
    if (so) so.value = "desc";
    renderMasterRecords();
  });

  // Standalone Payments & Reports
  addSafeListener("reportPeriodSelect", "change", renderReportsView);

  // Settings Form
  const settingsForm = document.getElementById("settingsForm");
  if (settingsForm) {
    settingsForm.addEventListener("submit", async (e) => {
      e.preventDefault();
      const readSettingNumber = (id, fallback) => {
        const raw = document.getElementById(id)?.value?.trim() ?? "";
        if (raw === "") return fallback;
        const value = Number(raw);
        return Number.isFinite(value) ? value : null;
      };
      const pricePerLitre = readSettingNumber("settingRatePerLitre", 33);
      const pricePerPav = readSettingNumber("settingRatePerPav", 8);
      const pavPerLitre = readSettingNumber("settingPavPerLitre", 4);
      const openingBalance = readSettingNumber("settingOpeningBalance", 0);

      if (
        pricePerLitre === null || pricePerLitre < 0 ||
        pricePerPav === null || pricePerPav < 0 ||
        !Number.isInteger(pavPerLitre) || pavPerLitre <= 0 ||
        openingBalance === null
      ) {
        showToast(
          "दर शून्य या उससे अधिक रखें और पाव प्रति लीटर में 1 या उससे अधिक पूर्णांक दर्ज करें।",
          "error",
        );
        return;
      }

      APP_STATE.settings.pricePerLitre = pricePerLitre;
      APP_STATE.settings.pricePerPav = pricePerPav;
      APP_STATE.settings.pavPerLitre = pavPerLitre;
      APP_STATE.settings.openingBalance = openingBalance;

      await saveSettingToDB("pricePerLitre", pricePerLitre);
      await saveSettingToDB("pricePerPav", pricePerPav);
      await saveSettingToDB("pavPerLitre", pavPerLitre);
      await saveSettingToDB("openingBalance", openingBalance);

      APP_STATE.entries = recalculateEditedLedger(APP_STATE.entries, true);
      await replaceEntriesInDB(APP_STATE.entries);
      showToast(
        "सेटिंग्स सुरक्षित हुईं व नई दरों से हिसाब अद्यतन हुआ",
        "success",
      );
      renderTodayOverview();
      renderDashboard();
    });
  }

  addSafeListener("backupJsonBtn", "click", () => {
    const dataStr =
      "data:text/json;charset=utf-8," +
      encodeURIComponent(JSON.stringify(APP_STATE, null, 2));
    const a = document.createElement("a");
    a.href = dataStr;
    a.download = `DoodhHisab_Backup_${getIndiaDateString()}.json`;
    a.click();
    showToast("JSON बैकअप फ़ाइल डाउनलोड हो गई", "success");
  });

  addSafeListener("loadSampleDataBtn", "click", async () => {
    showConfirmModal(
      "डिफ़ॉल्ट डेटा लोड करें",
      "यह डेटाबेस में 20 से 30 अक्टूबर 2026 का मानक डेटा रीसेट कर देगा। जारी रखें?",
      async () => {
        const seeds = generateSeedData();
        APP_STATE.entries = recalculateLedger(seeds);
        await replaceEntriesInDB(APP_STATE.entries);
        showToast("नमूना परीक्षण डेटा लोड हो गया", "success");
        renderDashboard();
        runDateRangeCalculator();
      },
    );
  });

  addSafeListener("clearDatabaseBtn", "click", () => {
    showConfirmModal(
      "संपूर्ण डेटा मिटाएं",
      "चेतावनी: यह सारा डेटा मिटा देगा! क्या आप वाकई आगे बढ़ना चाहते हैं?",
      async () => {
        await clearAllDataFromDB();
        APP_STATE.entries = [];
        APP_STATE.settings = {
          pricePerLitre: 33,
          pricePerPav: 8,
          pavPerLitre: 4,
          openingBalance: 0,
          currency: "₹",
        };
        document.getElementById("settingRatePerLitre").value = 33;
        document.getElementById("settingRatePerPav").value = 8;
        document.getElementById("settingPavPerLitre").value = 4;
        document.getElementById("settingOpeningBalance").value = 0;
        applySystemTheme();
        showToast("सभी रिकॉर्ड और सेटिंग्स साफ़ कर दी गईं", "warn");
        renderTodayOverview();
        renderDashboard();
        renderMasterRecords();
        renderPaymentsLedger();
        renderReportsView();
        runDateRangeCalculator();
      },
    );
  });

  // Follow the operating system / device appearance preference.
  const colorSchemeQuery = window.matchMedia("(prefers-color-scheme: dark)");
  function applySystemTheme() {
    document.documentElement.classList.toggle("dark", colorSchemeQuery.matches);
  }
  applySystemTheme();
  if (typeof colorSchemeQuery.addEventListener === "function") {
    colorSchemeQuery.addEventListener("change", applySystemTheme);
  } else if (typeof colorSchemeQuery.addListener === "function") {
    // Older Safari support.
    colorSchemeQuery.addListener(applySystemTheme);
  }

  const footerYear = document.getElementById("footerYear");
  if (footerYear) footerYear.textContent = new Date().getFullYear();

  // Initial View Setup
  renderDashboard();
  runDateRangeCalculator();
  fetchFromGoogleSheet();
});

if ("serviceWorker" in navigator && location.protocol !== "file:") {
  window.addEventListener("load", () => {
    navigator.serviceWorker
      .register("./service-worker.js")
      .catch((error) =>
        console.error("Service worker registration failed:", error),
      );
  });
}
