// WhatsApp Follow-up Dashboard
// Local frontend version

const $ = (id) => document.getElementById(id);

let customers = [];
let templates = [];
let selectedId = null;
let activeFilter = "all";
let currentPage = 1;
const pageSize = 10;

let campaignTimer = null;
let campaignRunning = false;
let campaignPaused = false;
let lastOpenedId = null;

const STORAGE_KEY = "wa_followup_dashboard_v1";

const defaultTemplates = [
  {
    name: "Payment Reminder",
    message: "Hello sir,\n\nThis is a gentle reminder regarding your outstanding amount of ₹{Amount}.\n\nKindly clear the outstanding amount and confirm once the payment has been completed.\n\nThank you for your cooperation."
  },
  {
    name: "PTP Follow-up",
    message: "Hello sir,\n\nAs discussed, you had committed to making the payment today. Kindly confirm whether the payment has been completed.\n\nThank you."
  },
  {
    name: "Payment Confirmation",
    message: "Hello sir,\n\nPlease confirm once your payment has been completed.\n\nThank you."
  }
];

function showToast(message) {
  const toast = $("toast");
  toast.textContent = message;
  toast.classList.add("show");

  clearTimeout(showToast.timeout);
  showToast.timeout = setTimeout(() => {
    toast.classList.remove("show");
  }, 3500);
}

function saveData() {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify({
      customers,
      templates,
      selectedId
    }));
  } catch (error) {
    showToast("Unable to save data in this browser.");
  }
}

function loadData() {
  try {
    const saved = JSON.parse(localStorage.getItem(STORAGE_KEY));
    if (saved) {
      customers = Array.isArray(saved.customers) ? saved.customers : [];
      templates = Array.isArray(saved.templates) ? saved.templates : [];
      selectedId = saved.selectedId ?? null;
    }
  } catch (error) {
    customers = [];
    templates = [];
  }

  if (!templates.length) {
    templates = [...defaultTemplates];
  }
}

function makeId() {
  return Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
}

function cleanPhone(value) {
  return String(value ?? "").replace(/\D/g, "");
}

function formatAmount(value) {
  if (value === "" || value === null || value === undefined) return "—";

  const number = Number(value);
  if (!Number.isFinite(number)) return String(value);

  return number.toLocaleString("en-IN", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  });
}

function parseExcelDate(value) {
  if (value === null || value === undefined || value === "") return "";

  if (value instanceof Date && !isNaN(value)) {
    return `${value.getFullYear()}-${String(value.getMonth() + 1).padStart(2, "0")}-${String(value.getDate()).padStart(2, "0")}`;
  }

  if (typeof value === "number" && window.XLSX) {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (parsed) {
      return `${parsed.y}-${String(parsed.m).padStart(2, "0")}-${String(parsed.d).padStart(2, "0")}`;
    }
  }

  const text = String(value).trim();

  // Supports DD-MM-YYYY, DD/MM/YYYY and YYYY-MM-DD
  let match = text.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{4})$/);
  if (match) {
    return `${match[3]}-${String(match[2]).padStart(2, "0")}-${String(match[1]).padStart(2, "0")}`;
  }

  match = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (match) {
    return `${match[1]}-${String(match[2]).padStart(2, "0")}-${String(match[3]).padStart(2, "0")}`;
  }

  const date = new Date(text);
  if (!isNaN(date)) {
    return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`;
  }

  return "";
}

function parseExcelTime(value) {
  if (value === null || value === undefined || value === "") return "";

  if (typeof value === "number") {
    const totalMinutes = Math.round(value * 24 * 60);
    const hours = Math.floor(totalMinutes / 60) % 24;
    const minutes = totalMinutes % 60;
    return `${String(hours).padStart(2, "0")}:${String(minutes).padStart(2, "0")}`;
  }

  const text = String(value).trim();
  let match = text.match(/^(\d{1,2}):(\d{2})(?::\d{2})?\s*(AM|PM)?$/i);

  if (!match) return "";

  let hours = Number(match[1]);
  const minutes = Number(match[2]);
  const meridiem = (match[3] || "").toUpperCase();

  if (meridiem === "PM" && hours < 12) hours += 12;
  if (meridiem === "AM" && hours === 12) hours = 0;

  if (hours > 23 || minutes > 59) return "";

  return `${String(hours).padStart(2, "0")}:${String(minutes).padStart(2, "0")}`;
}

function formatDate(dateValue) {
  if (!dateValue) return "—";
  const [year, month, day] = dateValue.split("-");
  return `${day} ${new Date(Number(year), Number(month) - 1).toLocaleString("en-IN", { month: "short" })} ${year}`;
}

function formatTime(timeValue) {
  if (!timeValue) return "—";
  const [hours, minutes] = timeValue.split(":").map(Number);
  const suffix = hours >= 12 ? "PM" : "AM";
  const h = hours % 12 || 12;
  return `${String(h).padStart(2, "0")}:${String(minutes).padStart(2, "0")} ${suffix}`;
}

function findTemplate(name) {
  return templates.find(t => t.name.toLowerCase() === String(name || "").trim().toLowerCase());
}

function buildMessage(customer) {
  const template = findTemplate(customer.template);
  if (!template) return "No matching template selected.";

  let message = template.message;
  const amount = customer.amount;

  message = message.replace(/\{Amount\}/gi, amount === "" ? "[AMOUNT REQUIRED]" : `₹${formatAmount(amount)}`);
  message = message.replace(/\{Mobile\}/gi, customer.mobile || "");
  message = message.replace(/\{Date\}/gi, formatDate(customer.date) === "—" ? "" : formatDate(customer.date));
  message = message.replace(/\{Time\}/gi, formatTime(customer.time) === "—" ? "" : formatTime(customer.time));

  return message;
}

function getStatus(customer) {
  if (customer.status === "Completed") return "completed";
  if (customer.status === "Skipped") return "skipped";
  if (customer.status === "Opened") return "opened";
  if (customer.status === "Pending") {
    if (!customer.date || !customer.time) return "pending";

    const dueAt = new Date(`${customer.date}T${customer.time}:00`);
    if (dueAt.getTime() <= Date.now()) return "due";
    return "scheduled";
  }
  return "pending";
}

function getStatusLabel(customer) {
  const status = getStatus(customer);
  const labels = {
    due: "Due Now",
    scheduled: "Scheduled",
    completed: "Completed",
    skipped: "Skipped",
    opened: "Opened",
    pending: "Unscheduled"
  };
  return labels[status] || "Pending";
}

function isEligible(customer) {
  if (customer.status === "Completed" || customer.status === "Skipped") return false;

  // No date/time means normal queue.
  if (!customer.date || !customer.time) return true;

  const dueAt = new Date(`${customer.date}T${customer.time}:00`);
  return dueAt.getTime() <= Date.now();
}

function getCounts() {
  const due = customers.filter(c => getStatus(c) === "due").length;
  const scheduled = customers.filter(c => getStatus(c) === "scheduled").length;
  const completed = customers.filter(c => c.status === "Completed").length;
  const skipped = customers.filter(c => c.status === "Skipped").length;

  return {
    total: customers.length,
    due,
    scheduled,
    completed,
    skipped
  };
}

function updateStats() {
  const c = getCounts();

  $("totalCount").textContent = c.total;
  $("scheduledCount").textContent = c.scheduled;
  $("dueCount").textContent = c.due;
  $("completedCount").textContent = c.completed;
  $("skippedCount").textContent = c.skipped;

  $("allTabCount").textContent = c.total;
  $("dueTabCount").textContent = c.due;
  $("scheduledTabCount").textContent = c.scheduled;
  $("completedTabCount").textContent = c.completed;
  $("skippedTabCount").textContent = c.skipped;

  $("campaignQueueSummary").textContent =
    `${c.total} total · ${c.due} due · ${c.scheduled} scheduled · ${c.completed} completed · ${c.skipped} skipped`;

  $("reportSummary").textContent =
    `Total: ${c.total}, completed: ${c.completed}, skipped: ${c.skipped}, scheduled: ${c.scheduled}, due now: ${c.due}`;
}

function getFilteredCustomers() {
  const search = $("searchInput").value.trim().toLowerCase();

  return customers.filter(customer => {
    const status = getStatus(customer);

    const matchesSearch =
      customer.mobile.toLowerCase().includes(search) ||
      String(customer.template || "").toLowerCase().includes(search);

    let matchesFilter = true;

    if (activeFilter === "due") matchesFilter = status === "due";
    if (activeFilter === "scheduled") matchesFilter = status === "scheduled";
    if (activeFilter === "completed") matchesFilter = status === "completed";
    if (activeFilter === "skipped") matchesFilter = status === "skipped";

    return matchesSearch && matchesFilter;
  });
}

function renderTable() {
  const tbody = $("customerTableBody");
  const filtered = getFilteredCustomers();
  const maxPages = Math.max(1, Math.ceil(filtered.length / pageSize));

  currentPage = Math.min(currentPage, maxPages);
  const start = (currentPage - 1) * pageSize;
  const visible = filtered.slice(start, start + pageSize);

  tbody.innerHTML = "";

  visible.forEach(customer => {
    const status = getStatus(customer);
    const tr = document.createElement("tr");

    const checkboxCell = document.createElement("td");
    const checkbox = document.createElement("input");
    checkbox.type = "checkbox";
    checkbox.dataset.id = customer.id;
    checkboxCell.appendChild(checkbox);
    tr.appendChild(checkboxCell);

    const values = [
      customer.mobile,
      formatAmount(customer.amount),
      customer.template || "—",
      formatDate(customer.date),
      formatTime(customer.time)
    ];

    values.forEach((value, index) => {
      const td = document.createElement("td");
      td.textContent = value;

      if (index === 2) {
        const tag = document.createElement("span");
        tag.className = "template-tag";
        tag.textContent = value;
        td.textContent = "";
        td.appendChild(tag);
      }

      tr.appendChild(td);
    });

    const statusCell = document.createElement("td");
    const pill = document.createElement("span");
    pill.className = `status-pill ${status}`;
    pill.textContent = getStatusLabel(customer);
    statusCell.appendChild(pill);
    tr.appendChild(statusCell);

    const actionCell = document.createElement("td");
    const openButton = document.createElement("button");
    openButton.className = "table-open-button";
    openButton.textContent = "◉ Open WhatsApp";
    openButton.addEventListener("click", () => {
      selectCustomer(customer.id);
      openWhatsApp(customer);
    });

    const moreButton = document.createElement("button");
    moreButton.className = "icon-button";
    moreButton.textContent = "⋮";
    moreButton.title = "Mark as skipped";
    moreButton.addEventListener("click", () => {
      if (confirm(`Mark ${customer.mobile} as skipped?`)) {
        customer.status = "Skipped";
        saveData();
        renderAll();
      }
    });

    actionCell.append(openButton, moreButton);
    tr.appendChild(actionCell);

    tr.addEventListener("click", (event) => {
      if (event.target.closest("button") || event.target.type === "checkbox") return;
      selectCustomer(customer.id);
    });

    tbody.appendChild(tr);
  });

  $("tableSummary").textContent = filtered.length
    ? `Showing ${start + 1} - ${Math.min(start + pageSize, filtered.length)} of ${filtered.length} customers`
    : "No customers found";

  renderPagination(maxPages);
}

function renderPagination(maxPages) {
  const container = $("pagination");
  container.innerHTML = "";

  const prev = document.createElement("button");
  prev.textContent = "‹";
  prev.disabled = currentPage <= 1;
  prev.onclick = () => {
    currentPage--;
    renderTable();
  };
  container.appendChild(prev);

  for (let page = 1; page <= maxPages; page++) {
    if (maxPages > 7 && page > 3 && page < maxPages - 2 && Math.abs(page - currentPage) > 1) {
      if (page === 4) {
        const dots = document.createElement("span");
        dots.textContent = "…";
        container.appendChild(dots);
      }
      continue;
    }

    const btn = document.createElement("button");
    btn.textContent = page;
    if (page === currentPage) btn.classList.add("active");
    btn.onclick = () => {
      currentPage = page;
      renderTable();
    };
    container.appendChild(btn);
  }

  const next = document.createElement("button");
  next.textContent = "›";
  next.disabled = currentPage >= maxPages;
  next.onclick = () => {
    currentPage++;
    renderTable();
  };
  container.appendChild(next);
}

function selectCustomer(id) {
  selectedId = id;
  const customer = customers.find(c => c.id === id);
  if (!customer) return;

  $("selectedMobile").textContent = customer.mobile;
  $("selectedAmount").textContent = customer.amount === "" ? "—" : `₹ ${formatAmount(customer.amount)}`;
  $("selectedTemplate").textContent = customer.template || "No template";
  $("selectedSchedule").textContent =
    customer.date && customer.time
      ? `${formatDate(customer.date)} | ${formatTime(customer.time)}`
      : "Unscheduled";

  $("selectedStatus").textContent = getStatusLabel(customer);
  $("selectedStatus").className = `status-pill ${getStatus(customer)}`;
  $("messagePreview").textContent = buildMessage(customer);

  saveData();
}

function openWhatsApp(customer) {
  if (!customer) {
    showToast("Please select a customer first.");
    return;
  }

  const phone = cleanPhone(customer.mobile);
  if (!phone || phone.length < 10) {
    showToast("Invalid mobile number.");
    return;
  }

  const template = findTemplate(customer.template);
  if (!template) {
    showToast("Select a valid template for this customer.");
    return;
  }

  if (/\{Amount\}/i.test(template.message) && (customer.amount === "" || customer.amount === null)) {
    showToast("This template requires an amount. Please add it first.");
    return;
  }

  const message = buildMessage(customer);
  const url = `https://wa.me/${phone}?text=${encodeURIComponent(message)}`;

  // Open from a user action where possible.
  const tab = window.open(url, "_blank");

  if (!tab) {
    showToast("Popup blocked. Allow popups or use the Open WhatsApp button.");
    return false;
  }

  customer.status = "Opened";
  customer.openedAt = new Date().toISOString();
  lastOpenedId = customer.id;

  saveData();
  renderAll();
  selectCustomer(customer.id);
  showToast("WhatsApp chat opened. Review and send the message manually.");
  return true;
}

function markSelectedAsSent() {
  const customer = customers.find(c => c.id === selectedId);
  if (!customer) {
    showToast("Select a customer first.");
    return;
  }

  customer.status = "Completed";
  customer.completedAt = new Date().toISOString();

  saveData();
  renderAll();
  selectCustomer(customer.id);
  showToast("Marked as sent-confirmed.");
}

function getNextEligibleCustomer() {
  const eligible = customers.filter(isEligible);

  eligible.sort((a, b) => {
    const aTime = a.date && a.time ? new Date(`${a.date}T${a.time}:00`).getTime() : 0;
    const bTime = b.date && b.time ? new Date(`${b.date}T${b.time}:00`).getTime() : 0;
    return aTime - bTime;
  });

  return eligible[0] || null;
}

function campaignTick() {
  if (!campaignRunning || campaignPaused) return;

  const next = getNextEligibleCustomer();

  if (!next) {
    stopCampaign();
    showToast("No eligible customers remaining.");
    return;
  }

  selectCustomer(next.id);

  // Automatic opening may be blocked by browser popup policies.
  const success = openWhatsApp(next);

  if (!success) {
    $("campaignStatus").textContent =
      "Waiting for manual action. Browser may have blocked automatic opening.";
    pauseCampaign();
  }
}

function startCampaign() {
  const interval = Number($("intervalInput").value);

  if (!Number.isFinite(interval) || interval < 10 || interval > 3600) {
    showToast("Set an interval between 10 and 3600 seconds.");
    return;
  }

  if (campaignRunning && campaignPaused) {
    campaignPaused = false;
    $("campaignStatus").textContent = "Campaign resumed";
    showToast("Campaign resumed.");
    return;
  }

  if (campaignRunning) {
    showToast("Campaign is already running.");
    return;
  }

  campaignRunning = true;
  campaignPaused = false;
  $("campaignStatus").textContent = "Campaign running";

  // First opening is triggered by the Start button click.
  campaignTick();

  clearInterval(campaignTimer);
  campaignTimer = setInterval(campaignTick, interval * 1000);

  showToast("Campaign started. Browser may block automatic future tab openings.");
}

function pauseCampaign() {
  if (!campaignRunning) return;
  campaignPaused = true;
  $("campaignStatus").textContent = "Campaign paused";
  showToast("Campaign paused.");
}

function stopCampaign() {
  campaignRunning = false;
  campaignPaused = false;
  clearInterval(campaignTimer);
  campaignTimer = null;
  $("campaignStatus").textContent = "Campaign stopped";
}

function readSheetRows(workbook, sheetName) {
  const sheet = workbook.Sheets[sheetName];
  if (!sheet) return [];
  return XLSX.utils.sheet_to_json(sheet, { defval: "" });
}

function normalizeKey(key) {
  return String(key).trim().toLowerCase().replace(/[\s_-]+/g, "");
}

function findValue(row, aliases) {
  const keys = Object.keys(row);
  for (const alias of aliases) {
    const found = keys.find(k => normalizeKey(k) === normalizeKey(alias));
    if (found !== undefined) return row[found];
  }
  return "";
}

function importWorkbook(file) {
  if (!window.XLSX) {
    showToast("Excel library could not load. Check your internet connection.");
    return;
  }

  const reader = new FileReader();

  reader.onload = (event) => {
    try {
      const workbook = XLSX.read(event.target.result, { type: "array", cellDates: true });

      const workflowRows = readSheetRows(workbook, "Message Workflow");
      const templateRows = readSheetRows(workbook, "Templates");
      const customerRows = readSheetRows(workbook, "Customer List");

      if (!workflowRows.length && !customerRows.length) {
        showToast("No usable rows found. Check the required sheet names.");
        return;
      }

      if (templateRows.length) {
        templates = templateRows.map(row => ({
          name: String(findValue(row, ["Template Name", "Template"]) || "").trim(),
          message: String(findValue(row, ["Message", "Template Message"]) || "").trim()
        })).filter(t => t.name && t.message);
      }

      const sourceRows = workflowRows.length ? workflowRows : customerRows;

      const imported = sourceRows.map(row => {
        const mobile = cleanPhone(findValue(row, ["Mobile Number", "Mobile", "Phone", "Phone Number"]));
        const amount = findValue(row, ["Amount", "Outstanding Amount", "Bare Minimum"]);
        const template = String(findValue(row, ["Template", "Template Name"]) || "").trim();
        const date = parseExcelDate(findValue(row, ["Follow-up Date", "Follow Up Date", "Date"]));
        const time = parseExcelTime(findValue(row, ["Follow-up Time", "Follow Up Time", "Time"]));
        const rawStatus = String(findValue(row, ["Status"]) || "Pending").trim();

        return {
          id: makeId(),
          mobile,
          amount: amount === "" ? "" : Number(amount),
          template,
          date,
          time,
          status: ["Completed", "Skipped", "Opened"].includes(rawStatus) ? rawStatus : "Pending"
        };
      }).filter(c => c.mobile);

      if (!imported.length) {
        showToast("No valid mobile numbers found in the workbook.");
        return;
      }

      const replace = confirm(
        `Found ${imported.length} workflow records.\n\nOK = Replace current customer list\nCancel = Merge and skip duplicate mobile/date/template rows`
      );

      if (replace) {
        customers = imported;
      } else {
        const existing = new Set(customers.map(c => `${c.mobile}|${c.date}|${c.time}|${c.template}`));
        const newRows = imported.filter(c => !existing.has(`${c.mobile}|${c.date}|${c.time}|${c.template}`));
        customers.push(...newRows);
      }

      currentPage = 1;
      selectedId = customers[0]?.id || null;
      saveData();
      renderAll();

      if (selectedId) selectCustomer(selectedId);
      showToast(`Imported ${imported.length} workflow rows.`);
    } catch (error) {
      console.error(error);
      showToast("Unable to read workbook. Verify the Excel sheet names and format.");
    }
  };

  reader.readAsArrayBuffer(file);
}

function importTemplateFile(file) {
  if (!window.XLSX) {
    showToast("Excel library could not load.");
    return;
  }

  const reader = new FileReader();

  reader.onload = (event) => {
    try {
      const workbook = XLSX.read(event.target.result, { type: "array" });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });

      const incoming = rows.map(row => ({
        name: String(findValue(row, ["Template Name", "Template"]) || "").trim(),
        message: String(findValue(row, ["Message", "Template Message"]) || "").trim()
      })).filter(t => t.name && t.message);

      if (!incoming.length) {
        showToast("No valid templates found.");
        return;
      }

      const replace = confirm(
        `Found ${incoming.length} templates.\n\nOK = Replace template library\nCancel = Merge with existing templates`
      );

      if (replace) {
        templates = incoming;
      } else {
        incoming.forEach(item => {
          const index = templates.findIndex(t => t.name.toLowerCase() === item.name.toLowerCase());
          if (index >= 0) {
            const update = confirm(`Template "${item.name}" already exists. Replace it?`);
            if (update) templates[index] = item;
          } else {
            templates.push(item);
          }
        });
      }

      saveData();
      renderTemplates();
      renderAll();
      showToast(`${incoming.length} templates processed.`);
    } catch (error) {
      console.error(error);
      showToast("Could not import templates.");
    }
  };

  reader.readAsArrayBuffer(file);
}

function renderTemplates() {
  const tbody = $("templateTableBody");
  tbody.innerHTML = "";

  templates.forEach(template => {
    const tr = document.createElement("tr");
    const name = document.createElement("td");
    const message = document.createElement("td");

    name.textContent = template.name;
    message.textContent = template.message;

    tr.append(name, message);
    tbody.appendChild(tr);
  });
}

function exportWorkflowCsv() {
  const headers = ["Mobile Number", "Amount", "Template", "Follow-up Date", "Follow-up Time", "Status"];
  const rows = customers.map(c => [
    c.mobile,
    c.amount,
    c.template,
    c.date,
    c.time,
    c.status
  ]);

  const csv = [headers, ...rows]
    .map(row => row.map(value => `"${String(value ?? "").replace(/"/g, '""')}"`).join(","))
    .join("\r\n");

  const blob = new Blob(["\ufeff" + csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "WhatsApp_Message_Workflow.csv";
  a.click();
  URL.revokeObjectURL(url);
  showToast("Workflow CSV exported.");
}

function updateClock() {
  const now = new Date();

  $("currentDate").textContent = now.toLocaleDateString("en-IN", {
    day: "2-digit",
    month: "short",
    year: "numeric"
  });

  $("currentTime").textContent = now.toLocaleTimeString("en-IN", {
    hour: "2-digit",
    minute: "2-digit"
  });
}

function showPage(pageName) {
  document.querySelectorAll(".page-content").forEach(page => page.classList.add("hidden"));
  const target = $(`${pageName}Page`);
  if (target) target.classList.remove("hidden");

  document.querySelectorAll(".nav-item").forEach(button => {
    button.classList.toggle("active", button.dataset.page === pageName);
  });
}

function renderAll() {
  updateStats();
  renderTable();
  renderTemplates();

  const selected = customers.find(c => c.id === selectedId);
  if (selected) selectCustomer(selected.id);
}

function bindEvents() {
  document.querySelectorAll(".nav-item").forEach(button => {
    button.addEventListener("click", () => showPage(button.dataset.page));
  });

  document.querySelectorAll(".filter-tab").forEach(button => {
    button.addEventListener("click", () => {
      activeFilter = button.dataset.filter;
      currentPage = 1;

      document.querySelectorAll(".filter-tab").forEach(tab => tab.classList.remove("active"));
      button.classList.add("active");

      renderTable();
    });
  });

  $("searchInput").addEventListener("input", () => {
    currentPage = 1;
    renderTable();
  });

  $("workbookInput").addEventListener("change", event => {
    const file = event.target.files[0];
    if (file) importWorkbook(file);
    event.target.value = "";
  });

  $("workbookInput2").addEventListener("change", event => {
    const file = event.target.files[0];
    if (file) importWorkbook(file);
    event.target.value = "";
  });

  $("templateInput").addEventListener("change", event => {
    const file = event.target.files[0];
    if (file) importTemplateFile(file);
    event.target.value = "";
  });

  $("openSelected").addEventListener("click", () => {
    const customer = customers.find(c => c.id === selectedId);
    openWhatsApp(customer);
  });

  $("markSent").addEventListener("click", markSelectedAsSent);
  $("startCampaign").addEventListener("click", startCampaign);
  $("campaignStartAlt").addEventListener("click", startCampaign);
  $("pauseCampaign").addEventListener("click", pauseCampaign);
  $("stopCampaign").addEventListener("click", stopCampaign);
  $("exportCsv").addEventListener("click", exportWorkflowCsv);

  $("selectAll").addEventListener("change", event => {
    document.querySelectorAll("#customerTableBody input[type=checkbox]")
      .forEach(input => input.checked = event.target.checked);
  });

  $("clearData").addEventListener("click", () => {
    if (!confirm("Clear all locally saved customers and templates?")) return;
    stopCampaign();
    customers = [];
    templates = [...defaultTemplates];
    selectedId = null;
    localStorage.removeItem(STORAGE_KEY);
    renderAll();
    $("selectedMobile").textContent = "No customer selected";
    $("selectedAmount").textContent = "—";
    $("selectedTemplate").textContent = "—";
    $("selectedSchedule").textContent = "—";
    $("messagePreview").textContent = "Select a customer to preview the message.";
    showToast("Local data cleared.");
  });
}

function init() {
  loadData();
  bindEvents();
  updateClock();
  setInterval(updateClock, 1000);
  renderAll();

  if (selectedId) selectCustomer(selectedId);

  // Refresh due status periodically while the page is open.
  setInterval(() => {
    updateStats();
    renderTable();
  }, 15000);
}

document.addEventListener("DOMContentLoaded", init);