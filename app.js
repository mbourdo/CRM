const STORAGE_KEY = "annuity-crm-leads-v1";
const AD_SPEND_STORAGE_KEY = "annuity-crm-adspend-v1";
const API_LEADS_PATH = "/api/leads";
const SUPABASE_URL = (window.CRM_CONFIG?.supabaseUrl || "").trim();
const SUPABASE_ANON_KEY = (window.CRM_CONFIG?.supabaseAnonKey || "").trim();
const SUPABASE_STATE_TABLE = "crm_state";
const SUPABASE_ENABLED = Boolean(SUPABASE_URL && SUPABASE_ANON_KEY);
const APPOINTMENT_STATUSES = ["notBooked", "bookedUnconfirmed", "bookedConfirmed"];
const LEAD_STATUSES = ["active", "inactive"];
const ACTIVITY_TYPES = [
  "callLiveContact",
  "callVM",
  "callNC",
  "textSent",
  "textReceived",
  "emailSent",
  "emailReceived",
  "dial",
  "live",
];
const FIRST_APPOINTMENT_STATUSES = ["notScheduled", "scheduled", "completed", "noShow"];
const SECOND_APPOINTMENT_STATUSES = ["notBooked", "scheduled", "completed", "noShow"];
const THIRD_APPOINTMENT_STATUSES = ["notBooked", "scheduled", "completed", "noShow"];
const DOCUMENT_STATUSES = ["notRequested", "waiting", "received"];
const CARRIER_STATUSES = ["notStarted", "contacted", "quoted"];
const TASK_ASSIGNEES = ["JAKE", "COLBY"];
const DEFAULT_TASK_ASSIGNEE = "JAKE";

const STAGES = [
  "lead",
  "firstAppointmentBooked",
  "firstAppointmentNoShow",
  "firstAppointmentRescheduled",
  "firstAppointmentConfirmed",
  "secondAppointmentBooked",
  "secondAppointmentConfirmed",
  "secondAppointmentRescheduled",
  "secondAppointmentAttended",
  "thirdAppointmentBooked",
  "thirdAppointmentConfirmed",
  "thirdAppointmentRescheduled",
  "thirdAppointmentAttended",
  "pendingWon",
  "pendingLost",
  "won",
  "lost",
  "unqualified",
  "notReady",
];

const OPEN_STAGES = [
  "lead",
  "firstAppointmentBooked",
  "firstAppointmentNoShow",
  "firstAppointmentRescheduled",
  "firstAppointmentConfirmed",
  "secondAppointmentBooked",
  "secondAppointmentConfirmed",
  "secondAppointmentRescheduled",
  "secondAppointmentAttended",
  "thirdAppointmentBooked",
  "thirdAppointmentConfirmed",
  "thirdAppointmentRescheduled",
  "thirdAppointmentAttended",
  "pendingWon",
  "pendingLost",
  "unqualified",
  "notReady",
];
const BOOKED_FLOW_STAGES = [
  "firstAppointmentBooked",
  "firstAppointmentNoShow",
  "firstAppointmentRescheduled",
  "firstAppointmentConfirmed",
  "secondAppointmentBooked",
  "secondAppointmentConfirmed",
  "secondAppointmentRescheduled",
  "secondAppointmentAttended",
  "thirdAppointmentBooked",
  "thirdAppointmentConfirmed",
  "thirdAppointmentRescheduled",
  "thirdAppointmentAttended",
  "pendingWon",
  "pendingLost",
  "won",
];
const SLA_HOURS_BY_STAGE = {
  lead: 24,
  firstAppointmentBooked: 24,
  firstAppointmentNoShow: 24,
  firstAppointmentConfirmed: 48,
  secondAppointmentBooked: 72,
  secondAppointmentConfirmed: 72,
  secondAppointmentAttended: 72,
  thirdAppointmentBooked: 72,
  thirdAppointmentConfirmed: 72,
  thirdAppointmentAttended: 72,
  pendingWon: 96,
};

const stageAdvanceMap = {
  lead: "firstAppointmentBooked",
  firstAppointmentNoShow: "firstAppointmentRescheduled",
  firstAppointmentBooked: "firstAppointmentConfirmed",
  firstAppointmentRescheduled: "firstAppointmentBooked",
  firstAppointmentConfirmed: "secondAppointmentBooked",
  secondAppointmentBooked: "secondAppointmentConfirmed",
  secondAppointmentConfirmed: "secondAppointmentAttended",
  secondAppointmentRescheduled: "secondAppointmentBooked",
  secondAppointmentAttended: "thirdAppointmentBooked",
  thirdAppointmentBooked: "thirdAppointmentConfirmed",
  thirdAppointmentConfirmed: "thirdAppointmentAttended",
  thirdAppointmentRescheduled: "thirdAppointmentBooked",
  thirdAppointmentAttended: "pendingWon",
  pendingWon: "won",
  pendingLost: "lost",
  won: "won",
  lost: "lost",
  unqualified: "unqualified",
  notReady: "notReady",
};

const CSV_COLUMNS = [
  "id",
  "firstName",
  "nickname",
  "lastName",
  "email",
  "phone",
  "state",
  "stage",
  "leadStatus",
  "appointmentStatus",
  "confirmedPhone",
  "confirmedCalendar",
  "firstConfirmedPhone",
  "firstConfirmedCalendar",
  "secondConfirmedPhone",
  "secondConfirmedCalendar",
  "thirdConfirmedPhone",
  "thirdConfirmedCalendar",
  "dialAttempts",
  "liveContacts",
  "textsAttempted",
  "texted",
  "firstAppointmentAt",
  "firstMeetingRecordingUrl",
  "firstAppointmentSummary",
  "secondAppointmentAt",
  "secondMeetingRecordingUrl",
  "secondAppointmentSummary",
  "thirdAppointmentAt",
  "thirdMeetingRecordingUrl",
  "thirdAppointmentSummary",
  "firstAppointmentStatus",
  "secondAppointmentStatus",
  "thirdAppointmentStatus",
  "documentsStatus",
  "documentsNotes",
  "carrierStatus",
  "carrierNotes",
  "firstMeetingDone",
  "infoMeetingDone",
  "secondAppointmentBooked",
  "firstAppointmentNoShow",
  "secondAppointmentNoShow",
  "thirdAppointmentNoShow",
  "waitingOnDocuments",
  "carrierReachedOut",
  "qualifiedProspectMeetingAttended",
  "opportunitySize",
  "nextAction",
  "nextAppointmentAt",
  "followUpAt",
  "closeReason",
  "closedAt",
  "stageUpdatedAt",
  "statusUpdatedAt",
  "callLogs",
  "activityLogs",
  "tasks",
  "notes",
  "createdAt",
  "updatedAt",
];

let leads = [];
let editingLeadId = null;
let draggedLeadId = null;
let shouldSeedServer = false;
let shouldSeedSupabase = false;
let draftTasks = [];
let adSpend = 0;

const dialog = document.getElementById("leadDialog");
const form = document.getElementById("leadForm");
const formTitle = document.getElementById("formTitle");
const newLeadBtn = document.getElementById("newLeadBtn");
const cancelBtn = document.getElementById("cancelBtn");
const deleteLeadBtn = document.getElementById("deleteLeadBtn");
const exportCsvBtn = document.getElementById("exportCsvBtn");
const importCsvBtn = document.getElementById("importCsvBtn");
const importCsvInput = document.getElementById("importCsvInput");
const activitySection = document.getElementById("activitySection");
const formActivityItems = document.getElementById("formActivityItems");
const formTaskItems = document.getElementById("formTaskItems");
const formCpqpma = document.getElementById("formCpqpma");
const addTaskBtn = document.getElementById("addTaskBtn");
const openFirstRecordingBtn = document.getElementById("openFirstRecordingBtn");
const openSecondRecordingBtn = document.getElementById("openSecondRecordingBtn");
const openThirdRecordingBtn = document.getElementById("openThirdRecordingBtn");
const aggregateTasksList = document.getElementById("aggregateTasksList");
const aggregateTasksCount = document.getElementById("aggregateTasksCount");
const aggregateOpenTasksList = document.getElementById("aggregateOpenTasksList");
const aggregateOpenTasksCount = document.getElementById("aggregateOpenTasksCount");
const adSpendInput = document.getElementById("adSpendInput");
const cardTemplate = document.getElementById("leadCardTemplate");

newLeadBtn.addEventListener("click", openNewLeadForm);
cancelBtn.addEventListener("click", () => dialog.close());
deleteLeadBtn.addEventListener("click", deleteCurrentLead);
exportCsvBtn.addEventListener("click", exportCSV);
importCsvBtn.addEventListener("click", () => importCsvInput.click());
importCsvInput.addEventListener("change", importCSV);
form.addEventListener("submit", saveLeadFromForm);
addTaskBtn.addEventListener("click", addTaskFromFormInput);
openFirstRecordingBtn.addEventListener("click", () => openUrlInNewTab(form.elements.firstMeetingRecordingUrl.value));
openSecondRecordingBtn.addEventListener("click", () =>
  openUrlInNewTab(form.elements.secondMeetingRecordingUrl.value)
);
openThirdRecordingBtn.addEventListener("click", () =>
  openUrlInNewTab(form.elements.thirdMeetingRecordingUrl.value)
);
adSpendInput.addEventListener("input", previewAdSpendFromInput);
adSpendInput.addEventListener("blur", saveAdSpendFromInput);
form.elements.taskInput.addEventListener("keydown", (event) => {
  if (event.key === "Enter") {
    event.preventDefault();
    addTaskFromFormInput();
  }
});
document.querySelectorAll(".dialog-activity-btn").forEach((button) => {
  button.addEventListener("click", () => {
    if (!editingLeadId) return;
    logActivity(editingLeadId, button.dataset.activityType);
    refreshDialogActivity();
  });
});
form.elements.opportunitySize.addEventListener("blur", () => {
  form.elements.opportunitySize.value = formatCurrency(parseCurrency(form.elements.opportunitySize.value));
});

init();

async function init() {
  leads = await loadLeads();
  adSpend = loadAdSpend();
  adSpendInput.value = formatCurrency(adSpend);
  render();
  if (shouldSeedSupabase) {
    persist();
    shouldSeedSupabase = false;
  }
  if (shouldSeedServer) {
    persist();
    shouldSeedServer = false;
  }
}

async function loadLeads() {
  const localLeads = loadLeadsFromLocalStorage();

  if (SUPABASE_ENABLED) {
    try {
      const remoteLeads = await loadLeadsFromSupabase();
      if (remoteLeads.length === 0 && localLeads.length > 0) {
        shouldSeedSupabase = true;
        return localLeads;
      }
      return remoteLeads;
    } catch {
      return localLeads;
    }
  }

  try {
    const response = await fetch(API_LEADS_PATH, { cache: "no-store" });
    if (!response.ok) throw new Error("API unavailable");
    const data = await response.json();
    if (!Array.isArray(data)) throw new Error("Invalid data");
    const normalized = data.map(normalizeLead);

    if (normalized.length === 0 && localLeads.length > 0) {
      shouldSeedServer = true;
      return localLeads;
    }
    return normalized;
  } catch {
    return localLeads;
  }
}

function loadLeadsFromLocalStorage() {
  const raw = localStorage.getItem(STORAGE_KEY);
  if (!raw) return [];

  try {
    const parsed = JSON.parse(raw);
    if (!Array.isArray(parsed)) return [];
    return parsed.map(normalizeLead);
  } catch {
    return [];
  }
}

function loadAdSpend() {
  const raw = localStorage.getItem(AD_SPEND_STORAGE_KEY);
  if (!raw) return 0;
  return toNumber(raw);
}

function saveAdSpendFromInput() {
  adSpend = parseCurrency(adSpendInput.value);
  localStorage.setItem(AD_SPEND_STORAGE_KEY, String(adSpend));
  adSpendInput.value = formatCurrency(adSpend);
  renderKPIs();
  renderFormCpqpma();
}

function previewAdSpendFromInput() {
  adSpend = parseCurrency(adSpendInput.value);
  renderKPIs();
  renderFormCpqpma();
}

function getCostPerQualified() {
  const qualifiedAttended = leads.filter((lead) => lead.qualifiedProspectMeetingAttended).length;
  if (qualifiedAttended <= 0) return null;
  return adSpend / qualifiedAttended;
}

async function persist() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(leads));

  if (SUPABASE_ENABLED) {
    try {
      await saveLeadsToSupabase(leads);
      return;
    } catch {
      // Fall through to local API and localStorage fallback.
    }
  }

  try {
    await fetch(API_LEADS_PATH, {
      method: "PUT",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(leads),
    });
  } catch {
    // Keep localStorage as offline fallback.
  }
}

async function loadLeadsFromSupabase() {
  const endpoint = `${SUPABASE_URL}/rest/v1/${SUPABASE_STATE_TABLE}?id=eq.1&select=leads`;
  const response = await fetch(endpoint, {
    headers: supabaseHeaders(),
    cache: "no-store",
  });
  if (!response.ok) {
    throw new Error(`Supabase read failed: ${response.status}`);
  }

  const rows = await response.json();
  if (!Array.isArray(rows) || rows.length === 0) return [];
  const leadsValue = rows[0]?.leads;
  if (!Array.isArray(leadsValue)) return [];
  return leadsValue.map(normalizeLead);
}

async function saveLeadsToSupabase(nextLeads) {
  const endpoint = `${SUPABASE_URL}/rest/v1/${SUPABASE_STATE_TABLE}`;
  const response = await fetch(endpoint, {
    method: "POST",
    headers: {
      ...supabaseHeaders(),
      "Content-Type": "application/json",
      Prefer: "resolution=merge-duplicates,return=minimal",
    },
    body: JSON.stringify({
      id: 1,
      leads: nextLeads,
      updated_at: new Date().toISOString(),
    }),
  });
  if (!response.ok) {
    throw new Error(`Supabase write failed: ${response.status}`);
  }
}

function supabaseHeaders() {
  return {
    apikey: SUPABASE_ANON_KEY,
    Authorization: `Bearer ${SUPABASE_ANON_KEY}`,
  };
}

function render() {
  STAGES.forEach((stage) => {
    const container = document.getElementById(`leadList-${stage}`);
    container.innerHTML = "";

    leads
      .filter((lead) => lead.stage === stage)
      .forEach((lead) => container.appendChild(makeLeadCard(lead)));
  });

  renderKPIs();
  renderAggregateTasks();
  setupDragAndDrop();
}

function renderKPIs() {
  const totalLeads = leads.length;
  const booked = leads.filter((lead) => lead.appointmentStatus !== "notBooked").length;
  const pipeline = leads
    .filter((lead) => OPEN_STAGES.includes(lead.stage) && lead.leadStatus === "active")
    .reduce((sum, lead) => sum + lead.opportunitySize, 0);
  const commissionOpportunity = pipeline * 0.07;
  const qualifiedAttended = leads.filter((lead) => lead.qualifiedProspectMeetingAttended).length;
  const qualifiedClosedWon = leads.filter(
    (lead) => lead.qualifiedProspectMeetingAttended && lead.stage === "won"
  ).length;
  const qualifiedCloseRate = qualifiedAttended > 0 ? (qualifiedClosedWon / qualifiedAttended) * 100 : 0;
  const costPerQualified = getCostPerQualified();
  const wonCount = leads.filter((lead) => lead.stage === "won").length;
  const noShowCount = leads.filter(
    (lead) =>
      lead.firstAppointmentStatus === "noShow" ||
      lead.secondAppointmentStatus === "noShow" ||
      lead.thirdAppointmentStatus === "noShow"
  ).length;
  const inactiveCount = leads.filter((lead) => lead.leadStatus === "inactive").length;

  document.getElementById("kpiTotalLeads").textContent = totalLeads;
  document.getElementById("kpiBooked").textContent = booked;
  document.getElementById("kpiPipeline").textContent = formatCurrency(pipeline);
  document.getElementById("kpiPipelineCommission").textContent = formatCurrency(commissionOpportunity);
  document.getElementById("kpiPipelineTop").textContent = formatCurrency(pipeline);
  document.getElementById("kpiPipelineCommissionTop").textContent = formatCurrency(commissionOpportunity);
  document.getElementById("kpiQualifiedAttended").textContent = qualifiedAttended;
  document.getElementById("kpiQualifiedCloseRate").textContent =
    `Close rate: ${qualifiedCloseRate.toFixed(1)}% (${qualifiedClosedWon}/${qualifiedAttended})`;
  document.getElementById("kpiCostPerQualifiedTop").textContent =
    costPerQualified == null ? "n/a" : formatCurrency(costPerQualified);
  document.getElementById("kpiWon").textContent = wonCount;
  document.getElementById("kpiNoShow").textContent = noShowCount;
  document.getElementById("kpiInactive").textContent = inactiveCount;
}

function renderAggregateTasks() {
  const now = Date.now();
  const upcomingItems = [];
  const openTaskItems = [];
  leads.forEach((lead) => {
    [
      { label: "1st Appointment", at: lead.firstAppointmentAt },
      { label: "2nd Appointment", at: lead.secondAppointmentAt },
      { label: "3rd Appointment", at: lead.thirdAppointmentAt },
    ].forEach((appt) => {
      if (!appt.at) return;
      const date = new Date(appt.at);
      if (Number.isNaN(date.getTime())) return;
      if (date.getTime() < now) return;
      upcomingItems.push({
        leadId: lead.id,
        leadName: fullName(lead),
        label: appt.label,
        at: appt.at,
      });
    });

    getOpenTasks(lead).forEach((task) => {
      openTaskItems.push({
        leadId: lead.id,
        taskId: task.id,
        leadName: fullName(lead),
        title: task.title,
        assignee: normalizeTaskAssignee(task.assignee),
        dueAt: task.dueAt || "",
        createdAt: task.createdAt || 0,
      });
    });
  });

  aggregateTasksCount.textContent = `${upcomingItems.length} upcoming`;
  aggregateTasksList.innerHTML = "";

  if (upcomingItems.length === 0) {
    const empty = document.createElement("p");
    empty.className = "aggregate-empty";
    empty.textContent = "No upcoming appointments.";
    aggregateTasksList.appendChild(empty);
  }

  upcomingItems
    .sort((a, b) => new Date(a.at).getTime() - new Date(b.at).getTime())
    .slice(0, 20)
    .forEach((item) => {
      const row = document.createElement("div");
      row.className = "aggregate-task-row";

      const copy = document.createElement("div");
      copy.className = "aggregate-task-copy";
      const title = document.createElement("p");
      title.className = "aggregate-task-title";
      title.textContent = `${item.label}: ${formatDateTime(item.at)}`;
      const meta = document.createElement("p");
      meta.className = "aggregate-task-meta";
      meta.textContent = item.leadName;
      copy.appendChild(title);
      copy.appendChild(meta);

      const actions = document.createElement("div");
      actions.className = "aggregate-task-actions";

      const openBtn = document.createElement("button");
      openBtn.type = "button";
      openBtn.className = "btn btn-small btn-secondary";
      openBtn.textContent = "Open";
      openBtn.addEventListener("click", () => openEditLeadForm(item.leadId));

      actions.appendChild(openBtn);
      row.appendChild(copy);
      row.appendChild(actions);
      aggregateTasksList.appendChild(row);
    });

  if (!aggregateOpenTasksCount || !aggregateOpenTasksList) return;

  aggregateOpenTasksCount.textContent = `${openTaskItems.length} open`;
  aggregateOpenTasksList.innerHTML = "";

  if (openTaskItems.length === 0) {
    const empty = document.createElement("p");
    empty.className = "aggregate-empty";
    empty.textContent = "No open tasks.";
    aggregateOpenTasksList.appendChild(empty);
    return;
  }

  openTaskItems
    .sort((a, b) => compareTasksForDisplay(a, b))
    .slice(0, 20)
    .forEach((item) => {
      const row = document.createElement("div");
      row.className = `aggregate-task-row ${taskAssigneeClass(item.assignee)}`;

      const copy = document.createElement("div");
      copy.className = "aggregate-task-copy";
      const title = document.createElement("p");
      title.className = "aggregate-task-title";
      title.textContent = `${item.title}${taskDueLabel(item.dueAt)}`;
      const meta = document.createElement("p");
      meta.className = "aggregate-task-meta";
      meta.textContent = item.leadName;
      const badge = document.createElement("span");
      badge.className = `task-assignee-badge ${taskAssigneeClass(item.assignee)}`;
      badge.textContent = item.assignee;
      copy.appendChild(title);
      copy.appendChild(meta);
      copy.appendChild(badge);

      const actions = document.createElement("div");
      actions.className = "aggregate-task-actions";

      const doneBtn = document.createElement("button");
      doneBtn.type = "button";
      doneBtn.className = "btn btn-small btn-success";
      doneBtn.textContent = "Done";
      doneBtn.addEventListener("click", () => markLeadTaskDone(item.leadId, item.taskId));

      const openBtn = document.createElement("button");
      openBtn.type = "button";
      openBtn.className = "btn btn-small btn-secondary";
      openBtn.textContent = "Open";
      openBtn.addEventListener("click", () => openEditLeadForm(item.leadId));

      actions.appendChild(doneBtn);
      actions.appendChild(openBtn);
      row.appendChild(copy);
      row.appendChild(actions);
      aggregateOpenTasksList.appendChild(row);
    });
}

function makeLeadCard(lead) {
  const node = cardTemplate.content.firstElementChild.cloneNode(true);
  const prospectNotesEl = node.querySelector(".prospect-notes");
  const nextPreviewEl = node.querySelector(".next-preview");
  const taskPreviewEl = node.querySelector(".task-preview");
  const lastTouchPreviewEl = node.querySelector(".last-touch-preview");
  node.draggable = true;
  node.dataset.leadId = lead.id;

  node.addEventListener("dragstart", (event) => {
    draggedLeadId = lead.id;
    event.dataTransfer.effectAllowed = "move";
    event.dataTransfer.setData("text/plain", lead.id);
    node.classList.add("dragging");
  });
  node.addEventListener("dragend", () => {
    draggedLeadId = null;
    node.classList.remove("dragging");
    document.querySelectorAll(".column").forEach((column) => column.classList.remove("drop-target"));
  });

  const fullLeadName = fullName(lead);
  node.querySelector(".lead-name").textContent = fullLeadName;
  node.querySelector(".lead-email").textContent = lead.email || "No email";
  node.querySelector(".lead-phone").textContent = lead.phone || "No phone";
  const leadStateEl = node.querySelector(".lead-state");
  if (lead.state) {
    leadStateEl.textContent = lead.state;
    leadStateEl.classList.remove("hidden");
  } else {
    leadStateEl.textContent = "";
    leadStateEl.classList.add("hidden");
  }

  bindCopyButton(node.querySelector(".copy-email-btn"), lead.email, "Copy email");
  bindCopyButton(node.querySelector(".copy-phone-btn"), lead.phone, "Copy phone");
  const firstAppointmentEl = node.querySelector(".first-appointment-time");
  const secondAppointmentEl = node.querySelector(".second-appointment-time");
  const thirdAppointmentEl = node.querySelector(".third-appointment-time");
  renderAppointmentTime(
    firstAppointmentEl,
    "1st Appointment",
    lead.firstAppointmentAt,
    lead.firstAppointmentStatus
  );
  renderAppointmentTime(
    secondAppointmentEl,
    "2nd Appointment",
    lead.secondAppointmentAt,
    lead.secondAppointmentStatus
  );
  renderAppointmentTime(
    thirdAppointmentEl,
    "3rd Appointment",
    lead.thirdAppointmentAt,
    lead.thirdAppointmentStatus
  );
  node.querySelector(".money").textContent = `Est. premium: ${formatCurrency(lead.opportunitySize)}`;

  const openTasks = getOpenTasks(lead);
  const topTask = openTasks.slice().sort(compareTasksForDisplay)[0];

  const appointmentPill = node.querySelector(".appointment");
  const statusPill = node.querySelector(".status");
  statusPill.textContent = `Status: ${leadStatusLabel(lead.leadStatus)}`;
  statusPill.classList.remove("pill-status-active", "pill-status-inactive");
  statusPill.classList.add(lead.leadStatus === "inactive" ? "pill-status-inactive" : "pill-status-active");
  appointmentPill.textContent = openTasks.length > 0 ? `Open Tasks: ${openTasks.length}` : "Open Tasks: 0";

  nextPreviewEl.textContent = lead.nextAction ? trim(lead.nextAction, 84) : "Not set";
  taskPreviewEl.textContent = topTask
    ? `${trim(topTask.title, 58)} (${taskDueHint(topTask.dueAt)})`
    : "No open task";
  lastTouchPreviewEl.textContent = buildLastTouchPreview(lead);
  if (lead.notes) {
    prospectNotesEl.textContent = `Context: ${trim(lead.notes, 92)}`;
    prospectNotesEl.classList.remove("hidden");
  } else {
    prospectNotesEl.textContent = "";
    prospectNotesEl.classList.add("hidden");
  }

  const sla = getSLAStatus(lead);
  if (sla.overdue) {
    node.classList.add("sla-warning");
  }

  const editBtn = node.querySelector(".edit-btn");
  const advanceBtn = node.querySelector(".advance-btn");

  editBtn.addEventListener("click", () => openEditLeadForm(lead.id));
  advanceBtn.addEventListener("click", () => advanceStage(lead.id));

  if (lead.stage === "won" || lead.stage === "lost") {
    advanceBtn.classList.add("hidden");
  }

  return node;
}

function bindCopyButton(button, value, defaultLabel) {
  if (!button) return;
  button.addEventListener("mousedown", (event) => event.stopPropagation());
  button.addEventListener("click", async (event) => {
    event.preventDefault();
    event.stopPropagation();
    const text = String(value || "").trim();
    if (!text) return;

    const copied = await copyText(text);
    button.title = copied ? "Copied!" : "Copy failed";
    button.setAttribute("aria-label", button.title);
    window.setTimeout(() => {
      button.title = defaultLabel;
      button.setAttribute("aria-label", defaultLabel);
    }, 900);
  });
}

async function copyText(text) {
  try {
    if (navigator.clipboard && window.isSecureContext) {
      await navigator.clipboard.writeText(text);
      return true;
    }
  } catch {
    // Fallback below.
  }

  try {
    const area = document.createElement("textarea");
    area.value = text;
    area.setAttribute("readonly", "");
    area.style.position = "fixed";
    area.style.opacity = "0";
    document.body.appendChild(area);
    area.focus();
    area.select();
    const ok = document.execCommand("copy");
    document.body.removeChild(area);
    return ok;
  } catch {
    return false;
  }
}

function taskDueHint(value) {
  const normalized = normalizeDateOnlyString(value);
  if (!normalized) return "No due date";

  const dueDate = new Date(`${normalized}T23:59:59`);
  if (Number.isNaN(dueDate.getTime())) return `Due ${formatDateOnly(normalized)}`;

  const now = new Date();
  const todayStart = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const tomorrowStart = new Date(now.getFullYear(), now.getMonth(), now.getDate() + 1);

  if (dueDate < todayStart) return "Overdue";
  if (dueDate < tomorrowStart) return "Due today";
  return `Due ${formatDateOnly(normalized)}`;
}

function buildLastTouchPreview(lead) {
  const last = Array.isArray(lead.activityLogs) ? lead.activityLogs[0] : null;
  if (!last || !last.at || !last.type) return "No activity logged";
  return `${activityTypeLabel(last.type)} - ${formatDateTime(last.at)}`;
}

function openUrlInNewTab(rawUrl) {
  const trimmed = String(rawUrl || "").trim();
  if (!trimmed) return;
  const url = /^(https?:)?\/\//i.test(trimmed) ? trimmed : `https://${trimmed}`;
  try {
    window.open(url, "_blank", "noopener,noreferrer");
  } catch {
    // no-op
  }
}

function openNewLeadForm() {
  editingLeadId = null;
  formTitle.textContent = "New Lead";
  form.reset();
  form.elements.stage.value = "lead";
  form.elements.opportunitySize.value = "$0";
  form.elements.leadStatus.value = "active";
  form.elements.firstAppointmentStatus.value = "notScheduled";
  form.elements.secondAppointmentStatus.value = "notBooked";
  form.elements.thirdAppointmentStatus.value = "notBooked";
  form.elements.firstConfirmedPhone.checked = false;
  form.elements.firstConfirmedCalendar.checked = false;
  form.elements.secondConfirmedPhone.checked = false;
  form.elements.secondConfirmedCalendar.checked = false;
  form.elements.thirdConfirmedPhone.checked = false;
  form.elements.thirdConfirmedCalendar.checked = false;
  form.elements.documentsStatus.value = "notRequested";
  form.elements.documentsNotes.value = "";
  form.elements.carrierStatus.value = "notStarted";
  form.elements.carrierNotes.value = "";
  form.elements.firstAppointmentAt.value = "";
  form.elements.firstMeetingRecordingUrl.value = "";
  form.elements.firstAppointmentSummary.value = "";
  form.elements.secondAppointmentAt.value = "";
  form.elements.secondMeetingRecordingUrl.value = "";
  form.elements.secondAppointmentSummary.value = "";
  form.elements.thirdAppointmentAt.value = "";
  form.elements.thirdMeetingRecordingUrl.value = "";
  form.elements.thirdAppointmentSummary.value = "";
  form.elements.state.value = "";
  activitySection.classList.add("hidden");
  formActivityItems.innerHTML = "";
  form.elements.nickname.value = "";
  form.elements.nextAction.value = "";
  form.elements.closeReason.value = "";
  form.elements.notes.value = "";
  form.elements.taskInput.value = "";
  form.elements.taskDueAt.value = "";
  form.elements.taskAssignee.value = DEFAULT_TASK_ASSIGNEE;
  form.elements.qualifiedProspectMeetingAttended.checked = false;
  draftTasks = [];
  renderFormTasks();
  renderFormCpqpma();
  deleteLeadBtn.classList.add("hidden");
  dialog.showModal();
}

function openEditLeadForm(id) {
  const lead = leads.find((item) => item.id === id);
  if (!lead) return;

  editingLeadId = id;
  formTitle.textContent = `Edit Lead: ${fullName(lead)}`;
  deleteLeadBtn.classList.remove("hidden");

  form.elements.firstName.value = lead.firstName;
  form.elements.nickname.value = lead.nickname;
  form.elements.lastName.value = lead.lastName;
  form.elements.email.value = lead.email;
  form.elements.phone.value = lead.phone;
  form.elements.state.value = lead.state || "";
  form.elements.stage.value = lead.stage;
  form.elements.leadStatus.value = lead.leadStatus;
  form.elements.opportunitySize.value = formatCurrency(lead.opportunitySize);
  form.elements.firstConfirmedPhone.checked = lead.firstConfirmedPhone;
  form.elements.firstConfirmedCalendar.checked = lead.firstConfirmedCalendar;
  form.elements.secondConfirmedPhone.checked = lead.secondConfirmedPhone;
  form.elements.secondConfirmedCalendar.checked = lead.secondConfirmedCalendar;
  form.elements.thirdConfirmedPhone.checked = lead.thirdConfirmedPhone;
  form.elements.thirdConfirmedCalendar.checked = lead.thirdConfirmedCalendar;
  form.elements.firstAppointmentStatus.value = lead.firstAppointmentStatus;
  form.elements.secondAppointmentStatus.value = lead.secondAppointmentStatus;
  form.elements.thirdAppointmentStatus.value = lead.thirdAppointmentStatus;
  form.elements.documentsStatus.value = lead.documentsStatus;
  form.elements.documentsNotes.value = lead.documentsNotes || "";
  form.elements.carrierStatus.value = lead.carrierStatus;
  form.elements.carrierNotes.value = lead.carrierNotes || "";
  form.elements.firstAppointmentAt.value = toDateTimeLocalValue(lead.firstAppointmentAt);
  form.elements.firstMeetingRecordingUrl.value = lead.firstMeetingRecordingUrl || "";
  form.elements.firstAppointmentSummary.value = lead.firstAppointmentSummary || "";
  form.elements.secondAppointmentAt.value = toDateTimeLocalValue(lead.secondAppointmentAt);
  form.elements.secondMeetingRecordingUrl.value = lead.secondMeetingRecordingUrl || "";
  form.elements.secondAppointmentSummary.value = lead.secondAppointmentSummary || "";
  form.elements.thirdAppointmentAt.value = toDateTimeLocalValue(lead.thirdAppointmentAt);
  form.elements.thirdMeetingRecordingUrl.value = lead.thirdMeetingRecordingUrl || "";
  form.elements.thirdAppointmentSummary.value = lead.thirdAppointmentSummary || "";
  activitySection.classList.remove("hidden");
  form.elements.nextAction.value = lead.nextAction;
  form.elements.closeReason.value = lead.closeReason;
  form.elements.notes.value = lead.notes || "";
  form.elements.qualifiedProspectMeetingAttended.checked = Boolean(
    lead.qualifiedProspectMeetingAttended
  );
  form.elements.taskInput.value = "";
  form.elements.taskDueAt.value = "";
  form.elements.taskAssignee.value = DEFAULT_TASK_ASSIGNEE;
  draftTasks = normalizeTasks(lead.tasks || []);
  renderFormTasks();
  renderFormCpqpma();

  dialog.showModal();
  refreshDialogActivity();
}

function saveLeadFromForm(event) {
  event.preventDefault();

  const now = Date.now();
  const rawPayload = normalizeLead({
    firstName: form.elements.firstName.value.trim(),
    nickname: form.elements.nickname.value.trim(),
    lastName: form.elements.lastName.value.trim(),
    email: form.elements.email.value.trim(),
    phone: form.elements.phone.value.trim(),
    state: form.elements.state.value.trim().toUpperCase(),
    stage: form.elements.stage.value,
    leadStatus: form.elements.leadStatus.value,
    opportunitySize: parseCurrency(form.elements.opportunitySize.value),
    firstConfirmedPhone: form.elements.firstConfirmedPhone.checked,
    firstConfirmedCalendar: form.elements.firstConfirmedCalendar.checked,
    secondConfirmedPhone: form.elements.secondConfirmedPhone.checked,
    secondConfirmedCalendar: form.elements.secondConfirmedCalendar.checked,
    thirdConfirmedPhone: form.elements.thirdConfirmedPhone.checked,
    thirdConfirmedCalendar: form.elements.thirdConfirmedCalendar.checked,
    firstAppointmentStatus: form.elements.firstAppointmentStatus.value,
    secondAppointmentStatus: form.elements.secondAppointmentStatus.value,
    thirdAppointmentStatus: form.elements.thirdAppointmentStatus.value,
    documentsStatus: form.elements.documentsStatus.value,
    documentsNotes: form.elements.documentsNotes.value.trim(),
    carrierStatus: form.elements.carrierStatus.value,
    carrierNotes: form.elements.carrierNotes.value.trim(),
    firstAppointmentAt: form.elements.firstAppointmentAt.value || "",
    firstMeetingRecordingUrl: form.elements.firstMeetingRecordingUrl.value.trim(),
    firstAppointmentSummary: form.elements.firstAppointmentSummary.value.trim(),
    secondAppointmentAt: form.elements.secondAppointmentAt.value || "",
    secondMeetingRecordingUrl: form.elements.secondMeetingRecordingUrl.value.trim(),
    secondAppointmentSummary: form.elements.secondAppointmentSummary.value.trim(),
    thirdAppointmentAt: form.elements.thirdAppointmentAt.value || "",
    thirdMeetingRecordingUrl: form.elements.thirdMeetingRecordingUrl.value.trim(),
    thirdAppointmentSummary: form.elements.thirdAppointmentSummary.value.trim(),
    qualifiedProspectMeetingAttended: form.elements.qualifiedProspectMeetingAttended.checked,
    nextAction: form.elements.nextAction.value.trim(),
    closeReason: form.elements.closeReason.value.trim(),
    notes: form.elements.notes.value.trim(),
    tasks: sanitizeTasks(draftTasks),
    updatedAt: now,
  });
  const payload = editingLeadId
    ? withAutoConfirmTasks(leads.find((lead) => lead.id === editingLeadId), rawPayload)
    : withAutoConfirmTasks(null, rawPayload);

  if (editingLeadId) {
    leads = leads.map((lead) =>
      lead.id === editingLeadId ? mergeLeadUpdate(lead, payload, now) : lead
    );
  } else {
    leads.unshift({
      ...payload,
      id: crypto.randomUUID(),
      createdAt: now,
      stageUpdatedAt: now,
      updatedAt: now,
    });
  }

  persist();
  render();
  dialog.close();
  draftTasks = [];
}

function withAutoConfirmTasks(existingLead, payload) {
  const next = { ...payload };
  const tasks = sanitizeTasks(next.tasks || []);
  const firstWasScheduled = existingLead?.firstAppointmentStatus === "scheduled";
  const secondWasScheduled = existingLead?.secondAppointmentStatus === "scheduled";
  const thirdWasScheduled = existingLead?.thirdAppointmentStatus === "scheduled";
  const firstNowScheduled = next.firstAppointmentStatus === "scheduled";
  const secondNowScheduled = next.secondAppointmentStatus === "scheduled";
  const thirdNowScheduled = next.thirdAppointmentStatus === "scheduled";
  const firstHadTime = hasDateTimeValue(existingLead?.firstAppointmentAt);
  const secondHadTime = hasDateTimeValue(existingLead?.secondAppointmentAt);
  const thirdHadTime = hasDateTimeValue(existingLead?.thirdAppointmentAt);
  const firstHasTimeNow = hasDateTimeValue(next.firstAppointmentAt);
  const secondHasTimeNow = hasDateTimeValue(next.secondAppointmentAt);
  const thirdHasTimeNow = hasDateTimeValue(next.thirdAppointmentAt);
  const firstNewlyScheduled = firstNowScheduled && !firstWasScheduled;
  const secondNewlyScheduled = secondNowScheduled && !secondWasScheduled;
  const thirdNewlyScheduled = thirdNowScheduled && !thirdWasScheduled;
  const firstTimeAdded = firstHasTimeNow && !firstHadTime;
  const secondTimeAdded = secondHasTimeNow && !secondHadTime;
  const thirdTimeAdded = thirdHasTimeNow && !thirdHadTime;
  const firstNeedsConfirm =
    (firstNewlyScheduled || firstTimeAdded) &&
    next.firstAppointmentStatus !== "completed" &&
    next.firstAppointmentStatus !== "noShow";
  const secondNeedsConfirm =
    (secondNewlyScheduled || secondTimeAdded) &&
    next.secondAppointmentStatus !== "completed" &&
    next.secondAppointmentStatus !== "noShow";
  const thirdNeedsConfirm =
    (thirdNewlyScheduled || thirdTimeAdded) &&
    next.thirdAppointmentStatus !== "completed" &&
    next.thirdAppointmentStatus !== "noShow";

  if (firstNeedsConfirm) {
    maybeAddConfirmTask(tasks, "Confirm 1st appointment", next.firstAppointmentAt);
  }
  if (secondNeedsConfirm) {
    maybeAddConfirmTask(tasks, "Confirm 2nd appointment", next.secondAppointmentAt);
  }
  if (thirdNeedsConfirm) {
    maybeAddConfirmTask(tasks, "Confirm 3rd appointment", next.thirdAppointmentAt);
  }

  next.tasks = tasks;
  return next;
}

function hasDateTimeValue(value) {
  if (!value) return false;
  const parsed = new Date(value);
  return !Number.isNaN(parsed.getTime());
}

function maybeAddConfirmTask(tasks, title, appointmentAt) {
  const normalizedTitle = title.toLowerCase();
  const existingOpenTask = tasks.find(
    (task) =>
      !taskIsDone(task) &&
      String(task.title || "")
        .trim()
        .toLowerCase() === normalizedTitle
  );
  const dueAt = getDayBeforeDateOnly(appointmentAt);
  if (existingOpenTask) {
    if (dueAt && existingOpenTask.dueAt !== dueAt) {
      existingOpenTask.dueAt = dueAt;
    }
    return;
  }

  tasks.unshift(createTask(title, dueAt, DEFAULT_TASK_ASSIGNEE));
}

function getDayBeforeDateOnly(appointmentAt) {
  if (!appointmentAt) return "";
  const parsed = new Date(appointmentAt);
  if (Number.isNaN(parsed.getTime())) return "";
  const localDate = new Date(parsed.getFullYear(), parsed.getMonth(), parsed.getDate());
  localDate.setDate(localDate.getDate() - 1);
  const year = localDate.getFullYear();
  const month = String(localDate.getMonth() + 1).padStart(2, "0");
  const day = String(localDate.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function deleteCurrentLead() {
  if (!editingLeadId) return;
  leads = leads.filter((lead) => lead.id !== editingLeadId);
  editingLeadId = null;
  activitySection.classList.add("hidden");
  formActivityItems.innerHTML = "";
  draftTasks = [];
  renderFormTasks();
  persist();
  render();
  dialog.close();
}

function refreshDialogActivity() {
  if (!editingLeadId) return;
  const lead = leads.find((item) => item.id === editingLeadId);
  if (!lead) return;
  renderActivityLog(formActivityItems, lead);
}

function addTaskFromFormInput() {
  const title = form.elements.taskInput.value.trim();
  if (!title) return;
  draftTasks.unshift(
    createTask(title, form.elements.taskDueAt.value, form.elements.taskAssignee.value)
  );
  form.elements.taskInput.value = "";
  form.elements.taskDueAt.value = "";
  form.elements.taskAssignee.value = DEFAULT_TASK_ASSIGNEE;
  syncTasksForEditedLead();
  renderFormTasks();
}

function renderFormTasks() {
  if (!formTaskItems) return;
  formTaskItems.innerHTML = "";
  if (!Array.isArray(draftTasks) || draftTasks.length === 0) {
    const empty = document.createElement("p");
    empty.className = "task-empty";
    empty.textContent = "No tasks yet.";
    formTaskItems.appendChild(empty);
    return;
  }

  const openTasks = draftTasks.filter((task) => !taskIsDone(task)).sort(compareTasksForDisplay);
  const doneTasks = draftTasks
    .filter((task) => taskIsDone(task))
    .sort((a, b) => (b.completedAt || 0) - (a.completedAt || 0));

  renderOpenTaskGroup(openTasks);
  renderCompletedTaskGroup(doneTasks);
}

function renderOpenTaskGroup(tasks) {
  const heading = document.createElement("p");
  heading.className = "task-group-heading";
  heading.textContent = `Open: ${tasks.length}`;
  formTaskItems.appendChild(heading);

  if (tasks.length === 0) {
    const empty = document.createElement("p");
    empty.className = "task-empty";
    empty.textContent = "No open tasks.";
    formTaskItems.appendChild(empty);
    return;
  }

  tasks.forEach((task) => {
    const row = document.createElement("div");
    row.className = `task-open-item ${taskAssigneeClass(task.assignee)}`;

    const check = document.createElement("input");
    check.type = "checkbox";
    check.checked = false;
    check.addEventListener("change", () => {
      markDraftTaskDone(task.id);
    });

    const copy = document.createElement("div");
    copy.className = "task-open-copy";
    const title = document.createElement("p");
    title.className = "task-open-title";
    title.textContent = task.title;
    const meta = document.createElement("p");
    meta.className = "task-open-meta";
    meta.textContent = `Due: ${taskDueText(task.dueAt)}`;
    const dueEditRow = document.createElement("div");
    dueEditRow.className = "task-open-due-row";
    const dueInput = document.createElement("input");
    dueInput.type = "date";
    dueInput.className = "task-open-due-input";
    dueInput.value = normalizeDateOnlyString(task.dueAt);
    dueInput.setAttribute("aria-label", `Due date for ${task.title}`);
    dueInput.addEventListener("change", () => {
      updateDraftTaskDueAt(task.id, dueInput.value);
    });
    dueEditRow.appendChild(dueInput);
    const badge = document.createElement("span");
    badge.className = `task-assignee-badge ${taskAssigneeClass(task.assignee)}`;
    badge.textContent = normalizeTaskAssignee(task.assignee);
    copy.appendChild(title);
    copy.appendChild(meta);
    copy.appendChild(dueEditRow);
    copy.appendChild(badge);

    const removeBtn = document.createElement("button");
    removeBtn.type = "button";
    removeBtn.className = "btn btn-small btn-danger";
    removeBtn.textContent = "Delete";
    removeBtn.addEventListener("click", () => {
      removeDraftTask(task.id);
    });

    row.appendChild(check);
    row.appendChild(copy);
    row.appendChild(removeBtn);
    formTaskItems.appendChild(row);
  });
}

function renderCompletedTaskGroup(tasks) {
  const heading = document.createElement("p");
  heading.className = "task-group-heading";
  heading.textContent = `Completed (Logged): ${tasks.length}`;
  formTaskItems.appendChild(heading);

  if (tasks.length === 0) {
    const empty = document.createElement("p");
    empty.className = "task-empty";
    empty.textContent = "No completed tasks logged.";
    formTaskItems.appendChild(empty);
    return;
  }

  tasks.forEach((task) => {
    const row = document.createElement("div");
    row.className = `task-log-row ${taskAssigneeClass(task.assignee)}`;

    const copy = document.createElement("div");
    copy.className = "task-log-copy";

    const title = document.createElement("p");
    title.className = "task-open-title";
    title.textContent = task.title;

    const meta = document.createElement("p");
    meta.className = "task-open-meta";
    const completed = task.completedAt ? formatDateTime(task.completedAt) : "n/a";
    meta.textContent = `Due: ${taskDueText(task.dueAt)} | Completed: ${completed}`;
    const badge = document.createElement("span");
    badge.className = `task-assignee-badge ${taskAssigneeClass(task.assignee)}`;
    badge.textContent = normalizeTaskAssignee(task.assignee);

    const deleteBtn = document.createElement("button");
    deleteBtn.type = "button";
    deleteBtn.className = "btn btn-small btn-danger";
    deleteBtn.textContent = "Delete";
    deleteBtn.addEventListener("click", () => removeDraftTask(task.id));

    copy.appendChild(title);
    copy.appendChild(meta);
    copy.appendChild(badge);
    row.appendChild(copy);
    row.appendChild(deleteBtn);
    formTaskItems.appendChild(row);
  });
}

function createTask(title, dueAtInput, assigneeInput = DEFAULT_TASK_ASSIGNEE) {
  return {
    id: crypto.randomUUID(),
    title: title.trim(),
    dueAt: normalizeDateOnlyString(dueAtInput || ""),
    assignee: normalizeTaskAssignee(assigneeInput),
    status: "open",
    done: false,
    createdAt: Date.now(),
    completedAt: null,
  };
}

function compareTasksForDisplay(a, b) {
  const aDue = a.dueAt ? new Date(a.dueAt).getTime() : Number.MAX_SAFE_INTEGER;
  const bDue = b.dueAt ? new Date(b.dueAt).getTime() : Number.MAX_SAFE_INTEGER;
  if (aDue !== bDue) return aDue - bDue;
  return (b.createdAt || 0) - (a.createdAt || 0);
}

function taskIsDone(task) {
  if (!task) return false;
  if ((task.status || "").toString().trim().toLowerCase() === "done") return true;
  return toBool(task.done);
}

function taskDueText(value) {
  const normalized = normalizeDateOnlyString(value);
  if (!normalized) return "not set";
  return formatDateOnly(normalized);
}

function markDraftTaskDone(taskId) {
  const now = Date.now();
  draftTasks = draftTasks.map((task) =>
    task.id === taskId
      ? {
          ...task,
          status: "done",
          done: true,
          completedAt: now,
        }
      : task
  );
  syncTasksForEditedLead();
  renderFormTasks();
}

function removeDraftTask(taskId) {
  draftTasks = draftTasks.filter((task) => task.id !== taskId);
  syncTasksForEditedLead();
  renderFormTasks();
}

function updateDraftTaskDueAt(taskId, dueAtInput) {
  const normalizedDueAt = normalizeDateOnlyString(dueAtInput);
  draftTasks = draftTasks.map((task) =>
    task.id === taskId
      ? {
          ...task,
          dueAt: normalizedDueAt,
        }
      : task
  );
  syncTasksForEditedLead();
  renderFormTasks();
}

function markLeadTaskDone(leadId, taskId) {
  const now = Date.now();
  leads = leads.map((lead) => {
    if (lead.id !== leadId) return lead;
    const tasks = (lead.tasks || []).map((task) => {
      if (task.id !== taskId) return task;
      return {
        ...task,
        status: "done",
        done: true,
        completedAt: now,
      };
    });
    return normalizeLead({
      ...lead,
      tasks,
      updatedAt: now,
    });
  });
  persist();
  render();
  if (editingLeadId === leadId) {
    const lead = leads.find((item) => item.id === leadId);
    draftTasks = normalizeTasks(lead?.tasks || []);
    renderFormTasks();
  }
}

function syncTasksForEditedLead() {
  if (!editingLeadId) return;
  const now = Date.now();
  leads = leads.map((lead) => {
    if (lead.id !== editingLeadId) return lead;
    return normalizeLead({
      ...lead,
      tasks: sanitizeTasks(draftTasks),
      updatedAt: now,
    });
  });
  persist();
  render();
}

function advanceStage(id) {
  leads = leads.map((lead) => {
    if (lead.id !== id) return lead;

    const newStage = stageAdvanceMap[lead.stage] || "lead";
    return normalizeLead({
      ...lead,
      stage: newStage,
      stageUpdatedAt: Date.now(),
      updatedAt: Date.now(),
    });
  });

  persist();
  render();
}

function setOutcome(id, outcome) {
  leads = leads.map((lead) => {
    if (lead.id !== id) return lead;
    return normalizeLead({
      ...lead,
      stage: outcome,
      closedAt: Date.now(),
      stageUpdatedAt: Date.now(),
      closeReason:
        lead.closeReason ||
        (outcome === "won" ? "Closed won." : "Closed lost. No reason recorded."),
      updatedAt: Date.now(),
    });
  });

  persist();
  render();
}

function setLeadStatus(id, status) {
  const normalizedStatus = normalizeLeadStatus(status);
  leads = leads.map((lead) => {
    if (lead.id !== id) return lead;
    return normalizeLead({
      ...lead,
      leadStatus: normalizedStatus,
      statusUpdatedAt: Date.now(),
      updatedAt: Date.now(),
    });
  });

  persist();
  render();
}

function renderFormCpqpma() {
  const value = getCostPerQualified();
  formCpqpma.textContent =
    value == null
      ? "Cost per qualified prospect meeting attended: n/a"
      : `Cost per qualified prospect meeting attended: ${formatCurrency(value)}`;
}

function setFirstAppointmentNoShow(id) {
  const now = Date.now();
  leads = leads.map((lead) => {
    if (lead.id !== id) return lead;
    return normalizeLead({
      ...lead,
      stage: "firstAppointmentNoShow",
      leadStatus: "active",
      firstAppointmentStatus: "noShow",
      stageUpdatedAt: now,
      statusUpdatedAt: now,
      updatedAt: now,
    });
  });

  persist();
  render();
}

function setupDragAndDrop() {
  document.querySelectorAll(".column").forEach((column) => {
    if (column.dataset.dndBound === "true") return;
    column.dataset.dndBound = "true";

    column.addEventListener("dragover", (event) => {
      event.preventDefault();
      event.dataTransfer.dropEffect = "move";
      column.classList.add("drop-target");
    });

    column.addEventListener("dragleave", () => {
      column.classList.remove("drop-target");
    });

    column.addEventListener("drop", (event) => {
      event.preventDefault();
      column.classList.remove("drop-target");
      const targetStage = column.dataset.stage;
      if (!targetStage) return;
      const droppedId = event.dataTransfer.getData("text/plain") || draggedLeadId;
      if (!droppedId) return;
      moveLeadToStage(droppedId, targetStage);
    });
  });
}

function moveLeadToStage(id, stage) {
  if (!STAGES.includes(stage)) return;

  let didChange = false;
  leads = leads.map((lead) => {
    if (lead.id !== id) return lead;
    if (lead.stage === stage) return lead;
    didChange = true;

    const now = Date.now();
    const base = {
      ...lead,
      stage,
      stageUpdatedAt: now,
      updatedAt: now,
    };

    if (stage === "won" || stage === "lost") {
      return normalizeLead({
        ...base,
        closedAt: now,
        closeReason:
          lead.closeReason ||
          (stage === "won" ? "Closed won." : "Closed lost. No reason recorded."),
      });
    }

    return normalizeLead(base);
  });

  if (!didChange) return;
  persist();
  render();
}

function logActivity(id, type) {
  if (!ACTIVITY_TYPES.includes(type)) return;
  const now = Date.now();
  leads = leads.map((lead) => {
    if (lead.id !== id) return lead;

    const entry = {
      id: crypto.randomUUID(),
      at: now,
      type: normalizeActivityType(type),
      note: "",
    };
    const activityLogs = [entry, ...(lead.activityLogs || [])].slice(0, 100);

    return normalizeLead({
      ...lead,
      activityLogs,
      updatedAt: now,
    });
  });

  persist();
  render();
}

function saveActivityNote(leadId, logId, note) {
  const now = Date.now();
  leads = leads.map((lead) => {
    if (lead.id !== leadId) return lead;
    const activityLogs = (lead.activityLogs || []).map((entry) =>
      entry.id === logId ? { ...entry, note: note.trim() } : entry
    );
    return normalizeLead({
      ...lead,
      activityLogs,
      updatedAt: now,
    });
  });

  persist();
  render();
  refreshDialogActivity();
}

function deleteActivityLog(leadId, logId) {
  const now = Date.now();
  leads = leads.map((lead) => {
    if (lead.id !== leadId) return lead;
    const activityLogs = (lead.activityLogs || []).filter((entry) => entry.id !== logId);
    return normalizeLead({
      ...lead,
      activityLogs,
      updatedAt: now,
    });
  });

  persist();
  render();
  refreshDialogActivity();
}

function mergeLeadUpdate(existingLead, payload, now) {
  const stageChanged = existingLead.stage !== payload.stage;
  const statusChanged = existingLead.leadStatus !== payload.leadStatus;
  const isClosed = payload.stage === "won" || payload.stage === "lost";
  return {
    ...existingLead,
    ...payload,
    id: existingLead.id,
    createdAt: existingLead.createdAt,
    activityLogs: existingLead.activityLogs,
    closedAt: isClosed ? existingLead.closedAt || payload.closedAt || now : null,
    stageUpdatedAt: stageChanged ? now : existingLead.stageUpdatedAt,
    statusUpdatedAt: statusChanged ? now : existingLead.statusUpdatedAt,
    updatedAt: now,
  };
}

function normalizeLead(input) {
  const lead = { ...input };
  const stage = normalizeStageValue(lead.stage);
  const leadStatus = normalizeLeadStatus(lead.leadStatus);
  const firstConfirmedPhone =
    lead.firstConfirmedPhone !== undefined ? toBool(lead.firstConfirmedPhone) : toBool(lead.confirmedPhone);
  const firstConfirmedCalendar =
    lead.firstConfirmedCalendar !== undefined
      ? toBool(lead.firstConfirmedCalendar)
      : toBool(lead.confirmedCalendar);
  const secondConfirmedPhone = toBool(lead.secondConfirmedPhone);
  const secondConfirmedCalendar = toBool(lead.secondConfirmedCalendar);
  const thirdConfirmedPhone = toBool(lead.thirdConfirmedPhone);
  const thirdConfirmedCalendar = toBool(lead.thirdConfirmedCalendar);

  const firstAppointmentStatus = inferFirstAppointmentStatus(lead, stage);
  const secondAppointmentStatus = inferSecondAppointmentStatus(lead, stage);
  const thirdAppointmentStatus = inferThirdAppointmentStatus(lead, stage);
  const documentsStatus = inferDocumentsStatus(lead);
  const carrierStatus = inferCarrierStatus(lead, stage);
  const appointmentStatus = inferPrimaryAppointmentStatus(
    firstAppointmentStatus,
    firstConfirmedPhone,
    firstConfirmedCalendar
  );
  const confirmedPhone = firstConfirmedPhone;
  const confirmedCalendar = firstConfirmedCalendar;

  const firstMeetingDone = firstAppointmentStatus === "completed";
  const infoMeetingDone = secondAppointmentStatus === "completed";
  const secondAppointmentBooked =
    secondAppointmentStatus === "scheduled" ||
    secondAppointmentStatus === "completed" ||
    secondAppointmentStatus === "noShow";
  const firstAppointmentNoShow = firstAppointmentStatus === "noShow";
  const secondAppointmentNoShow = secondAppointmentStatus === "noShow";
  const thirdAppointmentNoShow = thirdAppointmentStatus === "noShow";
  const waitingOnDocuments = documentsStatus === "waiting";
  const carrierReachedOut = carrierStatus === "contacted" || carrierStatus === "quoted";

  let closedAt = lead.closedAt ? toNumber(lead.closedAt) : null;
  if ((stage === "won" || stage === "lost") && !closedAt) {
    closedAt = Date.now();
  }
  if (stage !== "won" && stage !== "lost") {
    closedAt = null;
  }

  let closeReason = (lead.closeReason || "").toString().trim();
  if (stage === "lost" && !closeReason) {
    closeReason = "No reason recorded.";
  }
  if (stage === "won" && !closeReason) {
    closeReason = "Closed won.";
  }

  const activityLogs = normalizeActivityLogs(lead.activityLogs || lead.callLogs);
  const tasks = normalizeTasks(lead.tasks);
  const metrics = computeActivityMetrics(activityLogs, lead);

  return {
    id: (lead.id || "").toString(),
    firstName: (lead.firstName || "").toString().trim(),
    nickname: (lead.nickname || "").toString().trim(),
    lastName: (lead.lastName || "").toString().trim(),
    email: (lead.email || "").toString().trim(),
    phone: (lead.phone || "").toString().trim(),
    state: (lead.state || "").toString().trim().toUpperCase(),
    stage,
    leadStatus,
    appointmentStatus,
    confirmedPhone,
    confirmedCalendar,
    firstConfirmedPhone,
    firstConfirmedCalendar,
    secondConfirmedPhone,
    secondConfirmedCalendar,
    thirdConfirmedPhone,
    thirdConfirmedCalendar,
    dialAttempts: metrics.callCount,
    liveContacts: metrics.liveCount,
    textsAttempted: metrics.textSentCount,
    texted: metrics.textSentCount > 0,
    callCount: metrics.callCount,
    callVmCount: metrics.callVmCount,
    callNcCount: metrics.callNcCount,
    textSentCount: metrics.textSentCount,
    textReceivedCount: metrics.textReceivedCount,
    firstAppointmentStatus,
    secondAppointmentStatus,
    thirdAppointmentStatus,
    documentsStatus,
    documentsNotes: (lead.documentsNotes || "").toString().trim(),
    carrierStatus,
    carrierNotes: (lead.carrierNotes || "").toString().trim(),
    firstAppointmentAt: normalizeDateTimeString(lead.firstAppointmentAt || lead.nextAppointmentAt),
    firstMeetingRecordingUrl: (lead.firstMeetingRecordingUrl || "").toString().trim(),
    firstAppointmentSummary: (lead.firstAppointmentSummary || "").toString().trim(),
    secondAppointmentAt: normalizeDateTimeString(lead.secondAppointmentAt),
    secondMeetingRecordingUrl: (lead.secondMeetingRecordingUrl || "").toString().trim(),
    secondAppointmentSummary: (lead.secondAppointmentSummary || "").toString().trim(),
    thirdAppointmentAt: normalizeDateTimeString(lead.thirdAppointmentAt),
    thirdMeetingRecordingUrl: (lead.thirdMeetingRecordingUrl || "").toString().trim(),
    thirdAppointmentSummary: (lead.thirdAppointmentSummary || "").toString().trim(),
    firstMeetingDone,
    infoMeetingDone,
    secondAppointmentBooked,
    firstAppointmentNoShow,
    secondAppointmentNoShow,
    thirdAppointmentNoShow,
    waitingOnDocuments,
    carrierReachedOut,
    qualifiedProspectMeetingAttended: toBool(lead.qualifiedProspectMeetingAttended),
    opportunitySize: toNumber(lead.opportunitySize),
    nextAction: (lead.nextAction || "").toString().trim(),
    followUpAt: normalizeDateTimeString(lead.followUpAt),
    closeReason,
    closedAt,
    stageUpdatedAt: toNumber(lead.stageUpdatedAt || lead.updatedAt || lead.createdAt || Date.now()),
    statusUpdatedAt: toNumber(lead.statusUpdatedAt || lead.updatedAt || lead.createdAt || Date.now()),
    activityLogs,
    callLogs: activityLogs,
    tasks,
    notes: (lead.notes || "").toString().trim(),
    createdAt: toNumber(lead.createdAt || Date.now()),
    updatedAt: toNumber(lead.updatedAt || Date.now()),
  };
}

function normalizeLeadStatus(value) {
  const status = (value || "").toString().trim();
  if (status === "inactive") return "inactive";
  if (status === "noShow" || status === "rescheduled" || status === "doNotContact") {
    return "inactive";
  }
  return "active";
}

function normalizeStageValue(value) {
  const raw = (value || "").toString().trim();
  const mapped = mapLegacyStage(raw);
  if (STAGES.includes(mapped)) return mapped;
  return "lead";
}

function mapLegacyStage(stage) {
  if (stage === "booked") return "firstAppointmentBooked";
  if (stage === "firstNoShow") return "firstAppointmentNoShow";
  if (stage === "firstMeeting") return "secondAppointmentBooked";
  if (stage === "infoMeeting") return "secondAppointmentAttended";
  if (stage === "carrierOutreach") return "pendingWon";
  if (stage === "proposal") return "pendingWon";
  return stage;
}

function normalizeTasks(value) {
  if (Array.isArray(value)) {
    return sanitizeTasks(value);
  }
  if (typeof value === "string" && value.trim()) {
    try {
      const parsed = JSON.parse(value);
      if (!Array.isArray(parsed)) return [];
      return sanitizeTasks(parsed);
    } catch {
      return [];
    }
  }
  return [];
}

function sanitizeTasks(list) {
  if (!Array.isArray(list)) return [];
  return list
    .map((task) => {
      const isTextTask = typeof task === "string";
      const source = isTextTask ? {} : (task || {});
      const title = isTextTask
        ? task
        : (source.title || source.text || source.name || "").toString();
      const done = toBool(source.done) || (source.status || "").toString().toLowerCase() === "done";
      return {
        id: (source.id || crypto.randomUUID()).toString(),
        title: title.trim(),
        dueAt: normalizeDateOnlyString(source.dueAt || source.dueDate || source.due),
        assignee: normalizeTaskAssignee(source.assignee),
        status: done ? "done" : "open",
        done,
        createdAt: toNumber(source.createdAt || Date.now()),
        completedAt: done ? toNumber(source.completedAt || Date.now()) : null,
      };
    })
    .filter((task) => task.title.length > 0)
    .slice(0, 100);
}

function normalizeTaskAssignee(value) {
  const normalized = (value || "").toString().trim().toUpperCase();
  return TASK_ASSIGNEES.includes(normalized) ? normalized : DEFAULT_TASK_ASSIGNEE;
}

function taskAssigneeClass(value) {
  return normalizeTaskAssignee(value) === "COLBY" ? "assignee-colby" : "assignee-jake";
}

function getOpenTasks(lead) {
  return (lead.tasks || []).filter((task) => !taskIsDone(task));
}

function taskDueLabel(value, plain = false) {
  if (!value) return plain ? "Due: not set" : " | Due not set";
  return plain ? `Due: ${formatDateOnly(value)}` : ` | Due ${formatDateOnly(value)}`;
}

function normalizeDateOnlyString(value) {
  if (!value) return "";
  const text = String(value).trim();
  const match = text.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (match) return `${match[1]}-${match[2]}-${match[3]}`;
  const parsed = new Date(text);
  if (Number.isNaN(parsed.getTime())) return "";
  return parsed.toISOString().slice(0, 10);
}

function formatDateOnly(value) {
  const normalized = normalizeDateOnlyString(value);
  if (!normalized) return "n/a";
  const [year, month, day] = normalized.split("-");
  return `${month}/${day}/${year}`;
}

function inferPrimaryAppointmentStatus(firstStatus, firstPhoneConfirmed, firstCalendarConfirmed) {
  if (firstStatus === "notScheduled") return "notBooked";
  if (firstPhoneConfirmed || firstCalendarConfirmed) return "bookedConfirmed";
  return "bookedUnconfirmed";
}

function inferAppointmentStatus(lead, stage) {
  if (toBool(lead.confirmedPhone) || toBool(lead.confirmedCalendar)) {
    return "bookedConfirmed";
  }
  if (BOOKED_FLOW_STAGES.includes(stage)) return "bookedUnconfirmed";
  return "notBooked";
}

function appointmentLabel(lead) {
  if (lead.appointmentStatus === "notBooked") return "Appointment: Not booked";
  if (lead.appointmentStatus === "bookedUnconfirmed") return "Appointment: Booked (unconfirmed)";

  const channels = [];
  if (lead.confirmedPhone) channels.push("Phone");
  if (lead.confirmedCalendar) channels.push("Calendar");
  const channelLabel = channels.length > 0 ? channels.join(" + ") : "Confirmed";
  return `Appointment: Confirmed (${channelLabel})`;
}

function leadStatusLabel(status) {
  if (status === "inactive") return "Inactive";
  return "Active";
}

function carrierStatusLabel(status) {
  if (status === "contacted") return "Contacted";
  if (status === "quoted") return "Quoted";
  return "Pending";
}

function firstAppointmentStatusLabel(status) {
  if (status === "scheduled") return "Scheduled";
  if (status === "completed") return "Completed";
  if (status === "noShow") return "No show";
  return "Not scheduled";
}

function secondAppointmentStatusLabel(status) {
  if (status === "scheduled") return "Scheduled";
  if (status === "completed") return "Completed";
  if (status === "noShow") return "No show";
  return "Not booked";
}

function documentsStatusLabel(status) {
  if (status === "waiting") return "Waiting on documents";
  if (status === "received") return "Received";
  return "Not requested";
}

function renderAppointmentTime(element, label, value, status = "") {
  element.classList.remove("appt-empty", "appt-upcoming", "appt-soon", "appt-overdue", "appt-completed", "appt-noshow");

  if (status === "completed") {
    element.classList.add("appt-completed");
    element.textContent = `${label}: Attended${value ? ` (${formatDateTime(value)})` : ""}`;
    return;
  }

  if (status === "noShow") {
    element.classList.add("appt-noshow");
    element.textContent = `${label}: No show${value ? ` (${formatDateTime(value)})` : ""}`;
    return;
  }

  if (!value) {
    element.classList.add("appt-empty");
    element.textContent = `${label}: Not set`;
    return;
  }

  const date = new Date(value);
  if (Number.isNaN(date.getTime())) {
    element.classList.add("appt-empty");
    element.textContent = `${label}: Not set`;
    return;
  }

  const now = Date.now();
  const diffMs = date.getTime() - now;
  const diffHours = diffMs / (1000 * 60 * 60);

  if (diffHours < 0) {
    element.classList.add("appt-overdue");
    element.textContent = `${label}: ${formatDateTime(value)} (past due)`;
    return;
  }
  if (diffHours <= 24) {
    element.classList.add("appt-soon");
    element.textContent = `${label}: ${formatDateTime(value)} (within 24h)`;
    return;
  }

  element.classList.add("appt-upcoming");
  element.textContent = `${label}: ${formatDateTime(value)}`;
}

function inferFirstAppointmentStatus(lead, stage) {
  if (lead.firstAppointmentStatus === "noShow") return "noShow";
  if (lead.firstAppointmentStatus === "completed") return "completed";
  if (stage === "firstAppointmentNoShow") return "noShow";
  if (toBool(lead.firstAppointmentNoShow)) return "noShow";
  if (
    toBool(lead.firstMeetingDone) ||
    stage === "secondAppointmentBooked" ||
    stage === "secondAppointmentConfirmed" ||
    stage === "secondAppointmentRescheduled" ||
    stage === "secondAppointmentAttended" ||
    stage === "thirdAppointmentBooked" ||
    stage === "thirdAppointmentConfirmed" ||
    stage === "thirdAppointmentRescheduled" ||
    stage === "thirdAppointmentAttended" ||
    stage === "pendingWon" ||
    stage === "pendingLost" ||
    stage === "won" ||
    stage === "lost"
  ) {
    return "completed";
  }
  if (lead.firstAppointmentStatus === "scheduled") return "scheduled";
  if (lead.appointmentStatus === "bookedUnconfirmed" || lead.appointmentStatus === "bookedConfirmed") {
    return "scheduled";
  }
  return "notScheduled";
}

function inferSecondAppointmentStatus(lead, stage) {
  if (lead.secondAppointmentStatus === "noShow") return "noShow";
  if (lead.secondAppointmentStatus === "completed") return "completed";
  if (toBool(lead.secondAppointmentNoShow)) return "noShow";
  if (
    toBool(lead.infoMeetingDone) ||
    stage === "secondAppointmentAttended" ||
    stage === "thirdAppointmentBooked" ||
    stage === "thirdAppointmentConfirmed" ||
    stage === "thirdAppointmentRescheduled" ||
    stage === "thirdAppointmentAttended" ||
    stage === "pendingWon" ||
    stage === "pendingLost" ||
    stage === "won" ||
    stage === "lost"
  ) {
    return "completed";
  }
  if (lead.secondAppointmentStatus === "scheduled") return "scheduled";
  if (toBool(lead.secondAppointmentBooked)) return "scheduled";
  return "notBooked";
}

function inferThirdAppointmentStatus(lead, stage) {
  if (THIRD_APPOINTMENT_STATUSES.includes(lead.thirdAppointmentStatus)) {
    return lead.thirdAppointmentStatus;
  }
  if (toBool(lead.thirdAppointmentNoShow)) return "noShow";
  if (stage === "thirdAppointmentAttended" || stage === "pendingWon" || stage === "pendingLost" || stage === "won" || stage === "lost") {
    return "completed";
  }
  if (
    stage === "thirdAppointmentBooked" ||
    stage === "thirdAppointmentConfirmed" ||
    stage === "thirdAppointmentRescheduled"
  ) {
    return "scheduled";
  }
  return "notBooked";
}

function inferDocumentsStatus(lead) {
  if (DOCUMENT_STATUSES.includes(lead.documentsStatus)) return lead.documentsStatus;
  if (toBool(lead.waitingOnDocuments)) return "waiting";
  return "notRequested";
}

function inferCarrierStatus(lead, stage) {
  if (CARRIER_STATUSES.includes(lead.carrierStatus)) return lead.carrierStatus;
  if (toBool(lead.carrierReachedOut) || stage === "pendingWon" || stage === "pendingLost") {
    return "contacted";
  }
  if (stage === "won" || stage === "lost") {
    return "quoted";
  }
  return "notStarted";
}

function fullName(lead) {
  if (lead.nickname) return `${lead.firstName} "${lead.nickname}" ${lead.lastName}`.trim();
  return `${lead.firstName} ${lead.lastName}`.trim();
}

function exportCSV() {
  const rows = [CSV_COLUMNS.join(",")];
  leads.forEach((lead) => {
    const row = CSV_COLUMNS.map((column) => {
      if (column === "callLogs" || column === "activityLogs") {
        return escapeCSV(JSON.stringify(lead.activityLogs || []));
      }
      if (column === "tasks") return escapeCSV(JSON.stringify(lead.tasks || []));
      return escapeCSV(lead[column]);
    });
    rows.push(row.join(","));
  });

  const csv = rows.join("\n");
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = `annuity-crm-${new Date().toISOString().slice(0, 10)}.csv`;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}

async function importCSV(event) {
  const file = event.target.files?.[0];
  if (!file) return;

  try {
    const text = await file.text();
    const records = parseCSV(text);
    if (records.length === 0) {
      window.alert("No rows found in CSV.");
      return;
    }

    const byId = new Map(leads.map((lead) => [lead.id, lead]));
    let importedCount = 0;

    records.forEach((record) => {
      const normalized = normalizeLead({
        ...record,
        id: record.id || crypto.randomUUID(),
        confirmedPhone: toBool(record.confirmedPhone),
        confirmedCalendar: toBool(record.confirmedCalendar),
        firstConfirmedPhone: toBool(record.firstConfirmedPhone),
        firstConfirmedCalendar: toBool(record.firstConfirmedCalendar),
        secondConfirmedPhone: toBool(record.secondConfirmedPhone),
        secondConfirmedCalendar: toBool(record.secondConfirmedCalendar),
        thirdConfirmedPhone: toBool(record.thirdConfirmedPhone),
        thirdConfirmedCalendar: toBool(record.thirdConfirmedCalendar),
        texted: toBool(record.texted),
        firstMeetingDone: toBool(record.firstMeetingDone),
        infoMeetingDone: toBool(record.infoMeetingDone),
        secondAppointmentBooked: toBool(record.secondAppointmentBooked),
        firstAppointmentNoShow: toBool(record.firstAppointmentNoShow),
        secondAppointmentNoShow: toBool(record.secondAppointmentNoShow),
        thirdAppointmentNoShow: toBool(record.thirdAppointmentNoShow),
        waitingOnDocuments: toBool(record.waitingOnDocuments),
        carrierReachedOut: toBool(record.carrierReachedOut),
        activityLogs: parseActivityLogs(record.activityLogs || record.callLogs),
      });

      if (!normalized.id) normalized.id = crypto.randomUUID();
      byId.set(normalized.id, normalized);
      importedCount += 1;
    });

    leads = Array.from(byId.values()).sort((a, b) => b.updatedAt - a.updatedAt);
    persist();
    render();
    window.alert(`Imported ${importedCount} lead row(s).`);
  } catch {
    window.alert("Could not import CSV. Check format and try again.");
  } finally {
    event.target.value = "";
  }
}

function parseCSV(text) {
  const rows = [];
  let currentRow = [];
  let currentField = "";
  let inQuotes = false;

  for (let i = 0; i < text.length; i += 1) {
    const char = text[i];
    const next = text[i + 1];

    if (char === '"') {
      if (inQuotes && next === '"') {
        currentField += '"';
        i += 1;
      } else {
        inQuotes = !inQuotes;
      }
      continue;
    }

    if (char === "," && !inQuotes) {
      currentRow.push(currentField);
      currentField = "";
      continue;
    }

    if ((char === "\n" || char === "\r") && !inQuotes) {
      if (char === "\r" && next === "\n") i += 1;
      currentRow.push(currentField);
      currentField = "";
      if (currentRow.length > 1 || currentRow[0] !== "") {
        rows.push(currentRow);
      }
      currentRow = [];
      continue;
    }

    currentField += char;
  }

  currentRow.push(currentField);
  if (currentRow.length > 1 || currentRow[0] !== "") {
    rows.push(currentRow);
  }

  if (rows.length < 2) return [];
  const headers = rows[0].map((h) => h.trim());
  const dataRows = rows.slice(1);

  return dataRows.map((row) => {
    const record = {};
    headers.forEach((header, index) => {
      record[header] = (row[index] || "").trim();
    });
    return record;
  });
}

function escapeCSV(value) {
  const stringValue = value == null ? "" : String(value);
  if (
    stringValue.includes(",") ||
    stringValue.includes('"') ||
    stringValue.includes("\n") ||
    stringValue.includes("\r")
  ) {
    return `"${stringValue.replace(/"/g, '""')}"`;
  }
  return stringValue;
}

function getSLAStatus(lead) {
  if (lead.stage === "won" || lead.stage === "lost") {
    return { overdue: false, label: "" };
  }
  const hoursAllowed = SLA_HOURS_BY_STAGE[lead.stage];
  if (!hoursAllowed) return { overdue: false, label: "" };

  const elapsedHours = (Date.now() - lead.stageUpdatedAt) / (1000 * 60 * 60);
  if (elapsedHours <= hoursAllowed) return { overdue: false, label: "" };

  const overdueHours = Math.floor(elapsedHours - hoursAllowed);
  return { overdue: true, label: `${overdueHours}h overdue` };
}

function normalizeDateTimeString(value) {
  if (!value) return "";
  const parsed = new Date(value);
  if (Number.isNaN(parsed.getTime())) return "";
  return toDateTimeLocalValue(parsed.toISOString());
}

function toDateTimeLocalValue(value) {
  if (!value) return "";
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return "";
  const offsetMs = date.getTimezoneOffset() * 60 * 1000;
  const localDate = new Date(date.getTime() - offsetMs);
  return localDate.toISOString().slice(0, 16);
}

function formatDateTime(value) {
  if (!value) return "n/a";
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return "n/a";
  return date.toLocaleString("en-US", {
    month: "short",
    day: "numeric",
    hour: "numeric",
    minute: "2-digit",
  });
}

function renderActivityLog(container, lead) {
  container.innerHTML = "";
  const logs = (lead.activityLogs || []).slice(0, 8);
  if (logs.length === 0) {
    const empty = document.createElement("p");
    empty.className = "activity-empty";
    empty.textContent = "No activity logged yet.";
    container.appendChild(empty);
    return;
  }

  logs.forEach((entry) => {
    const row = document.createElement("div");
    row.className = "activity-item";

    const head = document.createElement("div");
    head.className = "activity-head";

    const top = document.createElement("p");
    top.className = "activity-line";
    top.textContent = `${activityTypeLabel(entry.type)} | ${formatDateTime(entry.at)}`;

    const menu = document.createElement("details");
    menu.className = "activity-menu";
    const summary = document.createElement("summary");
    summary.className = "activity-menu-trigger";
    summary.textContent = "⋮";
    summary.setAttribute("aria-label", "Activity actions");
    const menuList = document.createElement("div");
    menuList.className = "activity-menu-list";

    const editBtn = document.createElement("button");
    editBtn.type = "button";
    editBtn.className = "activity-menu-item";
    editBtn.textContent = entry.note ? "Edit Note" : "Add Note";

    const deleteNoteBtn = document.createElement("button");
    deleteNoteBtn.type = "button";
    deleteNoteBtn.className = "activity-menu-item";
    deleteNoteBtn.textContent = "Clear Note";
    if (!entry.note) deleteNoteBtn.classList.add("hidden");
    deleteNoteBtn.addEventListener("click", () => {
      saveActivityNote(lead.id, entry.id, "");
      menu.open = false;
    });

    const deleteLogBtn = document.createElement("button");
    deleteLogBtn.type = "button";
    deleteLogBtn.className = "activity-menu-item activity-menu-item-danger";
    deleteLogBtn.textContent = "Delete Log";
    deleteLogBtn.addEventListener("click", () => {
      deleteActivityLog(lead.id, entry.id);
      menu.open = false;
    });

    menuList.appendChild(editBtn);
    menuList.appendChild(deleteNoteBtn);
    menuList.appendChild(deleteLogBtn);
    menu.appendChild(summary);
    menu.appendChild(menuList);
    head.appendChild(top);
    head.appendChild(menu);

    const preview = document.createElement("p");
    preview.className = "activity-preview";
    preview.textContent = entry.note || "No note";

    const editor = document.createElement("div");
    editor.className = "activity-note-row hidden";

    const input = document.createElement("input");
    input.type = "text";
    input.className = "activity-note-input";
    input.placeholder = "Add note";
    input.value = entry.note || "";

    const saveBtn = document.createElement("button");
    saveBtn.type = "button";
    saveBtn.className = "btn btn-small btn-primary activity-mini-btn hidden";
    saveBtn.textContent = "Save";
    saveBtn.addEventListener("click", () => saveActivityNote(lead.id, entry.id, input.value));

    const cancelBtn = document.createElement("button");
    cancelBtn.type = "button";
    cancelBtn.className = "btn btn-small btn-secondary activity-mini-btn";
    cancelBtn.textContent = "Close";
    cancelBtn.addEventListener("click", () => {
      input.value = entry.note || "";
      saveBtn.classList.add("hidden");
      editor.classList.add("hidden");
    });

    input.addEventListener("input", () => {
      if (input.value.trim() === (entry.note || "").trim()) {
        saveBtn.classList.add("hidden");
      } else {
        saveBtn.classList.remove("hidden");
      }
    });

    editBtn.addEventListener("click", () => {
      editor.classList.toggle("hidden");
      menu.open = false;
    });

    editor.appendChild(input);
    editor.appendChild(saveBtn);
    editor.appendChild(cancelBtn);

    row.appendChild(head);
    row.appendChild(preview);
    row.appendChild(editor);
    container.appendChild(row);
  });
}

function activityTypeLabel(type) {
  if (type === "callLiveContact") return "Call Live Contact";
  if (type === "callVM") return "Call VM";
  if (type === "callNC") return "Call NC";
  if (type === "textSent") return "Text Sent";
  if (type === "textReceived") return "Text Received";
  if (type === "emailSent") return "Email Sent";
  if (type === "emailReceived") return "Email Received";
  return "Activity";
}

function normalizeActivityLogs(value) {
  const logs = parseActivityLogs(value);
  return logs
    .filter((entry) => entry && entry.at && entry.type)
    .sort((a, b) => b.at - a.at)
    .slice(0, 100);
}

function parseActivityLogs(value) {
  if (Array.isArray(value)) {
    return value.map((entry) => ({
      id: entry.id || crypto.randomUUID(),
      at: toNumber(entry.at),
      type: normalizeActivityType(entry.type),
      note: (entry.note || "").toString().trim(),
    }));
  }
  if (typeof value === "string" && value.trim()) {
    try {
      const parsed = JSON.parse(value);
      if (!Array.isArray(parsed)) return [];
      return parsed.map((entry) => ({
        id: entry.id || crypto.randomUUID(),
        at: toNumber(entry.at),
        type: normalizeActivityType(entry.type),
        note: (entry.note || "").toString().trim(),
      }));
    } catch {
      return [];
    }
  }
  return [];
}

function normalizeActivityType(type) {
  if (type === "live") return "callLiveContact";
  if (type === "dial") return "callNC";
  if (ACTIVITY_TYPES.includes(type)) return type;
  return "callNC";
}

function computeActivityMetrics(logs, lead) {
  const liveCount = logs.filter((entry) => entry.type === "callLiveContact").length;
  const callVmCount = logs.filter((entry) => entry.type === "callVM").length;
  const callNcCount = logs.filter((entry) => entry.type === "callNC").length;
  const textSentCount = logs.filter((entry) => entry.type === "textSent").length;
  const textReceivedCount = logs.filter((entry) => entry.type === "textReceived").length;
  const callCount = liveCount + callVmCount + callNcCount;

  if (logs.length > 0) {
    return { callCount, liveCount, callVmCount, callNcCount, textSentCount, textReceivedCount };
  }

  return {
    callCount: toNumber(lead.dialAttempts),
    liveCount: toNumber(lead.liveContacts),
    callVmCount: 0,
    callNcCount: Math.max(0, toNumber(lead.dialAttempts) - toNumber(lead.liveContacts)),
    textSentCount: toNumber(lead.textsAttempted),
    textReceivedCount: 0,
  };
}

function buildActivitySummary(lead) {
  const logs = Array.isArray(lead.activityLogs) ? lead.activityLogs : [];
  const emailSentCount = logs.filter((entry) => entry.type === "emailSent").length;
  const emailReceivedCount = logs.filter((entry) => entry.type === "emailReceived").length;
  const callCount = toNumber(lead.callCount || lead.dialAttempts);
  const liveCount = toNumber(lead.liveContacts);
  const callVmCount = toNumber(lead.callVmCount);
  const callNcCount = toNumber(lead.callNcCount);
  const textSentCount = toNumber(lead.textSentCount || lead.textsAttempted);
  const textReceivedCount = toNumber(lead.textReceivedCount);

  if (
    callCount <= 0 &&
    liveCount <= 0 &&
    callVmCount <= 0 &&
    callNcCount <= 0 &&
    textSentCount <= 0 &&
    textReceivedCount <= 0 &&
    emailSentCount <= 0 &&
    emailReceivedCount <= 0
  ) {
    return "Activity: none logged yet";
  }

  return `Calls: ${callCount} (Live/VM/NC ${liveCount}/${callVmCount}/${callNcCount}) | Text S/R ${textSentCount}/${textReceivedCount} | Email S/R ${emailSentCount}/${emailReceivedCount}`;
}

function formatDate(timestamp) {
  if (!timestamp) return "n/a";
  return new Date(timestamp).toLocaleDateString("en-US", {
    year: "numeric",
    month: "short",
    day: "numeric",
  });
}

function toBool(value) {
  if (typeof value === "boolean") return value;
  if (typeof value === "string") return ["true", "1", "yes"].includes(value.toLowerCase());
  return Boolean(value);
}

function parseCurrency(value) {
  if (typeof value === "number") return toNumber(value);
  const cleaned = String(value || "").replace(/[^0-9.-]/g, "");
  const numeric = Number(cleaned);
  if (!Number.isFinite(numeric)) return 0;
  if (numeric < 0) return 0;
  return Math.round(numeric);
}

function toNumber(value) {
  const numeric = Number(value);
  if (!Number.isFinite(numeric)) return 0;
  if (numeric < 0) return 0;
  return Math.round(numeric);
}

function formatCurrency(value) {
  return new Intl.NumberFormat("en-US", {
    style: "currency",
    currency: "USD",
    maximumFractionDigits: 0,
  }).format(value || 0);
}

function trim(value, limit) {
  if (value.length <= limit) return value;
  return `${value.slice(0, limit - 1)}...`;
}
