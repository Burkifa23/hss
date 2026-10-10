import { initializeApp } from "https://www.gstatic.com/firebasejs/10.12.2/firebase-app.js";
import {
  getDatabase,
  ref,
  set,
  onValue,
} from "https://www.gstatic.com/firebasejs/10.12.2/firebase-database.js";
import {
  getAuth,
  signInWithEmailAndPassword,
  signOut,
  onAuthStateChanged,
} from "https://www.gstatic.com/firebasejs/10.12.2/firebase-auth.js";

// ============================================================================
// 1. CONFIGURATION
// ============================================================================
const CONFIG = {
  firebase: {
    apiKey: "AIzaSyAWdmNxekMefIGAkFDSdUU-Paq7xqwwNig",
    authDomain: "hss-volunteer-6546a.firebaseapp.com",
    databaseURL:
      "https://hss-volunteer-6546a-default-rtdb.europe-west1.firebasedatabase.app",
    projectId: "hss-volunteer-6546a",
    storageBucket: "hss-volunteer-6546a.firebasestorage.app",
    messagingSenderId: "278719960058",
    appId: "1:278719960058:web:92eb1bb11e8e6ad6a08a43",
    adminUid: "se6U6S4Iv3c2z27mdDLlTz3zrIG3",
    dbNode: "hss_roster_data_final",
  },
  keys: {
    session: "ashesi_stu_session",
    cache: "ashesi_offline_cache",
  },
};

// ============================================================================
// 2. CORE SERVICES
// ============================================================================
const fbApp = initializeApp(CONFIG.firebase);
const db = getDatabase(fbApp);
const auth = getAuth(fbApp);
const dbRef = ref(db, CONFIG.firebase.dbNode);

const Security = {
  sanitize: (str) => window.DOMPurify.sanitize(String(str).trim()),
  validateId: (id) => /^\d{8}$/.test(id),
  attr: (str) => String(str ?? "").replace(/"/g, "&quot;"),
};

// "HH:MM" helpers for duration calculation and validation
const Time = {
  toMins: (t) => {
    if (!t) return null;
    const [h, m] = t.split(":").map(Number);
    return h * 60 + m;
  },
  // Minutes between activity start and end, or null if incomplete/invalid
  duration: (v) => {
    const s = Time.toMins(v.start);
    const e = Time.toMins(v.end);
    return s !== null && e !== null && e > s ? e - s : null;
  },
  fmt: (mins) =>
    mins === null ? "-" : `${Math.floor(mins / 60)}h ${mins % 60}m`,
  hours: (mins) => (mins === null ? "" : Math.round((mins / 60) * 100) / 100),
};

const Auth = {
  admin: false,
  student: null,

  init: function (callback) {
    const stored = localStorage.getItem(CONFIG.keys.session);
    if (stored) this.student = JSON.parse(stored);

    onAuthStateChanged(auth, (user) => {
      this.admin = user && user.uid === CONFIG.firebase.adminUid;
      callback();
    });
  },
  loginStudent: function (name, id) {
    if (!Security.validateId(id))
      throw new Error("Student ID must be exactly 8 digits.");
    this.student = { name: Security.sanitize(name), id: Security.sanitize(id) };
    localStorage.setItem(CONFIG.keys.session, JSON.stringify(this.student));
  },
  logoutStudent: function () {
    this.student = null;
    localStorage.removeItem(CONFIG.keys.session);
  },
  loginAdmin: async function (email, pw) {
    return signInWithEmailAndPassword(auth, email, pw);
  },
  logoutAdmin: function () {
    signOut(auth);
  },
};

const UI = {
  toastTimer: null,
  toast: function (msg, type = "success") {
    const t = document.getElementById("app-toast");
    t.querySelector(".toast-msg").textContent = msg;
    t.className = `toast show ${type}`;
    t.querySelector(".toast-icon").className =
      type === "error"
        ? "ph ph-warning-circle toast-icon"
        : "ph ph-check-circle toast-icon";
    clearTimeout(this.toastTimer);
    this.toastTimer = setTimeout(() => {
      t.classList.remove("show");
    }, 3000);
  },
  openModal: function (html) {
    this.closeModal();
    const bg = document.createElement("div");
    bg.className = "modal-bg";
    bg.id = "app-modal";
    bg.innerHTML = `<div class="modal">${html}</div>`;
    bg.addEventListener("click", (e) => {
      if (e.target === bg) this.closeModal();
    });
    document.getElementById("modals-container").appendChild(bg);
  },
  closeModal: function () {
    const m = document.getElementById("app-modal");
    if (m) m.remove();
  },
  switchTab: function (tabId) {
    document
      .querySelectorAll(".nav-tab")
      .forEach((el) => el.classList.remove("active"));
    document
      .querySelector(`.nav-tab[data-target="${tabId}"]`)
      .classList.add("active");
    document
      .querySelectorAll(".app-view")
      .forEach((el) => el.classList.remove("active"));
    document.getElementById(`view-${tabId}`).classList.add("active");
  },
  // Generic Dynamic Form Builder for CRUD Operations
  openFormModal: function (title, fields, submitText, onSubmit) {
    let fieldsHtml = fields
      .map(
        (f) => `
      <div class="form-group">
        <label>${f.label}</label>
        <input type="${f.type}" id="${f.id}" value="${Security.attr(f.value)}" placeholder="${Security.attr(f.placeholder)}" ${f.attrs || ""}>
      </div>
    `,
      )
      .join("");

    this.openModal(`
      <h3 class="modal-title">${title}</h3>
      ${fieldsHtml}
      <div class="modal-footer">
        <button class="btn-outline" id="btn-cancel-modal">Cancel</button>
        <button class="btn-primary" id="btn-submit-modal">${submitText}</button>
      </div>
    `);

    document.getElementById("btn-cancel-modal").onclick = () =>
      this.closeModal();
    document.getElementById("btn-submit-modal").onclick = async () => {
      let values = {};
      let valid = true;
      fields.forEach((f) => {
        const val = Security.sanitize(document.getElementById(f.id).value);
        if (f.required && !val) valid = false;
        values[f.id] = val;
      });
      if (!valid) return this.toast("Please fill in required fields.", "error");
      // onSubmit returns false to keep the modal open (validation failed)
      if ((await onSubmit(values)) === false) return;
      this.closeModal();
    };
  },
};

const ExcelService = {
  exportAOA: function (aoaData, sheetName, fileName) {
    const ws = XLSX.utils.aoa_to_sheet(aoaData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, sheetName);
    XLSX.writeFile(wb, `${fileName}.xlsx`);
  },
};

// ============================================================================
// 3. APPLICATION STATE & DATA FETCHING
// ============================================================================
let AppState = {
  events: [],
  activeEventId: null,
  trackerActiveId: null,
  masterFilter: { event: "", team: "" },
};
function genId() {
  return (
    "id_" + Date.now().toString(36) + Math.random().toString(36).slice(2, 6)
  );
}

// Firebase drops empty arrays, so restore them before rendering
function normalizeEvents(events) {
  return (events || []).map((ev) => ({
    ...ev,
    teams: (ev.teams || []).map((t) => ({
      ...t,
      volunteers: t.volunteers || [],
      leads: t.leads || [],
    })),
  }));
}

function currentEvent() {
  return (
    AppState.events.find((e) => e.id === AppState.activeEventId) ||
    AppState.events[0]
  );
}

function isLead(team) {
  return !!Auth.student && team.leads.includes(Auth.student.id);
}

// Only admins and the team's designated leads may log times
function canTrack(team) {
  return Auth.admin || isLead(team);
}

onValue(dbRef, (snapshot) => {
  const data = snapshot.val() || {};
  AppState.events = normalizeEvents(data.events);
  AppState.activeEventId = data.activeEventId || null;
  localStorage.setItem(CONFIG.keys.cache, JSON.stringify(data));
  renderAll();
});

async function syncState() {
  try {
    await set(dbRef, {
      events: AppState.events,
      activeEventId: AppState.activeEventId,
    });
  } catch (e) {
    UI.toast("Error saving data. Check permissions.", "error");
  }
}

function renderAll() {
  renderDashboard();
  renderPicker();
  renderTracker();
  renderMaster();
}

// ============================================================================
// 4. VIEW RENDERERS
// ============================================================================

function renderDashboard() {
  const badge = document.getElementById("user-status-badge");
  if (Auth.admin)
    badge.innerHTML = `<i class="ph ph-shield-check"></i> Administrator Access`;
  else if (Auth.student)
    badge.innerHTML = `<i class="ph ph-user"></i> ${Auth.student.name} (${Auth.student.id})`;
  else
    badge.innerHTML = `<i class="ph ph-user-minus"></i> Guest Mode - Sign in to join`;
}

function renderPicker() {
  const actions = document.getElementById("picker-hdr-actions");
  if (Auth.admin) {
    actions.innerHTML = `
      <button class="btn-outline btn-sm" data-action="new-event"><i class="ph ph-plus"></i> Create Task / Event</button>
      <button class="btn-danger btn-sm" data-action="logout-admin"><i class="ph ph-sign-out"></i> Exit Admin</button>
    `;
  } else if (Auth.student) {
    actions.innerHTML = `
      <button class="btn-outline btn-sm" data-action="logout-student"><i class="ph ph-arrows-left-right"></i> Switch Student</button>
      <button class="btn-outline btn-sm" data-action="login-modal"><i class="ph ph-lock-key"></i> Admin Login</button>
    `;
  } else {
    actions.innerHTML = `<button class="btn-outline btn-sm" data-action="login-modal"><i class="ph ph-sign-in"></i> Login to Join</button>`;
  }

  if (!AppState.events || AppState.events.length === 0) {
    document.getElementById("picker-main").innerHTML =
      `<p style="padding:1rem;">No events or tasks available. Admin must create one.</p>`;
    document.getElementById("picker-event-bar").innerHTML = "";
    return;
  }

  const ev = currentEvent();

  let evOpts = AppState.events
    .map(
      (e) =>
        `<option value="${e.id}" ${e.id === ev.id ? "selected" : ""}>${e.name}</option>`,
    )
    .join("");
  document.getElementById("picker-event-bar").innerHTML = `
    <div class="event-bar-group">
      <strong>Select:</strong>
      <select class="event-bar-select" id="event-selector">${evOpts}</select>
    </div>
    ${
      Auth.admin
        ? `
      <div class="event-bar-actions">
        <button class="icon-btn" data-action="edit-event" data-evid="${ev.id}" title="Edit Event"><i class="ph ph-pencil-simple"></i></button>
        <button class="icon-btn danger" data-action="delete-event" data-evid="${ev.id}" title="Delete Event"><i class="ph ph-trash"></i></button>
      </div>
    `
        : ""
    }
  `;

  let html = `<div class="team-grid">`;
  (ev.teams || []).forEach((team) => {
    const isFull = team.volunteers.length >= team.max;

    let volHtml =
      team.volunteers.length === 0
        ? `<p style="font-size:0.85rem;color:var(--text-muted)">No sign-ups yet.</p>`
        : team.volunteers
            .map((v) => {
              const lead = team.leads.includes(v.id);
              const adminBtns = Auth.admin
                ? `<button class="btn-chip ${lead ? "on" : ""}" data-action="toggle-lead" data-tid="${team.id}" data-vid="${v.id}" title="${lead ? "Remove Lead" : "Make Group Lead"}"><i class="ph ${lead ? "ph-star-fill" : "ph-star"}"></i></button>
                   <button class="btn-chip" data-action="edit-vol" data-tid="${team.id}" data-vid="${v.id}" title="Edit Volunteer"><i class="ph ph-pencil-simple"></i></button>
                   <button class="btn-remove" data-action="rm-vol" data-tid="${team.id}" data-vid="${v.id}" title="Remove Volunteer">&times;</button>`
                : "";
              return `<span class="vol-chip ${lead ? "lead" : ""}">${lead ? `<i class="ph ph-star-fill lead-icon" title="Group Lead"></i>` : ""}${v.name} (${v.id}) ${adminBtns}</span>`;
            })
            .join("");

    const isMember =
      Auth.student && team.volunteers.some((v) => v.id === Auth.student.id);
    let joinHtml;
    if (isMember)
      joinHtml = `<button class="btn-join btn-leave" data-action="leave-team" data-tid="${team.id}"><i class="ph ph-sign-out"></i> Leave Team</button>`;
    else if (isFull)
      joinHtml = `<p style="color:var(--ashesi-red); font-weight:bold;"><i class="ph ph-prohibit"></i> Team Full</p>`;
    else
      joinHtml = `<button class="btn-join" data-action="join-team" data-tid="${team.id}">Join Team</button>`;
    if (Auth.admin)
      joinHtml += `<button class="btn-outline btn-sm" style="margin-top:0.5rem;" data-action="add-vol" data-tid="${team.id}"><i class="ph ph-user-plus"></i> Add Volunteer</button>`;

    html += `
      <div class="team-card">
        <div class="card-top-bar"></div>
        <div class="card-head">
          <div class="team-name">${team.name}</div>
          <div style="font-size:0.85rem; color:var(--text-muted); margin-bottom:0.5rem;">${team.desc}</div>
          <div class="badge-col">${team.volunteers.length} / ${team.max} Slots Filled</div>
          ${
            Auth.admin
              ? `
            <div class="team-actions">
              <button class="icon-btn" data-action="edit-team" data-tid="${team.id}"><i class="ph ph-pencil-simple"></i></button>
              <button class="icon-btn danger" data-action="delete-team" data-tid="${team.id}"><i class="ph ph-trash"></i></button>
            </div>
          `
              : ""
          }
        </div>
        <div class="roster-body">
          <div style="margin-bottom: 1rem;">${volHtml}</div>
          ${joinHtml}
        </div>
      </div>
    `;
  });
  if (Auth.admin) {
    html += `
      <div class="team-card" style="display:flex;align-items:center;justify-content:center;min-height:150px; border: 2px dashed var(--border-color); background:transparent; box-shadow:none;">
        <button class="btn-outline" data-action="new-team"><i class="ph ph-plus"></i> Add Team / Group</button>
      </div>`;
  }
  html += `</div>`;
  document.getElementById("picker-main").innerHTML = html;
}

function renderTracker() {
  const grid = document.getElementById("tracker-tasks-grid");
  const detail = document.getElementById("tracker-detail-view");

  if (Auth.admin) {
    document.getElementById("tracker-hdr-actions").innerHTML =
      `<button class="btn-primary" data-action="new-event"><i class="ph ph-plus"></i> Create Trackable Task</button>`;
  } else {
    document.getElementById("tracker-hdr-actions").innerHTML = "";
  }

  if (!AppState.trackerActiveId) {
    grid.style.display = "grid";
    detail.style.display = "none";

    if (!AppState.events || AppState.events.length === 0) {
      grid.innerHTML = "<p>No tasks to track yet.</p>";
      return;
    }

    grid.innerHTML = AppState.events
      .map(
        (ev) => `
      <div class="tracker-card" data-action="open-tracker" data-evid="${ev.id}">
        <h3>${ev.name}</h3>
        <p style="color:var(--text-muted); font-size:0.85rem;"><i class="ph ph-calendar-blank"></i> ${ev.date || "TBD"}</p>
        <p style="color:var(--text-muted); font-size:0.85rem;"><i class="ph ph-users"></i> ${(ev.teams || []).length} Sub-Teams</p>
      </div>
    `,
      )
      .join("");
  } else {
    const ev = AppState.events.find((e) => e.id === AppState.trackerActiveId);
    if (!ev) {
      // Event was deleted while open
      AppState.trackerActiveId = null;
      return renderTracker();
    }
    grid.style.display = "none";
    detail.style.display = "block";

    document.getElementById("tr-detail-name").innerText = ev.name;
    document.getElementById("tr-breadcrumb-name").innerText = `/ ${ev.name}`;

    const editable = ev.teams.filter(canTrack).length;
    document.getElementById("tr-perm-note").innerHTML = Auth.admin
      ? ""
      : editable
        ? `<i class="ph ph-star-fill"></i> You are a Group Lead for ${editable} team(s) here. You can log times for those teams only.`
        : `<i class="ph ph-lock"></i> Read-only. Only Group Leads (assigned by an Admin) can log times.`;

    let rowsHtml = "";
    let counter = 1;
    ev.teams.forEach((team) => {
      const dis = canTrack(team) ? "" : "disabled";
      team.volunteers.forEach((v) => {
        const lead = team.leads.includes(v.id);
        rowsHtml += `
          <tr>
            <td>${counter++}</td>
            <td><strong>${v.name}</strong>${lead ? ` <i class="ph ph-star-fill lead-icon" title="Group Lead"></i>` : ""}</td>
            <td>${v.id}</td>
            <td><span class="badge-col" style="background:var(--ashesi-red-light); color:var(--ashesi-red);">${team.name}</span></td>
            <td><input type="time" class="rep-time" data-tid="${team.id}" data-vid="${v.id}" value="${v.reportTime || ""}" ${dis}></td>
            <td><input type="time" class="start-time" data-tid="${team.id}" data-vid="${v.id}" value="${v.start || ""}" ${dis}></td>
            <td><input type="time" class="end-time" data-tid="${team.id}" data-vid="${v.id}" value="${v.end || ""}" ${dis}></td>
            <td><strong>${Time.fmt(Time.duration(v))}</strong></td>
          </tr>
        `;
      });
    });
    document.getElementById("tr-members-tbody").innerHTML =
      rowsHtml ||
      '<tr><td colspan="8">No volunteers assigned to this task.</td></tr>';
  }
}

// Volunteer records visible under the current Master Sheet filters
function masterRecords() {
  const f = AppState.masterFilter;
  const out = [];
  AppState.events.forEach((ev) => {
    if (f.event && ev.id !== f.event) return;
    ev.teams.forEach((t) => {
      if (f.team && t.name !== f.team) return;
      t.volunteers.forEach((v) => out.push({ ev, t, v }));
    });
  });
  return out;
}

function renderMaster() {
  const thead = document.getElementById("master-head");
  const tbody = document.getElementById("master-body");

  const f = AppState.masterFilter;

  // Drop filters pointing at deleted events/teams
  if (f.event && !AppState.events.some((e) => e.id === f.event)) f.event = "";
  const teamNames = [
    ...new Set(
      AppState.events
        .filter((e) => !f.event || e.id === f.event)
        .flatMap((e) => e.teams.map((t) => t.name)),
    ),
  ].sort();
  if (f.team && !teamNames.includes(f.team)) f.team = "";

  const evOpts = AppState.events
    .map(
      (e) =>
        `<option value="${e.id}" ${e.id === f.event ? "selected" : ""}>${e.name}</option>`,
    )
    .join("");
  const tmOpts = teamNames
    .map(
      (n) =>
        `<option value="${Security.attr(n)}" ${n === f.team ? "selected" : ""}>${n}</option>`,
    )
    .join("");
  document.getElementById("master-filters").innerHTML = `
    <div class="filter-group">
      <label for="master-filter-event">Event</label>
      <select id="master-filter-event"><option value="">All Events</option>${evOpts}</select>
    </div>
    <div class="filter-group">
      <label for="master-filter-team">Team</label>
      <select id="master-filter-team"><option value="">All Teams</option>${tmOpts}</select>
    </div>
  `;

  const records = masterRecords();
  const totalMins = records.reduce(
    (sum, r) => sum + (Time.duration(r.v) || 0),
    0,
  );
  const logged = records.filter((r) => Time.duration(r.v) !== null).length;
  const students = new Set(records.map((r) => r.v.id)).size;
  document.getElementById("master-stats").innerHTML = `
    <div class="stat-card"><div class="stat-val">${records.length}</div><div class="stat-lbl">Sign-ups</div></div>
    <div class="stat-card"><div class="stat-val">${students}</div><div class="stat-lbl">Unique Students</div></div>
    <div class="stat-card"><div class="stat-val">${Time.hours(totalMins) || 0}</div><div class="stat-lbl">Total Hours Logged</div></div>
    <div class="stat-card"><div class="stat-val">${logged} / ${records.length}</div><div class="stat-lbl">Shifts Completed</div></div>
  `;

  thead.innerHTML = `<tr><th>Event / Task</th><th>Team</th><th>Volunteer Name</th><th>Student ID</th><th>Report Time</th><th>Start Time</th><th>End Time</th><th>Total</th></tr>`;
  const rows = records
    .map(
      ({ ev, t, v }) => `<tr>
          <td><strong>${ev.name}</strong></td>
          <td>${t.name}</td>
          <td>${v.name}</td>
          <td>${v.id}</td>
          <td>${v.reportTime || "-"}</td>
          <td>${v.start || "-"}</td>
          <td>${v.end || "-"}</td>
          <td><strong>${Time.fmt(Time.duration(v))}</strong></td>
        </tr>`,
    )
    .join("");
  tbody.innerHTML =
    rows || '<tr><td colspan="8">No matching records.</td></tr>';
}

// ============================================================================
// 5. GLOBAL EVENT DELEGATOR
// ============================================================================
document.addEventListener("DOMContentLoaded", () => {
  Auth.init(renderAll);

  document.body.addEventListener("click", async (e) => {
    // Tab Switching
    const tabTarget = e.target.closest("[data-target]");
    if (
      tabTarget &&
      (tabTarget.classList.contains("nav-tab") ||
        tabTarget.classList.contains("dash-card"))
    ) {
      UI.switchTab(tabTarget.getAttribute("data-target"));
    }

    const actionBtn = e.target.closest("[data-action]");
    if (!actionBtn) return;
    const action = actionBtn.getAttribute("data-action");

    // --- AUTHENTICATION ---
    if (action === "login-modal") {
      UI.openModal(`
        <h3 class="modal-title">Authentication</h3>
        <div style="display:flex; gap:10px; margin-bottom:1.5rem;">
          <button class="btn-primary" id="btn-show-student">Student</button>
          <button class="btn-outline" id="btn-show-admin">Administrator</button>
        </div>
        <div id="auth-forms">
          <div class="form-group"><label>Full Name</label><input type="text" id="stu-name" placeholder="John Doe"></div>
          <div class="form-group"><label>Student ID (8 digits)</label><input type="text" id="stu-id" placeholder="12345678"></div>
          <button class="btn-primary" style="width:100%" id="do-stu-login"><i class="ph ph-sign-in"></i> Login as Student</button>
        </div>
      `);

      document.getElementById("btn-show-admin").onclick = () => {
        document.getElementById("auth-forms").innerHTML = `
          <div class="form-group"><label>Email</label><input type="email" id="adm-email"></div>
          <div class="form-group"><label>Password</label><input type="password" id="adm-pw"></div>
          <button class="btn-primary" style="width:100%" id="do-adm-login"><i class="ph ph-lock-key"></i> Login as Admin</button>
        `;
        document.getElementById("do-adm-login").onclick = async () => {
          try {
            await Auth.loginAdmin(
              document.getElementById("adm-email").value,
              document.getElementById("adm-pw").value,
            );
            UI.closeModal();
            UI.toast("Admin Logged In");
          } catch (err) {
            UI.toast("Invalid Credentials", "error");
          }
        };
      };

      document.getElementById("do-stu-login").onclick = () => {
        try {
          Auth.loginStudent(
            document.getElementById("stu-name").value,
            document.getElementById("stu-id").value,
          );
          UI.closeModal();
          renderAll();
          UI.toast(`Welcome ${Auth.student.name}`);
        } catch (err) {
          UI.toast(err.message, "error");
        }
      };
    }

    if (action === "logout-admin") {
      Auth.logoutAdmin();
    }
    if (action === "logout-student") {
      Auth.logoutStudent();
      renderAll();
      UI.toast("Logged out.");
    }

    // --- VOLUNTEER ACTIONS ---
    if (action === "join-team") {
      if (!Auth.student)
        return UI.toast("You must be logged in as a student to join.", "error");
      const team = currentEvent().teams.find(
        (t) => t.id === actionBtn.getAttribute("data-tid"),
      );

      if (team.volunteers.some((v) => v.id === Auth.student.id))
        return UI.toast("You are already in this team.", "error");
      if (team.volunteers.length >= team.max)
        return UI.toast("This team is full.", "error");

      team.volunteers.push({
        id: Auth.student.id,
        name: Auth.student.name,
        reportTime: "",
        start: "",
        end: "",
      });
      await syncState();
      UI.toast(`Successfully joined ${team.name}`);
    }

    if (action === "leave-team") {
      if (!Auth.student) return;
      const team = currentEvent().teams.find(
        (t) => t.id === actionBtn.getAttribute("data-tid"),
      );
      if (!confirm(`Leave ${team.name}? Any logged times will be lost.`))
        return;
      team.volunteers = team.volunteers.filter(
        (v) => v.id !== Auth.student.id,
      );
      team.leads = team.leads.filter((id) => id !== Auth.student.id);
      await syncState();
      UI.toast(`You left ${team.name}`);
    }

    if (action === "rm-vol") {
      if (!Auth.admin) return;
      const team = currentEvent().teams.find(
        (t) => t.id === actionBtn.getAttribute("data-tid"),
      );
      const vId = actionBtn.getAttribute("data-vid");
      if (!confirm("Remove this volunteer from the team?")) return;
      team.volunteers = team.volunteers.filter((v) => v.id !== vId);
      team.leads = team.leads.filter((id) => id !== vId);
      await syncState();
      UI.toast("Volunteer Removed");
    }

    // Volunteer Create (admin adds a student manually)
    if (action === "add-vol") {
      if (!Auth.admin) return;
      const team = currentEvent().teams.find(
        (t) => t.id === actionBtn.getAttribute("data-tid"),
      );
      UI.openFormModal(
        `Add Volunteer to ${team.name}`,
        [
          { id: "vol-name", label: "Full Name", type: "text", required: true },
          {
            id: "vol-id",
            label: "Student ID (8 digits)",
            type: "text",
            required: true,
          },
        ],
        "Add Volunteer",
        async (vals) => {
          if (!Security.validateId(vals["vol-id"])) {
            UI.toast("Student ID must be exactly 8 digits.", "error");
            return false;
          }
          if (team.volunteers.some((v) => v.id === vals["vol-id"])) {
            UI.toast("That student is already in this team.", "error");
            return false;
          }
          if (
            team.volunteers.length >= team.max &&
            !confirm("Team is at capacity. Add anyway?")
          )
            return false;
          team.volunteers.push({
            id: vals["vol-id"],
            name: vals["vol-name"],
            reportTime: "",
            start: "",
            end: "",
          });
          await syncState();
          UI.toast("Volunteer Added");
        },
      );
    }

    // Volunteer Update
    if (action === "edit-vol") {
      if (!Auth.admin) return;
      const team = currentEvent().teams.find(
        (t) => t.id === actionBtn.getAttribute("data-tid"),
      );
      const vol = team.volunteers.find(
        (v) => v.id === actionBtn.getAttribute("data-vid"),
      );
      UI.openFormModal(
        "Edit Volunteer",
        [
          {
            id: "vol-name",
            label: "Full Name",
            type: "text",
            value: vol.name,
            required: true,
          },
          {
            id: "vol-id",
            label: "Student ID (8 digits)",
            type: "text",
            value: vol.id,
            required: true,
          },
        ],
        "Save Changes",
        async (vals) => {
          const newId = vals["vol-id"];
          if (!Security.validateId(newId)) {
            UI.toast("Student ID must be exactly 8 digits.", "error");
            return false;
          }
          if (
            newId !== vol.id &&
            team.volunteers.some((v) => v.id === newId)
          ) {
            UI.toast("Another volunteer in this team has that ID.", "error");
            return false;
          }
          team.leads = team.leads.map((id) => (id === vol.id ? newId : id));
          vol.id = newId;
          vol.name = vals["vol-name"];
          await syncState();
          UI.toast("Volunteer Updated");
        },
      );
    }

    // Group Lead assignment
    if (action === "toggle-lead") {
      if (!Auth.admin) return;
      const team = currentEvent().teams.find(
        (t) => t.id === actionBtn.getAttribute("data-tid"),
      );
      const vId = actionBtn.getAttribute("data-vid");
      const wasLead = team.leads.includes(vId);
      team.leads = wasLead
        ? team.leads.filter((id) => id !== vId)
        : [...team.leads, vId];
      await syncState();
      UI.toast(wasLead ? "Group Lead removed" : "Group Lead assigned");
    }

    // --- FULL ADMIN CRUD OPERATIONS ---
    // Event Create
    if (action === "new-event") {
      UI.openFormModal(
        "Create New Task / Event",
        [
          { id: "ev-name", label: "Event Name", type: "text", required: true },
          { id: "ev-date", label: "Date", type: "date" },
        ],
        "Create Event",
        async (vals) => {
          const newEv = {
            id: genId(),
            name: vals["ev-name"],
            date: vals["ev-date"],
            teams: [],
          };
          if (!AppState.events) AppState.events = [];
          AppState.events.push(newEv);
          AppState.activeEventId = newEv.id;
          await syncState();
          UI.toast("Task / Event Created");
        },
      );
    }

    // Event Update
    if (action === "edit-event") {
      const evId = actionBtn.getAttribute("data-evid");
      const ev = AppState.events.find((e) => e.id === evId);
      UI.openFormModal(
        "Edit Task / Event",
        [
          {
            id: "ev-name",
            label: "Event Name",
            type: "text",
            value: ev.name,
            required: true,
          },
          { id: "ev-date", label: "Date", type: "date", value: ev.date || "" },
        ],
        "Save Changes",
        async (vals) => {
          ev.name = vals["ev-name"];
          ev.date = vals["ev-date"];
          await syncState();
          UI.toast("Event Updated");
        },
      );
    }

    // Event Delete
    if (action === "delete-event") {
      if (
        !confirm(
          "Are you sure? This will delete the event and ALL team data inside it.",
        )
      )
        return;
      const evId = actionBtn.getAttribute("data-evid");
      AppState.events = AppState.events.filter((e) => e.id !== evId);
      if (AppState.activeEventId === evId)
        AppState.activeEventId = AppState.events.length
          ? AppState.events[0].id
          : null;
      await syncState();
      UI.toast("Event Deleted");
    }

    // Team Create
    if (action === "new-team") {
      UI.openFormModal(
        "Add Team / Group",
        [
          { id: "tm-name", label: "Team Name", type: "text", required: true },
          { id: "tm-desc", label: "Description", type: "text" },
          {
            id: "tm-max",
            label: "Capacity (Max Volunteers)",
            type: "number",
            value: "5",
            required: true,
          },
        ],
        "Create Team",
        async (vals) => {
          const max = parseInt(vals["tm-max"]);
          if (!(max > 0)) {
            UI.toast("Capacity must be at least 1.", "error");
            return false;
          }
          currentEvent().teams.push({
            id: genId(),
            name: vals["tm-name"],
            desc: vals["tm-desc"],
            max,
            volunteers: [],
            leads: [],
          });
          await syncState();
          UI.toast("Team Created");
        },
      );
    }

    // Team Update
    if (action === "edit-team") {
      const tId = actionBtn.getAttribute("data-tid");
      const team = currentEvent().teams.find((t) => t.id === tId);
      UI.openFormModal(
        "Edit Team",
        [
          {
            id: "tm-name",
            label: "Team Name",
            type: "text",
            value: team.name,
            required: true,
          },
          {
            id: "tm-desc",
            label: "Description",
            type: "text",
            value: team.desc || "",
          },
          {
            id: "tm-max",
            label: "Capacity (Max Volunteers)",
            type: "number",
            value: team.max,
            required: true,
          },
        ],
        "Save Changes",
        async (vals) => {
          const max = parseInt(vals["tm-max"]);
          if (!(max > 0)) {
            UI.toast("Capacity must be at least 1.", "error");
            return false;
          }
          team.name = vals["tm-name"];
          team.desc = vals["tm-desc"];
          team.max = max;
          await syncState();
          UI.toast("Team Updated");
        },
      );
    }

    // Team Delete
    if (action === "delete-team") {
      if (!confirm("Delete this entire team?")) return;
      const tId = actionBtn.getAttribute("data-tid");
      const ev = currentEvent();
      ev.teams = ev.teams.filter((t) => t.id !== tId);
      await syncState();
      UI.toast("Team Deleted");
    }

    // --- TRACKER NAVIGATION ---
    if (action === "open-tracker") {
      AppState.trackerActiveId = actionBtn.getAttribute("data-evid");
      renderTracker();
    }

    if (action === "close-tracker") {
      AppState.trackerActiveId = null;
      renderTracker();
    }

    // --- EXCEL (.XLSX) EXPORTS ---
    if (action === "export-tracker") {
      const ev = AppState.events.find((e) => e.id === AppState.trackerActiveId);
      let aoa = [
        [
          "No.",
          "Name",
          "Student ID",
          "Team",
          "Report Time",
          "Start Time",
          "End Time",
          "Total Mins",
          "Total Hours",
        ],
      ];
      let count = 1;
      ev.teams.forEach((t) => {
        t.volunteers.forEach((v) => {
          const mins = Time.duration(v);
          aoa.push([
            count++,
            v.name,
            v.id,
            t.name,
            v.reportTime || "",
            v.start || "",
            v.end || "",
            mins ?? "",
            Time.hours(mins),
          ]);
        });
      });
      ExcelService.exportAOA(
        aoa,
        "Task Tracker",
        `Tracker_${ev.name.replace(/\s+/g, "_")}`,
      );
      UI.toast("XLSX Downloaded!");
    }

    if (action === "export-master") {
      let aoa = [
        [
          "Event Name",
          "Team Name",
          "Volunteer Name",
          "Student ID",
          "Report Time",
          "Start",
          "End",
          "Total Mins",
          "Total Hours",
        ],
      ];
      // Exports what the current filters show
      masterRecords().forEach(({ ev, t, v }) => {
        const mins = Time.duration(v);
        aoa.push([
          ev.name,
          t.name,
          v.name,
          v.id,
          v.reportTime || "",
          v.start || "",
          v.end || "",
          mins ?? "",
          Time.hours(mins),
        ]);
      });
      ExcelService.exportAOA(aoa, "Master Roster", "Ashesi_Master_Roster");
      UI.toast("Master XLSX Downloaded!");
    }
  });

  // Track Input changes in Tracker Table using Event Delegation (No Inline OnChange)
  document.body.addEventListener("change", async (e) => {
    // Handle Dropdown change
    if (e.target.id === "event-selector") {
      AppState.activeEventId = e.target.value;
      await syncState();
      return;
    }

    // Master Sheet filters (local view state only, not synced)
    if (e.target.id === "master-filter-event") {
      AppState.masterFilter.event = e.target.value;
      return renderMaster();
    }
    if (e.target.id === "master-filter-team") {
      AppState.masterFilter.team = e.target.value;
      return renderMaster();
    }

    // Handle Time Input changes inside Tracker
    if (
      e.target.classList.contains("rep-time") ||
      e.target.classList.contains("start-time") ||
      e.target.classList.contains("end-time")
    ) {
      const ev = AppState.events.find(
        (ev) => ev.id === AppState.trackerActiveId,
      );
      const team = ev.teams.find(
        (t) => t.id === e.target.getAttribute("data-tid"),
      );
      const vol = team.volunteers.find(
        (v) => v.id === e.target.getAttribute("data-vid"),
      );

      if (!canTrack(team)) {
        UI.toast("Only this team's Group Lead can log times.", "error");
        return renderTracker();
      }

      const val = Security.sanitize(e.target.value);
      const field = e.target.classList.contains("rep-time")
        ? "reportTime"
        : e.target.classList.contains("start-time")
          ? "start"
          : "end";

      const next = { ...vol, [field]: val };
      if (
        next.start &&
        next.end &&
        Time.toMins(next.end) <= Time.toMins(next.start)
      ) {
        UI.toast("End time must be later than start time.", "error");
        e.target.value = vol[field] || "";
        return;
      }

      vol[field] = val;
      await syncState();
    }
  });
});
