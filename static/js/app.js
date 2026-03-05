/**
 * EXO Tools – Frontend JavaScript
 * Handles the People Picker, permission search, filtering, and CSV export.
 */

"use strict";

// ---------------------------------------------------------------------------
// State
// ---------------------------------------------------------------------------
let selectedUser   = null;   // { id, displayName, upn }
let allPermissions = [];     // raw results from the API
let activeType     = "all";  // current permission-type filter
let activeText     = "";     // current text filter

// ---------------------------------------------------------------------------
// DOM ready
// ---------------------------------------------------------------------------
document.addEventListener("DOMContentLoaded", () => {
  initPeoplePicker();
  initPermissionCheck();
  initFilters();
  initExport();
});

// ---------------------------------------------------------------------------
// People Picker
// ---------------------------------------------------------------------------
function initPeoplePicker() {
  const input    = document.getElementById("userSearch");
  const dropdown = document.getElementById("searchDropdown");
  const spinner  = document.getElementById("searchSpinner");

  let debounceTimer = null;

  input.addEventListener("input", () => {
    const q = input.value.trim();
    clearTimeout(debounceTimer);

    if (q.length < 2) {
      hideDropdown();
      return;
    }

    debounceTimer = setTimeout(async () => {
      spinner.classList.remove("d-none");
      try {
        const res  = await fetch(`/api/search-users?q=${encodeURIComponent(q)}`);
        const data = await res.json();
        renderDropdown(data.users || []);
      } catch {
        hideDropdown();
      } finally {
        spinner.classList.add("d-none");
      }
    }, 300);
  });

  // Keyboard navigation inside dropdown
  input.addEventListener("keydown", (e) => {
    const items = [...dropdown.querySelectorAll(".ac-item")];
    const idx   = items.indexOf(document.activeElement);

    if (e.key === "ArrowDown") {
      e.preventDefault();
      (items[idx + 1] || items[0])?.focus();
    } else if (e.key === "ArrowUp") {
      e.preventDefault();
      (items[idx - 1] || items[items.length - 1])?.focus();
    } else if (e.key === "Escape") {
      hideDropdown();
      input.focus();
    }
  });

  // Close on outside click
  document.addEventListener("click", (e) => {
    if (!input.contains(e.target) && !dropdown.contains(e.target)) {
      hideDropdown();
    }
  });
}

function renderDropdown(users) {
  const dropdown = document.getElementById("searchDropdown");

  if (!users.length) {
    dropdown.innerHTML = `
      <li>
        <div class="ac-item text-muted py-3 justify-content-center">
          <i class="bi bi-search me-2"></i>No users found
        </div>
      </li>`;
    dropdown.classList.remove("d-none");
    return;
  }

  dropdown.innerHTML = users.map((u) => `
    <li>
      <button
        class="ac-item"
        role="option"
        data-id="${escHtml(u.id)}"
        data-name="${escHtml(u.displayName)}"
        data-upn="${escHtml(u.userPrincipalName)}"
      >
        <div class="avatar">${escHtml(u.displayName.charAt(0).toUpperCase())}</div>
        <div class="flex-grow-1 overflow-hidden">
          <div class="fw-semibold text-truncate">${escHtml(u.displayName)}</div>
          <div class="text-muted small text-truncate">${escHtml(u.userPrincipalName)}</div>
          ${u.jobTitle ? `<div class="text-muted small text-truncate">${escHtml(u.jobTitle)}${u.department ? " · " + escHtml(u.department) : ""}</div>` : ""}
        </div>
      </button>
    </li>
  `).join("");

  // Attach click listeners
  dropdown.querySelectorAll(".ac-item").forEach((btn) => {
    btn.addEventListener("click", () => {
      selectUser(btn.dataset.id, btn.dataset.name, btn.dataset.upn);
    });
    btn.addEventListener("keydown", (e) => {
      if (e.key === "Enter" || e.key === " ") {
        e.preventDefault();
        selectUser(btn.dataset.id, btn.dataset.name, btn.dataset.upn);
      }
    });
  });

  dropdown.classList.remove("d-none");
}

function hideDropdown() {
  document.getElementById("searchDropdown").classList.add("d-none");
}

function selectUser(id, displayName, upn) {
  selectedUser = { id, displayName, upn };

  document.getElementById("userSearch").value = displayName;
  hideDropdown();

  document.getElementById("pillName").textContent = displayName;
  document.getElementById("pillUpn").textContent  = `<${upn}>`;
  document.getElementById("selectedPill").classList.remove("d-none");
  document.getElementById("checkBtn").disabled = false;

  // Reset results
  resetResults();
}

// ---------------------------------------------------------------------------
// Permission Check
// ---------------------------------------------------------------------------
function initPermissionCheck() {
  document.getElementById("checkBtn").addEventListener("click", runPermissionCheck);

  document.getElementById("clearBtn").addEventListener("click", () => {
    selectedUser = null;
    document.getElementById("userSearch").value = "";
    document.getElementById("selectedPill").classList.add("d-none");
    document.getElementById("checkBtn").disabled = true;
    resetResults();
  });
}

async function runPermissionCheck() {
  if (!selectedUser) return;

  showLoading();

  try {
    const res  = await fetch("/api/get-permissions", {
      method:  "POST",
      headers: { "Content-Type": "application/json" },
      body:    JSON.stringify({ userPrincipalName: selectedUser.upn }),
    });
    const data = await res.json();

    if (!res.ok || data.error) {
      showError(data.error || "Failed to retrieve permissions. Please try again.");
      return;
    }

    allPermissions = data.permissions || [];
    showResults();
  } catch {
    showError("A network error occurred. Please try again.");
  }
}

// ---------------------------------------------------------------------------
// Filters
// ---------------------------------------------------------------------------
function initFilters() {
  document.getElementById("textFilter").addEventListener("input", (e) => {
    activeText = e.target.value.toLowerCase();
    renderTable();
  });

  document.getElementById("typeFilter").addEventListener("click", (e) => {
    const btn = e.target.closest("[data-ptype]");
    if (!btn) return;
    activeType = btn.dataset.ptype;
    document
      .querySelectorAll("#typeFilter [data-ptype]")
      .forEach((b) => b.classList.toggle("active", b === btn));
    renderTable();
  });
}

function filterByType(type) {
  activeType = type;
  document
    .querySelectorAll("#typeFilter [data-ptype]")
    .forEach((b) => b.classList.toggle("active", b.dataset.ptype === type));
  renderTable();
}

// ---------------------------------------------------------------------------
// Rendering
// ---------------------------------------------------------------------------
function showLoading() {
  document.getElementById("loadingCard").classList.remove("d-none");
  document.getElementById("errorAlert").classList.add("d-none");
  document.getElementById("resultsCard").classList.add("d-none");
}

function showError(msg) {
  document.getElementById("loadingCard").classList.add("d-none");
  document.getElementById("errorAlert").classList.remove("d-none");
  document.getElementById("resultsCard").classList.add("d-none");
  document.getElementById("errorMsg").textContent = msg;
}

function resetResults() {
  document.getElementById("loadingCard").classList.add("d-none");
  document.getElementById("errorAlert").classList.add("d-none");
  document.getElementById("resultsCard").classList.add("d-none");
  allPermissions = [];
}

function showResults() {
  document.getElementById("loadingCard").classList.add("d-none");
  document.getElementById("errorAlert").classList.add("d-none");
  document.getElementById("resultsCard").classList.remove("d-none");

  document.getElementById("resultUserName").textContent = selectedUser.displayName;
  document.getElementById("totalBadge").textContent =
    `${allPermissions.length} permission${allPermissions.length !== 1 ? "s" : ""}`;

  // Reset filters
  activeType = "all";
  activeText = "";
  document.getElementById("textFilter").value = "";
  document.querySelectorAll("#typeFilter [data-ptype]").forEach((b) => {
    b.classList.toggle("active", b.dataset.ptype === "all");
  });

  renderSummaryStrip();
  renderTable();
}

function renderSummaryStrip() {
  const counts = {
    FullAccess:    allPermissions.filter((p) => p.PermissionType === "FullAccess").length,
    SendAs:        allPermissions.filter((p) => p.PermissionType === "SendAs").length,
    SendOnBehalf:  allPermissions.filter((p) => p.PermissionType === "SendOnBehalf").length,
  };

  const items = [
    { type: "FullAccess",   icon: "bi-folder2-open", label: "Full Access",    color: "danger",  count: counts.FullAccess   },
    { type: "SendAs",       icon: "bi-send",         label: "Send As",        color: "warning", count: counts.SendAs       },
    { type: "SendOnBehalf", icon: "bi-send-check",   label: "Send on Behalf", color: "info",    count: counts.SendOnBehalf },
  ];

  const strip = document.getElementById("summaryStrip");
  strip.innerHTML = items.map((item, i) => `
    <div
      class="col-md-4 summary-col ${item.count > 0 ? "" : "text-muted"}"
      role="${item.count > 0 ? "button" : ""}"
      data-filter-type="${escHtml(item.type)}"
      title="${item.count > 0 ? `Filter to ${escHtml(item.label)}` : ""}"
      ${i < items.length - 1 ? 'style="border-right:1px solid rgba(0,0,0,.07);"' : ""}
    >
      <i class="bi ${item.icon} text-${item.color}" style="font-size:1.4rem;"></i>
      <div class="fw-bold fs-4 lh-1 mt-1">${item.count}</div>
      <div class="small text-muted mt-1">${escHtml(item.label)}</div>
    </div>
  `).join("");

  // Attach click listeners after injection
  strip.querySelectorAll("[data-filter-type]").forEach((el) => {
    if (el.getAttribute("role") === "button") {
      el.addEventListener("click", () => filterByType(el.dataset.filterType));
    }
  });
}

const TYPE_BADGE = {
  FullAccess:   "badge-fullaccess",
  SendAs:       "badge-sendas",
  SendOnBehalf: "badge-sendonbehalf",
};

function renderTable() {
  let rows = allPermissions;

  if (activeType !== "all") {
    rows = rows.filter((p) => p.PermissionType === activeType);
  }
  if (activeText) {
    rows = rows.filter(
      (p) =>
        (p.MailboxDisplayName || "").toLowerCase().includes(activeText) ||
        (p.MailboxUPN || "").toLowerCase().includes(activeText)
    );
  }

  const tbody          = document.getElementById("permTableBody");
  const emptyState     = document.getElementById("emptyState");
  const filterEmpty    = document.getElementById("filterEmptyState");
  const tableContainer = tbody.closest(".table-responsive");

  // Hide all conditional sections first
  emptyState.classList.add("d-none");
  filterEmpty.classList.add("d-none");
  tableContainer.classList.remove("d-none");

  if (allPermissions.length === 0) {
    tableContainer.classList.add("d-none");
    emptyState.classList.remove("d-none");
    tbody.innerHTML = "";
    return;
  }

  if (rows.length === 0) {
    tableContainer.classList.add("d-none");
    filterEmpty.classList.remove("d-none");
    tbody.innerHTML = "";
    return;
  }

  tbody.innerHTML = rows.map((p) => `
    <tr>
      <td>
        <div class="fw-semibold">${escHtml(p.MailboxDisplayName || p.MailboxUPN || "—")}</div>
      </td>
      <td class="text-muted small">${escHtml(p.MailboxUPN || "")}</td>
      <td>
        <span class="badge rounded-pill px-2 py-1 ${TYPE_BADGE[p.PermissionType] || "bg-secondary"}">
          ${escHtml(p.PermissionType || "")}
        </span>
      </td>
      <td class="text-muted small">${escHtml(p.MailboxType || "")}</td>
    </tr>
  `).join("");
}

// ---------------------------------------------------------------------------
// CSV Export
// ---------------------------------------------------------------------------
function initExport() {
  document.getElementById("exportBtn").addEventListener("click", () => {
    if (!allPermissions.length) return;

    const headers = ["MailboxDisplayName", "MailboxUPN", "MailboxType", "PermissionType", "AccessRights"];
    const csvRows = [
      headers.join(","),
      ...allPermissions.map((p) =>
        headers.map((h) => `"${(p[h] || "").replace(/"/g, '""')}"`).join(",")
      ),
    ];
    const blob = new Blob([csvRows.join("\r\n")], { type: "text/csv;charset=utf-8;" });
    const url  = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href     = url;
    link.download = `mailbox-permissions-${selectedUser?.upn || "export"}-${new Date().toISOString().slice(0, 10)}.csv`;
    link.click();
    URL.revokeObjectURL(url);
  });
}

// ---------------------------------------------------------------------------
// Utility
// ---------------------------------------------------------------------------
function escHtml(str) {
  return String(str ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}
