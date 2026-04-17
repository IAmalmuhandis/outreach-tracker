const orgList = document.getElementById("orgList");
const orgForm = document.getElementById("orgForm");
const orgDetails = document.getElementById("orgDetails");
const searchInput = document.getElementById("searchInput");
const statusFilter = document.getElementById("statusFilter");
const priorityFilter = document.getElementById("priorityFilter");
const categoryFilter = document.getElementById("categoryFilter");
const subcategoryFilter = document.getElementById("subcategoryFilter");
const countryFilter = document.getElementById("countryFilter");
const tierFilter = document.getElementById("tierFilter");
const documentsFilter = document.getElementById("documentsFilter");
const formCategory = document.getElementById("formCategory");
const formSubcategory = document.getElementById("formSubcategory");
const clearFiltersBtn = document.getElementById("clearFiltersBtn");
const logoutBtn = document.getElementById("logoutBtn");
const importDocxBtn = document.getElementById("importDocxBtn");
const importDocxInput = document.getElementById("importDocxInput");
const deleteAllBtn = document.getElementById("deleteAllBtn");
const orgItemTemplate = document.getElementById("orgItemTemplate");
const statsRow = document.getElementById("statsRow");
const statusBanner = document.getElementById("statusBanner");
const guideChecklist = document.getElementById("guideChecklist");
const orgModal = document.getElementById("orgModal");
const orgModalBackdrop = document.getElementById("orgModalBackdrop");
const closeOrgModalBtn = document.getElementById("closeOrgModalBtn");
const orgModalTitle = document.getElementById("orgModalTitle");
const openGuideModalBtn = document.getElementById("openGuideModalBtn");
const guideModal = document.getElementById("guideModal");
const guideModalBackdrop = document.getElementById("guideModalBackdrop");
const closeGuideModalBtn = document.getElementById("closeGuideModalBtn");
const guideSearchInput = document.getElementById("guideSearchInput");
const guideModalContent = document.getElementById("guideModalContent");
const confirmModal = document.getElementById("confirmModal");
const confirmMessage = document.getElementById("confirmMessage");
const confirmCancelBtn = document.getElementById("confirmCancelBtn");
const confirmApproveBtn = document.getElementById("confirmApproveBtn");
const confirmBackdrop = confirmModal.querySelector(".modal-backdrop");

const GUIDE_SECTIONS = [
  {
    id: "prep",
    title: "Preparation",
    items: [
      { id: "onePager", label: "Carry one-page YARE brief" },
      { id: "demoReady", label: "Demo video/platform ready offline" },
      { id: "businessCards", label: "Business cards ready" },
      { id: "orgResearch", label: "Researched this organization's latest campaign" },
    ],
  },
  {
    id: "meeting",
    title: "Meeting",
    items: [
      { id: "introducedMission", label: "Introduced YARE with mission first" },
      { id: "askedCurrentProcess", label: "Asked current translation workflow" },
      { id: "pitchedTrial", label: "Requested trial project (not full contract)" },
      { id: "gotFollowupChannel", label: "Got best follow-up channel (WhatsApp/email)" },
    ],
  },
  {
    id: "followup",
    title: "Follow-up",
    items: [
      { id: "sentThankYou", label: "Sent thank-you message within 24 hours" },
      { id: "loggedMeeting", label: "Logged meeting summary and next step" },
      { id: "setNextDate", label: "Set next follow-up date" },
      { id: "sharedValue", label: "Shared value before asking again" },
    ],
  },
];

const FULL_GUIDE = [
  {
    title: "Before You Reach Out",
    summary: "Get clarity and context before first contact.",
    points: [
      "Confirm organization profile, mandate, and current campaigns.",
      "Match one clear YARE use-case to their real workflow.",
      "Prepare one simple outcome for the first meeting.",
      "Have quick proof ready: sample, short deck, or demo clip.",
    ],
    script: "Hi [Name], we support teams handling multilingual programs. We noticed your work on [initiative] and can help with [specific challenge]. Can we schedule a short intro this week?",
  },
  {
    title: "First Meeting Flow",
    summary: "Keep meetings structured and short.",
    points: [
      "Start with their current process, not your product pitch.",
      "Ask where translation delays, quality gaps, or cost issues appear.",
      "Share one focused value proposition tied to their pain point.",
      "Close with a concrete next step and owner on both sides.",
    ],
    script: "If we could reduce turnaround and keep language quality consistent for your next program, would a pilot on one document set be useful?",
  },
  {
    title: "Follow-up Execution",
    summary: "Fast follow-up wins more than long proposals.",
    points: [
      "Send summary and agreed actions within 24 hours.",
      "Attach only relevant material for that stakeholder.",
      "Set a date for next touchpoint while they are still warm.",
      "Log every response and update organization status immediately.",
    ],
    script: "Thanks again for today. As agreed, we will send a pilot scope by [date]. Please confirm your preferred review window so we can align delivery.",
  },
  {
    title: "Decision Stage",
    summary: "De-risk adoption for procurement and leadership.",
    points: [
      "Define pilot scope, timeline, and acceptance criteria.",
      "Map decision makers: technical, program, procurement, finance.",
      "Address compliance and quality assurance questions early.",
      "Offer phased rollout instead of all-at-once implementation.",
    ],
    script: "To keep risk low, let us run a controlled pilot for one unit first. If outcomes meet your benchmark, we can scale in phases.",
  },
  {
    title: "Relationship Management",
    summary: "Pipeline health depends on consistency.",
    points: [
      "Review priorities weekly and push stalled deals forward.",
      "Share useful insights even when not asking for a meeting.",
      "Escalate high-potential accounts with clear internal notes.",
      "Track wins/losses and refine talking points continuously.",
    ],
    script: "We noticed an update relevant to your language program and thought this may help your team’s planning this quarter.",
  },
];

let organizations = [];
let selectedOrgId = null;
let confirmResolver = null;

function formatDate(dateString) {
  if (!dateString) return "N/A";
  return new Date(dateString).toLocaleDateString();
}

function escapeHtml(value) {
  return String(value || "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll("\"", "&quot;")
    .replaceAll("'", "&#039;");
}

function formatOrgLocation(org) {
  const parts = [org.address, org.city, org.state, org.country]
    .map((item) => String(item || "").trim())
    .filter(Boolean);
  if (parts.length > 0) return parts.join(", ");
  return org.location || "No location";
}

function getDocumentType(doc) {
  const fileName = String(doc?.originalName || "").toLowerCase();
  const mime = String(doc?.mimeType || "").toLowerCase();
  if (fileName.endsWith(".pdf") || mime.includes("pdf")) return { label: "PDF", icon: "DOC", className: "doc-badge-pdf" };
  if (fileName.endsWith(".doc") || fileName.endsWith(".docx") || mime.includes("wordprocessingml") || mime.includes("msword")) {
    return { label: "Word", icon: "DOC", className: "doc-badge-word" };
  }
  if (fileName.endsWith(".xls") || fileName.endsWith(".xlsx") || mime.includes("spreadsheetml") || mime.includes("ms-excel")) {
    return { label: "Excel", icon: "XLS", className: "doc-badge-excel" };
  }
  if (fileName.endsWith(".png") || fileName.endsWith(".jpg") || fileName.endsWith(".jpeg") || mime.startsWith("image/")) {
    return { label: "Image", icon: "IMG", className: "doc-badge-image" };
  }
  if (fileName.endsWith(".txt") || mime.includes("text/plain")) return { label: "Text", icon: "TXT", className: "doc-badge-text" };
  return { label: "File", icon: "FILE", className: "doc-badge-generic" };
}

async function api(path, options = {}) {
  const headers = { ...(options.headers || {}) };
  if (!(options.body instanceof FormData) && !headers["Content-Type"]) {
    headers["Content-Type"] = "application/json";
  }

  const response = await fetch(path, {
    headers,
    ...options,
  });

  if (!response.ok) {
    const error = await response.json().catch(() => ({}));
    if (response.status === 401 && !path.startsWith("/api/auth/")) {
      window.location.replace("/auth");
    }
    throw new Error(error.message || "Request failed");
  }

  if (response.status === 204) return null;
  return response.json();
}

function showBanner(message) {
  statusBanner.hidden = false;
  statusBanner.textContent = message;
}

function clearBanner() {
  statusBanner.hidden = true;
  statusBanner.textContent = "";
}

async function ensureAuthenticated() {
  try {
    await api("/api/auth/session");
  } catch (_error) {
    window.location.replace("/auth");
    throw _error;
  }
}

function openOrgModal() {
  orgModal.hidden = false;
  syncModalBodyLock();
}

function closeOrgModal() {
  orgModal.hidden = true;
  syncModalBodyLock();
}

function openGuideModal() {
  renderGuideModalContent();
  guideModal.hidden = false;
  syncModalBodyLock();
}

function closeGuideModal() {
  guideModal.hidden = true;
  syncModalBodyLock();
}

function syncModalBodyLock() {
  const isAnyModalOpen = !orgModal.hidden || !guideModal.hidden || !confirmModal.hidden;
  if (isAnyModalOpen) {
    document.body.classList.add("modal-open");
  } else {
    document.body.classList.remove("modal-open");
  }
}

function showConfirm(message) {
  confirmMessage.textContent = message;
  confirmModal.hidden = false;
  syncModalBodyLock();
  return new Promise((resolve) => {
    confirmResolver = resolve;
  });
}

function closeConfirm(result) {
  confirmModal.hidden = true;
  syncModalBodyLock();
  if (confirmResolver) {
    confirmResolver(result);
    confirmResolver = null;
  }
}

function setSelectOptions(selectEl, values, allLabel) {
  const current = selectEl.value || "all";
  const options = [`<option value="all">${allLabel}</option>`];
  values.forEach((value) => {
    options.push(`<option value="${escapeHtml(value)}">${escapeHtml(value)}</option>`);
  });
  selectEl.innerHTML = options.join("");
  selectEl.value = values.includes(current) ? current : "all";
}

function setFormSelectOptions(selectEl, values, placeholder) {
  const current = selectEl.value || "";
  const options = [`<option value="">${placeholder}</option>`];
  values.forEach((value) => {
    options.push(`<option value="${escapeHtml(value)}">${escapeHtml(value)}</option>`);
  });
  selectEl.innerHTML = options.join("");
  selectEl.value = values.includes(current) ? current : "";
}

function renderGuideModalContent() {
  const query = guideSearchInput.value.trim().toLowerCase();
  const visibleSections = FULL_GUIDE.filter((section) => {
    if (!query) return true;
    const haystack = `${section.title} ${section.summary} ${section.script} ${section.points.join(" ")}`.toLowerCase();
    return haystack.includes(query);
  });

  if (visibleSections.length === 0) {
    guideModalContent.innerHTML = '<p class="muted">No guide section matched your search.</p>';
    return;
  }

  guideModalContent.innerHTML = visibleSections
    .map(
      (section, index) => `
      <details class="guide-modal-section" ${index === 0 ? "open" : ""}>
        <summary>
          <span>${escapeHtml(section.title)}</span>
          <small>${escapeHtml(section.summary)}</small>
        </summary>
        <ul>
          ${section.points.map((point) => `<li>${escapeHtml(point)}</li>`).join("")}
        </ul>
        <div class="guide-script">
          <strong>Suggested line</strong>
          <p>${escapeHtml(section.script)}</p>
        </div>
      </details>
    `
    )
    .join("");
}

async function fetchFilterOptions() {
  const data = await api("/api/organizations/filters");
  setSelectOptions(categoryFilter, data.categories || [], "All categories");
  setSelectOptions(subcategoryFilter, data.subcategories || [], "All subcategories");
  setSelectOptions(countryFilter, data.countries || [], "All countries");
  setSelectOptions(tierFilter, data.priorityTiers || [], "All priority tiers");
  setFormSelectOptions(formCategory, data.categories || [], "Select category");
  setFormSelectOptions(formSubcategory, data.subcategories || [], "Select subcategory");
}

async function fetchOrganizations() {
  const params = new URLSearchParams({
    search: searchInput.value.trim(),
    status: statusFilter.value,
    priority: priorityFilter.value,
    category: categoryFilter.value,
    subcategory: subcategoryFilter.value,
    country: countryFilter.value,
    priorityTier: tierFilter.value,
    hasDocuments: documentsFilter.value,
  });
  organizations = await api(`/api/organizations?${params.toString()}`);
  renderStats();
  renderOrganizations();
  renderSelectedOrganization();
}

async function fetchHealth() {
  try {
    const health = await api("/api/health");
    if (health.mode === "memory-fallback") {
      showBanner(`MongoDB is unreachable, app is running in local memory mode. Reason: ${health.reason || "Unknown"}`);
    } else {
      clearBanner();
    }
  } catch {
    showBanner("Could not reach API health endpoint.");
  }
}

async function refreshData() {
  await Promise.all([fetchOrganizations(), fetchFilterOptions()]);
}

function renderStats() {
  const total = organizations.length;
  const contacted = organizations.filter((item) => ["Contacted", "Meeting Scheduled", "In Discussion", "Won"].includes(item.status)).length;
  const won = organizations.filter((item) => item.status === "Won").length;
  const high = organizations.filter((item) => item.priority === "High").length;
  const withDocuments = organizations.filter((item) => Array.isArray(item.documents) && item.documents.length > 0).length;
  statsRow.innerHTML = `
    <div class="stat-card"><div class="stat-label">Total Organizations</div><div class="stat-value">${total}</div></div>
    <div class="stat-card"><div class="stat-label">Active Outreach</div><div class="stat-value">${contacted}</div></div>
    <div class="stat-card"><div class="stat-label">Won</div><div class="stat-value">${won}</div></div>
    <div class="stat-card"><div class="stat-label">High Priority</div><div class="stat-value">${high}</div></div>
    <div class="stat-card"><div class="stat-label">With Documents</div><div class="stat-value">${withDocuments}</div></div>
  `;
}

function renderOrganizations() {
  orgList.innerHTML = "";

  if (organizations.length === 0) {
    orgList.innerHTML = '<p class="muted">No organizations found for current filters.</p>';
    return;
  }

  organizations.forEach((org) => {
    const node = orgItemTemplate.content.firstElementChild.cloneNode(true);
    node.dataset.orgId = org._id;
    if (org._id === selectedOrgId) {
      node.classList.add("active");
    }

    node.querySelector("h3").textContent = org.name;
    const idText = org.organizationId ? `${org.organizationId} | ` : "";
    const categoryText = org.category || org.sector || "No category";
    node.querySelector(".meta").textContent = `${idText}${categoryText} | ${org.subcategory || "No subcategory"}`;
    node.querySelector(".notes").textContent = org.relevanceToYare || org.description || org.notes || "No description yet.";
    node.querySelector(".chips").innerHTML = `
      <span class="chip">${org.priority}</span>
      <span class="chip">${org.status}</span>
      <span class="chip">${org.priorityTier || "No tier"}</span>
      <span class="chip">${org.country || "No country"}</span>
      <span class="chip">${org.contacts.length} contact(s)</span>
      <span class="chip">${Array.isArray(org.documents) ? org.documents.length : 0} document(s)</span>
    `;

    node.addEventListener("click", () => {
      selectedOrgId = org._id;
      renderOrganizations();
      renderSelectedOrganization();
      openOrgModal();
    });

    const deleteBtn = node.querySelector(".delete-btn");
    deleteBtn.addEventListener("click", async (event) => {
      event.stopPropagation();
      const approved = await showConfirm(`Delete ${org.name}? This action cannot be undone.`);
      if (!approved) return;
      await api(`/api/organizations/${org._id}`, { method: "DELETE" });
      if (selectedOrgId === org._id) {
        selectedOrgId = null;
        closeOrgModal();
      }
      await refreshData();
    });

    orgList.appendChild(node);
  });
}

function renderSelectedOrganization() {
  const org = organizations.find((item) => item._id === selectedOrgId);
  if (!org) {
    orgModalTitle.textContent = "Organization";
    orgDetails.className = "stack muted";
    orgDetails.textContent = "Pick an organization to manage contacts and outreach logs.";
    renderGuideChecklist(null);
    return;
  }

  orgModalTitle.textContent = org.name;
  orgDetails.className = "stack";
  const orgDocuments = Array.isArray(org.documents) ? org.documents : [];
  const safeDescription = escapeHtml(org.description || org.notes || "");
  const safeRelevance = escapeHtml(org.relevanceToYare || "");
  const safeContactsToMeet = escapeHtml(org.contactsToMeet || "");
  const safeTalkingPoint = escapeHtml(org.talkingPoint || "");
  orgDetails.innerHTML = `
    <div class="stack soft-card">
      <h4>Structured Profile</h4>
      <div><strong>Organization ID:</strong> ${org.organizationId || "N/A"}</div>
      <div><strong>Category:</strong> ${org.category || "N/A"}</div>
      <div><strong>Subcategory:</strong> ${org.subcategory || "N/A"}</div>
      <div><strong>Address:</strong> ${formatOrgLocation(org)}</div>
      <div><strong>Priority Tier:</strong> ${org.priorityTier || "N/A"}</div>
      <div><strong>Description:</strong><br />${safeDescription || "N/A"}</div>
      <div><strong>Relevance to YARE:</strong><br />${safeRelevance || "N/A"}</div>
      <div><strong>Contacts to Meet:</strong><br />${safeContactsToMeet || "N/A"}</div>
      <div><strong>Talking Point:</strong><br />${safeTalkingPoint || "N/A"}</div>
    </div>

    <div class="stack soft-card">
      <h4>Edit Organization Details</h4>
      <input id="orgNameEdit" value="${org.name}" />
      <input id="orgCategoryEdit" value="${org.category || ""}" placeholder="Category" />
      <input id="orgSubcategoryEdit" value="${org.subcategory || ""}" placeholder="Subcategory" />
      <input id="orgAddressEdit" value="${org.address || ""}" placeholder="Address" />
      <input id="orgCityEdit" value="${org.city || ""}" placeholder="City" />
      <input id="orgStateEdit" value="${org.state || ""}" placeholder="State" />
      <input id="orgCountryEdit" value="${org.country || ""}" placeholder="Country" />
      <input id="orgPriorityTierEdit" value="${org.priorityTier || ""}" placeholder="Priority tier (e.g. Tier 1)" />
      <input id="orgSectorEdit" value="${org.sector || ""}" placeholder="Sector" />
      <input id="orgLocationEdit" value="${org.location || ""}" placeholder="Location" />
      <input id="orgWebsiteEdit" value="${org.website || ""}" placeholder="Website" />
      <input id="orgFollowUpEdit" type="date" value="${org.nextFollowUpDate ? new Date(org.nextFollowUpDate).toISOString().slice(0, 10) : ""}" />
      <textarea id="orgDescriptionEdit" rows="3" placeholder="Description">${org.description || ""}</textarea>
      <textarea id="orgRelevanceEdit" rows="3" placeholder="Relevance to YARE">${org.relevanceToYare || ""}</textarea>
      <textarea id="orgContactsToMeetEdit" rows="3" placeholder="Contacts to meet">${org.contactsToMeet || ""}</textarea>
      <textarea id="orgTalkingPointEdit" rows="3" placeholder="Talking point">${org.talkingPoint || ""}</textarea>
      <textarea id="orgNotesEdit" rows="3" placeholder="Notes">${org.notes || ""}</textarea>
      <button class="btn" id="saveOrgDetailsBtn" type="button">Save Organization Details</button>
    </div>

    <div class="stack soft-card">
      <label>Update status</label>
      <select id="statusUpdate">
        ${["Not Started", "Researching", "Contacted", "Meeting Scheduled", "In Discussion", "Won", "Paused"]
          .map((item) => `<option value="${item}" ${item === org.status ? "selected" : ""}>${item}</option>`)
          .join("")}
      </select>
      <button class="btn" id="saveStatusBtn" type="button">Save Status</button>
    </div>

    <div class="stack soft-card">
      <h4>Add Contact</h4>
      <input id="contactName" placeholder="Full name *" />
      <input id="contactRole" placeholder="Role/title" />
      <input id="contactPhone" placeholder="Phone number" />
      <input id="contactEmail" type="email" placeholder="Email" />
      <input id="contactPreferred" placeholder="Preferred channel" />
      <textarea id="contactNotes" rows="3" placeholder="Contact notes"></textarea>
      <button class="btn" id="addContactBtn" type="button">Add Contact</button>
      <p class="muted helper">Add multiple contacts at once. Format: <code>Name | Role | Phone | Email | Preferred Channel | Notes</code></p>
      <textarea id="bulkContacts" rows="4" placeholder="Jane Doe | Communications Lead | +234... | jane@org.com | WhatsApp | Abuja office&#10;John Aliyu | Procurement | +234... | john@org.com | Email | Responds fast"></textarea>
      <button class="btn" id="addBulkContactsBtn" type="button">Add Multiple Contacts</button>
    </div>

    <div class="stack soft-card">
      <h4>Add Outreach Log</h4>
      <input id="logType" placeholder="Activity type (Call, Meeting, Email) *" />
      <textarea id="logSummary" rows="3" placeholder="Summary *"></textarea>
      <input id="logNextAction" placeholder="Next action" />
      <button class="btn" id="addLogBtn" type="button">Add Log Entry</button>
    </div>

    <div class="stack soft-card">
      <h4>Contacts</h4>
      ${
        org.contacts.length
          ? org.contacts
              .map(
                (contact) => `
        <div class="contact-item">
          <strong>${contact.fullName}</strong> ${contact.role ? `(${contact.role})` : ""}
          <div>${contact.phone || "No phone"} | ${contact.email || "No email"}</div>
          <div class="muted">${contact.preferredChannel || "No preferred channel"}</div>
          <div class="muted">${contact.notes || ""}</div>
        </div>`
              )
              .join("")
          : '<p class="muted">No contacts yet.</p>'
      }
    </div>

    <div class="stack soft-card">
      <h4>Outreach History</h4>
      ${
        org.outreachLogs.length
          ? org.outreachLogs
              .map(
                (log) => `
        <div class="log-item">
          <strong>${log.activityType}</strong> on ${formatDate(log.date)}
          <div>${log.summary}</div>
          <div class="muted">${log.nextAction ? `Next: ${log.nextAction}` : "No next action set"}</div>
        </div>`
              )
              .join("")
          : '<p class="muted">No outreach logs yet.</p>'
      }
    </div>

    <div class="stack soft-card">
      <h4>Organization Documents</h4>
      <input id="orgDocumentFile" type="file" accept=".pdf,.doc,.docx,.xls,.xlsx,.txt,.png,.jpg,.jpeg" hidden />
      <input id="orgDocumentVersion" placeholder="Document version (e.g. v1.0, Draft 2)" />
      <textarea id="orgDocumentNotes" rows="2" placeholder="Document notes / label"></textarea>
      <button class="btn" id="uploadDocumentBtn" type="button">Choose File and Upload</button>
      ${
        orgDocuments.length
          ? orgDocuments
              .map((doc) => {
                const typeInfo = getDocumentType(doc);
                return `
        <div class="document-item">
          <div class="document-main">
            <div class="document-header">
              <span class="doc-icon">${typeInfo.icon}</span>
              <span class="doc-badge ${typeInfo.className}">${typeInfo.label}</span>
            </div>
            <a href="${doc.url}" target="_blank" rel="noreferrer">${doc.originalName}</a>
            <div class="muted">${(doc.size / 1024).toFixed(1)} KB | Uploaded ${formatDate(doc.uploadedAt)}</div>
            <div class="muted">${doc.version ? `Version: ${doc.version}` : "Version: Not set"}</div>
            <div class="muted">${doc.notes ? `Notes: ${doc.notes}` : "Notes: None"}</div>
          </div>
          <button class="delete-btn-inline" type="button" data-doc-id="${doc._id}">Delete</button>
        </div>`;
              })
              .join("")
          : '<p class="muted">No documents uploaded yet.</p>'
      }
    </div>

    <div class="muted">Last contacted: ${formatDate(org.lastContactedDate)} | Next follow-up: ${formatDate(org.nextFollowUpDate)}</div>
  `;

  document.getElementById("saveStatusBtn").onclick = async () => {
    try {
      const status = document.getElementById("statusUpdate").value;
      await api(`/api/organizations/${org._id}`, {
        method: "PUT",
        body: JSON.stringify({ status }),
      });
      await refreshData();
    } catch (error) {
      alert(error.message);
    }
  };

  document.getElementById("saveOrgDetailsBtn").onclick = async () => {
    try {
      const name = document.getElementById("orgNameEdit").value.trim();
      if (!name) {
        alert("Organization name is required.");
        return;
      }
      const nextFollowUpDate = document.getElementById("orgFollowUpEdit").value;
      await api(`/api/organizations/${org._id}`, {
        method: "PUT",
        body: JSON.stringify({
          name,
          category: document.getElementById("orgCategoryEdit").value.trim(),
          subcategory: document.getElementById("orgSubcategoryEdit").value.trim(),
          address: document.getElementById("orgAddressEdit").value.trim(),
          city: document.getElementById("orgCityEdit").value.trim(),
          state: document.getElementById("orgStateEdit").value.trim(),
          country: document.getElementById("orgCountryEdit").value.trim(),
          priorityTier: document.getElementById("orgPriorityTierEdit").value.trim(),
          sector: document.getElementById("orgSectorEdit").value.trim(),
          location: document.getElementById("orgLocationEdit").value.trim(),
          website: document.getElementById("orgWebsiteEdit").value.trim(),
          description: document.getElementById("orgDescriptionEdit").value.trim(),
          relevanceToYare: document.getElementById("orgRelevanceEdit").value.trim(),
          contactsToMeet: document.getElementById("orgContactsToMeetEdit").value.trim(),
          talkingPoint: document.getElementById("orgTalkingPointEdit").value.trim(),
          notes: document.getElementById("orgNotesEdit").value.trim(),
          nextFollowUpDate: nextFollowUpDate ? new Date(nextFollowUpDate).toISOString() : null,
        }),
      });
      await refreshData();
    } catch (error) {
      alert(error.message);
    }
  };

  document.getElementById("addContactBtn").onclick = async () => {
    try {
      const fullName = document.getElementById("contactName").value.trim();
      if (!fullName) {
        alert("Contact name is required.");
        return;
      }
      const payload = {
        fullName,
        role: document.getElementById("contactRole").value.trim(),
        phone: document.getElementById("contactPhone").value.trim(),
        email: document.getElementById("contactEmail").value.trim(),
        preferredChannel: document.getElementById("contactPreferred").value.trim(),
        notes: document.getElementById("contactNotes").value.trim(),
      };
      await api(`/api/organizations/${org._id}/contacts`, {
        method: "POST",
        body: JSON.stringify(payload),
      });
      await refreshData();
    } catch (error) {
      alert(error.message);
    }
  };

  document.getElementById("addBulkContactsBtn").onclick = async () => {
    try {
      const raw = document.getElementById("bulkContacts").value.trim();
      if (!raw) {
        alert("Please enter at least one contact line.");
        return;
      }

      const contacts = raw
        .split("\n")
        .map((line) => line.trim())
        .filter(Boolean)
        .map((line) => {
          const parts = line.split("|").map((part) => part.trim());
          return {
            fullName: parts[0] || "",
            role: parts[1] || "",
            phone: parts[2] || "",
            email: parts[3] || "",
            preferredChannel: parts[4] || "",
            notes: parts[5] || "",
          };
        });

      const result = await api(`/api/organizations/${org._id}/contacts/bulk`, {
        method: "POST",
        body: JSON.stringify({ contacts }),
      });
      alert(`Added ${result.added} contact(s).`);
      await refreshData();
    } catch (error) {
      alert(error.message);
    }
  };

  document.getElementById("addLogBtn").onclick = async () => {
    try {
      const activityType = document.getElementById("logType").value.trim();
      const summary = document.getElementById("logSummary").value.trim();
      if (!activityType || !summary) {
        alert("Activity type and summary are required.");
        return;
      }
      const nextAction = document.getElementById("logNextAction").value.trim();
      await api(`/api/organizations/${org._id}/logs`, {
        method: "POST",
        body: JSON.stringify({ activityType, summary, nextAction }),
      });
      await api(`/api/organizations/${org._id}`, {
        method: "PUT",
        body: JSON.stringify({ lastContactedDate: new Date().toISOString() }),
      });
      await refreshData();
    } catch (error) {
      alert(error.message);
    }
  };

  const fileInput = document.getElementById("orgDocumentFile");
  const uploadButton = document.getElementById("uploadDocumentBtn");

  uploadButton.onclick = () => {
    fileInput.click();
  };

  fileInput.onchange = async () => {
    try {
      const file = fileInput.files[0];
      if (!file) {
        return;
      }
      const formData = new FormData();
      formData.append("document", file);
      formData.append("version", document.getElementById("orgDocumentVersion").value.trim());
      formData.append("notes", document.getElementById("orgDocumentNotes").value.trim());
      await api(`/api/organizations/${org._id}/documents`, {
        method: "POST",
        body: formData,
      });
      fileInput.value = "";
      document.getElementById("orgDocumentVersion").value = "";
      document.getElementById("orgDocumentNotes").value = "";
      await refreshData();
    } catch (error) {
      alert(error.message);
    }
  };

  orgDetails.querySelectorAll(".delete-btn-inline[data-doc-id]").forEach((button) => {
    button.onclick = async () => {
      try {
        const approved = await showConfirm("Delete this document?");
        if (!approved) return;
        await api(`/api/organizations/${org._id}/documents/${button.dataset.docId}`, {
          method: "DELETE",
        });
        await refreshData();
      } catch (error) {
        alert(error.message);
      }
    };
  });

  renderGuideChecklist(org);
}

function renderGuideChecklist(org) {
  if (!org) {
    guideChecklist.className = "stack muted";
    guideChecklist.textContent = "Select an organization to start ticking guide items.";
    return;
  }

  const saved = org.guideChecklist || {};
  guideChecklist.className = "stack";
  guideChecklist.innerHTML = `
    <h3>${org.name}</h3>
    ${GUIDE_SECTIONS.map(
      (section) => `
      <div class="guide-section">
        <h4>${section.title}</h4>
        ${section.items
          .map(
            (item) => `
          <label class="guide-item">
            <input type="checkbox" data-guide-key="${section.id}.${item.id}" ${saved[`${section.id}.${item.id}`] ? "checked" : ""} />
            <span>${item.label}</span>
          </label>
        `
          )
          .join("")}
      </div>
    `
    ).join("")}
    <button class="btn" id="saveGuideChecklistBtn" type="button">Save Guide Checklist</button>
  `;

  document.getElementById("saveGuideChecklistBtn").onclick = async () => {
    try {
      const nextChecklist = {};
      guideChecklist.querySelectorAll("input[data-guide-key]").forEach((el) => {
        nextChecklist[el.dataset.guideKey] = el.checked;
      });
      await api(`/api/organizations/${org._id}/guide-checklist`, {
        method: "PUT",
        body: JSON.stringify({ guideChecklist: nextChecklist }),
      });
      await refreshData();
    } catch (error) {
      alert(error.message);
    }
  };
}

orgForm.addEventListener("submit", async (event) => {
  try {
    event.preventDefault();
    const data = new FormData(orgForm);
    const payload = Object.fromEntries(data.entries());
    await api("/api/organizations", {
      method: "POST",
      body: JSON.stringify(payload),
    });
    orgForm.reset();
    await refreshData();
  } catch (error) {
    alert(error.message);
  }
});

logoutBtn.addEventListener("click", async () => {
  try {
    await api("/api/auth/logout", { method: "POST" });
  } catch (_error) {
    // Ignore logout errors and redirect anyway.
  } finally {
    window.location.replace("/auth");
  }
});

const debouncedFetch = (() => {
  let timer;
  return () => {
    clearTimeout(timer);
    timer = setTimeout(fetchOrganizations, 250);
  };
})();

searchInput.addEventListener("input", debouncedFetch);
[statusFilter, priorityFilter, categoryFilter, subcategoryFilter, countryFilter, tierFilter, documentsFilter].forEach((el) => {
  el.addEventListener("change", fetchOrganizations);
});

clearFiltersBtn.addEventListener("click", async () => {
  searchInput.value = "";
  statusFilter.value = "all";
  priorityFilter.value = "all";
  categoryFilter.value = "all";
  subcategoryFilter.value = "all";
  countryFilter.value = "all";
  tierFilter.value = "all";
  documentsFilter.value = "all";
  await fetchOrganizations();
});

importDocxBtn.addEventListener("click", () => {
  importDocxInput.click();
});

importDocxInput.addEventListener("change", async () => {
  const file = importDocxInput.files?.[0];
  if (!file) return;
  importDocxBtn.disabled = true;
  importDocxBtn.textContent = "Importing...";
  try {
    const formData = new FormData();
    formData.append("document", file);
    const result = await api("/api/import/docx", { method: "POST", body: formData });
    alert(
      `Import complete. Detected: ${result.detected}, Imported: ${result.imported}, Skipped: ${result.skipped}, Mode: ${result.mode}`
    );
    await refreshData();
  } catch (error) {
    alert(`Import failed: ${error.message}`);
  } finally {
    importDocxBtn.disabled = false;
    importDocxBtn.textContent = "Import Organizations from DOCX";
    importDocxInput.value = "";
  }
});

deleteAllBtn.addEventListener("click", async () => {
  const approved = await showConfirm("Delete ALL organizations? This cannot be undone.");
  if (!approved) return;

  try {
    await api("/api/organizations", { method: "DELETE" });
    selectedOrgId = null;
    closeOrgModal();
    await refreshData();
  } catch (error) {
    alert(`Delete all failed: ${error.message}`);
  }
});

closeOrgModalBtn.addEventListener("click", closeOrgModal);
orgModalBackdrop.addEventListener("click", closeOrgModal);
openGuideModalBtn.addEventListener("click", openGuideModal);
closeGuideModalBtn.addEventListener("click", closeGuideModal);
guideModalBackdrop.addEventListener("click", closeGuideModal);
guideSearchInput.addEventListener("input", renderGuideModalContent);

confirmCancelBtn.addEventListener("click", () => closeConfirm(false));
confirmApproveBtn.addEventListener("click", () => closeConfirm(true));
confirmBackdrop.addEventListener("click", () => closeConfirm(false));

document.addEventListener("keydown", (event) => {
  if (event.key === "Escape") {
    if (!confirmModal.hidden) {
      closeConfirm(false);
      return;
    }
    if (!orgModal.hidden) {
      closeOrgModal();
      return;
    }
    if (!guideModal.hidden) {
      closeGuideModal();
    }
  }
});

(async () => {
  try {
    await ensureAuthenticated();
    await fetchHealth();
    await refreshData();
  } catch (_error) {
    // Redirect is already handled in ensureAuthenticated/api.
  }
})();
