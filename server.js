const express = require("express");
const mongoose = require("mongoose");
const cors = require("cors");
const dotenv = require("dotenv");
const fs = require("fs/promises");
const fsSync = require("fs");
const path = require("path");
const mammoth = require("mammoth");
const { randomUUID } = require("crypto");
const multer = require("multer");
const bcrypt = require("bcryptjs");
const jwt = require("jsonwebtoken");
const cookieParser = require("cookie-parser");

dotenv.config();

const app = express();
const PORT = process.env.PORT || 4000;
const MONGODB_URI =
  process.env.MONGODB_URI ||
  "mongodb+srv://iamalmuhandis_db_user:6RoNDtXdDgOAZw0R@outreach.ttmbj3v.mongodb.net/?appName=Outreach";
const JWT_SECRET = process.env.JWT_SECRET || "change-this-secret-in-env";
const AUTH_COOKIE_NAME = "yare_auth";
const AUTH_COOKIE_MAX_AGE_MS = 1000 * 60 * 60 * 24 * 7;

app.use(cors({ origin: true, credentials: true }));
app.use(cookieParser());
app.use(express.json({ limit: "2mb" }));
app.use(express.static(path.join(__dirname, "public")));

const uploadDir = path.join(__dirname, "uploads", "organization-docs");
fsSync.mkdirSync(uploadDir, { recursive: true });
app.use("/uploads", express.static(path.join(__dirname, "uploads")));

const contactSchema = new mongoose.Schema(
  {
    _id: { type: String, default: () => randomUUID() },
    fullName: { type: String, required: true, trim: true },
    role: { type: String, trim: true, default: "" },
    phone: { type: String, trim: true, default: "" },
    email: { type: String, trim: true, default: "" },
    preferredChannel: { type: String, trim: true, default: "" },
    notes: { type: String, trim: true, default: "" },
  },
  { _id: false }
);

const outreachLogSchema = new mongoose.Schema(
  {
    _id: { type: String, default: () => randomUUID() },
    activityType: { type: String, required: true, trim: true },
    summary: { type: String, required: true, trim: true },
    date: { type: Date, default: Date.now },
    nextAction: { type: String, trim: true, default: "" },
  },
  { _id: false }
);

const organizationDocumentSchema = new mongoose.Schema(
  {
    _id: { type: String, default: () => randomUUID() },
    originalName: { type: String, required: true, trim: true },
    storedName: { type: String, required: true, trim: true },
    mimeType: { type: String, required: true, trim: true },
    version: { type: String, trim: true, default: "" },
    notes: { type: String, trim: true, default: "" },
    size: { type: Number, required: true },
    url: { type: String, required: true, trim: true },
    uploadedAt: { type: Date, default: Date.now },
  },
  { _id: false }
);

const organizationSchema = new mongoose.Schema(
  {
    ownerId: { type: String, required: true, index: true },
    organizationId: { type: String, trim: true, default: "" },
    name: { type: String, required: true, trim: true },
    category: { type: String, trim: true, default: "" },
    subcategory: { type: String, trim: true, default: "" },
    sector: { type: String, trim: true, default: "" },
    address: { type: String, trim: true, default: "" },
    city: { type: String, trim: true, default: "" },
    state: { type: String, trim: true, default: "" },
    country: { type: String, trim: true, default: "" },
    description: { type: String, trim: true, default: "" },
    relevanceToYare: { type: String, trim: true, default: "" },
    contactsToMeet: { type: String, trim: true, default: "" },
    talkingPoint: { type: String, trim: true, default: "" },
    priorityTier: { type: String, trim: true, default: "" },
    status: {
      type: String,
      enum: ["Not Started", "Researching", "Contacted", "Meeting Scheduled", "In Discussion", "Won", "Paused"],
      default: "Not Started",
    },
    priority: {
      type: String,
      enum: ["High", "Medium", "Low"],
      default: "Medium",
    },
    website: { type: String, trim: true, default: "" },
    notes: { type: String, trim: true, default: "" },
    location: { type: String, trim: true, default: "" },
    nextFollowUpDate: { type: Date, default: null },
    lastContactedDate: { type: Date, default: null },
    guideChecklist: { type: mongoose.Schema.Types.Mixed, default: {} },
    contacts: { type: [contactSchema], default: [] },
    outreachLogs: { type: [outreachLogSchema], default: [] },
    documents: { type: [organizationDocumentSchema], default: [] },
  },
  { timestamps: true }
);

const Organization = mongoose.model("Organization", organizationSchema);
const userSchema = new mongoose.Schema(
  {
    _id: { type: String, default: () => randomUUID() },
    fullName: { type: String, required: true, trim: true },
    email: { type: String, required: true, trim: true, lowercase: true, unique: true, index: true },
    passwordHash: { type: String, required: true },
  },
  { timestamps: true }
);
const User = mongoose.model("User", userSchema);
let useMemoryStore = false;
let mongoFallbackReason = "";
const memoryOrganizations = [];
const memoryUsers = [];
const allowedDocumentMimeTypes = new Set([
  "application/pdf",
  "application/msword",
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  "application/vnd.ms-excel",
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  "text/plain",
  "image/jpeg",
  "image/png",
]);
const allowedDocumentExtensions = new Set([".pdf", ".doc", ".docx", ".xls", ".xlsx", ".txt", ".png", ".jpg", ".jpeg"]);
const allowedImportDocMimeTypes = new Set([
  "application/msword",
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  "application/octet-stream",
]);
const allowedImportDocExtensions = new Set([".doc", ".docx"]);
const upload = multer({
  storage: multer.diskStorage({
    destination: (_req, _file, cb) => cb(null, uploadDir),
    filename: (_req, file, cb) => {
      const ext = path.extname(file.originalname || "");
      cb(null, `${Date.now()}-${randomUUID()}${ext}`);
    },
  }),
  limits: { fileSize: 10 * 1024 * 1024 },
  fileFilter: (_req, file, cb) => {
    const extension = path.extname(file.originalname || "").toLowerCase();
    if (allowedDocumentMimeTypes.has(file.mimetype) || allowedDocumentExtensions.has(extension)) {
      cb(null, true);
      return;
    }
    cb(new Error("Unsupported file type. Upload PDF, DOC/DOCX, XLS/XLSX, TXT, PNG, or JPG."));
  },
});
const importUpload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 20 * 1024 * 1024 },
  fileFilter: (_req, file, cb) => {
    const extension = path.extname(file.originalname || "").toLowerCase();
    if (allowedImportDocMimeTypes.has(file.mimetype) || allowedImportDocExtensions.has(extension)) {
      cb(null, true);
      return;
    }
    cb(new Error("Unsupported import file type. Upload a DOC or DOCX file."));
  },
});

function normalizeString(value) {
  return String(value || "")
    .trim()
    .toLowerCase();
}

function signAuthToken(user) {
  return jwt.sign({ sub: user._id, email: user.email }, JWT_SECRET, { expiresIn: "7d" });
}

function setAuthCookie(res, token) {
  res.cookie(AUTH_COOKIE_NAME, token, {
    httpOnly: true,
    sameSite: "lax",
    secure: process.env.NODE_ENV === "production",
    maxAge: AUTH_COOKIE_MAX_AGE_MS,
  });
}

function clearAuthCookie(res) {
  res.clearCookie(AUTH_COOKIE_NAME, {
    httpOnly: true,
    sameSite: "lax",
    secure: process.env.NODE_ENV === "production",
  });
}

function sanitizeUser(user) {
  return {
    id: user._id,
    fullName: user.fullName,
    email: user.email,
  };
}

function getMemoryUserByEmail(email) {
  const normalized = normalizeString(email);
  return memoryUsers.find((user) => normalizeString(user.email) === normalized);
}

function getMemoryUserById(id) {
  return memoryUsers.find((user) => user._id === id);
}

function getUserMemoryOrganizations(ownerId) {
  return memoryOrganizations.filter((item) => item.ownerId === ownerId);
}

function getMemoryOrgById(id, ownerId) {
  return memoryOrganizations.find((item) => item._id === id && item.ownerId === ownerId);
}

async function deleteUploadedFileByUrl(url) {
  if (!url || typeof url !== "string") return;
  const relativeUrl = url.startsWith("/") ? url.slice(1) : url;
  const filePath = path.join(__dirname, relativeUrl);
  try {
    await fs.unlink(filePath);
  } catch (error) {
    if (error.code !== "ENOENT") {
      console.warn(`Failed to delete uploaded file ${filePath}: ${error.message}`);
    }
  }
}

async function deleteOrganizationDocuments(org) {
  const docs = Array.isArray(org?.documents) ? org.documents : [];
  await Promise.all(docs.map((doc) => deleteUploadedFileByUrl(doc.url)));
}

function createMemoryOrganization(payload, ownerId) {
  return {
    _id: randomUUID(),
    ownerId,
    organizationId: payload.organizationId?.trim() || "",
    name: payload.name?.trim() || "",
    category: payload.category?.trim() || "",
    subcategory: payload.subcategory?.trim() || "",
    sector: payload.sector?.trim() || "",
    address: payload.address?.trim() || "",
    city: payload.city?.trim() || "",
    state: payload.state?.trim() || "",
    country: payload.country?.trim() || "",
    description: payload.description?.trim() || "",
    relevanceToYare: payload.relevanceToYare?.trim() || "",
    contactsToMeet: payload.contactsToMeet?.trim() || "",
    talkingPoint: payload.talkingPoint?.trim() || "",
    priorityTier: payload.priorityTier?.trim() || "",
    status: payload.status || "Not Started",
    priority: payload.priority || "Medium",
    website: payload.website?.trim() || "",
    notes: payload.notes?.trim() || "",
    location: payload.location?.trim() || "",
    nextFollowUpDate: payload.nextFollowUpDate || null,
    lastContactedDate: payload.lastContactedDate || null,
    guideChecklist: payload.guideChecklist || {},
    contacts: Array.isArray(payload.contacts) ? payload.contacts : [],
    outreachLogs: Array.isArray(payload.outreachLogs) ? payload.outreachLogs : [],
    documents: Array.isArray(payload.documents) ? payload.documents : [],
    createdAt: new Date().toISOString(),
    updatedAt: new Date().toISOString(),
  };
}

const STRUCTURED_DOC_LABELS = [
  "ORGANISATION_ID",
  "ORGANISATION_NAME",
  "CATEGORY",
  "SUBCATEGORY",
  "ADDRESS",
  "CITY",
  "STATE",
  "COUNTRY",
  "DESCRIPTION",
  "RELEVANCE_TO_YARE",
  "CONTACTS_TO_MEET",
  "TALKING_POINT",
  "PRIORITY_TIER",
];

const LABEL_TO_FIELD_MAP = {
  ORGANISATION_ID: "organizationId",
  ORGANISATION_NAME: "name",
  CATEGORY: "category",
  SUBCATEGORY: "subcategory",
  ADDRESS: "address",
  CITY: "city",
  STATE: "state",
  COUNTRY: "country",
  DESCRIPTION: "description",
  RELEVANCE_TO_YARE: "relevanceToYare",
  CONTACTS_TO_MEET: "contactsToMeet",
  TALKING_POINT: "talkingPoint",
  PRIORITY_TIER: "priorityTier",
};

function normalizeTierToPriority(priorityTier) {
  const normalized = normalizeString(priorityTier);
  if (!normalized) return "Medium";
  if (normalized.includes("high") || normalized.includes("tier 1") || normalized.includes("tier i")) return "High";
  if (normalized.includes("low") || normalized.includes("tier 3") || normalized.includes("tier iii")) return "Low";
  return "Medium";
}

function buildOrganizationPayload(entry) {
  const addressParts = [entry.address, entry.city, entry.state, entry.country]
    .map((item) => String(item || "").trim())
    .filter(Boolean);
  return {
    organizationId: String(entry.organizationId || "").trim(),
    name: String(entry.name || "").trim(),
    category: String(entry.category || "").trim(),
    subcategory: String(entry.subcategory || "").trim(),
    sector: String(entry.subcategory || entry.category || "").trim(),
    address: String(entry.address || "").trim(),
    city: String(entry.city || "").trim(),
    state: String(entry.state || "").trim(),
    country: String(entry.country || "").trim(),
    description: String(entry.description || "").trim(),
    relevanceToYare: String(entry.relevanceToYare || "").trim(),
    contactsToMeet: String(entry.contactsToMeet || "").trim(),
    talkingPoint: String(entry.talkingPoint || "").trim(),
    priorityTier: String(entry.priorityTier || "").trim(),
    priority: normalizeTierToPriority(entry.priorityTier),
    status: "Not Started",
    notes: [entry.description, entry.relevanceToYare, entry.contactsToMeet, entry.talkingPoint]
      .map((item) => String(item || "").trim())
      .filter(Boolean)
      .join("\n\n"),
    location: addressParts.join(", "),
  };
}

function parseStructuredDocxEntries(rawText) {
  const lines = String(rawText || "")
    .split(/\r?\n/)
    .map((line) => line.replace(/\u00A0/g, " ").trim());

  const entries = [];
  let current = null;
  let pendingLabel = null;
  let currentCategoryHeader = "";

  const pushCurrent = () => {
    if (!current) return;
    if (!current.organizationId || !current.name) return;
    const startsWithOrg = /^ORG-\d+/i.test(current.organizationId);
    if (!startsWithOrg) return;
    if (!current.category && currentCategoryHeader) {
      current.category = currentCategoryHeader;
    }
    entries.push(buildOrganizationPayload(current));
  };

  for (const line of lines) {
    if (!line) continue;

    const categoryHeader = line.match(/^CATEGORY\s*\d+\s*[—-]\s*(.+)$/i);
    if (categoryHeader) {
      currentCategoryHeader = categoryHeader[1].trim();
      continue;
    }

    if (pendingLabel) {
      current = current || {};
      current[LABEL_TO_FIELD_MAP[pendingLabel]] = line;
      pendingLabel = null;
      continue;
    }

    const labeledLine = line.match(/^([A-Z_]+)\s*[:\-]\s*(.*)$/);
    if (labeledLine && LABEL_TO_FIELD_MAP[labeledLine[1]]) {
      const label = labeledLine[1];
      const field = LABEL_TO_FIELD_MAP[label];
      const value = labeledLine[2].trim();
      if (label === "ORGANISATION_ID") {
        pushCurrent();
        current = {};
      } else {
        current = current || {};
      }
      if (value) {
        current[field] = value;
      } else {
        pendingLabel = label;
      }
      continue;
    }

    if (STRUCTURED_DOC_LABELS.includes(line)) {
      if (line === "ORGANISATION_ID") {
        pushCurrent();
        current = {};
      } else {
        current = current || {};
      }
      pendingLabel = line;
    }
  }

  pushCurrent();
  const deduped = new Map();
  for (const entry of entries) {
    const key = normalizeString(entry.organizationId) || normalizeString(entry.name);
    if (key && !deduped.has(key)) {
      deduped.set(key, entry);
    }
  }
  return [...deduped.values()];
}

function parseLegacyDocxOrgs(rawText) {
  const lines = rawText
    .split("\n")
    .map((line) => line.trim())
    .filter(Boolean);
  const organizations = [];

  for (const line of lines) {
    if (line.length < 3) {
      continue;
    }

    const cleanLine = line.replace(/^\d+[\.\)]\s*/, "").trim();
    if (!cleanLine) {
      continue;
    }

    const parts = cleanLine.split(/\s[-–|]\s|,\s(?=[A-Z])/).map((part) => part.trim());
    const name = parts[0];
    if (!name || name.length < 2) {
      continue;
    }

    organizations.push({
      name,
      sector: parts[1] || "",
      notes: parts.slice(2).join(" | "),
      priority: "Medium",
      status: "Not Started",
    });
  }

  const uniqueByName = new Map();
  for (const org of organizations) {
    if (!uniqueByName.has(org.name.toLowerCase())) {
      uniqueByName.set(org.name.toLowerCase(), org);
    }
  }

  return [...uniqueByName.values()];
}

async function requireAuth(req, res, next) {
  try {
    const token = req.cookies?.[AUTH_COOKIE_NAME];
    if (!token) {
      return res.status(401).json({ message: "Authentication required" });
    }
    const payload = jwt.verify(token, JWT_SECRET);
    const userId = String(payload.sub || "");
    if (!userId) {
      clearAuthCookie(res);
      return res.status(401).json({ message: "Invalid session" });
    }

    if (useMemoryStore) {
      const user = getMemoryUserById(userId);
      if (!user) {
        clearAuthCookie(res);
        return res.status(401).json({ message: "Invalid session" });
      }
      req.authUserId = user._id;
      req.authUser = sanitizeUser(user);
      return next();
    }

    const user = await User.findById(userId).lean();
    if (!user) {
      clearAuthCookie(res);
      return res.status(401).json({ message: "Invalid session" });
    }
    req.authUserId = String(user._id);
    req.authUser = sanitizeUser(user);
    return next();
  } catch (_error) {
    clearAuthCookie(res);
    return res.status(401).json({ message: "Session expired. Please sign in again." });
  }
}

app.post("/api/auth/signup", async (req, res) => {
  try {
    const fullName = String(req.body?.fullName || "").trim();
    const email = String(req.body?.email || "").trim().toLowerCase();
    const password = String(req.body?.password || "");

    if (!fullName) return res.status(400).json({ message: "Full name is required." });
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return res.status(400).json({ message: "Valid email is required." });
    if (password.length < 8) return res.status(400).json({ message: "Password must be at least 8 characters." });

    if (useMemoryStore) {
      if (getMemoryUserByEmail(email)) {
        return res.status(409).json({ message: "Account already exists for this email." });
      }
      const user = {
        _id: randomUUID(),
        fullName,
        email,
        passwordHash: await bcrypt.hash(password, 10),
      };
      memoryUsers.push(user);
      const token = signAuthToken(user);
      setAuthCookie(res, token);
      return res.status(201).json({ user: sanitizeUser(user) });
    }

    const existingUser = await User.findOne({ email });
    if (existingUser) {
      return res.status(409).json({ message: "Account already exists for this email." });
    }
    const user = await User.create({
      _id: randomUUID(),
      fullName,
      email,
      passwordHash: await bcrypt.hash(password, 10),
    });
    const token = signAuthToken(user);
    setAuthCookie(res, token);
    return res.status(201).json({ user: sanitizeUser(user) });
  } catch (error) {
    return res.status(400).json({ message: "Failed to sign up", error: error.message });
  }
});

app.post("/api/auth/signin", async (req, res) => {
  try {
    const email = String(req.body?.email || "").trim().toLowerCase();
    const password = String(req.body?.password || "");
    if (!email || !password) {
      return res.status(400).json({ message: "Email and password are required." });
    }

    const user = useMemoryStore ? getMemoryUserByEmail(email) : await User.findOne({ email });
    if (!user) {
      return res.status(401).json({ message: "Invalid email or password." });
    }
    const matches = await bcrypt.compare(password, user.passwordHash);
    if (!matches) {
      return res.status(401).json({ message: "Invalid email or password." });
    }
    const token = signAuthToken(user);
    setAuthCookie(res, token);
    return res.json({ user: sanitizeUser(user) });
  } catch (error) {
    return res.status(400).json({ message: "Failed to sign in", error: error.message });
  }
});

app.post("/api/auth/logout", (req, res) => {
  clearAuthCookie(res);
  res.status(204).send();
});

app.get("/api/auth/session", requireAuth, (req, res) => {
  res.json({ user: req.authUser });
});

app.get("/api/health", (_req, res) => {
  res.json({
    ok: true,
    message: "Outreach assistant API is running.",
    mode: useMemoryStore ? "memory-fallback" : "mongodb",
    reason: useMemoryStore ? mongoFallbackReason : "",
  });
});

app.get("/api/organizations/filters", requireAuth, async (req, res) => {
  try {
    if (useMemoryStore) {
      const unique = (field) =>
        [...new Set(getUserMemoryOrganizations(req.authUserId).map((item) => String(item[field] || "").trim()).filter(Boolean))].sort((a, b) =>
          a.localeCompare(b)
        );
      return res.json({
        categories: unique("category"),
        subcategories: unique("subcategory"),
        countries: unique("country"),
        priorityTiers: unique("priorityTier"),
      });
    }

    const [categories, subcategories, countries, priorityTiers] = await Promise.all([
      Organization.distinct("category", { ownerId: req.authUserId, category: { $nin: [null, ""] } }),
      Organization.distinct("subcategory", { ownerId: req.authUserId, subcategory: { $nin: [null, ""] } }),
      Organization.distinct("country", { ownerId: req.authUserId, country: { $nin: [null, ""] } }),
      Organization.distinct("priorityTier", { ownerId: req.authUserId, priorityTier: { $nin: [null, ""] } }),
    ]);

    return res.json({
      categories: categories.sort((a, b) => a.localeCompare(b)),
      subcategories: subcategories.sort((a, b) => a.localeCompare(b)),
      countries: countries.sort((a, b) => a.localeCompare(b)),
      priorityTiers: priorityTiers.sort((a, b) => a.localeCompare(b)),
    });
  } catch (error) {
    return res.status(500).json({ message: "Failed to load filter values", error: error.message });
  }
});

app.get("/api/organizations", requireAuth, async (req, res) => {
  try {
    const {
      search = "",
      status = "all",
      priority = "all",
      category = "all",
      subcategory = "all",
      country = "all",
      priorityTier = "all",
      hasDocuments = "all",
    } = req.query;
    let organizations;

    if (useMemoryStore) {
      const needle = normalizeString(search);
      organizations = getUserMemoryOrganizations(req.authUserId)
        .filter((org) => {
          if (status !== "all" && org.status !== status) return false;
          if (priority !== "all" && org.priority !== priority) return false;
          if (category !== "all" && normalizeString(org.category) !== normalizeString(category)) return false;
          if (subcategory !== "all" && normalizeString(org.subcategory) !== normalizeString(subcategory)) return false;
          if (country !== "all" && normalizeString(org.country) !== normalizeString(country)) return false;
          if (priorityTier !== "all" && normalizeString(org.priorityTier) !== normalizeString(priorityTier)) return false;
          if (hasDocuments === "yes" && (!Array.isArray(org.documents) || org.documents.length === 0)) return false;
          if (hasDocuments === "no" && Array.isArray(org.documents) && org.documents.length > 0) return false;
          if (!needle) return true;
          return (
            normalizeString(org.organizationId).includes(needle) ||
            normalizeString(org.name).includes(needle) ||
            normalizeString(org.category).includes(needle) ||
            normalizeString(org.subcategory).includes(needle) ||
            normalizeString(org.sector).includes(needle) ||
            normalizeString(org.description).includes(needle) ||
            normalizeString(org.relevanceToYare).includes(needle) ||
            normalizeString(org.contactsToMeet).includes(needle) ||
            normalizeString(org.talkingPoint).includes(needle) ||
            normalizeString(org.notes).includes(needle) ||
            org.contacts.some((contact) => normalizeString(contact.fullName).includes(needle))
          );
        })
        .sort((a, b) => new Date(b.updatedAt) - new Date(a.updatedAt));
    } else {
      const query = { ownerId: req.authUserId };
      if (status !== "all") query.status = status;
      if (priority !== "all") query.priority = priority;
      if (category !== "all") query.category = category;
      if (subcategory !== "all") query.subcategory = subcategory;
      if (country !== "all") query.country = country;
      if (priorityTier !== "all") query.priorityTier = priorityTier;
      if (hasDocuments === "yes") query["documents.0"] = { $exists: true };
      if (hasDocuments === "no") query["documents.0"] = { $exists: false };
      if (search) {
        query.$or = [
          { organizationId: { $regex: search, $options: "i" } },
          { name: { $regex: search, $options: "i" } },
          { category: { $regex: search, $options: "i" } },
          { subcategory: { $regex: search, $options: "i" } },
          { sector: { $regex: search, $options: "i" } },
          { description: { $regex: search, $options: "i" } },
          { relevanceToYare: { $regex: search, $options: "i" } },
          { contactsToMeet: { $regex: search, $options: "i" } },
          { talkingPoint: { $regex: search, $options: "i" } },
          { notes: { $regex: search, $options: "i" } },
          { "contacts.fullName": { $regex: search, $options: "i" } },
        ];
      }
      organizations = await Organization.find(query).sort({ updatedAt: -1 });
    }

    res.json(organizations);
  } catch (error) {
    res.status(500).json({ message: "Failed to fetch organizations", error: error.message });
  }
});

app.post("/api/organizations", requireAuth, async (req, res) => {
  try {
    let organization;
    if (useMemoryStore) {
      organization = createMemoryOrganization(req.body, req.authUserId);
      if (!organization.name) {
        return res.status(400).json({ message: "Organization name is required" });
      }
      memoryOrganizations.push(organization);
    } else {
      organization = await Organization.create({ ...req.body, ownerId: req.authUserId });
    }
    res.status(201).json(organization);
  } catch (error) {
    res.status(400).json({ message: "Failed to create organization", error: error.message });
  }
});

app.put("/api/organizations/:id", requireAuth, async (req, res) => {
  try {
    const allowedFields = [
      "organizationId",
      "name",
      "category",
      "subcategory",
      "sector",
      "address",
      "city",
      "state",
      "country",
      "description",
      "relevanceToYare",
      "contactsToMeet",
      "talkingPoint",
      "priorityTier",
      "status",
      "priority",
      "website",
      "notes",
      "location",
      "nextFollowUpDate",
      "lastContactedDate",
      "guideChecklist",
    ];
    const updates = {};
    for (const field of allowedFields) {
      if (Object.prototype.hasOwnProperty.call(req.body, field)) {
        updates[field] = req.body[field];
      }
    }

    let organization;
    if (useMemoryStore) {
      organization = getMemoryOrgById(req.params.id, req.authUserId);
      if (!organization) {
        return res.status(404).json({ message: "Organization not found" });
      }
      Object.assign(organization, updates, { updatedAt: new Date().toISOString() });
    } else {
      organization = await Organization.findOneAndUpdate({ _id: req.params.id, ownerId: req.authUserId }, updates, {
        new: true,
        runValidators: true,
      });
    }
    if (!organization) {
      return res.status(404).json({ message: "Organization not found" });
    }
    res.json(organization);
  } catch (error) {
    res.status(400).json({ message: "Failed to update organization", error: error.message });
  }
});

app.delete("/api/organizations/:id", requireAuth, async (req, res) => {
  try {
    let result;
    if (useMemoryStore) {
      const index = memoryOrganizations.findIndex((item) => item._id === req.params.id && item.ownerId === req.authUserId);
      if (index >= 0) {
        const [deletedOrg] = memoryOrganizations.splice(index, 1);
        await deleteOrganizationDocuments(deletedOrg);
        result = true;
      } else {
        result = null;
      }
    } else {
      result = await Organization.findOneAndDelete({ _id: req.params.id, ownerId: req.authUserId });
      if (result) {
        await deleteOrganizationDocuments(result);
      }
    }
    if (!result) {
      return res.status(404).json({ message: "Organization not found" });
    }
    res.status(204).send();
  } catch (error) {
    res.status(400).json({ message: "Failed to delete organization", error: error.message });
  }
});

app.delete("/api/organizations", requireAuth, async (req, res) => {
  try {
    if (useMemoryStore) {
      for (const org of getUserMemoryOrganizations(req.authUserId)) {
        await deleteOrganizationDocuments(org);
      }
      for (let index = memoryOrganizations.length - 1; index >= 0; index -= 1) {
        if (memoryOrganizations[index].ownerId === req.authUserId) {
          memoryOrganizations.splice(index, 1);
        }
      }
    } else {
      const organizationsToDelete = await Organization.find({ ownerId: req.authUserId }, { documents: 1 });
      for (const org of organizationsToDelete) {
        await deleteOrganizationDocuments(org);
      }
      await Organization.deleteMany({ ownerId: req.authUserId });
    }
    res.status(204).send();
  } catch (error) {
    res.status(400).json({ message: "Failed to delete all organizations", error: error.message });
  }
});

app.post("/api/organizations/:id/contacts", requireAuth, async (req, res) => {
  try {
    const contact = {
      _id: randomUUID(),
      fullName: req.body.fullName?.trim() || "",
      role: req.body.role?.trim() || "",
      phone: req.body.phone?.trim() || "",
      email: req.body.email?.trim() || "",
      preferredChannel: req.body.preferredChannel?.trim() || "",
      notes: req.body.notes?.trim() || "",
    };
    if (!contact.fullName) {
      return res.status(400).json({ message: "Contact name is required" });
    }

    let org;
    if (useMemoryStore) {
      org = getMemoryOrgById(req.params.id, req.authUserId);
      if (!org) {
        return res.status(404).json({ message: "Organization not found" });
      }
      org.contacts.push(contact);
      org.updatedAt = new Date().toISOString();
    } else {
      org = await Organization.findOne({ _id: req.params.id, ownerId: req.authUserId });
      if (!org) {
        return res.status(404).json({ message: "Organization not found" });
      }
      org.contacts.push(contact);
      await org.save();
    }
    if (!org) {
      return res.status(404).json({ message: "Organization not found" });
    }
    res.status(201).json(org);
  } catch (error) {
    res.status(400).json({ message: "Failed to add contact", error: error.message });
  }
});

app.post("/api/organizations/:id/contacts/bulk", requireAuth, async (req, res) => {
  try {
    const contacts = Array.isArray(req.body.contacts) ? req.body.contacts : [];
    const preparedContacts = contacts
      .map((item) => ({
        _id: randomUUID(),
        fullName: item.fullName?.trim() || "",
        role: item.role?.trim() || "",
        phone: item.phone?.trim() || "",
        email: item.email?.trim() || "",
        preferredChannel: item.preferredChannel?.trim() || "",
        notes: item.notes?.trim() || "",
      }))
      .filter((item) => item.fullName);

    if (preparedContacts.length === 0) {
      return res.status(400).json({ message: "Please provide at least one valid contact." });
    }

    let org;
    if (useMemoryStore) {
      org = getMemoryOrgById(req.params.id, req.authUserId);
      if (!org) {
        return res.status(404).json({ message: "Organization not found" });
      }
      org.contacts.push(...preparedContacts);
      org.updatedAt = new Date().toISOString();
    } else {
      org = await Organization.findOne({ _id: req.params.id, ownerId: req.authUserId });
      if (!org) {
        return res.status(404).json({ message: "Organization not found" });
      }
      org.contacts.push(...preparedContacts);
      await org.save();
    }

    res.status(201).json({ organization: org, added: preparedContacts.length });
  } catch (error) {
    res.status(400).json({ message: "Failed to add contacts in bulk", error: error.message });
  }
});

app.post("/api/organizations/:id/logs", requireAuth, async (req, res) => {
  try {
    const log = {
      _id: randomUUID(),
      activityType: req.body.activityType?.trim() || "",
      summary: req.body.summary?.trim() || "",
      nextAction: req.body.nextAction?.trim() || "",
      date: new Date().toISOString(),
    };
    if (!log.activityType || !log.summary) {
      return res.status(400).json({ message: "Activity type and summary are required." });
    }

    let org;
    if (useMemoryStore) {
      org = getMemoryOrgById(req.params.id, req.authUserId);
      if (!org) {
        return res.status(404).json({ message: "Organization not found" });
      }
      org.outreachLogs.unshift(log);
      org.updatedAt = new Date().toISOString();
    } else {
      org = await Organization.findOne({ _id: req.params.id, ownerId: req.authUserId });
      if (!org) {
        return res.status(404).json({ message: "Organization not found" });
      }
      org.outreachLogs.unshift(log);
      await org.save();
    }
    if (!org) {
      return res.status(404).json({ message: "Organization not found" });
    }
    res.status(201).json(org);
  } catch (error) {
    res.status(400).json({ message: "Failed to add outreach log", error: error.message });
  }
});

app.put("/api/organizations/:id/guide-checklist", requireAuth, async (req, res) => {
  try {
    const guideChecklist = req.body.guideChecklist && typeof req.body.guideChecklist === "object"
      ? req.body.guideChecklist
      : {};

    let org;
    if (useMemoryStore) {
      org = getMemoryOrgById(req.params.id, req.authUserId);
      if (!org) {
        return res.status(404).json({ message: "Organization not found" });
      }
      org.guideChecklist = guideChecklist;
      org.updatedAt = new Date().toISOString();
    } else {
      org = await Organization.findOneAndUpdate(
        { _id: req.params.id, ownerId: req.authUserId },
        { guideChecklist },
        { new: true, runValidators: true }
      );
    }

    if (!org) {
      return res.status(404).json({ message: "Organization not found" });
    }

    res.json(org);
  } catch (error) {
    res.status(400).json({ message: "Failed to update guide checklist", error: error.message });
  }
});

app.post("/api/organizations/:id/documents", requireAuth, upload.single("document"), async (req, res) => {
  try {
    const file = req.file;
    if (!file) {
      return res.status(400).json({ message: "No document uploaded. Use form field 'document'." });
    }
    const version = String(req.body?.version || "").trim();
    const notes = String(req.body?.notes || "").trim();

    const documentEntry = {
      _id: randomUUID(),
      originalName: file.originalname,
      storedName: file.filename,
      mimeType: file.mimetype,
      version,
      notes,
      size: file.size,
      url: `/uploads/organization-docs/${file.filename}`,
      uploadedAt: new Date().toISOString(),
    };

    let org;
    if (useMemoryStore) {
      org = getMemoryOrgById(req.params.id, req.authUserId);
      if (!org) {
        await deleteUploadedFileByUrl(documentEntry.url);
        return res.status(404).json({ message: "Organization not found" });
      }
      org.documents.push(documentEntry);
      org.updatedAt = new Date().toISOString();
    } else {
      org = await Organization.findOne({ _id: req.params.id, ownerId: req.authUserId });
      if (!org) {
        await deleteUploadedFileByUrl(documentEntry.url);
        return res.status(404).json({ message: "Organization not found" });
      }
      org.documents.push(documentEntry);
      await org.save();
    }

    res.status(201).json({ organization: org, document: documentEntry });
  } catch (error) {
    if (req.file?.filename) {
      await deleteUploadedFileByUrl(`/uploads/organization-docs/${req.file.filename}`);
    }
    res.status(400).json({ message: "Failed to upload document", error: error.message });
  }
});

app.delete("/api/organizations/:id/documents/:documentId", requireAuth, async (req, res) => {
  try {
    let org;
    let document;
    if (useMemoryStore) {
      org = getMemoryOrgById(req.params.id, req.authUserId);
      if (!org) {
        return res.status(404).json({ message: "Organization not found" });
      }
      const index = org.documents.findIndex((item) => item._id === req.params.documentId);
      if (index < 0) {
        return res.status(404).json({ message: "Document not found" });
      }
      [document] = org.documents.splice(index, 1);
      org.updatedAt = new Date().toISOString();
    } else {
      org = await Organization.findOne({ _id: req.params.id, ownerId: req.authUserId });
      if (!org) {
        return res.status(404).json({ message: "Organization not found" });
      }
      document = org.documents.id(req.params.documentId);
      if (!document) {
        return res.status(404).json({ message: "Document not found" });
      }
      document.deleteOne();
      await org.save();
    }

    await deleteUploadedFileByUrl(document.url);
    res.status(204).send();
  } catch (error) {
    res.status(400).json({ message: "Failed to delete document", error: error.message });
  }
});

app.use((error, _req, res, next) => {
  if (error instanceof multer.MulterError) {
    if (error.code === "LIMIT_FILE_SIZE") {
      return res.status(400).json({ message: "Document exceeds 10MB upload limit." });
    }
    return res.status(400).json({ message: error.message });
  }
  if (
    error &&
    error.message &&
    (error.message.includes("Unsupported file type") || error.message.includes("Unsupported import file type"))
  ) {
    return res.status(400).json({ message: error.message });
  }
  return next(error);
});

app.post("/api/import/docx", requireAuth, importUpload.single("document"), async (req, res) => {
  try {
    const file = req.file;
    if (!file) {
      return res.status(400).json({ message: "No DOCX file uploaded. Use form field 'document'." });
    }

    const result = await mammoth.extractRawText({ buffer: file.buffer });
    const parsedOrganizations = parseStructuredDocxEntries(result.value);
    const organizationsToImport = parsedOrganizations;
    if (organizationsToImport.length === 0) {
      return res.status(400).json({
        message:
          "No structured organizations found. Ensure DOCX uses labels like ORGANISATION_ID, ORGANISATION_NAME, CATEGORY, SUBCATEGORY, ADDRESS, CITY, STATE, COUNTRY, DESCRIPTION, RELEVANCE_TO_YARE, CONTACTS_TO_MEET, TALKING_POINT, PRIORITY_TIER.",
      });
    }

    const existingKeys = useMemoryStore
      ? new Set(
          getUserMemoryOrganizations(req.authUserId).map((item) => normalizeString(item.organizationId) || normalizeString(item.name))
        )
      : new Set(
          (await Organization.find({ ownerId: req.authUserId }, { organizationId: 1, name: 1 })).map(
            (item) => normalizeString(item.organizationId) || normalizeString(item.name)
          )
        );
    const newOrganizations = organizationsToImport.filter((org) => {
      const key = normalizeString(org.organizationId) || normalizeString(org.name);
      return key && !existingKeys.has(key);
    });
    if (newOrganizations.length > 0) {
      if (useMemoryStore) {
        newOrganizations.forEach((org) => memoryOrganizations.push(createMemoryOrganization(org, req.authUserId)));
      } else {
        await Organization.insertMany(newOrganizations.map((org) => ({ ...org, ownerId: req.authUserId })));
      }
    }

    res.json({
      detected: organizationsToImport.length,
      imported: newOrganizations.length,
      skipped: organizationsToImport.length - newOrganizations.length,
      mode: "structured-labeled-import",
    });
  } catch (error) {
    res.status(500).json({ message: "Failed to import DOCX list", error: error.message });
  }
});

app.get("/guide", (_req, res) => {
  res.sendFile(path.join(__dirname, "public", "guide.html"));
});

app.get("/auth", (_req, res) => {
  res.sendFile(path.join(__dirname, "public", "auth.html"));
});

app.get("/", (_req, res) => {
  res.sendFile(path.join(__dirname, "public", "index.html"));
});

app.get(/^\/(?!api\/).*/, (req, res, next) => {
  if (req.path === "/auth" || req.path === "/guide") {
    return next();
  }
  return res.sendFile(path.join(__dirname, "public", "index.html"));
});

async function start() {
  try {
    await mongoose.connect(MONGODB_URI);
    useMemoryStore = false;
    mongoFallbackReason = "";
    console.log("Connected to MongoDB Atlas.");
  } catch (error) {
    useMemoryStore = true;
    mongoFallbackReason = error.message;
    console.warn(`MongoDB unavailable, starting in memory mode: ${error.message}`);
  } finally {
    app.listen(PORT, () => {
      const mode = useMemoryStore ? "memory-fallback" : "mongodb";
      console.log(`Server running at http://localhost:${PORT} (${mode})`);
    });
  }
}

module.exports = app;

if (require.main === module) {
  start();
}
