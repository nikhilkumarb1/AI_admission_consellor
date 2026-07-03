import axios from "axios";
import express from "express";
import cors from "cors";
import dotenv from "dotenv";
import OpenAI from "openai";
import { google } from "googleapis";



dotenv.config();

const app = express();
app.use(cors());
app.use(express.json());

const PORT = process.env.PORT || 3000;

const openai = new OpenAI({
  apiKey: process.env.OPENAI_API_KEY,
});

const conversations = new Map();

async function getSheetsClient() {
  const auth = new google.auth.GoogleAuth({
    credentials: {
      client_email: process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL,
      private_key: process.env.GOOGLE_PRIVATE_KEY.replace(/\\n/g, "\n"),
    },
    scopes: ["https://www.googleapis.com/auth/spreadsheets"],
  });

  return google.sheets({ version: "v4", auth });
}

let sheetCache = null;
let sheetCacheTime = 0;
const CACHE_DURATION = 5 * 60 * 1000;

async function loadSheetData() {
  const now = Date.now();

  // Return cached data if it's still fresh
  if (sheetCache && (now - sheetCacheTime) < CACHE_DURATION) {
    console.log("📦 Using cached Google Sheet data");
    return sheetCache;
  }

  console.log("📥 Loading Google Sheet from API...");

  const sheets = await getSheetsClient();

  const response = await sheets.spreadsheets.values.get({
    spreadsheetId: process.env.GOOGLE_SHEET_ID,
    range: "'Amity University'!A:Z",
  });

  const rows = response.data.values || [];

  if (!rows.length) return [];

  const headers = rows[0];

  const data = rows.slice(1).map((row) => {
    const obj = {};

    headers.forEach((header, index) => {
      obj[header.trim()] = row[index] || "";
    });

    return obj;
  });

  // Save to cache
  sheetCache = data;
  sheetCacheTime = now;

  return data;
}
async function saveLead({ name, phone, email, campus, course, message }) {
  const sheets = await getSheetsClient();

  await sheets.spreadsheets.values.append({
    spreadsheetId: process.env.GOOGLE_SHEET_ID,
    range: "Leads!A:H",
    valueInputOption: "USER_ENTERED",
    requestBody: {
      values: [
        [
          new Date().toLocaleString("en-IN"),
          name || "",
          phone || "",
          email || "",
          campus || "",
          course || "",
          message || "",
          "New Lead",
        ],
      ],
    },
  });
}

async function updateLeadStatus(phone, status) {
  const sheets = await getSheetsClient();

  const response = await sheets.spreadsheets.values.get({
    spreadsheetId: process.env.GOOGLE_SHEET_ID,
    range: "Leads!A:H",
  });

  const rows = response.data.values || [];

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];

    if (row[2] === phone) {
      await sheets.spreadsheets.values.update({
        spreadsheetId: process.env.GOOGLE_SHEET_ID,
        range: `Leads!H${i + 1}`,
        valueInputOption: "USER_ENTERED",
        requestBody: {
          values: [[status]],
        },
      });

      console.log(`Lead ${phone} updated to ${status}`);
      return;
    }
  }

  console.log("Lead not found");
}

function getSessionId(phone = "", fallback = "test-user") {
  return phone || fallback;
}

function getConversation(sessionId) {
  if (!conversations.has(sessionId)) {
    conversations.set(sessionId, {
      campus: "",
      course: "",
      name: "",
      email: "",
      phone: "",
      lastMessages: [],
      stage: "new",
      documents: {
      tenth: false,
      twelfth: false,
      aadhaar: false,
      photo: false,
      entrance: false
      }
    });
  }

  return conversations.get(sessionId);
}

function updateConversation(sessionId, message, lead = {}) {
  const memory = getConversation(sessionId);

  const campus = detectCampus(message);
  const course = detectCourse(message);
  const email = extractEmail(message);
  const phone = extractPhone(message);

  if (campus) memory.campus = campus;
  if (course) memory.course = course;
  if (email) memory.email = email;
  if (phone) memory.phone = phone;

  if (lead.name) memory.name = lead.name;
  if (lead.email) memory.email = lead.email;
  if (lead.phone) memory.phone = lead.phone;
  if (lead.campus) memory.campus = lead.campus;
  if (lead.course) memory.course = lead.course;

  memory.lastMessages.push(message);

  if (memory.lastMessages.length > 5) {
    memory.lastMessages.shift();
  }

  conversations.set(sessionId, memory);

  return memory;
}

function normalize(text) {
  return String(text || "")
    .toLowerCase()
    .replace(/\./g, "")
    .replace(/-/g, " ")
    .replace(/:/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function detectCampus(message) {
  const text = normalize(message);
   const campusAliases = {
    noida: ["noida"],
    bangalore: ["bangalore", "banglore", "bengaluru", "bengalooru"],
    lucknow: ["lucknow"],
    jaipur: ["jaipur"],
    gurgaon: ["gurgaon", "gurugram"],
    gwalior: ["gwalior"],
    mumbai: ["mumbai"],
    raipur: ["raipur"],
    mohali: ["mohali"],
    hyderabad: ["hyderabad"]
  };
   for (const [campus, aliases] of Object.entries(campusAliases)) {
    if (aliases.some(alias => text.includes(alias))) {
      return campus;
    }
  }

  return "";
}
function detectTopic(message) {
  const text = normalize(message);

  if (
    text.includes("fee") ||
    text.includes("fees") ||
    text.includes("cost") ||
    text.includes("charges") ||
    text.includes("price")
  ) return "Fees";

  if (
    text.includes("course") ||
    text.includes("courses") ||
    text.includes("program") ||
    text.includes("programs")
  ) return "Courses";

  if (text.includes("hostel")) return "Hostel";
  if (text.includes("placement") || text.includes("package")) return "Placement";
  if (text.includes("scholarship") || text.includes("schollarship")) return "Scholarship";
  if (text.includes("eligibility")) return "Eligibility";
  if (text.includes("application fee")) return "Application_Fee";
  if (text.includes("admission") || text.includes("apply") || text.includes("process")) return "Admission_Process";

  return "General";
}

function detectCourse(message) {
  const text = normalize(message);

  const courses = [
    "btech cse", "btech it", "btech", "bba", "bca", "mba",
    "mca", "ba llb", "bcom", "bsc", "msc"
  ];

  return courses.find((course) => text.includes(course)) || "";
}

function isApplyIntent(message) {
  const text = normalize(message);

  return (
    text.includes("apply") ||
    text.includes("admission") ||
    text.includes("register") ||
    text.includes("proceed") ||
    text.includes("interested") ||
    text.includes("call me")
  );
}

function extractEmail(message) {
  const match = message.match(/[^\s@]+@[^\s@]+\.[^\s@]+/);
  return match ? match[0] : "";
}

function extractPhone(message) {
  const match = message.match(/(?:\+91[-\s]?)?[6-9]\d{9}/);
  return match ? match[0] : "";
}

async function extractLeadDetails(message) {
  const prompt = `
Extract lead details from the student message.

Return only valid JSON with these keys:
name, phone, email, campus, course

If missing, keep value empty.

Message:
${message}
`;

  const response = await openai.chat.completions.create({
    model: "gpt-4.1-mini",
    temperature: 0,
    messages: [{ role: "user", content: prompt }],
  });

  try {
    return JSON.parse(response.choices[0].message.content);
  } catch {
    return {
      name: "",
      phone: extractPhone(message),
      email: extractEmail(message),
      campus: detectCampus(message),
      course: detectCourse(message),
    };
  }
}

function findRelevantRows(data, message, memory, entities = {}) {
  const campus = entities.campus || detectCampus(message) || memory.campus || "";
  const course = entities.course || detectCourse(message) || memory.course || "";

  const detectedTopic = entities.topic || detectTopic(message);

  const finalTopic =
    detectedTopic && detectedTopic !== "General"
      ? detectedTopic
      : memory.lastTopic || "General";

  let filtered = data;

  if (campus) {
    filtered = filtered.filter((row) => {
      const campusText = normalize(row.Campus || "");
      return (
        campusText.includes(campus) ||
        (campus === "gurgaon" && campusText.includes("gurugram")) ||
        (campus === "gurugram" && campusText.includes("gurgaon")) ||
        (campus === "bangalore" && campusText.includes("bengaluru")) ||
        (campus === "bengaluru" && campusText.includes("bangalore"))
      );
    });
  }

const topicNeedsCourse = {
  Fees: true,
  Courses: false,
  Eligibility: true,
  Application_Fee: true,
  Hostel: false,
  Placement: false,
  Scholarship: false,
  Admission_Process: false,
  General: false
};

if (course && topicNeedsCourse[finalTopic]) {
    const courseWords = normalize(course).split(" ");

    filtered = filtered.filter((row) => {
      const fullRow = normalize(Object.values(row).join(" "));
      return courseWords.every((word) => fullRow.includes(word));
    });
  }

  if (!filtered.length) {
    filtered = data.filter((row) => {
      const fullRow = normalize(Object.values(row).join(" "));
      const words = normalize(message).split(" ").filter((w) => w.length > 2);
      return words.some((word) => fullRow.includes(word));
    });
  }

  return {
    campus,
    course,
    topic: finalTopic,
    rows: filtered.slice(0, 5),
  };
}
async function formatAnswer(message, searchResult) {
  const prompt = `
You are an Amity University admission counselor on WhatsApp.

Rules:
1. Answer only using matched data.
2. Do not invent details.
3. Keep answer short and WhatsApp-friendly.
4. If exact answer is unavailable, say you will connect them with a counselor.

Student question:
${message}

Detected:
Campus: ${searchResult.campus || "Not detected"}
Course: ${searchResult.course || "Not detected"}
Topic: ${searchResult.topic}

Matched data:
${JSON.stringify(searchResult.rows, null, 2)}
`;

  const response = await openai.chat.completions.create({
    model: "gpt-4.1-mini",
    temperature: 0.2,
    messages: [
      { role: "system", content: "You are a helpful Amity admission counselor." },
      { role: "user", content: prompt },
    ],
  });

  return response.choices[0].message.content;
}



app.get("/", (req, res) => {
  res.send("Amity Bot Running 🚀");
});

function detectDocument(message) {
  const text = normalize(message);

  if (text.includes("10th") || text.includes("tenth")) return "tenth";
  if (text.includes("12th") || text.includes("twelfth")) return "twelfth";
  if (text.includes("aadhaar") || text.includes("aadhar")) return "aadhaar";
  if (text.includes("photo") || text.includes("passport")) return "photo";
  if (text.includes("jee") || text.includes("cuet") || text.includes("scorecard")) return "entrance";

  return "";
}

function getMissingDocuments(memory) {
  const required = ["tenth", "twelfth", "aadhaar", "photo"];

  return required.filter((doc) => !memory.documents?.[doc]);
}

function documentLabel(doc) {
  const labels = {
    tenth: "10th Marksheet",
    twelfth: "12th Marksheet",
    aadhaar: "Aadhaar Card",
    photo: "Passport Size Photo",
    entrance: "CUET/JEE Scorecard"
  };

  return labels[doc] || doc;
}

async function extractEntities(message, memory = {}) {
  const prompt = `
Extract entities from this student message for an Amity admission chatbot.

Return ONLY valid JSON.

Allowed campuses:
noida, bangalore, lucknow, jaipur, gurgaon, gwalior, mumbai, raipur, mohali, hyderabad

Allowed topics:
Fees, Courses, Hostel, Placement, Scholarship, Eligibility, Admission_Process, Application_Fee, General

Rules:
- Correct spelling mistakes like banglore -> bangalore.
- If user asks "what about noida", keep previous course/topic from memory.
- If something is missing, use memory.
- If still unknown, keep empty.

Memory:
${JSON.stringify(memory)}

Student message:
${message}

Return format:
{
  "campus": "",
  "course": "",
  "topic": "",
  "intent": "information"
}
`;

  const response = await openai.chat.completions.create({
    model: "gpt-4o-mini",
    temperature: 0,
    messages: [{ role: "user", content: prompt }],
  });

  try {
    return JSON.parse(response.choices[0].message.content);
  } catch {
    return {
      campus: "",
      course: "",
      topic: "General",
      intent: "information",
    };
  }
}


const FORM_LINK = "https://forms.gle/2XEo8vgP79yWVAF3A";

async function handleStudentMessage(message, phone = "") {
  const sessionId = getSessionId(phone);
  let memory = getConversation(sessionId);

  updateConversation(sessionId, message);

  memory = getConversation(sessionId);

  const entities = await extractEntities(message, memory);

if (entities.campus) memory.campus = entities.campus;
if (entities.course) memory.course = entities.course;
if (entities.topic && entities.topic !== "General") memory.lastTopic = entities.topic;

const newCampus = detectCampus(message);
const newCourse = detectCourse(message);
const newTopic = detectTopic(message);

if (newCampus) memory.campus = newCampus;
if (newCourse) memory.course = newCourse;
if (newTopic !== "General") memory.lastTopic = newTopic;

  const documentType = detectDocument(message);

if (documentType && memory.stage === "documents_pending") {
  memory.documents[documentType] = true;

  const missingDocs = getMissingDocuments(memory);

  if (missingDocs.length === 0) {
    memory.stage = "ready_for_counselor";

    await updateLeadStatus(
    memory.phone || phone,
    "Ready for Counselor"
);

    return {
      type: "documents_complete",
      memory,
      answer:
        "All required documents received ✅\n\nYour application is now ready for counselor verification. Our counselor will contact you shortly."
    };
  }

  return {
    type: "document_received",
    memory,
    answer:
      `${documentLabel(documentType)} received ✅\n\n` +
      `Please upload remaining documents:\n` +
      missingDocs.map((doc) => `• ${documentLabel(doc)}`).join("\n")
  };
}

  if (isApplyIntent(message)) {
    const details = await extractLeadDetails(message);

    memory = updateConversation(sessionId, message, details);

    const lead = {
      name: memory.name || "",
      phone: memory.phone || phone || "",
      email: memory.email || "",
      campus: memory.campus || "",
      course: memory.course || "",
      message,
    };

    await saveLead(lead);


    memory.stage = "documents_pending";

    return {
      type: "lead",
      memory,
      answer:
  `Great! Your admission interest has been saved ✅\n\n` +
  `Details received:\n` +
  `Name: ${lead.name || "Not provided"}\n` +
  `Phone: ${lead.phone || "Not provided"}\n` +
  `Email: ${lead.email || "Not provided"}\n` +
  `Campus: ${lead.campus || "Not provided"}\n` +
  `Course: ${lead.course || "Not provided"}\n\n` +
   `Please upload your admission documents using the secure Google Form:\n\n` +
  `${FORM_LINK}\n\n` +
  `• 10th Marksheet\n` +
  `• 12th Marksheet\n` +
  `• Aadhaar Card\n` +
  `• Passport Size Photo\n\n` +
  `Optional: CUET/JEE Scorecard`
    };
  }

  const data = await loadSheetData();

  const messageWithMemory = `
Student current message: ${message}

Known conversation memory:
Campus: ${memory.campus || "Not provided"}
Course: ${memory.course || "Not provided"}
Topic: ${memory.lastTopic || "Not provided"}
Name: ${memory.name || "Not provided"}
Email: ${memory.email || "Not provided"}
Phone: ${memory.phone || "Not provided"}

Use this memory if the student asks short questions like "fees?", "hostel?", "placement?".
Use memory if student asks follow-ups like:
- what about noida?
- and jaipur?
- hostel?
- fees?
`;

const searchMessage = `
Student current message: ${message}

Final context:
Campus: ${memory.campus || "Not provided"}
Course: ${memory.course || "Not provided"}
Topic: ${memory.lastTopic || "Not provided"}
`;


console.log("\n========== MEMORY ==========");
console.log(memory);

console.log("\n========== USER MESSAGE ==========");
console.log(message);

const searchResult = findRelevantRows(data, message, memory, entities);
console.log("\n========== SEARCH ==========");
console.log(searchResult);
const answer = await formatAnswer(searchMessage, searchResult);

  return {
    type: "answer",
    memory,
    detected: {
      campus: searchResult.campus,
      course: searchResult.course,
      topic: searchResult.topic,
    },
    matched_rows: searchResult.rows.length,
    answer,
  };
}

app.post("/test", async (req, res) => {
  try {
    const { message, phone } = req.body;

    const result = await handleStudentMessage(message, phone);

    res.json({
      success: true,
      ...result,
    });
  } catch (error) {
    res.status(500).json({
      success: false,
      error: error.message,
    });
  }
});

app.get("/webhook/aisensy", (req, res) => {
  return res.status(200).send("OK");
});

function cleanTemplateText(text) {
  return String(text || "")
    .replace(/\n/g, " ")
    .replace(/\t/g, " ")
    .replace(/\s{4,}/g, " ")
    .trim();
}

async function sendAiSensyReply(phone, message) {
  const response = await axios.post(
    `https://apis.aisensy.com/project-apis/v1/project/${process.env.AISENSY_PROJECT_ID}/messages`,
    {
      to: phone,
      type: "text",
      recipient_type: "individual",
      text: {
        body: message,
      },
    },
    {
      headers: {
        Accept: "application/json",
        "Content-Type": "application/json",
        "X-AiSensy-Project-API-Pwd": process.env.AISENSY_PROJECT_API_PWD,
      },
    }
  );

  return response.data;
}

app.post("/webhook/aisensy", async (req, res) => {
  try {
    const body = req.body;
    const topic = body.topic;

    if (topic !== "message.sender.user") {
      return res.status(200).send("OK");
    }

    const message =
      body.data?.message?.message_content?.text ||
      body.message ||
      body.text ||
      "";

    const phone =
      body.data?.message?.phone_number ||
      body.phone ||
      "";

    if (!message || !phone) {
      return res.status(200).send("OK");
    }

    const result = await handleStudentMessage(message, phone);

   await sendAiSensyReply(phone, result.answer);
   console.log("AI Reply Sent:", result.answer);

    return res.status(200).send("OK");
  } catch (error) {
    console.error("Webhook Error:", error.response?.data || error.message);
    return res.status(200).send("OK");
  }
});

app.listen(PORT, () => {
  console.log(`Server running on ${PORT}`);
});