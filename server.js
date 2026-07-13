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
    text === "yes" ||
    text === "yes interested" ||
    text === "yes i am interested" ||
    text.includes("apply") ||
    text.includes("admission") ||
    text.includes("register") ||
    text.includes("proceed") ||
    text.includes("interested") ||
    text.includes("call me")
    );
}

function isCounselorIntent(message) {
  const text = normalize(message);

  return (
    text.includes("connect to counsellor") ||
    text.includes("connect to counselor") ||
    text.includes("connect me to counsellor") ||
    text.includes("connect me to counselor") ||
    text.includes("talk to counsellor") ||
    text.includes("talk to counselor") ||
    text.includes("speak to counsellor") ||
    text.includes("speak to counselor") ||
    text.includes("counsellor call") ||
    text.includes("counselor call") ||
    text.includes("counseller") ||
    text.includes("counsellor") ||
    text.includes("counselor") ||
    text.includes("human agent") ||
    text.includes("talk to human") ||
    text.includes("connect with human") ||
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

function isNegativeIntent(message) {
  const text = normalize(message);
  return ["no", "nope", "not now", "later", "maybe later"].includes(text);
}

function formatWhatsAppPhone(phone) {
  let digits = String(phone || "").replace(/\D/g, "");

  // If Indian 10 digit number, add 91
  if (digits.length === 10) {
    digits = "91" + digits;
  }

  return digits;
}


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

  if (memory.campus === "noida") {
  return {
    type: "noida_reserved",
    memory,
    answer:
      `Sorry, seats for Amity Noida are currently reserved.\n\n` +
      `You can choose another campus:\n` +
      `• Lucknow\n` +
      `• Jaipur\n` +
      `• Mumbai\n` +
      `• Bengaluru\n` +
      `• Raipur\n` +
      `• Mohali\n` +
      `• Gurgaon\n` +
      `• Gwalior\n` +
      `• Hyderabad`
  };
}

// If already handed over, stop AI replies
if (memory.stage === "counselor_handoff") {
  return {
    type: "counselor_handoff_hold",
    memory,
    answer: ""
  };
}

// Student asks for counselor
if (isCounselorIntent(message)) {
  memory.stage = "counselor_handoff";

  const lead = {
    name: memory.name || "",
    phone: memory.phone || phone || "",
    email: memory.email || "",
    campus: memory.campus || "",
    course: memory.course || "",
    message: `[Counselor Requested] ${message}`,
  };

  await saveLead(lead);

  await updateLeadStatus(memory.phone || phone, "Counselor Requested").catch(() => {});

  const counselorMessage =
    `🚨 Counselor Requested\n\n` +
    `A student wants to talk to a counselor.\n\n` +
    `Student Phone: ${memory.phone || phone || "Not available"}\n` +
    `Name: ${memory.name || "Not provided"}\n` +
    `Campus: ${memory.campus || "Not provided"}\n` +
    `Course: ${memory.course || "Not provided"}\n` +
    `Message: ${message}\n\n` +
    `Please check AiSensy chat / Leads Google Sheet and follow up.`;

 const counselorPhone = formatWhatsAppPhone(process.env.COUNSELOR_PHONE);

console.log("Sending counselor alert to:", counselorPhone);

if (counselorPhone && counselorPhone.length >= 12) {
  await sendAiSensyReply(counselorPhone, counselorMessage).catch((err) => {
    console.log("Counselor alert failed:", err.response?.data || err.message);
  });
} else {
  console.log("Invalid COUNSELOR_PHONE in .env");
}

  return {
    type: "counselor_handoff",
    memory,
    answer:
      `Sure, I’m connecting you with our admission counselor. ✅\n\n` +
      `Our counselor will contact you shortly for further assistance.`
  };
}

  // Step 3: after details saved, user says Yes/Proceed => send form link
  if (memory.stage === "details_received" && isApplyIntent(message)) {
    memory.stage = "documents_pending";

    return {
      type: "form_link",
      memory,
      answer:
        `Perfect! ✅\n\n` +
        `Please upload your admission documents using the link below:\n\n` +
        `${FORM_LINK}\n\n` +
        `Required Documents:\n` +
        `• 10th Marksheet\n` +
        `• 12th Marksheet\n` +
        `• Aadhaar Card\n` +
        `• Passport Size Photo\n\n` +
        `Our counselor will verify your documents and contact you shortly.`
    };
  }

  // Step 3 alternate: user says No/Later
  if (memory.stage === "details_received" && isNegativeIntent(message)) {
    memory.stage = "counselor_followup";

    return {
      type: "handoff",
      memory,
      answer:
        `No problem. Our counselor will call you shortly for more queries.`
    };
  }

  // Step 2: user has shared details after bot asked
  if (memory.stage === "collecting_details") {
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

    memory.stage = "details_received";

    return {
      type: "lead_saved",
      memory,
      answer:
        `Thank you! ✅\n\n` +
        `Your details have been received successfully.\n\n` +
        `Can we begin with the admission process?`
    };
  }

  // Step 1: user says Yes/Interested/Apply => ask for details only
  if (isApplyIntent(message)) {
    memory.stage = "collecting_details";

    return {
      type: "collecting_details",
      memory,
      answer:
        `Great! 🎉\n\n` +
        `Please share the following details:\n\n` +
        `👤 Full Name\n` +
        `📞 Phone Number\n` +
        `📧 Email ID\n` +
        `🏫 Preferred Campus\n` +
        `🎓 Preferred Course\n\n` +
        `Example:\n` +
        `My name is John Doe, phone 9876543210, email john.doe@gmail.com, campus Lucknow, course BTech CSE`
    };
  }

  const data = await loadSheetData();

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

  // Ask for campus if missing
if (
  searchResult.topic !== "General" &&
  !searchResult.campus &&
  ["Fees", "Hostel", "Placement", "Scholarship", "Admission_Process"].includes(searchResult.topic)
) {
  return {
    type: "missing_campus",
    memory,
    answer:
      `The information depends on the campus.\n\n` +
      `Could you please tell me which Amity campus you're interested in?\n\n` +
      `• Noida\n` +
      `• Lucknow\n` +
      `• Jaipur\n` +
      `• Mumbai\n` +
      `• Bengaluru\n` +
      `• Raipur\n` +
      `• Mohali\n` +
      `• Gwalior\n` +
      `• Gurgaon\n` +
      `• Hyderabad`
  };
}


// Ask for course if missing
if (
  searchResult.topic === "Fees" &&
  searchResult.campus &&
  !searchResult.course
) {
  return {
    type: "missing_course",
    memory,
    answer:
      `Sure! Which course are you interested in?\n\n` +
      `• B.Tech CSE\n` +
      `• B.Tech AI & ML\n` +
      `• B.Tech IT\n` +
      `• BBA\n` +
      `• B.Sc\n` +
      `• BCA\n` +
      `• MBA\n` +
      `• MCA`
  };
}

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

   if (result.answer && result.answer.trim()) {
  await sendAiSensyReply(phone, result.answer);
  console.log("AI Reply Sent:", result.answer);
} else {
  console.log("AI reply skipped because student is handed over to counselor");
}

    return res.status(200).send("OK");
  } catch (error) {
    console.error("Webhook Error:", error.response?.data || error.message);
    return res.status(200).send("OK");
  }
});

app.listen(PORT, () => {
  console.log(`Server running on ${PORT}`);
});