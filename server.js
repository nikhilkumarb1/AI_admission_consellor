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

  const indiaTimestamp = new Date().toLocaleString("en-IN",{
  timeZone: "Asia/Kolkata",
  day: "2-digit",
  month: "2-digit",
  year: "numeric",
  hour: "2-digit",
  minute: "2-digit",
  second: "2-digit",
  hour12: true,
});

await sheets.spreadsheets.values.append({
  spreadsheetId: process.env.GOOGLE_SHEET_ID,
  range: "Leads!A:H",
  valueInputOption: "USER_ENTERED",
  requestBody: {
    values: [
      [
        indiaTimestamp,
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
    text.includes("i want to apply") ||
    text.includes("want to apply") ||
    text.includes("apply now") ||
    text.includes("start application") ||
    text.includes("start admission process") ||
    text.includes("begin admission process") ||
    text.includes("proceed with admission") ||
    text.includes("proceed with application") ||
    text.includes("i am interested") ||
    text.includes("yes interested") ||
    text.includes("register me")
  );
}

function isCounselorIntent(message) {
  const text = normalize(message);

  return (
    text.includes("connect to counsellor") ||
    text.includes("connect to counselor") ||
    text.includes("connect to councelor") ||
    text.includes("connect me to counsellor") ||
    text.includes("connect me to counselor") ||
    text.includes("connect me to councelor") ||

    text.includes("connect me with counsellor") ||
    text.includes("connect me with counselor") ||
    text.includes("connect me with councelor") ||
    text.includes("connect with counsellor") ||
    text.includes("connect with counselor") ||
    text.includes("connect with councelor") ||

    text.includes("talk to counsellor") ||
    text.includes("talk to counselor") ||
    text.includes("talk to councelor") ||
    text.includes("speak to counsellor") ||
    text.includes("speak to counselor") ||
    text.includes("speak to councelor") ||

    text.includes("counsellor") ||
    text.includes("counselor") ||
    text.includes("councelor") ||
    text.includes("counseller") ||

    text.includes("contact number") ||
    text.includes("contact no") ||
    text.includes("any contact") ||
    text.includes("any number") ||
    text.includes("admission counselor number") ||
    text.includes("admission counsellor number") ||
    text.includes("admission contact") ||
    text.includes("call me") ||
    text.includes("call back") ||
    text.includes("human agent") ||
    text.includes("talk to human")
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

  if (digits.length === 10) {
    digits = "91" + digits;
  }

  return digits;
}

function getCounselorPhones() {
  const phones = process.env.COUNSELOR_PHONES || process.env.COUNSELOR_PHONE || "";

  return phones
    .split(",")
    .map((phone) => formatWhatsAppPhone(phone))
    .filter((phone) => phone && phone.length >= 12);
}

function isStopIntent(message) {
  const text = normalize(message);

  return (
    text === "stop" ||
    text === "unsubscribe" ||
    text.includes("not interested") ||
    text.includes("not intrested") ||
    text.includes("i am not interested") ||
    text.includes("i'm not interested") ||
    text.includes("already taken admission") ||
    text.includes("already take admission") ||
    text.includes("i have taken admission") ||
    text.includes("i already took admission") ||
    text.includes("i already take admission") ||
    text.includes("admission ho gaya") ||
    text.includes("mera admission ho gaya") ||
    text.includes("dont message") ||
    text.includes("don't message") ||
    text.includes("do not message") ||
    text.includes("remove me")
  );
}

function isFeeObjection(message) {
  const text = normalize(message);

  return (
    text.includes("fee is high") ||
    text.includes("fees is high") ||
    text.includes("fees are high") ||
    text.includes("too high") ||
    text.includes("very high") ||
    text.includes("expensive") ||
    text.includes("costly") ||
    text.includes("zyada hai") ||
    text.includes("bahut zyada") ||
    text.includes("budget se bahar") ||
    text.includes("out of budget") ||
    text.includes("not affordable") ||
    text.includes("affordable")
  );
}

function extractBudgetInLakhs(message) {
  const text = normalize(message).replace(/,/g, "");

  let match = text.match(/(\d+(\.\d+)?)\s*(lakh|lakhs|lac|lacs|lak)/);
  if (match) return Number(match[1]);

  match = text.match(/(\d+(\.\d+)?)\s*(cr|crore|crores)/);
  if (match) return Number(match[1]) * 100;

  match = text.match(/(\d+(\.\d+)?)/);
  if (match) {
    const value = Number(match[1]);

    // If student says 800000, convert to 8 lakh
    if (value >= 100000) return value / 100000;

    // If student says 8 or 10, assume lakh
    if (value <= 100) return value;
  }

  return null;
}

function extractFeeInLakhs(feeText) {
  const text = normalize(feeText).replace(/,/g, "");

  let match = text.match(/(\d+(\.\d+)?)\s*(lakh|lakhs|lac|lacs|lak)/);
  if (match) return Number(match[1]);

  match = text.match(/(\d+(\.\d+)?)\s*(cr|crore|crores)/);
  if (match) return Number(match[1]) * 100;

  return null;
}

function getAffordableOptions(data, budgetLakhs, course = "") {
  const options = [];

  for (const row of data) {
    const fee = extractFeeInLakhs(row.Fees || row.Fee || "");

    if (!fee) continue;

    const rowText = normalize(Object.values(row).join(" "));
    const courseText = normalize(course);

    if (courseText) {
      const courseWords = courseText.split(" ").filter((w) => w.length > 1);
      const courseMatched = courseWords.every((word) => rowText.includes(word));

      if (!courseMatched) continue;
    }

    // Show options under budget or slightly above budget by 10%
    if (fee <= budgetLakhs * 1.1) {
      options.push({
        college: row.College || "Amity University",
        campus: row.Campus || "",
        course: row.Course || row.Courses || "",
        fee,
        feeText: row.Fees || "",
      });
    }
  }

  return options
    .sort((a, b) => a.fee - b.fee)
    .slice(0, 5);
}

function isThanksIntent(message) {
  const text = normalize(message);

  return (
    text === "thanks" ||
    text === "thank you" ||
    text === "thankyou" ||
    text === "ok thanks" ||
    text === "okay thanks" ||
    text === "ok" ||
    text === "okay"
  );
}

function isFriendAdmissionQuery(message) {
  const text = normalize(message);

  return (
    text.includes("my friend") ||
    text.includes("one of my friends") ||
    text.includes("for my friend") ||
    text.includes("not my admission") ||
    text.includes("its not my admission") ||
    text.includes("it's not my admission") ||
    text.includes("for a friend")
  );
}

function isCompartmentQuery(message) {
  const text = normalize(message);

  return (
    text.includes("compartment") ||
    text.includes("backlog") ||
    text.includes("failed") ||
    text.includes("fail in") ||
    text.includes("physics compartment") ||
    text.includes("maths compartment") ||
    text.includes("chemistry compartment")
  );
}

function isPositiveReply(message) {
  const text = normalize(message);

  return (
    text === "yes" ||
    text === "yes please" ||
    text === "please" ||
    text === "sure" ||
    text === "ok" ||
    text === "okay"
  );
}

function isAccountsPaymentIssue(message) {
  const text = normalize(message);

  return (
    text.includes("upi") ||
    text.includes("net banking") ||
    text.includes("debit card") ||
    text.includes("credit card") ||
    text.includes("payment limit") ||
    text.includes("max payment") ||
    text.includes("maximum payment") ||
    text.includes("not able to pay") ||
    text.includes("unable to pay") ||
    text.includes("can't pay") ||
    text.includes("cannot pay") ||
    text.includes("payment failed") ||
    text.includes("payment deducted") ||
    text.includes("receipt not received") ||
    text.includes("transaction") ||
    text.includes("accounts office") ||
    text.includes("account office") ||
    text.includes("fee office")
  );
}

function isNoNeedIntent(message) {
  const text = normalize(message);

  return (
    text === "no" ||
    text === "nope" ||
    text === "nothing" ||
    text === "nothing else" ||
    text === "not now" ||
    text === "no thanks" ||
    text === "no thank you" ||
    text === "nahi" ||
    text === "kuch nahi" ||
    text === "abhi nahi"
  );
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


// 1. Stop/Not interested check
if (memory.stage === "stopped") {
  return {
    type: "stopped",
    memory,
    answer: ""
  };
}

if (isStopIntent(message)) {
  memory.stage = "stopped";

  const status = normalize(message).includes("admission")
    ? "Already Admitted"
    : "Not Interested";

  const lead = {
    name: memory.name || "",
    phone: memory.phone || phone || "",
    email: memory.email || "",
    campus: memory.campus || "",
    course: memory.course || "",
    message: `[${status}] ${message}`,
  };

  await saveLead(lead);

  await updateLeadStatus(memory.phone || phone, status).catch(() => {});

  return {
    type: "student_stopped",
    memory,
    answer:
      status === "Already Admitted"
        ? `Thank you for letting us know. ✅\n\nWe have updated your status as already admitted.`
        : `Thank you for your response. I completely understand. If you ever plan to pursue higher education in the future, feel free to contact me. Wishing you all the best!.`
  };
}

// If student says no / nothing, don't greet again
if (isNoNeedIntent(message)) {
  memory.stage = "new";

  return {
    type: "no_need",
    memory,
    answer:
      `No problem. 😊\n\n` +
      `If you need help later with fees, eligibility, hostel, placements, or admission process, just message me.`
  };
}


// If bot suggested counselor and student says yes
if (memory.stage === "counselor_suggested" && isPositiveReply(message)) {
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

  return {
    type: "counselor_handoff",
    memory,
    answer:
      `Sure, I’m connecting you with our admission counselor. ✅\n\n` +
      `Kindly call our admission counselor\n\n` +
      `Mr. Rahul Diwan - 76693 15881 - Admission Manager\n` +
      `Ms. Jaya - 87965 32575 - Education Counselor`
  };
}


// If student is asking for a friend, don't start application flow directly
if (isFriendAdmissionQuery(message)) {
  memory.forWhom = "friend";

  return {
    type: "friend_query",
    memory,
    answer:
      `Sure, I can help with your friend's admission query. ✅\n\n` +
      `Please tell me:\n` +
      `• Preferred campus\n` +
      `• Course name\n` +
      `• 12th subjects/percentage\n\n` +
      `For example: Lucknow, B.Pharma, PCB with 60%.`
  };
}


// If student says thanks, don't collect lead or send form
if (isThanksIntent(message)) {
  return {
    type: "thanks",
    memory,
    answer:
      `You're welcome. 😊\n\n` +
      `Let me know if you need help with fees, eligibility, hostel, placement, or admission process.`
  };
}


// If student asks about compartment/backlog, suggest counselor
if (isCompartmentQuery(message)) {
  memory.stage = "counselor_suggested";

  return {
    type: "counselor_suggested",
    memory,
    answer:
      `In case of a compartment/backlog, eligibility depends on the university rules and final document verification.\n\n` +
      `It would be better to confirm this with an admission counselor.\n\n` +
      `Would you like me to connect you with a counselor?`
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

  const lastChat = memory.lastMessages
  ?.slice(-5)
  .map((msg, index) => `${index + 1}. ${msg}`)
  .join("\n");

const studentPhone = formatWhatsAppPhone(memory.phone || phone);

const counselorMessage =
  `🚨 Counselor Requested\n\n` +
  `A student wants to talk to a counselor.\n\n` +
  `Student Phone: ${studentPhone || "Not available"}\n` +
  `Name: ${memory.name || "Not provided"}\n` +
  `Campus: ${memory.campus || "Not provided"}\n` +
  `Course: ${memory.course || "Not provided"}\n\n` +
  `Student Query:\n${message}\n\n` +
  `Last Chat:\n${lastChat || "No previous chat available"}\n\n` +
  `Action:\nOpen AiSensy Inbox and reply to this student manually.`;

const counselorPhones = getCounselorPhones();

if (counselorPhones.length) {
  for (const counselorPhone of counselorPhones) {
    console.log("Sending counselor alert to:", counselorPhone);

    await sendAiSensyReply(counselorPhone, counselorMessage).catch((err) => {
      console.log(
        `Counselor alert failed for ${counselorPhone}:`,
        err.response?.data || err.message
      );
    });
  }
} else {
  console.log("No valid counselor phone numbers found in .env");
}

  return {
    type: "counselor_handoff",
    memory,
    answer:
      `Sure, I’m connecting you with our admission counselor. ✅\n\n` +
      `Kindly call our admission counselor\n\n` +
      `Mr. Rahul Diwan - 76693 15881 - Admission Manager\n` +
      `Ms. Jaya - 87965 32575 - Education Counselor`
  };
}

// Payment / UPI / transaction / accounts office issue
if (isAccountsPaymentIssue(message)) {
  memory.stage = "accounts_payment_issue";

  const lead = {
    name: memory.name || "",
    phone: memory.phone || phone || "",
    email: memory.email || "",
    campus: memory.campus || "",
    course: memory.course || "",
    message: `[Accounts/Payment Issue] ${message}`,
  };

  await saveLead(lead);
  await updateLeadStatus(memory.phone || phone, "Accounts/Payment Issue").catch(() => {});

  return {
    type: "accounts_payment_issue",
    memory,
    answer:
      `I understand your concern.\n\n` +
      `I can help with admission and course-related queries like fees, eligibility, hostel, placements, and admission process.\n\n` +
      `For payment mode, UPI limit, transaction, receipt, or fee payment issues, please contact the college accounts/fee office directly.`
  };
}


// Student says fee is high / expensive
if (isFeeObjection(message)) {
  memory.stage = "asking_budget";

  return {
    type: "budget_ask",
    memory,
    answer:
      `I understand, fees can be a concern.\n\n` +
      `What is your approximate budget for the course?\n\n` +
      `Example:\n` +
      `• 8 Lakhs\n` +
      `• 10 Lakhs\n` +
      `• 12 Lakhs`
  };
}

// Student gives budget after fee objection
if (memory.stage === "asking_budget") {
  const budgetLakhs = extractBudgetInLakhs(message);

  if (!budgetLakhs) {
    return {
      type: "budget_missing",
      memory,
      answer:
        `Please share your approximate budget in lakhs.\n\n` +
        `Example: 8 Lakhs, 10 Lakhs, 12 Lakhs`
    };
  }

  const data = await loadSheetData();

  const affordableOptions = getAffordableOptions(
    data,
    budgetLakhs,
    memory.course || ""
  );

  memory.stage = "budget_suggested";

  if (!affordableOptions.length) {
    return {
      type: "no_budget_options",
      memory,
      answer:
        `Based on your budget of ${budgetLakhs} Lakhs, I couldn't find an exact matching option in the current data.\n\n` +
        `But I can connect you with a counselor who can suggest scholarships, flexible options, or nearby alternatives.`
    };
  }

  const optionText = affordableOptions
    .map((item, index) => {
      return (
        `${index + 1}. ${item.campus || item.college}\n` +
        `Course: ${item.course || "Available course"}\n` +
        `Fees: ${item.feeText || item.fee + " Lakhs"}`
      );
    })
    .join("\n\n");

  return {
    type: "budget_options",
    memory,
    answer:
      `Based on your budget of around ${budgetLakhs} Lakhs, these options may be suitable:\n\n` +
      `${optionText}\n\n` +
      `Would you like to connect with a counselor for scholarship or admission guidance?`
  };
}


  // Step 2: user has shared details after bot asked
  if (memory.stage === "collecting_details") {
    if (isThanksIntent(message) || isFriendAdmissionQuery(message)) {
    memory.stage = "new";

    return {
      type: "not_collecting_details",
      memory,
      answer:
        `No problem. 😊\n\n` +
        `Please ask your query directly, like course, eligibility, fees, hostel, or placement details.`
    };
  }
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

   memory.stage = "counselor_followup";

return {
  type: "lead_saved",
  memory,
  answer:
    `Thank you! ✅\n\n` +
    `Your details have been received successfully.\n\n` +
    `Our admission counselor will contact you shortly for further guidance.`
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
        `My name is John Doe, phone 9999999999, email john.doe@gmail.com, campus Lucknow, course BTech CSE`
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
      console.log("AI reply skipped because student is stopped/handed over");
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