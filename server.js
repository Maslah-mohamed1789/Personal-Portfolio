// Simple backend that receives contact-form submissions and
// appends each one as a new row in an Excel file (messages.xlsx).
//
// Run:
//   1. npm install
//   2. node server.js
//
// The server listens on http://localhost:3000
// Every submission is saved to server/messages.xlsx — just open
// that file in Excel any time to see your messages.

const express = require("express");
const cors = require("cors");
const XLSX = require("xlsx");
const path = require("path");
const fs = require("fs");

const app = express();
const PORT = process.env.PORT || 3000;
const FILE_PATH = path.join(__dirname, "messages.xlsx");
const SHEET_NAME = "Messages";

app.use(cors());
app.use(express.json());

// Create the Excel file with headers if it doesn't exist yet
function ensureFileExists() {
  if (!fs.existsSync(FILE_PATH)) {
    const worksheet = XLSX.utils.aoa_to_sheet([["Date", "Name", "Email", "Message"]]);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, SHEET_NAME);
    XLSX.writeFile(workbook, FILE_PATH);
  }
}

// Append one row to the Excel file
function appendMessage({ name, email, message }) {
  ensureFileExists();

  const workbook = XLSX.readFile(FILE_PATH);
  const worksheet = workbook.Sheets[SHEET_NAME];
  const existingRows = XLSX.utils.sheet_to_json(worksheet);

  existingRows.push({
    Date: new Date().toLocaleString(),
    Name: name,
    Email: email,
    Message: message,
  });

  const newWorksheet = XLSX.utils.json_to_sheet(existingRows);
  workbook.Sheets[SHEET_NAME] = newWorksheet;
  XLSX.writeFile(workbook, FILE_PATH);
}

app.post("/submit", (req, res) => {
  const { name, email, message } = req.body;

  if (!name || !email || !message) {
    return res.status(400).json({ error: "Name, email, and message are all required." });
  }

  try {
    appendMessage({ name, email, message });
    res.status(200).json({ status: "ok" });
  } catch (err) {
    console.error("Failed to save message:", err);
    res.status(500).json({ error: "Something went wrong saving your message." });
  }
});

// Optional: quick way to check the server is alive
app.get("/", (req, res) => {
  res.send("Contact form server is running. POST to /submit to save a message.");
});

app.listen(PORT, () => {
  console.log(`Contact form server running at http://localhost:${PORT}`);
  console.log(`Messages will be saved to ${FILE_PATH}`);
});