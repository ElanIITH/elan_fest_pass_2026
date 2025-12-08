const { google } = require("googleapis");
const nodemailer = require("nodemailer");
const fs = require("fs");
const path = require("path");
const handlebars = require("handlebars");
const bwipjs = require("bwip-js");
require("dotenv").config();

// Google Sheets Auth
const auth = new google.auth.GoogleAuth({
  keyFile: "credentials.json",
  scopes: ["https://www.googleapis.com/auth/spreadsheets.readonly"],
});

// Email Transporter (Gmail SMTP)
const transporter = nodemailer.createTransport({
  host: "smtp.gmail.com",
  port: 587,
  secure: false,
  auth: {
    user: process.env.EMAIL_USER,
    pass: process.env.EMAIL_APP_PASSWORD,
  },
});

// Validate Email Config
async function verifyEmailConfig() {
  try {
    await transporter.verify();
    console.log("SMTP Verified Successfully");
    return true;
  } catch (err) {
    console.error("SMTP Verification Failed:", err.message);
    return false;
  }
}

// Load pass template
const template = handlebars.compile(
  fs.readFileSync(path.join(__dirname, "pass.html"), "utf8")
);

// Barcode Generator
async function generateBarcode(email) {
  const text = `ELAN_26_${email}_GUEST`;
  return new Promise((resolve, reject) => {
    bwipjs.toBuffer(
      {
        bcid: "code128",
        text,
        scale: 2,
        height: 14,
        includetext: true,
        textxalign: "center",
      },
      (err, png) => (err ? reject(err) : resolve(png))
    );
  });
}

// Email Sender
async function sendPass(participant) {
  try {
    const barcodeBuffer = await generateBarcode(participant.email);

    const htmlContent = template({
      Name: participant.name,
      College: participant.college,
      City: participant.city,
      barcode: "cid:barcodeImage",
    });

    await transporter.sendMail({
      from: process.env.EMAIL_USER,
      to: participant.email,
      subject: "Registration confirmed | Elan & nVision Fest Pass",
      html: htmlContent,
      attachments: [
        {
          filename: "barcode.png",
          content: barcodeBuffer,
          cid: "barcodeImage",
        },
      ],
    });

    console.log(`SENT: ${participant.email}`);
  } catch (err) {
    console.error(`ERROR sending ${participant.email}:`, err.message);
  }
}

// Polling + Sheet Logic
let lastProcessedRow = 0;

async function initializeLastProcessedRow() {
  const sheets = google.sheets({ version: "v4", auth });
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: process.env.FORM_SHEET_ID,
    range: "A2:H",
  });

  const rows = res.data.values || [];
  lastProcessedRow = rows.length;
  console.log(`Startup: Skipping ${lastProcessedRow} existing rows`);
}

async function checkNewRegistrations() {
  const sheets = google.sheets({ version: "v4", auth });
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: process.env.FORM_SHEET_ID,
    range: "A2:H",
  });

  const rows = res.data.values || [];
  if (!rows.length) return;

  for (let i = lastProcessedRow; i < rows.length; i++) {
    const row = rows[i];
    const participant = {
      name: row[1],
      phone: row[2],
      email: row[3],
      college: row[4],
      city: row[6],
    };

    if (participant.email) {
      console.log(`NEW: ${participant.name} - ${participant.email}`);
      await sendPass(participant);
      await new Promise((r) => setTimeout(r, 1500));
    }

    lastProcessedRow = i + 1;
  }
}

// Main Runner
async function main() {
  console.log("Starting ELAN Pass Automation...");

  if (!(await verifyEmailConfig())) return;

  await initializeLastProcessedRow();

  setInterval(checkNewRegistrations, 60000);
}

main().catch(console.error);
