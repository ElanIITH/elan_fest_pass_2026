const { google } = require("googleapis");
const nodemailer = require("nodemailer");
const fs = require("fs");
const path = require("path");
const handlebars = require("handlebars");
const bwipjs = require("bwip-js");
require("dotenv").config();

const auth = new google.auth.GoogleAuth({
  keyFile: "credentials.json",
  scopes: ["https://www.googleapis.com/auth/spreadsheets.readonly"],
});

const transporter = nodemailer.createTransport({
  host: "smtp.gmail.com",
  port: 587,
  secure: false,
  auth: {
    user: process.env.EMAIL_USER,
    pass: process.env.EMAIL_APP_PASSWORD,
  },
});

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

const template = handlebars.compile(
  fs.readFileSync(path.join(__dirname, "pass.html"), "utf8")
);

async function generateBarcode(email) {
  const text = `${email}`.toUpperCase();
  return new Promise((resolve, reject) => {
    bwipjs.toBuffer(
      {
        bcid: "code128",
        text,
        scale: 4,
        height: 20,
        includetext: true,
        textxalign: "center",
        backgroundcolor: "ffffff",
      },
      (err, png) => (err ? reject(err) : resolve(png))
    );
  });
}

async function sendPass(participant) {
  try {
    const barcodeBuffer = await generateBarcode(participant.email);

    const formattedName =
      participant.name && participant.name.trim().length > 0
        ? participant.name
            .trim()
            .split(" ")[0]
            .toLowerCase()
            .replace(/^./, (c) => c.toUpperCase())
        : "Guest";

    const headerPath = path.join(__dirname, "pass_header_3.png");
    const hasHeader = fs.existsSync(headerPath);

    const htmlContent = template({
      Name: formattedName,
      Email: participant.email,
      barcodeImage: "cid:barcodeImage",
      headerImage: hasHeader ? "cid:headerImage" : null,
    });

    const attachments = [
      {
        filename: "barcode.png",
        content: barcodeBuffer,
        cid: "barcodeImage",
      },
      {
        filename: "elan-nvision-2026-barcode.png",
        content: barcodeBuffer,
      },
    ];

    if (hasHeader) {
      attachments.push({
        filename: "pass_header_3.png",
        path: headerPath,
        cid: "headerImage",
      });
    }

    await transporter.sendMail({
      from: process.env.EMAIL_USER,
      to: participant.email,
      subject: `${formattedName}, your Elan & nVision 2026 Pass is ready!`,
      html: htmlContent,
      attachments,
    });

    console.log(`SENT: ${participant.email}`);
  } catch (err) {
    console.error(`ERROR sending ${participant.email}:`, err.message);
  }
}

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
      timestamp: row[0],
      name: row[1],
      email: row[2],
      phone: row[3],
      college: row[4],
      age: row[5],
      city: row[6],
    };

    if (participant.email) {
      console.log(`NEW: ${participant.name} | ${participant.email}`);
      await sendPass(participant);
      await new Promise((resolve) => setTimeout(resolve, 1500));
    }

    lastProcessedRow = i + 1;
  }
}

async function main() {
  console.log("Starting ELAN Pass Automation...");

  if (!(await verifyEmailConfig())) return;

  await initializeLastProcessedRow();

  setInterval(checkNewRegistrations, 60000);
}

main().catch(console.error);
