const { google } = require("googleapis");
const nodemailer = require("nodemailer");
const fs = require("fs");
const path = require("path");
const handlebars = require("handlebars");
const bwipjs = require("bwip-js");
require("dotenv").config();

// Change this number (1-4) to switch between email accounts
const CURRENT_EMAIL_NUM = process.env.CURRENT_EMAIL_NUM || "1";
const STATE_FILE = path.join(__dirname, ".last_processed_row.txt");
let isProcessing = false;

/* ---------------- GOOGLE AUTH ---------------- */

const auth = new google.auth.GoogleAuth({
  keyFile: "credentials.json",
  scopes: ["https://www.googleapis.com/auth/spreadsheets"],
});

/* ---------------- SMTP ---------------- */

const transporter = nodemailer.createTransport({
  host: "smtp.gmail.com",
  port: 587,
  secure: false,
  auth: {
    user: process.env[`EMAIL_${CURRENT_EMAIL_NUM}`],
    pass: process.env[`EMAIL_${CURRENT_EMAIL_NUM}_PWD`],
  },
  pool: true,
  maxConnections: 1,
  maxMessages: 500,
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

/* ---------------- TEMPLATE ---------------- */

const template = handlebars.compile(
  fs.readFileSync(path.join(__dirname, "pass.html"), "utf8")
);

/* ---------------- BARCODE ---------------- */

async function generateBarcode(email) {
  return new Promise((resolve, reject) => {
    bwipjs.toBuffer(
      {
        bcid: "code128",
        text: email.toUpperCase(),
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

/* ---------------- STATE PERSISTENCE ---------------- */

function getLastProcessedRow() {
  try {
    if (fs.existsSync(STATE_FILE)) {
      const content = fs.readFileSync(STATE_FILE, "utf8").trim();
      const row = parseInt(content, 10);
      return isNaN(row) ? 1 : row;
    }
  } catch (err) {
    console.error("ERROR reading state file:", err.message);
  }
  return 1; // Start from row 2 (first data row)
}

function saveLastProcessedRow(row) {
  try {
    fs.writeFileSync(STATE_FILE, row.toString(), "utf8");
  } catch (err) {
    console.error("ERROR writing state file:", err.message);
  }
}

/* ---------------- SECURITY SHEET ---------------- */

async function addToSecSheet(participant) {
  try {
    const sheets = google.sheets({ version: "v4", auth });

    await sheets.spreadsheets.values.append({
      spreadsheetId: process.env.FORM_SHEET_ID_SEC,
      range: "A2:C",
      valueInputOption: "RAW",
      resource: {
        values: [[participant.name, participant.email, participant.phone]],
      },
    });

    console.log(
      `ADDED to security sheet: ${participant.name} | ${participant.email}`
    );
  } catch (err) {
    console.error("ERROR adding to security sheet:", err.message);
  }
}

/* ---------------- EMAIL SEND ---------------- */

async function sendPass(participant) {
  const barcodeBuffer = await generateBarcode(participant.email);

  const formattedName =
    participant.name && participant.name.trim()
      ? participant.name
          .trim()
          .split(" ")[0]
          .replace(/^./, (c) => c.toUpperCase())
      : "Guest";

  const headerPath = path.join(__dirname, "pass_header.jpg");
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
      filename: "pass_header.jpg",
      path: headerPath,
      cid: "headerImage",
    });
  }

  await transporter.sendMail({
    from: process.env[`EMAIL_${CURRENT_EMAIL_NUM}`],
    to: participant.email,
    subject: `${formattedName}, your Elan & nVision 2026 Pass is ready!`,
    html: htmlContent,
    attachments,
  });
}

/* ---------------- MAIN LOGIC ---------------- */

async function checkNewRegistrations() {
  if (isProcessing) return;
  isProcessing = true;

  const sheets = google.sheets({ version: "v4", auth });

  try {
    const lastProcessedRow = getLastProcessedRow();
    const startRow = lastProcessedRow + 1;

    // Fetch only rows after the last processed row
    // Using a reasonable batch size (e.g., 100 rows at a time)
    const res = await sheets.spreadsheets.values.get({
      spreadsheetId: process.env.FORM_SHEET_ID,
      range: `A${startRow}:J${startRow + 99}`,
    });

    const rows = res.data.values || [];
    if (!rows.length) {
      isProcessing = false;
      return;
    }

    // Process only the first unprocessed row with a valid email
    for (let i = 0; i < rows.length; i++) {
      const row = rows[i];
      const sheetRow = startRow + i;

      // Skip if no email present
      if (!row[8]) {
        // Move cursor forward even for empty rows to avoid getting stuck
        saveLastProcessedRow(sheetRow);
        continue;
      }

      // Skip if already sent (column J has a value)
      if (row[9]) {
        saveLastProcessedRow(sheetRow);
        continue;
      }

      const participant = {
        name: row[1],
        email: row[8],
        phone: row[3],
      };

      try {
        // LOCK
        await sheets.spreadsheets.values.update({
          spreadsheetId: process.env.FORM_SHEET_ID,
          range: `J${sheetRow}`,
          valueInputOption: "RAW",
          requestBody: { values: [["PROCESSING"]] },
        });

        await sendPass(participant);
        await addToSecSheet(participant);

        // MARK SENT
        await sheets.spreadsheets.values.update({
          spreadsheetId: process.env.FORM_SHEET_ID,
          range: `J${sheetRow}`,
          valueInputOption: "RAW",
          requestBody: { values: [[new Date().toISOString()]] },
        });

        console.log(`SENT: ${participant.email} (Row ${sheetRow})`);

        // Update state after successful processing
        saveLastProcessedRow(sheetRow);

        // Process only one email per iteration
        break;
      } catch (err) {
        // UNLOCK ON FAILURE (don't update state so we retry this row)
        await sheets.spreadsheets.values.update({
          spreadsheetId: process.env.FORM_SHEET_ID,
          range: `J${sheetRow}`,
          valueInputOption: "RAW",
          requestBody: { values: [[""]] },
        });

        console.error(
          `FAILED: ${participant.email} (Row ${sheetRow})`,
          err.message
        );
        // Don't update state on failure - retry this row next time
        break;
      }
    }
  } catch (err) {
    console.error("ERROR in checkNewRegistrations:", err.message);
  } finally {
    isProcessing = false;
  }
}

/* ---------------- BOOTSTRAP ---------------- */

async function main() {
  console.log("Starting ELAN Pass Automation...");
  console.log(`Last processed row: ${getLastProcessedRow()}`);

  if (!(await verifyEmailConfig())) return;

  setInterval(checkNewRegistrations, 30000);
}

main().catch(console.error);
