const { google } = require("googleapis");
const nodemailer = require("nodemailer");
const fs = require("fs");
const path = require("path");
const handlebars = require("handlebars");
const bwipjs = require("bwip-js");
require("dotenv").config();

const { Pool } = require("pg");

const pool = new Pool({
  connectionString: process.env.DATABASE_URL,
});

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
    user: process.env.EMAIL_USER,
    pass: process.env.EMAIL_APP_PASSWORD,
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
    from: process.env.EMAIL_2,
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
    const res = await sheets.spreadsheets.values.get({
      spreadsheetId: process.env.FORM_SHEET_ID,
      range: "A2:J",
    });

    const rows = res.data.values || [];
    if (!rows.length) return;

    for (let i = 0; i < rows.length; i++) {
      const row = rows[i];

      if (!row[8] || row[9]) continue; // no email OR already sent

      const participant = {
        name: row[1],
        email: row[8],
        phone: row[3],
      };

      const sheetRow = i + 2;

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

        // try {
        //   await pool.query("SELECT insert_user_for_elan($1, $2, $3)", [
        //     participant.name,
        //     participant.email,
        //     participant.phone,
        //   ]);

        //   console.log(`DB INSERT OK: ${participant.email}`);
        // } catch (err) {
        //   console.error(
        //     `DB INSERT FAILED: ${participant.email} | ${err.message}`
        //   );
        // }

        // MARK SENT
        await sheets.spreadsheets.values.update({
          spreadsheetId: process.env.FORM_SHEET_ID,
          range: `J${sheetRow}`,
          valueInputOption: "RAW",
          requestBody: { values: [[new Date().toISOString()]] },
        });

        console.log(`SENT: ${participant.email}`);
      } catch (err) {
        // UNLOCK ON FAILURE
        await sheets.spreadsheets.values.update({
          spreadsheetId: process.env.FORM_SHEET_ID,
          range: `J${sheetRow}`,
          valueInputOption: "RAW",
          requestBody: { values: [[""]] },
        });

        console.error(`FAILED: ${participant.email}`, err.message);
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

  if (!(await verifyEmailConfig())) return;

  setInterval(checkNewRegistrations, 30000);
}

main().catch(console.error);
