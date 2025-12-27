const { google } = require("googleapis");
const nodemailer = require("nodemailer");
const fs = require("fs");
const path = require("path");
const handlebars = require("handlebars");
const bwipjs = require("bwip-js");
require("dotenv").config();

const auth = new google.auth.GoogleAuth({
  keyFile: "credentials.json",
  scopes: ["https://www.googleapis.com/auth/spreadsheets"],
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

async function markEmailAsSent(rowIndex) {
  try {
    const sheets = google.sheets({ version: "v4", auth });
    const rowNumber = rowIndex + 2;
    await sheets.spreadsheets.values.update({
      spreadsheetId: process.env.FORM_SHEET_ID,
      range: `T${rowNumber}`,
      valueInputOption: "RAW",
      resource: {
        values: [[new Date().toISOString()]],
      },
    });
  } catch (err) {
    console.error(`ERROR marking row ${rowIndex} as sent:`, err.message);
  }
}

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
    console.error(`ERROR adding to security sheet:`, err.message);
  }
}

async function sendPass(participant, rowIndex) {
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
      from: process.env.EMAIL_USER,
      to: participant.email,
      subject: `${formattedName}, your Elan & nVision 2026 Pass is ready!`,
      html: htmlContent,
      attachments,
    });

    console.log(`SENT: ${participant.email}`);
    await markEmailAsSent(rowIndex);
  } catch (err) {
    console.error(`ERROR sending ${participant.email}:`, err.message);
  }
}

let lastProcessedRow = 0;

async function initializeLastProcessedRow() {
  const sheets = google.sheets({ version: "v4", auth });
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: process.env.FORM_SHEET_ID,
    range: "A2:I",
  });

  const rows = res.data.values || [];
  lastProcessedRow = 0;
  console.log(`Startup: Checking all ${rows.length} rows for unsent emails`);
}

async function checkNewRegistrations() {
  const sheets = google.sheets({ version: "v4", auth });

  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: process.env.FORM_SHEET_ID,
    range: "A2:T",
  });

  const rows = res.data.values || [];
  if (!rows.length) return;

  for (let i = lastProcessedRow; i < rows.length; i++) {
    const row = rows[i];

    const participant = {
      name: row[12],
      email: row[13],
      phone: row[14],
      emailSent: row[19],
    };

    if (participant.email && !participant.emailSent) {
      console.log(`NEW: ${participant.name} | ${participant.email}`);
      await sendPass(participant, i);
      await addToSecSheet(participant);
      await new Promise((resolve) => setTimeout(resolve, 1500));
    } else if (participant.email && participant.emailSent) {
      console.log(`SKIP: ${participant.email} (already sent)`);
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

// health check
const http = require("http");

http
  .createServer((req, res) => {
    if (req.url === "/health") {
      res.writeHead(200, { "Content-Type": "text/plain" });
      res.end("OK");
    } else {
      res.writeHead(404);
      res.end();
    }
  })
  .listen(process.env.PORT || 3000, () => {
    console.log("Health check server running");
  });
