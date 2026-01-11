const { google } = require("googleapis");
const nodemailer = require("nodemailer");
const fs = require("fs");
const path = require("path");
const handlebars = require("handlebars");
const bwipjs = require("bwip-js");
require("dotenv").config();

// Email account management
const ACCOUNT_STATE_FILE = path.join(__dirname, ".current_email_account.txt");
const MAX_ACCOUNTS = 4;
let currentEmailNum = null;
let activeTransporter = null;

const STATE_FILE = path.join(__dirname, ".last_processed_row.txt");
const BOT_ENABLED_FILE = path.join(__dirname, ".bot_enabled");
const BUFFER_SIZE = parseInt(process.env.BUFFER_SIZE || "10", 10);
let isProcessing = false;
let securitySheetBuffer = [];

/* ---------------- GOOGLE AUTH ---------------- */

const auth = new google.auth.GoogleAuth({
  keyFile: "credentials.json",
  scopes: ["https://www.googleapis.com/auth/spreadsheets"],
});

/* ---------------- SMTP ---------------- */

function getCurrentEmailAccount() {
  if (currentEmailNum !== null) return currentEmailNum;

  try {
    if (fs.existsSync(ACCOUNT_STATE_FILE)) {
      const content = fs.readFileSync(ACCOUNT_STATE_FILE, "utf8").trim();
      const num = parseInt(content, 10);
      if (num >= 1 && num <= MAX_ACCOUNTS) {
        currentEmailNum = num;
        return currentEmailNum;
      }
    }
  } catch (err) {
    console.error("ERROR reading account state:", err.message);
  }

  // Fallback to env or default
  currentEmailNum = parseInt(process.env.CURRENT_EMAIL_NUM || "1", 10);
  saveCurrentEmailAccount(currentEmailNum);
  return currentEmailNum;
}

function saveCurrentEmailAccount(accountNum) {
  try {
    fs.writeFileSync(ACCOUNT_STATE_FILE, accountNum.toString(), "utf8");
    currentEmailNum = accountNum;
  } catch (err) {
    console.error("ERROR saving account state:", err.message);
  }
}

function createTransporter(accountNum) {
  const user = process.env[`EMAIL_${accountNum}`];
  const pass = process.env[`EMAIL_${accountNum}_PWD`];

  if (!user || !pass) {
    throw new Error(`Email credentials not found for account ${accountNum}`);
  }

  return nodemailer.createTransport({
    host: "smtp.gmail.com",
    port: 587,
    secure: false,
    auth: { user, pass },
    pool: true,
    maxConnections: 1,
    maxMessages: 500,
  });
}

function getTransporter() {
  if (!activeTransporter) {
    const accountNum = getCurrentEmailAccount();
    activeTransporter = createTransporter(accountNum);
  }
  return activeTransporter;
}

function isQuotaError(err) {
  if (!err) return false;

  const message = err.message || "";
  const response = err.response || "";
  const code = err.code || "";
  const responseCode = err.responseCode || 0;

  // Gmail quota error signatures
  const quotaPatterns = [
    /daily.*sending.*limit.*exceeded/i,
    /daily user sending limit/i,
    /user has exceeded the allowed sending limits/i,
    /sending limits/i,
    /quota exceeded/i,
    /too many messages/i,
    /rate limit/i,
    /454[.\s-]4\.7\.0/,
    /550[.\s-]5\.4\.5/,
    /421[.\s-]4\.7\.0.*try again later/i,
  ];

  return (
    quotaPatterns.some(
      (pattern) => pattern.test(message) || pattern.test(response)
    ) || [454, 550, 421].includes(responseCode)
  );
}

async function switchToNextAccount(triedAccounts = new Set()) {
  const currentAccount = getCurrentEmailAccount();
  triedAccounts.add(currentAccount);

  // Try next account in sequence
  for (let i = 1; i <= MAX_ACCOUNTS; i++) {
    const nextAccount = (currentAccount % MAX_ACCOUNTS) + i;
    if (nextAccount > MAX_ACCOUNTS) continue;

    if (triedAccounts.has(nextAccount)) continue;

    // Check if credentials exist
    if (
      !process.env[`EMAIL_${nextAccount}`] ||
      !process.env[`EMAIL_${nextAccount}_PWD`]
    ) {
      triedAccounts.add(nextAccount);
      continue;
    }

    try {
      const testTransporter = createTransporter(nextAccount);
      await testTransporter.verify();

      // Success - switch to this account
      saveCurrentEmailAccount(nextAccount);
      activeTransporter = testTransporter;
      console.log(`✓ SWITCHED to EMAIL_${nextAccount}`);
      return true;
    } catch (err) {
      console.log(`✗ EMAIL_${nextAccount} verification failed:`, err.message);
      triedAccounts.add(nextAccount);
    }
  }

  return false; // All accounts exhausted
}

async function verifyEmailConfig() {
  try {
    const transporter = getTransporter();
    await transporter.verify();
    const accountNum = getCurrentEmailAccount();
    console.log(`SMTP Verified Successfully (EMAIL_${accountNum})`);
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

function isBotEnabled() {
  return fs.existsSync(BOT_ENABLED_FILE);
}

function enableBot() {
  try {
    fs.writeFileSync(BOT_ENABLED_FILE, "enabled", "utf8");
    console.log("✓ BOT ENABLED");
  } catch (err) {
    console.error("ERROR enabling bot:", err.message);
  }
}

function disableBot() {
  try {
    if (fs.existsSync(BOT_ENABLED_FILE)) {
      fs.unlinkSync(BOT_ENABLED_FILE);
    }
    console.log("✗ BOT DISABLED");
  } catch (err) {
    console.error("ERROR disabling bot:", err.message);
  }
}

/* ---------------- SECURITY SHEET ---------------- */

function addToBuffer(participant) {
  securitySheetBuffer.push([
    participant.name,
    participant.email,
    participant.phone,
  ]);
  console.log(
    `BUFFERED (${securitySheetBuffer.length}/${BUFFER_SIZE}): ${participant.name} | ${participant.email}`
  );
}

async function flushSecurityBuffer(sheets) {
  if (securitySheetBuffer.length === 0) return;

  try {
    await sheets.spreadsheets.values.append({
      spreadsheetId: process.env.FORM_SHEET_ID_SEC,
      range: "A2:C",
      valueInputOption: "RAW",
      resource: {
        values: securitySheetBuffer,
      },
    });

    console.log(
      `✓ FLUSHED ${securitySheetBuffer.length} rows to security sheet`
    );
    securitySheetBuffer = [];
  } catch (err) {
    console.error("ERROR flushing to security sheet:", err.message);
  }
}

/* ---------------- EMAIL SEND ---------------- */

async function sendPass(participant, triedAccounts = new Set()) {
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

  const currentAccount = getCurrentEmailAccount();
  const mailOptions = {
    from: process.env[`EMAIL_${currentAccount}`],
    to: participant.email,
    subject: `${formattedName}, your Elan & nVision 2026 Pass is ready!`,
    html: htmlContent,
    attachments,
  };

  try {
    const transporter = getTransporter();
    await transporter.sendMail(mailOptions);
  } catch (err) {
    // Check if quota error
    if (isQuotaError(err)) {
      console.log(`⚠️  QUOTA HIT on EMAIL_${currentAccount}`);

      // Try to switch to next account
      const switched = await switchToNextAccount(triedAccounts);

      if (switched) {
        // Retry with new account
        console.log(`🔄 RETRYING send for ${participant.email}`);
        return await sendPass(participant, triedAccounts);
      } else {
        // All accounts exhausted
        const exhaustedErr = new Error(
          `ALL_ACCOUNTS_EXHAUSTED: All ${MAX_ACCOUNTS} email accounts have hit quota limits`
        );
        exhaustedErr.code = "ALL_ACCOUNTS_EXHAUSTED";
        throw exhaustedErr;
      }
    }

    // Non-quota error, rethrow
    throw err;
  }
}

/* ---------------- MAIN LOGIC ---------------- */

async function checkNewRegistrations() {
  if (!isBotEnabled()) {
    console.log("Bot is disabled. Skipping email processing.");
    return;
  }

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
        await new Promise((resolve) => setTimeout(resolve, 1000));
        addToBuffer(participant);

        // Flush buffer if it reaches the limit
        if (securitySheetBuffer.length >= BUFFER_SIZE) {
          await flushSecurityBuffer(sheets);
        }

        // MARK SENT
        await sheets.spreadsheets.values.update({
          spreadsheetId: process.env.FORM_SHEET_ID,
          range: `J${sheetRow}`,
          valueInputOption: "RAW",
          requestBody: { values: [[new Date().toISOString()]] },
        });

        console.log(`SENT: ${participant.email} (Row ${sheetRow})`);
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
      }
      saveLastProcessedRow(sheetRow);
    }

    // Flush any remaining buffered rows
    await flushSecurityBuffer(sheets);
  } catch (err) {
    console.error("ERROR in checkNewRegistrations:", err.message);
  } finally {
    isProcessing = false;
  }
}

/* ---------------- BOOTSTRAP ---------------- */

async function main() {
  // Handle CLI commands
  const args = process.argv.slice(2);

  if (args.includes("on") || args.includes("enable")) {
    enableBot();
    return;
  }

  if (args.includes("off") || args.includes("disable")) {
    disableBot();
    return;
  }

  if (args.includes("status")) {
    console.log(`Bot status: ${isBotEnabled() ? "ENABLED ✓" : "DISABLED ✗"}`);
    console.log(`Last processed row: ${getLastProcessedRow()}`);
    return;
  }

  console.log("Starting ELAN Pass Automation...");
  console.log(`Bot status: ${isBotEnabled() ? "ENABLED ✓" : "DISABLED ✗"}`);
  console.log(`Active email account: EMAIL_${getCurrentEmailAccount()}`);
  console.log(`Last processed row: ${getLastProcessedRow()}`);
  console.log(`Security sheet buffer size: ${BUFFER_SIZE}`);

  if (!(await verifyEmailConfig())) return;

  setInterval(checkNewRegistrations, 30000);
}

main().catch(console.error);
