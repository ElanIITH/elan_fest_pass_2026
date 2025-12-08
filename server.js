const { google } = require("googleapis");
const nodemailer = require("nodemailer");
const fs = require("fs");
const path = require("path");
const handlebars = require("handlebars");
const bwipjs = require("bwip-js");
const axios = require("axios");
require("dotenv").config();

// Configure Google Sheets API
const auth = new google.auth.GoogleAuth({
  keyFile: "elan-pass-mailing-450613-dbf3999bb350.json", // Your Google Cloud service account key
  scopes: [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/spreadsheets",
  ],
});

// Add this near the top of your file, after the auth configuration
const serviceAccountCredentials = require("./elan-pass-mailing-450613-dbf3999bb350.json");
console.log("Service Account Email:", serviceAccountCredentials.client_email);

// Configure email transporter
const transporter = nodemailer.createTransport({
  host: "smtp.gmail.com",
  port: 587,
  secure: false, // Use TLS
  auth: {
    user: process.env.EMAIL_USER,
    pass: process.env.EMAIL_APP_PASSWORD,
  },
  tls: {
    rejectUnauthorized: false,
  },
});

// Add a verification step before processing
async function verifyEmailConfig() {
  try {
    await transporter.verify();
    console.log("Email configuration verified successfully");
    return true;
  } catch (error) {
    console.error("Email verification failed:", error);
    return false;
  }
}

// Read HTML template
const templatePath = path.join(__dirname, "..", "pass.html");
const template = handlebars.compile(fs.readFileSync(templatePath, "utf8"));

async function generateBarcode(email) {
  try {
    const barcodeText = `ELAN_25_${email}_GUEST`;

    // Generate barcode as PNG buffer
    const png = await new Promise((resolve, reject) => {
      bwipjs.toBuffer(
        {
          bcid: "code128", // Barcode type
          text: barcodeText, // Text to encode
          scale: 2, // 3x scaling factor
          height: 14, // Bar height, in millimeters
          includetext: true, // Show human-readable text
          textxalign: "center", // Center the text
        },
        function (err, png) {
          if (err) {
            reject(err);
          } else {
            resolve(png);
          }
        }
      );
    });

    // Convert to base64 for embedding in HTML
    return `data:image/png;base64,${png.toString("base64")}`;
  } catch (err) {
    console.error("Error generating barcode:", err);
    throw err;
  }
}

async function updateAccessSheet(participant) {
  try {
    const sheets = google.sheets({ version: "v4", auth });

    // Format the data according to the specified structure
    const barcodeText = `ELAN_25_${participant.email}_GUEST`.toUpperCase();
    const rowData = [
      [
        participant.name,
        "Elan 2025;https://www.elan.org.in/assets/white%20horizontal-CmMzaxkP.svg",
        barcodeText,
        "GUEST",
        "10-03-2025",
        participant.email,
      ],
    ];

    // First verify if we have access to the sheet
    try {
      await sheets.spreadsheets.get({
        spreadsheetId: "1oaAASGUlpfVnAAxcybypnDQRg3m7ZtzUIWyHct38B1g",
      });
    } catch (error) {
      if (error.status === 403) {
        throw new Error(
          `Service account does not have access to the spreadsheet. Please share the spreadsheet with the service account email address found in your credentials file.`
        );
      }
      throw error;
    }

    await sheets.spreadsheets.values.append({
      spreadsheetId: "1oaAASGUlpfVnAAxcybypnDQRg3m7ZtzUIWyHct38B1g",
      range: "Sheet1!A:F",
      valueInputOption: "RAW",
      insertDataOption: "INSERT_ROWS",
      resource: {
        values: rowData,
      },
    });

    console.log(`Access sheet updated for ${participant.email}`);
    return true;
  } catch (error) {
    console.error("Error updating access sheet:", error.message);
    return false;
  }
}

async function sendPass(participant) {
  try {
    // Generate barcode
    const barcodeBuffer = await new Promise((resolve, reject) => {
      bwipjs.toBuffer(
        {
          bcid: "code128",
          text: `ELAN_25_${participant.email}_GUEST`,
          scale: 2, // Reduced from 3 to 2
          height: 14, // Reduced from 10 to 8
          includetext: true,
          textxalign: "center",
          backgroundcolor: "FFFFFF",
          padding: 10,
        },
        function (err, png) {
          if (err) reject(err);
          else resolve(png);
        }
      );
    });

    // Read header image
    let headerBuffer;
    try {
      const headerImagePath = path.join(__dirname, "..", "header.png");
      headerBuffer = fs.readFileSync(headerImagePath);
    } catch (error) {
      console.warn("Header image not found:", error.message);
      headerBuffer = null;
    }

    // Generate HTML content with CID references
    const htmlContent = template({
      Name: participant.name,
      Pass: participant.passType,
      ALT: `ELAN_25_${participant.email}_GUEST`,
      barcode: "cid:barcodeImage", // Reference to content ID
      College: participant.college,
      City: participant.city,
      headerImage: headerBuffer ? "cid:headerImage" : null, // Reference to content ID
      useHeaderFallback: !headerBuffer,
    });

    // Configure email with attachments using content IDs
    const mailOptions = {
      from: process.env.EMAIL_USER,
      to: participant.email,
      subject: "Booking confirmed | Elan & nVision Fest Pass",
      html: htmlContent,
      attachments: [
        {
          filename: "barcode.png",
          content: barcodeBuffer,
          cid: "barcodeImage", // Content ID for barcode
        },
      ],
    };

    // Add header image attachment if available
    if (headerBuffer) {
      mailOptions.attachments.push({
        filename: "header.png",
        content: headerBuffer,
        cid: "headerImage", // Content ID for header
      });
    }

    await transporter.sendMail(mailOptions);
    console.log(`Pass sent successfully to ${participant.email}`);

    // Add the entry to the access sheet after successful email sending
    await updateAccessSheet(participant);

    return true;
  } catch (error) {
    console.error(`Error sending pass to ${participant.email}:`, error);
    return false;
  }
}

async function processRegistrations() {
  try {
    const sheets = google.sheets({ version: "v4", auth });

    // Updated range to match your sheet structure
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: "1UFSo5HVgpuVFi0gbseCmHNb1sFfdhuJR_EDIewXrPXc",
      range: "A2:H", // Simplified range format
    });

    const rows = response.data.values;
    if (!rows || rows.length === 0) {
      console.log("No data found.");
      return;
    }

    // Process each registration with updated column indices
    for (const row of rows) {
      const participant = {
        timestamp: row[0],
        name: row[1], // Name is in column 2
        phone: row[2], // Phone is in column 3
        email: row[3], // Email is in column 4
        college: row[4], // College name
        age: row[5], // Age
        city: row[6], // City
        source: row[7], // How they heard about the event
        passType: "General", // Default pass type if not specified
      };

      if (participant.email) {
        console.log(
          `Processing registration for ${participant.name} (${participant.email})`
        );
        await sendPass(participant);
        // Add delay to avoid hitting rate limits
        await new Promise((resolve) => setTimeout(resolve, 1000));
      }
    }
  } catch (error) {
    console.error("Error processing registrations:", error);
  }
}

// Replace the lastProcessedRow variable declaration with this:
let lastProcessedRow = 0;

// Add this new function to initialize the last processed row
async function initializeLastProcessedRow() {
  try {
    const sheets = google.sheets({ version: "v4", auth });
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: "1UFSo5HVgpuVFi0gbseCmHNb1sFfdhuJR_EDIewXrPXc",
      range: "A2:H", // Simplified range format
    });

    const rows = response.data.values;
    if (!rows || rows.length === 0) {
      lastProcessedRow = 0;
    } else {
      lastProcessedRow = rows.length;
      console.log(
        `Initialized: Will start processing from row ${
          lastProcessedRow + 2
        } (skipping ${lastProcessedRow} existing entries)`
      );
    }
  } catch (error) {
    console.error("Error initializing last processed row:", error);
    throw error;
  }
}

async function checkNewRegistrations() {
  try {
    const sheets = google.sheets({ version: "v4", auth });

    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: "1UFSo5HVgpuVFi0gbseCmHNb1sFfdhuJR_EDIewXrPXc",
      range: "A2:H", // Simplified range format
    });

    const rows = response.data.values;
    if (!rows || rows.length === 0) {
      console.log("No data found.");
      return;
    }

    // Process only new entries
    for (let i = lastProcessedRow; i < rows.length; i++) {
      const row = rows[i];
      const participant = {
        timestamp: row[0],
        name: row[1],
        phone: row[2],
        email: row[3],
        college: row[4],
        age: row[5],
        city: row[6],
        source: row[7],
        passType: "General",
      };

      if (participant.email) {
        console.log(
          `Processing new registration for ${participant.name} (${participant.email})`
        );
        await sendPass(participant);
        lastProcessedRow = i + 1;
        // Add delay to avoid hitting rate limits
        await new Promise((resolve) => setTimeout(resolve, 1000));
      }
    }
  } catch (error) {
    console.error("Error checking new registrations:", error);
  }
}

// Modify the main function
async function main() {
  console.log("Starting automatic email service...");

  // Verify email configuration first
  const emailConfigValid = await verifyEmailConfig();
  if (!emailConfigValid) {
    console.error(
      "Email configuration is invalid. Please check your credentials."
    );
    return;
  }

  // Initialize the last processed row to current sheet length
  await initializeLastProcessedRow();

  // Set up periodic checking (every 1 minute)
  const CHECK_INTERVAL = 1 * 60 * 1000;
  setInterval(async () => {
    console.log("Checking for new registrations...");
    await checkNewRegistrations();
  }, CHECK_INTERVAL);
}

// Keep the existing main() call at the bottom
// Main execution
main().catch(console.error);
