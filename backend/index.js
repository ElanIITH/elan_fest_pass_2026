// const { google } = require("googleapis");
// const nodemailer = require("nodemailer");
// const fs = require("fs");
// const path = require("path");
// const handlebars = require("handlebars");
// const bwipjs = require("bwip-js");
// require("dotenv").config();

// let isProcessing = false;

// const auth = new google.auth.GoogleAuth({
//   keyFile: "credentials.json",
//   scopes: ["https://www.googleapis.com/auth/spreadsheets"],
// });

// const transporter = nodemailer.createTransport({
//   host: "smtp.gmail.com",
//   port: 587,
//   secure: false,
//   auth: {
//     user: process.env.EMAIL_2,
//     pass: process.env.EMAIL_2_PWD,
//   },
// });

// async function verifyEmailConfig() {
//   try {
//     await transporter.verify();
//     console.log("SMTP Verified Successfully");
//     return true;
//   } catch (err) {
//     console.error("SMTP Verification Failed:", err.message);
//     return false;
//   }
// }

// const template = handlebars.compile(
//   fs.readFileSync(path.join(__dirname, "pass.html"), "utf8")
// );

// async function generateBarcode(email) {
//   const text = `${email}`.toUpperCase();
//   return new Promise((resolve, reject) => {
//     bwipjs.toBuffer(
//       {
//         bcid: "code128",
//         text,
//         scale: 4,
//         height: 20,
//         includetext: true,
//         textxalign: "center",
//         backgroundcolor: "ffffff",
//       },
//       (err, png) => (err ? reject(err) : resolve(png))
//     );
//   });
// }

// async function markEmailAsSent(rowIndex) {
//   try {
//     const sheets = google.sheets({ version: "v4", auth });
//     const rowNumber = rowIndex + 2;
//     await sheets.spreadsheets.values.update({
//       spreadsheetId: process.env.FORM_SHEET_ID,
//       range: `J${rowNumber}`,
//       valueInputOption: "RAW",
//       resource: {
//         values: [[new Date().toISOString()]],
//       },
//     });
//   } catch (err) {
//     console.error(`ERROR marking row ${rowIndex} as sent:`, err.message);
//   }
// }

// async function addToSecSheet(participant) {
//   try {
//     const sheets = google.sheets({ version: "v4", auth });
//     await sheets.spreadsheets.values.append({
//       spreadsheetId: process.env.FORM_SHEET_ID_SEC,
//       range: "A2:C",
//       valueInputOption: "RAW",
//       resource: {
//         values: [[participant.name, participant.email, participant.phone]],
//       },
//     });
//     console.log(
//       `ADDED to security sheet: ${participant.name} | ${participant.email}`
//     );
//   } catch (err) {
//     console.error(`ERROR adding to security sheet:`, err.message);
//   }
// }

// async function sendPass(participant, rowIndex) {
//   try {
//     const barcodeBuffer = await generateBarcode(participant.email);

//     const formattedName =
//       participant.name && participant.name.trim().length > 0
//         ? participant.name
//             .trim()
//             .split(" ")[0]
//             .toLowerCase()
//             .replace(/^./, (c) => c.toUpperCase())
//         : "Guest";

//     const headerPath = path.join(__dirname, "pass_header.jpg");
//     const hasHeader = fs.existsSync(headerPath);

//     const htmlContent = template({
//       Name: formattedName,
//       Email: participant.email,
//       barcodeImage: "cid:barcodeImage",
//       headerImage: hasHeader ? "cid:headerImage" : null,
//     });

//     const attachments = [
//       {
//         filename: "barcode.png",
//         content: barcodeBuffer,
//         cid: "barcodeImage",
//       },
//       {
//         filename: "elan-nvision-2026-barcode.png",
//         content: barcodeBuffer,
//       },
//     ];

//     if (hasHeader) {
//       attachments.push({
//         filename: "pass_header.jpg",
//         path: headerPath,
//         cid: "headerImage",
//       });
//     }

//     await transporter.sendMail({
//       from: process.env.EMAIL_USER,
//       to: participant.email,
//       subject: `${formattedName}, your Elan & nVision 2026 Pass is ready!`,
//       html: htmlContent,
//       attachments,
//     });

//     console.log(`SENT: ${participant.email}`);
//     // await markEmailAsSent(rowIndex);
//   } catch (err) {
//     console.error(`ERROR sending ${participant.email}:`, err.message);
//   }
// }

// // async function checkNewRegistrations() {
// //   if (isProcessing) {
// //     console.log("SKIP: Previous cycle still running");
// //     return;
// //   }

// //   isProcessing = true;

// //   try {
// //     const sheets = google.sheets({ version: "v4", auth });

// //     const res = await sheets.spreadsheets.values.get({
// //       spreadsheetId: process.env.FORM_SHEET_ID,
// //       range: "A2:J",
// //     });

// //     const rows = res.data.values || [];
// //     if (!rows.length) return;

// //     for (let i = rows.length - 1; i >= 0; i--) {
// //       const row = rows[i];

// //       const participant = {
// //         name: row[1],
// //         email: row[8],
// //         phone: row[3],
// //         emailSent: row[9],
// //       };

// //       if (participant.email && !participant.emailSent) {
// //         console.log(`NEW: ${participant.name} | ${participant.email}`);
// //         await sendPass(participant, i);
// //         await addToSecSheet(participant);
// //         await new Promise((resolve) => setTimeout(resolve, 1500));
// //       } else if (participant.email && participant.emailSent) {
// //         console.log(`SKIP: ${participant.email} (already sent)`);
// //         break;
// //       }
// //     }
// //   } catch (err) {
// //     console.error("ERROR in checkNewRegistrations:", err.message);
// //   } finally {
// //     isProcessing = false;
// //   }
// // }

// async function checkNewRegistrations() {
//   if (isProcessing) return;
//   isProcessing = true;

//   const sheets = google.sheets({ version: "v4", auth });

//   try {
//     const res = await sheets.spreadsheets.values.get({
//       spreadsheetId: process.env.FORM_SHEET_ID,
//       range: "A2:J",
//     });

//     const rows = res.data.values || [];
//     if (!rows.length) return;

//     for (let i = 0; i < rows.length; i++) {
//       const row = rows[i];

//       if (!row[8] || row[9]) continue;

//       const participant = {
//         name: row[1],
//         email: row[8],
//         phone: row[3],
//         emailSent: row[9],
//       };

//       const sheetRow = i + 2;

//       try {
//         // lock
//         await sheets.spreadsheets.values.update({
//           spreadsheetId: process.env.FORM_SHEET_ID,
//           range: `J${sheetRow}`,
//           valueInputOption: "RAW",
//           requestBody: { values: [["PROCESSING"]] },
//         });

//         await sendPass(participant, i);
//         await addToSecSheet(participant);

//         // mark sent
//         await sheets.spreadsheets.values.update({
//           spreadsheetId: process.env.FORM_SHEET_ID,
//           range: `J${sheetRow}`,
//           valueInputOption: "RAW",
//           requestBody: { values: [[new Date().toISOString()]] },
//         });

//         console.log(`SENT: ${participant.email}`);
//       } catch (err) {
//         // unlock on failure
//         await sheets.spreadsheets.values.update({
//           spreadsheetId: process.env.FORM_SHEET_ID,
//           range: `J${sheetRow}`,
//           valueInputOption: "RAW",
//           requestBody: { values: [[""]] },
//         });

//         console.error(`FAILED: ${participant.email}`, err.message);
//       }
//     }
//   } catch (err) {
//     console.error("ERROR in checkNewRegistrations:", err.message);
//   } finally {
//     isProcessing = false;
//   }
// }

// async function main() {
//   console.log("Starting ELAN Pass Automation...");

//   if (!(await verifyEmailConfig())) return;

//   setInterval(checkNewRegistrations, 30000);
// }

// main().catch(console.error);

const { google } = require("googleapis");
const nodemailer = require("nodemailer");
const fs = require("fs");
const path = require("path");
const handlebars = require("handlebars");
const bwipjs = require("bwip-js");
require("dotenv").config();

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
    user: process.env.EMAIL_2,
    pass: process.env.EMAIL_2_PWD,
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
