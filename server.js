require('dotenv').config();

const express = require("express");
const fs = require("fs");
const path = require("path");
const nodemailer = require("nodemailer");
const bodyParser = require("body-parser");
const XLSX = require("xlsx");
const { format, getYear, subMonths, getDaysInMonth } = require("date-fns");
const { PDFDocument } = require("pdf-lib");

const app = express();
app.use(bodyParser.json());
app.use(express.static("public"));

const TEMPLATE_PATH = path.join(__dirname, "templates", "timesheet-template.pdf");

function getCurrentMonthUpper() {
  return format(new Date(), "MMMM").toUpperCase();
}

function getWeekHeaders(month) {
  const monthNames = ["JANUARY","FEBRUARY","MARCH","APRIL","MAY","JUNE","JULY","AUGUST","SEPTEMBER","OCTOBER","NOVEMBER","DECEMBER"];
  const monthIndex = monthNames.indexOf(month);
  if (monthIndex === -1) return { week1: "Week 1", week2: "Week 2", week3: "Week 3", week4: "Week 4", week5: "Week 5" };

  const prevDate = subMonths(new Date(getYear(new Date()), monthIndex, 1), 1);
  const prevMonthAbbr = format(prevDate, "MMM").toUpperCase();
  const daysInPrevMonth = getDaysInMonth(prevDate);
  const adjustment = `26 – ${daysInPrevMonth} ${prevMonthAbbr}`;

  return {
    week1: `Week1(adjustment week ${adjustment})`,
    week2: "Week 2",
    week3: "Week 3",
    week4: "Week 4",
    week5: "Week 5",
  };
}

async function loadDataFromExcel() {
  const filePath = path.join(__dirname, "data.xlsx");
  if (!fs.existsSync(filePath)) {
    throw new Error("data.xlsx not found in project root.");
  }

  const workbook = XLSX.readFile(filePath);
  const sheetName = workbook.SheetNames.includes("Learners") ? "Learners" : workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];

  const rows = XLSX.utils.sheet_to_json(sheet);

  if (rows.length === 0) {
    throw new Error("No data rows found in the sheet.");
  }

  return rows.map(row => ({
    learnerName: (row.name || "Unknown Learner").toString().trim(),
    idNumber: (row.idNumber || "").toString().trim(),
    contact: (row.contact || "").toString().trim(),
    email: (row.email || "").toString().trim(),
    employer: (row.employer || "").toString().trim(),
    physicalAddress: (row.physicalAddress || "").toString().trim(),
    suburb: (row.suburb || "").toString().trim(),
    cityTown: (row.cityTown || "").toString().trim(),
    postalCode: (row.postalCode || "").toString().trim(),
    localMunicipality: (row.localMunicipality || "").toString().trim(),
    districtMunicipality: (row.districtMunicipality || "").toString().trim(),
    metropolitanMunicipality: (row.metropolitanMunicipality || "").toString().trim(),
    province: (row.province || "").toString().trim(),
    supervisorName: (row.supervisorName || "").toString().trim(),
    supervisorContact: (row.supervisorContact || "").toString().trim(),
    supervisorEmail: (row.supervisorEmail || "").toString().trim(),
    activitiesCount: (row.activitiesCount || "").toString().trim(),
    tvetCollege: (row.tvetCollege || "").toString().trim(),
    month: (row.month || getCurrentMonthUpper()).toString().toUpperCase().trim(),
  }));
}

async function generatePdf(data, outputPath) {
  const templateBytes = fs.readFileSync(TEMPLATE_PATH);
  const pdfDoc = await PDFDocument.load(templateBytes);
  const form = pdfDoc.getForm();

  const safeSet = (fieldName, value) => {
    try {
      if (value) {
        form.getTextField(fieldName).setText(value.toString());
        console.log(`Filled field: ${fieldName} with ${value}`);
      }
    } catch (e) {
      console.log(`Field not found or error filling: ${fieldName}`);
    }
  };

  safeSet("learnerName", data.learnerName);
  safeSet("idNumber", data.idNumber);
  safeSet("contact", data.contact);
  safeSet("email", data.email);
  safeSet("employer", data.employer);
  safeSet("physicalAddress", data.physicalAddress);
  safeSet("suburb", data.suburb);
  safeSet("cityTown", data.cityTown);
  safeSet("postalCode", data.postalCode);
  safeSet("localMunicipality", data.localMunicipality);
  safeSet("districtMunicipality", data.districtMunicipality);
  safeSet("metropolitanMunicipality", data.metropolitanMunicipality);
  safeSet("province", data.province);
  safeSet("supervisorName", data.supervisorName);
  safeSet("supervisorContact", data.supervisorContact);
  safeSet("supervisorEmail", data.supervisorEmail);
  safeSet("activitiesCount", data.activitiesCount);
  safeSet("tvetCollege", data.tvetCollege);
  safeSet("month", data.month);

  const weeks = getWeekHeaders(data.month);
  safeSet("week1", weeks.week1);
  safeSet("week2", weeks.week2);
  safeSet("week3", weeks.week3);
  safeSet("week4", weeks.week4);
  safeSet("week5", weeks.week5);

  const pdfBytes = await pdfDoc.save();
  fs.writeFileSync(outputPath, pdfBytes);
}

async function sendEmail(recipientEmail, attachmentPath, month) {
  const transporter = nodemailer.createTransport({
    host: process.env.EMAIL_HOST || "smtp.gmail.com",
    port: Number(process.env.EMAIL_PORT) || 587,
    secure: false,
    auth: {
      user: process.env.EMAIL_USER,
      pass: process.env.EMAIL_PASS,
    },
  });

  await transporter.sendMail({
    from: process.env.EMAIL_FROM || `"Timesheet Bot" <${process.env.EMAIL_USER}>`,
    to: recipientEmail,
    subject: `Your ${month} Fillable Timesheet`,
    text: `Hello,\n\nPlease find your pre-filled ${month} timesheet attached. You can fill in the remaining fields digitally.\n\nThank you!`,
    html: `<p>Hello,</p><p>Please find your pre-filled <strong>${month}</strong> timesheet attached.</p><p>You can fill in the remaining fields digitally in any PDF reader.</p><p>Thank you!</p>`,
    attachments: [{ filename: path.basename(attachmentPath), path: attachmentPath }]
  });

  console.log(`Email sent to: ${recipientEmail}`);
}

app.post("/generate", async (req, res) => {
  try {
    const learners = await loadDataFromExcel();
    const results = [];

    for (const learner of learners) {
      if (!learner.email) {
        results.push(`⚠️ ${learner.learnerName} skipped — no email`);
        continue;
      }

      const safeName = learner.learnerName.replace(/[^a-zA-Z0-9]/g, "_");
      const filename = `timesheet-${safeName}-${learner.month}.pdf`;
      const outputPath = path.join(__dirname, filename);

      await generatePdf(learner, outputPath);
      await sendEmail(learner.email, outputPath, learner.month);

      results.push(`${learner.learnerName} (${learner.month}) → generated & emailed to ${learner.email}`);
    }

    res.send(`✅ All done!<br><br>${results.join("<br>")}`);
  } catch (err) {
    console.error(err);
    res.status(500).send(`❌ Error: ${err.message}`);
  }
});

app.listen(5000, () => console.log("Server running at http://localhost:5000"));