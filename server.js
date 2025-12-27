require('dotenv').config();

const express = require("express");
const fs = require("fs");
const path = require("path");
const os = require("os");
const multer = require("multer");
const nodemailer = require("nodemailer");
const XLSX = require("xlsx");
const { format, getYear, subMonths, getDaysInMonth } = require("date-fns");
const { PDFDocument } = require("pdf-lib");

const app = express();
app.use(express.static("public"));

// Multer config: store uploaded file temporarily
const upload = multer({ dest: os.tmpdir() });

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

async function loadDataFromExcel(filePath) {
  if (!fs.existsSync(filePath)) {
    throw new Error("Excel file not found at specified path.");
  }

  const workbook = XLSX.readFile(filePath);
  const sheetName = workbook.SheetNames.includes("Learners") ? "Learners" : workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];

  const rows = XLSX.utils.sheet_to_json(sheet);

  if (rows.length === 0) {
    throw new Error("No data rows found in the Excel sheet.");
  }

  return rows.map(row => ({
    learnerName: (row.name || row["Learner Name"] || "Unknown Learner").toString().trim(),
    idNumber: (row.idNumber || row["ID Number"] || "").toString().trim(),
    contact: (row.contact || row["Contact"] || "").toString().trim(),
    email: (row.email || row["Email"] || "").toString().trim(),
    employer: (row.employer || row["Employer"] || "").toString().trim(),
    physicalAddress: (row.physicalAddress || row["Physical Address"] || "").toString().trim(),
    suburb: (row.suburb || row["Suburb"] || "").toString().trim(),
    cityTown: (row.cityTown || row["City/Town"] || "").toString().trim(),
    postalCode: (row.postalCode || row["Postal Code"] || "").toString().trim(),
    localMunicipality: (row.localMunicipality || row["Local Municipality"] || "").toString().trim(),
    districtMunicipality: (row.districtMunicipality || row["District Municipality"] || "").toString().trim(),
    metropolitanMunicipality: (row.metropolitanMunicipality || row["Metropolitan Municipality"] || "").toString().trim(),
    province: (row.province || row["Province"] || "").toString().trim(),
    supervisorName: (row.supervisorName || row["Supervisor Name"] || "").toString().trim(),
    supervisorContact: (row.supervisorContact || row["Supervisor Contact"] || "").toString().trim(),
    supervisorEmail: (row.supervisorEmail || row["Supervisor Email"] || "").toString().trim(),
    activitiesCount: (row.activitiesCount || "").toString().trim(),
    tvetCollege: (row.tvetCollege || row["TVET College"] || "").toString().trim(),
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
        console.log(`Filled field: ${fieldName} with "${value}"`);
      }
    } catch (e) {
      console.log(`Field not found: ${fieldName}`);
    }
  };

  // Fill all fields
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

  // Week headers
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
    from: process.env.EMAIL_FROM || `"ICM ShiftTrack" <${process.env.EMAIL_USER}>`,
    to: recipientEmail,
    subject: `Your ${month} Timesheet - ICM WIL`,
    text: `Hello,\n\nPlease find your pre-filled ${month} timesheet attached.\nFill in your hours and signatures digitally, then submit.\n\nThank you!\nICM Team`,
    html: `<p>Hello,</p>
           <p>Please find your pre-filled <strong>${month}</strong> timesheet attached.</p>
           <p>You can fill in hours and signatures digitally in any PDF reader.</p>
           <p>Thank you!<br><strong>ICM Team</strong></p>`,
    attachments: [{ filename: path.basename(attachmentPath), path: attachmentPath }]
  });

  console.log(`Email sent to: ${recipientEmail}`);
}

// Main route: supports both uploaded file and default data.xlsx
app.post("/generate", upload.single("dataFile"), async (req, res) => {
  let excelFilePath = null;
  let tempFileToDelete = null;

  try {
    if (req.file) {
      // User uploaded a file
      excelFilePath = req.file.path;
      tempFileToDelete = req.file.path;
      console.log(`Using uploaded file: ${req.file.originalname}`);
    } else if (fs.existsSync(path.join(__dirname, "data.xlsx"))) {
      // Fallback to default file
      excelFilePath = path.join(__dirname, "data.xlsx");
      console.log("Using default data.xlsx");
    } else {
      return res.status(400).send("❌ No file uploaded and no default 'data.xlsx' found in project root.");
    }

    const learners = await loadDataFromExcel(excelFilePath);
    const results = [];

    for (const learner of learners) {
      if (!learner.email) {
        results.push(`⚠️ ${learner.learnerName} skipped — no email provided`);
        continue;
      }

      const safeName = learner.learnerName.replace(/[^a-zA-Z0-9]/g, "_");
      const filename = `timesheet-${safeName}-${learner.month}.pdf`;
      const outputPath = path.join(__dirname, filename);

      await generatePdf(learner, outputPath);
      await sendEmail(learner.email, outputPath, learner.month);

      results.push(`${learner.learnerName} (${learner.month}) → emailed to ${learner.email}`);
    }

    res.send(`✅ All timesheets generated and sent!<br><br>${results.join("<br>")}`);

  } catch (err) {
    console.error("Error:", err);
    res.status(500).send(`❌ Error: ${err.message}`);
  } finally {
    // Clean up temporary uploaded file
    if (tempFileToDelete && fs.existsSync(tempFileToDelete)) {
      fs.unlinkSync(tempFileToDelete);
      console.log("Cleaned up temporary uploaded file");
    }
  }
});

app.listen(5000, () => console.log("ICM ShiftTrack server running at http://localhost:5000"));