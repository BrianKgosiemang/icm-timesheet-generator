require('dotenv').config();

const express = require("express");
const fs = require("fs");
const path = require("path");
const os = require("os");
const multer = require("multer");
const nodemailer = require("nodemailer");
const XLSX = require("xlsx");
const { format, getYear, subMonths, getDaysInMonth, addDays } = require("date-fns");
const { PDFDocument } = require("pdf-lib");

const app = express();
app.use(express.static("public"));

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
    month: "", // Will be overridden by frontend
    year: 0,   // Will be overridden by frontend
  }));
}

async function generatePdf(data, outputPath) {
  const templateBytes = fs.readFileSync(TEMPLATE_PATH);
  const pdfDoc = await PDFDocument.load(templateBytes);
  const form = pdfDoc.getForm();

  const safeSet = (fieldName, value) => {
    try {
      if (value !== undefined && value !== "") {
        form.getTextField(fieldName).setText(value.toString());
        console.log(`Filled: ${fieldName} = ${value}`);
      }
    } catch (e) {
      // Silent if field missing
    }
  };

  // Personal info
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

  // === ICM LAYOUT: DATES & DAYS ===
  const monthNames = ["JANUARY","FEBRUARY","MARCH","APRIL","MAY","JUNE","JULY","AUGUST","SEPTEMBER","OCTOBER","NOVEMBER","DECEMBER"];
  const monthIndex = monthNames.indexOf(data.month);
  if (monthIndex === -1) return;

  const year = data.year;

  const formatDate = (date) => format(date, "dd-MM-yy");
  const formatDay = (date) => format(date, "EEE");

  // 1. Week 1: 26th prev month → Sunday
  const prevMonthIndex = monthIndex === 0 ? 11 : monthIndex - 1;
  const prevYear = monthIndex === 0 ? year - 1 : year;
  let currentDate = new Date(prevYear, prevMonthIndex, 26);

  let dayIndex = 1;
  while (currentDate.getDay() !== 0 && dayIndex <= 7) { // Fill until Sunday (0 = Sunday)
    safeSet(`day_w1_d${dayIndex}`, formatDay(currentDate));
    safeSet(`date_w1_d${dayIndex}`, formatDate(currentDate));
    currentDate = addDays(currentDate, 1);
    dayIndex++;
  }
  // Fill the Sunday
  if (dayIndex <= 7) {
    safeSet(`day_w1_d${dayIndex}`, formatDay(currentDate));
    safeSet(`date_w1_d${dayIndex}`, formatDate(currentDate));
    currentDate = addDays(currentDate, 1); // Move to Monday for Week 2
  }

  // 2. Weeks 2–4: Full Monday–Sunday
  const fullWeeks = [2, 3, 4];
  fullWeeks.forEach(weekNum => {
    for (let d = 1; d <= 7; d++) {
      if (currentDate.getMonth() === monthIndex) {
        safeSet(`day_w${weekNum}_d${d}`, formatDay(currentDate));
        safeSet(`date_w${weekNum}_d${d}`, formatDate(currentDate));
      }
      currentDate = addDays(currentDate, 1);
    }
  });

  // 3. Week 5: Monday after Week 4 → 25th of current month
  dayIndex = 1;
  while (currentDate.getDate() <= 25 && dayIndex <= 7) {
    safeSet(`day_w5_d${dayIndex}`, formatDay(currentDate));
    safeSet(`date_w5_d${dayIndex}`, formatDate(currentDate));
    currentDate = addDays(currentDate, 1);
    dayIndex++;
  }

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

  const logoUrl = process.env.APP_URL 
    ? `${process.env.APP_URL.trim().replace(/\/$/, "")}/icm-logo.png`
    : "https://your-app.onrender.com/icm-logo.png";

  const htmlEmail = `
<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Your ${month} Timesheet</title>
  <style>
    body { font-family: 'Segoe UI', Arial, sans-serif; background:#f4f7fc; margin:0; padding:0; color:#333; }
    .container { max-width:600px; margin:20px auto; background:white; border-radius:12px; overflow:hidden; box-shadow:0 10px 30px rgba(0,51,102,0.1); border:1px solid #e0e7ff; }
    .header { background:linear-gradient(135deg,#003366 0%,#004488 100%); padding:30px; text-align:center; color:white; }
    .header img { height:80px; margin-bottom:15px; }
    .header h1 { margin:0; font-size:28px; font-weight:600; }
    .header p { margin:8px 0 0; font-size:16px; opacity:0.95; }
    .content { padding:35px; line-height:1.6; }
    .highlight { background:#e6f0ff; padding:20px; border-radius:10px; border-left:5px solid #0066cc; margin:20px 0; }
    .highlight strong { color:#003366; }
    .btn { display:inline-block; background:linear-gradient(135deg,#003366 0%,#0066cc 100%); color:white; padding:14px 32px; text-decoration:none; border-radius:50px; font-weight:600; margin:15px 0; box-shadow:0 6px 15px rgba(0,102,204,0.2); }
    .footer { background:#f8fbff; padding:25px; text-align:center; color:#666; font-size:14px; border-top:1px solid #e0e7ff; }
    .footer a { color:#0066cc; text-decoration:none; }
    @media (max-width:600px) { .container{margin:10px;} .content{padding:25px;} .header{padding:25px;} }
  </style>
</head>
<body>
  <div class="container">
    <div class="header">
      <img src="${logoUrl}" alt="ICM Logo">
      <h1>Your ${month} Timesheet</h1>
      <p>Institute of Certified Managers – Work Integrated Learning</p>
    </div>

    <div class="content">
      <p>Dear Learner,</p>

      <p>We hope you're having a productive and rewarding experience during your Work Integrated Learning placement.</p>

      <div class="highlight">
        <p><strong>Your pre-filled timesheet for <u>${month}</u> is attached.</strong></p>
        <p>Please complete the following:</p>
        <ul>
          <li>Daily <strong>Time In / Time Out</strong></li>
          <li>Your <strong>intern signature</strong> each day</li>
          <li>Supervisor sign-off weekly</li>
          <li>Submit by the deadline</li>
        </ul>
      </div>

      <p>You can edit this PDF digitally using any PDF reader (Adobe Acrobat Reader, browser, phone app, etc.).</p>

      <p style="text-align:center;">
        <a href="#" class="btn">Open Attached Timesheet</a>
      </p>

      <p>Thank you for your dedication and hard work!</p>

      <p>Best regards,<br>
      <strong>The ICM WIL Team</strong><br>
      Institute of Certified Managers</p>
    </div>

    <div class="footer">
      <p>This is an automated message from <strong>ICM ShiftTrack</strong>.</p>
      <p>Need help? Contact your WIL coordinator.</p>
    </div>
  </div>
</body>
</html>
  `;

  await transporter.sendMail({
    from: process.env.EMAIL_FROM || `"ICM ShiftTrack" <${process.env.EMAIL_USER}>`,
    to: recipientEmail,
    subject: `Your ${month} Timesheet - ICM WIL`,
    text: `Hello,

Your pre-filled ${month} timesheet is attached.

Please fill in your daily hours and signatures digitally, then submit.

Thank you!
ICM Team`,
    html: htmlEmail,
    attachments: [
      {
        filename: path.basename(attachmentPath),
        path: attachmentPath
      }
    ]
  });

  console.log(`Styled email sent to: ${recipientEmail}`);
}

app.post("/generate", upload.single("dataFile"), async (req, res) => {
  let excelFilePath = null;
  let tempFileToDelete = null;

  try {
    if (req.file) {
      excelFilePath = req.file.path;
      tempFileToDelete = req.file.path;
      console.log(`Using uploaded file: ${req.file.originalname}`);
    } else if (fs.existsSync(path.join(__dirname, "data.xlsx"))) {
      excelFilePath = path.join(__dirname, "data.xlsx");
      console.log("Using default data.xlsx");
    } else {
      return res.status(400).send("❌ No file uploaded and no default 'data.xlsx' found.");
    }

    let learners = await loadDataFromExcel(excelFilePath);
    const results = [];

    const testEmail = req.body.testEmail || req.body["testEmail"];
    const selectedMonth = req.body.month; // From dropdown
    const selectedYear = req.body.year ? parseInt(req.body.year) : new Date().getFullYear();

    // Override month and year if selected from frontend
    if (selectedMonth) {
      const upperMonth = selectedMonth.toUpperCase();
      learners = learners.map(learner => ({
        ...learner,
        month: upperMonth,
        year: selectedYear,
      }));
    }

    if (testEmail) {
      if (learners.length === 0) {
        return res.status(400).send("❌ No learner data available for test.");
      }

      const testLearner = { ...learners[0] };
      testLearner.email = testEmail;

      const safeName = testLearner.learnerName.replace(/[^a-zA-Z0-9]/g, "_");
      const filename = `timesheet-test-${safeName}-${testLearner.month}-${testLearner.year}.pdf`;
      const outputPath = path.join(__dirname, filename);

      await generatePdf(testLearner, outputPath);
      await sendEmail(testEmail, outputPath, testLearner.month);

      return res.send(`✅ Test timesheet for <strong>${testLearner.month} ${testLearner.year}</strong> sent to <strong>${testEmail}</strong>`);
    }

    // Normal mode: Send to all learners
    for (const learner of learners) {
      if (!learner.email) {
        results.push(`⚠️ ${learner.learnerName} skipped — no email`);
        continue;
      }

      const safeName = learner.learnerName.replace(/[^a-zA-Z0-9]/g, "_");
      const filename = `timesheet-${safeName}-${learner.month}-${learner.year}.pdf`;
      const outputPath = path.join(__dirname, filename);

      await generatePdf(learner, outputPath);
      await sendEmail(learner.email, outputPath, learner.month);

      results.push(`${learner.learnerName} (${learner.month} ${learner.year}) → emailed`);
    }

    const periodText = selectedMonth ? `${selectedMonth.toUpperCase()} ${selectedYear}` : "current period";
    res.send(`✅ Timesheets for <strong>${periodText}</strong> generated and sent!<br><br>${results.join("<br>")}`);

  } catch (err) {
    console.error("Error:", err);
    res.status(500).send(`❌ Error: ${err.message}`);
  } finally {
    if (tempFileToDelete && fs.existsSync(tempFileToDelete)) {
      fs.unlinkSync(tempFileToDelete);
      console.log("Cleaned up temporary file");
    }
  }
});

app.listen(5000, () => console.log("ICM ShiftTrack server running at http://localhost:5000"));