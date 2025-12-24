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
  return format(new Date(), "MMMM").toUpperCase(); // e.g., DECEMBER
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
    throw new Error("data.xlsx not found. Please create it with a 'Learners' sheet.");
  }
  const workbook = XLSX.readFile(filePath);
  const sheet = workbook.Sheets["Learners"];
  const rows = XLSX.utils.sheet_to_json(sheet);

  return rows.map(row => ({
    learnerName: row.name || row["Learner Name"] || "",
    idNumber: row.idNumber || row["ID Number"] || "",
    contact: row.contact || row["Contact"] || "",
    email: row.email || row["Email"] || "",
    employer: row.employer || row["Employer"] || "",
    physicalAddress: row.physicalAddress || row["Physical Address"] || "",
    suburb: row.suburb || row["Suburb"] || "",
    cityTown: row.cityTown || row["City/Town"] || "",
    postalCode: row.postalCode || row["Postal Code"] || "",
    localMunicipality: row.localMunicipality || row["Local Municipality"] || "",
    districtMunicipality: row.districtMunicipality || row["District Municipality"] || "",
    metropolitanMunicipality: row.metropolitanMunicipality || row["Metropolitan Municipality"] || "",
    province: row.province || row["Province"] || "",
    supervisorName: row.supervisorName || row["Supervisor Name"] || "",
    supervisorContact: row.supervisorContact || row["Supervisor Contact"] || "",
    supervisorEmail: row.supervisorEmail || row["Supervisor Email"] || "",
    activitiesCount: row.activitiesCount || "",
    tvetCollege: row.tvetCollege || row["TVET College"] || "",
    month: (row.month || getCurrentMonthUpper()).toUpperCase(),
  }));
}

async function generatePdf(data, outputPath) {
  const templateBytes = fs.readFileSync(TEMPLATE_PATH);
  const pdfDoc = await PDFDocument.load(templateBytes);
  const form = pdfDoc.getForm();

  // Fill the fields — only if the field exists in the PDF
  const safeSet = (fieldName, value) => {
    try {
      const field = form.getTextField(fieldName);
      field.setText(value || "");
    } catch (e) {
      // Field doesn't exist in template — that's okay
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

  // Week headers (if you created text fields for them)
  const weeks = getWeekHeaders(data.month);
  safeSet("week1", weeks.week1);
  safeSet("week2", weeks.week2);
  safeSet("week3", weeks.week3);
  safeSet("week4", weeks.week4);
  safeSet("week5", weeks.week5);

  // Optional: flatten to lock pre-filled data (keeps other fields fillable)
  // form.flatten();

  const pdfBytes = await pdfDoc.save();
  fs.writeFileSync(outputPath, pdfBytes);
}

async function sendEmail(recipientEmail, attachmentPath, month) {
  const transporter = nodemailer.createTransport({
    service: "gmail",
    auth: {
      user: "briankgosiemang@gmail.com",
      pass: "vremosugimlfrdmy" // Consider using environment variables later
    }
  });

  await transporter.sendMail({
    from: "Timesheet Bot <briankgosiemang@gmail.com>",
    to: recipientEmail,
    subject: `Your ${month} Fillable Timesheet`,
    text: "Please find your pre-filled timesheet attached. You can open it in any PDF reader and fill in the remaining fields digitally.",
    attachments: [{ filename: path.basename(attachmentPath), path: attachmentPath }]
  });
}

app.post("/generate", async (req, res) => {
  try {
    const learners = await loadDataFromExcel();
    const results = [];

    for (const learner of learners) {
      const safeName = learner.learnerName.replace(/[^a-zA-Z0-9]/g, "_");
      const filename = `timesheet-${safeName}-${learner.month}.pdf`;
      const outputPath = path.join(__dirname, filename);

      await generatePdf(learner, outputPath);
      await sendEmail(learner.email, outputPath, learner.month);

      results.push(`${learner.learnerName} (${learner.month}) → generated & emailed`);
    }

    res.send(`✅ Success!<br><br>${results.join("<br>")}`);
  } catch (err) {
    console.error(err);
    res.status(500).send(`❌ Error: ${err.message}`);
  }
});

app.listen(5000, () => console.log("Server running at http://localhost:5000"));