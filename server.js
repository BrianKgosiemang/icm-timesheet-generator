const express = require("express");
const fs = require("fs");
const path = require("path");
const xml2js = require("xml2js");
const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");
const nodemailer = require("nodemailer");
const bodyParser = require("body-parser");

const app = express();
app.use(bodyParser.json());
app.use(express.static("public"));

const TEMPLATE_PATH = path.join(__dirname, "templates/JUNE-template.docx");
const OUTPUT_PATH = path.join(__dirname, "output-timesheet.docx");
const XML_PATH = path.join(__dirname, "data.xml");

async function loadData() {
  const xml = fs.readFileSync(XML_PATH);
  const result = await xml2js.parseStringPromise(xml);
  const learner = result.learners.learner[0];
  return {
    learnerName: learner.name[0],
    idNumber: learner.id[0],
    contact: learner.contact[0],
    email: learner.email[0],
    employer: learner.employer[0],
    supervisorName: learner.supervisor[0].name[0],
    supervisorContact: learner.supervisor[0].contact[0],
    month: learner.month[0],
  };
}

function generateDocx(data) {
  const content = fs.readFileSync(TEMPLATE_PATH, "binary");
  const zip = new PizZip(content);
  const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true });

  doc.setData(data);
  doc.render();

  const buffer = doc.getZip().generate({ type: "nodebuffer" });
  fs.writeFileSync(OUTPUT_PATH, buffer);
}

async function sendEmail(recipientEmail) {
  const transporter = nodemailer.createTransport({
    service: "gmail",
    auth: {
      user: "briankgosiemang@gmail.com",
      pass: "vremosugimlfrdmy"
    }
  });

  await transporter.sendMail({
    from: "Timesheet Bot <youremail@gmail.com>",
    to: recipientEmail,
    subject: "Your June Timesheet",
    text: "Please find your June timesheet attached.",
    attachments: [{ filename: "timesheet.docx", path: OUTPUT_PATH }]
  });
}

app.post("/generate", async (req, res) => {
  try {
    const data = await loadData();
    generateDocx(data);
    await sendEmail(data.email);
    res.send("✅ Timesheet generated and emailed!");
  } catch (err) {
    console.error(err);
    res.status(500).send("❌ Failed to generate timesheet");
  }
});

app.listen(5000, () => console.log("🚀 Server running at http://localhost:5000"));
