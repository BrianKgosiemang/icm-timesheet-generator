const fs = require("fs");
const { Document, Packer, Paragraph, TextRun } = require("docx");

const tags = [
  "{{learnerName}}",
  "{{idNumber}}",
  "{{contact}}",
  "{{email}}",
  "{{employer}}",
  "{{supervisorName}}",
  "{{supervisorContact}}",
  "{{month}}"
];

const doc = new Document({
  sections: [{
    children: tags.map(tag => new Paragraph({ children: [new TextRun(tag)] }))
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync("templates/JUNE-template.docx", buffer);
  console.log("✅ Fresh Word template created: templates/JUNE-template.docx");
});
