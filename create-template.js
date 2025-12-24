const fs = require("fs");
const { Document, Packer, Paragraph, TextRun } = require("docx");

const tags = [
  "{learnerName}",
  "{idNumber}",
  "{contact}",
  "{email}",
  "{employer}",
  "{supervisorName}",
  "{supervisorContact}",
  "{month}"
];

const doc = new Document({
  sections: [{
    children: tags.map(tag => 
      new Paragraph({
        children: [
          new TextRun({
            text: `{${tag}}`,  // Wrap with single braces
            italics: true,     // optional: makes it clear it's a placeholder
          })
        ]
      })
    )
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync("templates/JUNE-template.docx", buffer);
  console.log("✅ Fresh Word template created with correct placeholders");
});