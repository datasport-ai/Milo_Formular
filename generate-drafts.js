const fs = require("fs");
const path = require("path");
const csv = require("csv-parser");

const OUT_DIR = "drafts";
fs.mkdirSync(OUT_DIR, { recursive: true });

function cleanFileName(str) {
  return String(str || "")
    .replace(/[<>:"/\\|?*]/g, "_")
    .slice(0, 80);
}

function encodeHeader(str) {
  return Buffer.from(str, "utf8").toString("base64");
}

fs.createReadStream("contacts.csv")
  .pipe(csv())
  .on("data", (row) => {
      const to = row.email?.trim() || "";
      const firstName = row.first_name?.trim() || "";
      const lastName = row.last_name?.trim() || "";
      const link = row.link?.trim() || "";
      if (!to || !link) {
            console.log(`Skipped: ${firstName} ${lastName}`);
                return;
            }

    const subject = "FIT for LIFE Coach - Formular";
    const html = `
<p>Hi ${firstName},</p>

<p>Please help us out and answer this form at this link:</p>

<p><a href="${link}">Answer the questionnaire</a></p>
`;

    const eml = [
      `To: ${to}`,
      `Subject: =?UTF-8?B?${encodeHeader(subject)}?=`,
      `MIME-Version: 1.0`,
      `Content-Type: text/html; charset="UTF-8"`,
      `Content-Transfer-Encoding: 8bit`,
      ``,
      html,
    ].join("\r\n");

    const fileName = cleanFileName(
          `${firstName}_${lastName}_${to}.eml`
    );
    fs.writeFileSync(path.join(OUT_DIR, fileName), eml, "utf8");

    console.log(`Created: ${fileName}`);
  })
  .on("end", () => {
    console.log("Done. EML files created in /drafts");
  });