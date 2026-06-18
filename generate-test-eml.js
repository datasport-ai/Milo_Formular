const fs = require("fs");
const csv = require("csv-parser");

let done = false;

fs.createReadStream("contacts.csv")
  .pipe(csv())
  .on("data", (row) => {
    if (done) return;
    done = true;

    const html = `
<p>Hi ${row.first_name},</p>

<p>Please help us out and answer this form at this link:</p>

<p><a href="${row.link}">Answer the questionnaire</a></p>
`;

    const eml = [
      `To: ${row.email}`,
      `Subject: Questionnaire`,
      `MIME-Version: 1.0`,
      `Content-Type: text/html; charset="UTF-8"`,
      ``,
      html,
    ].join("\r\n");

    fs.writeFileSync("test-email.eml", eml, "utf8");

    console.log("Test email created for:", row.email);
    console.log("Link:", row.link);
  });