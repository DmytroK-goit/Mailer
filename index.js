import "dotenv/config";
import { sendMail } from "./utils/sendMail.js";
import XLSX from "xlsx";

const htmlTemplate = `
  <p>
    Прохання встановити телеграм бот ГАЗМЕРЕЖІ 
    <a href="https://t.me/your_bot_link">Встановити бот</a>
  </p>
`;

const workbook = XLSX.readFile("./emails.xlsx");
const sheetName = workbook.SheetNames[0];
const worksheet = workbook.Sheets[sheetName];
const data = XLSX.utils.sheet_to_json(worksheet);

const recipients = data.map((row) => row.email).filter(Boolean);

async function sendBulkEmails() {
  for (let i = 0; i < recipients.length; i++) {
    const to = recipients[i];

    try {
      await sendMail({
        from: `"Газмережі" <${process.env.SMTP_USER}>`,
        to,
        subject: "Встановіть наш Telegram-бот",
        html: htmlTemplate,
      });
      console.log(`✅ Лист надіслано: ${to}`);
    } catch (err) {
      console.error(`❌ Помилка для ${to}:`, err);
    }

    await new Promise((res) => setTimeout(res, 1000));
  }

  console.log("📬 Розсилка завершена!");
}

sendBulkEmails();
