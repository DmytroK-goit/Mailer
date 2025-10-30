import "dotenv/config";
import { sendMail } from "./utils/sendMail.js";
import XLSX from "xlsx";

const htmlTemplate = `
  <div style="font-family: Arial, sans-serif; color: #333; line-height: 1.6; max-width: 600px; margin: auto; padding: 20px; border: 1px solid #ddd; border-radius: 8px; background: #fafafa;">
    <h1 style="font-size: 20px; text-align: center;" >Шановний клієнте</h1>
    <p style="font-size: 16px; text-align: center;">Наш контролер не застав Вас дома.</p>
    <p style="font-size: 16px; text-align: center;">Будь ласка, передайте показники лічильника в найкоротший термін</p>
    <h2 style="color: #004aad; text-align: center;">Чат-бот <b>MYGRMU_BOT</b></h2>

    <p style="font-size: 16px; text-align: center; margin-bottom: 20px;">
      📞 Кол-центр ГРМУ: <b>0 800 303 104</b><br>
      (дзвінки безкоштовні)
    </p>

    <p style="font-size: 16px; text-align: center;">
      Встанови телеграм-бот<br>
      <b>ТОВ "Газорозподільні мережі України" та передавай показники лічильника online</b>
    </p>

    <div style="text-align: center; margin: 20px 0;">
      <a href="https://t.me/MYGRMU_BOT" 
         style="display: inline-block; background: #004aad; color: #fff; text-decoration: none; padding: 12px 24px; border-radius: 6px; font-size: 16px; font-weight: bold;">
        🚀 Встановити бот
      </a>
    </div>

    <div style="text-align: center; margin: 20px 0;">
      <img src="https://res.cloudinary.com/dingybgqw/image/upload/v1758105978/products/cttlxzgmdlgxendaudm1.jpg" width="420" style="max-width: 100%; border-radius: 6px; border: 1px solid #ccc;" alt="Інструкція">
    </div>

    <p style="text-align: center; margin: 10px 0; font-size: 15px;">
      📲 Або скануй QR-код та реєструйся:
    </p>

    <div style="text-align: center;">
      <img src="https://res.cloudinary.com/dingybgqw/image/upload/v1758105948/products/k5zyka6pfm38jjap3ndv.jpg" width="120" style="border: 1px solid #ccc; padding: 6px; border-radius: 6px;" alt="QR код">
    </div>

  </div>
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
        from: "grmuvinnytsia@gmail.com",
        to,
        subject: "Встанови Telegram-бот ТОВ ГРМУ",
        html: htmlTemplate,
      });
      console.log(`✅ Лист надіслано: ${to}`);
    } catch (err) {
      console.error(`❌ Помилка для ${to}:`, err);
    }

    await new Promise((res) => setTimeout(res, 2000));
  }

  console.log("📬 Розсилка завершена!");
}

sendBulkEmails();
