import qrcode from "qrcode-terminal";
import { Client } from "whatsapp-web.js";
import xlsx from "xlsx";

const workbook = xlsx.readFile("./teste1.xlsx");

let worksheet = {};

for (const sheetName of workbook.SheetNames) {
  worksheet[sheetName] = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName]);
}

const sheet1 = worksheet.set22;

const client = new Client();

async function sendMessages(contacts) {
  try {
    for (let i = 0; i < contacts.length; i++) {
      for (let j = 0; j < sheet1.length; j++) {
        if (
          sheet1[j].Tutor !== undefined &&
          contacts[i].name === sheet1[j].Tutor
        ) {
          console.log(
            `Olá, ${sheet1[j].Tutor}, fizemos o fechamento do ${sheet1[j].Clientes}. O total foi de R$ ${sheet1[j].Total}`
          );
          await client.sendMessage(
            contacts[i].id._serialized,
            `Olá, ${sheet1[j].Tutor}, fizemos o fechamento do ${sheet1[j].Clientes}. O total foi de R$ ${sheet1[j].Total}`
          );
        }
      }
    }
  } catch (e) {
    console.log(e);
  }
}

client.on("qr", (qr) => {
  qrcode.generate(qr, { small: true });
});

client.on("ready", () => {
  console.log("Client is ready!");
  client.getContacts().then(async (contacts) => {
    await sendMessages(contacts);
  });
});

client.initialize();
