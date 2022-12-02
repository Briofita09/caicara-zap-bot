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
    const personContacts = [];
    for (let i = 0; i < contacts.length; i++) {
      if (contacts[i].isMyContact && !contacts[i].isGroup) {
        personContacts.push(contacts[i]);
      }
    }
    for (let j = 0; j < sheet1.length; j++) {
      for (let i = 0; i < personContacts.length; i++) {
        if (
          sheet1[j].Tutor !== undefined &&
          personContacts[i].name.trim() === sheet1[j].Tutor.trim() &&
          sheet1[j].Sexo.trim() === "Macho"
        ) {
          console.log(
            `Olá, ${sheet1[j].Tutor}, fizemos o fechamento do ${sheet1[j].Clientes}. O total foi de R$ ${sheet1[j].Total}`
          );
          await client.sendMessage(
            personContacts[i].id._serialized,
            `Olá, ${sheet1[j].Tutor}, fizemos o fechamento do ${sheet1[j].Clientes}. O total foi de R$ ${sheet1[j].Total}`
          );
        } else if (
          sheet1[j].Tutor !== undefined &&
          personContacts[i].name.trim() === sheet1[j].Tutor.trim() &&
          sheet1[j].Sexo.trim() === "Femea"
        ) {
          console.log(
            `Olá, ${sheet1[j].Tutor}, fizemos o fechamento da ${sheet1[j].Clientes}. O total foi de R$ ${sheet1[j].Total}`
          );
          await client.sendMessage(
            personContacts[i].id._serialized,
            `Olá, ${sheet1[j].Tutor}, fizemos o fechamento da ${sheet1[j].Clientes}. O total foi de R$ ${sheet1[j].Total}`
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
  console.log();
  console.log("Client is ready!");
  client
    .getContacts()
    .then(async (contacts) => {
      await sendMessages(contacts);
    })
    .catch((e) => {
      console.log(e);
    });
});

client.initialize();
