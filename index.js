const venom = require('venom-bot');
const XLSX = require('xlsx');
const axios = require('axios');
const fs = require('fs');

const excelUrl = "https://www.edicolamarlene.it/database.xlsm";
const localCsv = "database.csv";

async function downloadAndConvertExcel() {
    const response = await axios.get(excelUrl, { responseType: 'arraybuffer' });
    const workbook = XLSX.read(response.data, { type: 'buffer' });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const csv = XLSX.utils.sheet_to_csv(sheet);
    fs.writeFileSync(localCsv, csv);
}

function cercaCopertina(codice, numero) {
    const data = fs.readFileSync(localCsv, 'utf8').split("\n");
    const intestazioni = data[0].split(",");
    const codiceIndex = intestazioni.findIndex(h => h.trim().toLowerCase().includes("codice"));
    const numeroIndex = intestazioni.findIndex(h => h.trim().toLowerCase().includes("numero"));
    const linkIndex = intestazioni.findIndex(h => h.trim().toLowerCase().includes("foto") || h.trim().toLowerCase().includes("immagine"));

    let match = null;
    let precedenti = [];

    for (let i = 1; i < data.length; i++) {
        const riga = data[i].split(",");
        if (riga[codiceIndex] === codice) {
            if (riga[numeroIndex] === numero) {
                match = riga[linkIndex];
            } else if (precedenti.length < 3) {
                precedenti.push(riga[linkIndex]);
            }
        }
    }

    if (match) {
        return [match];
    } else {
        return precedenti;
    }
}

venom
  .create({
    session: 'copertine-bot',
    headless: true,
    useChrome: false,
    browserArgs: ['--no-sandbox']
  })
  .then(client => {
    client.onMessage(async message => {
      if (message.body && message.body.includes(" ")) {
        const [codice, numero] = message.body.trim().split(" ");
        if (codice && numero) {
          await downloadAndConvertExcel();
          const risultati = cercaCopertina(codice, numero);
          if (risultati.length === 0) {
            await client.sendText(message.from, "Foto non presente.");
          } else {
            for (let url of risultati) {
              if (!url) continue;
              let link = url.trim();
              try {
                const test = await axios.head(link);
                if (test.status !== 200) throw new Error();
              } catch {
                link = link.replace("http://www.edicoland.com/images/", "https://www.edicolamarlene.it/wp-content/uploads/EdiscanIMG/");
              }
              await client.sendImage(message.from, link, "copertina.jpg", "Ecco la copertina trovata");
            }
          }
        }
      } else {
        await client.sendText(message.from, "Ciao, sono il tuo chatbot personale per la ricerca delle copertine. Inviami codice e numero (es: 123456 25)");
      }
    });
  })
  .catch(e => {
    console.error('Errore nella creazione della sessione Venom:', e);
  });
