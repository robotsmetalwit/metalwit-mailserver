const express = require('express');
const nodemailer = require('nodemailer');
const cors = require('cors');
const bodyParser = require('body-parser');
const ExcelJS = require('exceljs');

const app = express();
const PORT = process.env.PORT || 3000;

// Konfiguracja z zmiennych środowiskowych (na Render.com ustawiasz w dashboardzie)
const GMAIL_USER = process.env.GMAIL_USER || 'robotsmetalwit@gmail.com';
const GMAIL_PASS = process.env.GMAIL_PASS || 'usjt bjqo wfxz epoi';
const DEFAULT_TO = process.env.DEFAULT_TO || 'robotsmetalwit@gmail.com';

// Middleware
app.use(cors());
app.use(bodyParser.json({ limit: '10mb' }));

// Jeden transporter dla całej aplikacji
const transporter = nodemailer.createTransport({
  service: 'gmail',
  auth: {
    user: GMAIL_USER,
    pass: GMAIL_PASS,
  },
});

// Health check (Render wymaga tego endpointu)
app.get('/', (req, res) => {
  res.json({ status: 'OK', message: 'MailServer działa' });
});

// Endpoint do wysyłania maili
app.post('/send-email', async (req, res) => {
  const { employeeName, itemsList } = req.body;

  if (!employeeName || !itemsList) {
    return res.status(400).json({ error: 'Brak wymaganych danych' });
  }

  const mailOptions = {
    from: GMAIL_USER,
    to: DEFAULT_TO,
    subject: employeeName,
    text: `${employeeName}\n\n${itemsList}`,
  };

  try {
    await transporter.sendMail(mailOptions);
    console.log(`Mail wysłany dla: ${employeeName}`);
    res.status(200).json({ message: 'OK' });
  } catch (error) {
    console.error('Błąd wysyłki:', error);
    res.status(500).json({ error: 'Błąd serwera poczty' });
  }
});


// Endpoint do wysyłki raportu (dane JSON)
app.post('/send-report', async (req, res) => {
  try {
    const { date, history } = req.body;

    if (!history || history.length === 0) {
      return res.status(400).json({ error: 'Brak danych do raportu' });
    }

    // Tworzenie pliku Excel
    const workbook = new ExcelJS.Workbook();

    // Arkusz 1: Podsumowanie
    const summarySheet = workbook.addWorksheet('Podsumowanie');
    summarySheet.columns = [
      { header: 'Przedmiot', key: 'item', width: 30 },
      { header: 'Ilość', key: 'total', width: 15 },
    ];

    // Sumowanie ilości dla każdego przedmiotu
    const summary = {};
    history.forEach(transaction => {
      transaction.items.forEach(item => {
        if (!summary[item.name]) {
          summary[item.name] = 0;
        }
        summary[item.name] += item.quantity;
      });
    });

    // Wypełnianie arkusza podsumowania
    Object.entries(summary).forEach(([item, total]) => {
      summarySheet.addRow({ item, total });
    });

    // Arkusz 2: Szczegółowa historia
    const detailsSheet = workbook.addWorksheet('Historia');
    detailsSheet.columns = [
      { header: 'Data', key: 'date', width: 12 },
      { header: 'Godzina', key: 'time', width: 10 },
      { header: 'Pracownik', key: 'employee', width: 25 },
      { header: 'Przedmiot', key: 'item', width: 30 },
      { header: 'Ilość', key: 'quantity', width: 10 },
    ];

    history.forEach(transaction => {
      const time = new Date(transaction.timestamp).toLocaleTimeString('pl-PL');
      transaction.items.forEach(item => {
        detailsSheet.addRow({
          date: transaction.date,
          time,
          employee: transaction.employeeName,
          item: item.name,
          quantity: item.quantity,
        });
      });
    });

    // Zapis pliku do bufora
    const buffer = await workbook.xlsx.writeBuffer();

    // Opcje maila z załącznikiem
    const toEmail = req.body.toEmail || DEFAULT_TO;
    const mailOptions = {
      from: GMAIL_USER,
      to: toEmail,
      subject: `Raport z dnia ${date}`,
      text: 'W załączniku raport z dwoma arkuszami: podsumowanie i szczegóły.',
      attachments: [
        {
          filename: `raport_${date}.xlsx`,
          content: buffer,
        },
      ],
    };

    await transporter.sendMail(mailOptions);
    res.json({ message: 'OK' });
  } catch (error) {
    console.error('Błąd wysyłki raportu:', error);
    res.status(500).json({ error: 'Błąd wysyłki' });
  }
});

// Start serwera
app.listen(PORT, () => {
  console.log(`Serwer działa na porcie ${PORT}`);
});