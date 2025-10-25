

const express = require('express');
const bodyParser = require('body-parser');
const XLSX = require('xlsx');
const cors = require('cors');
const fs = require('fs');

const app = express();
const PORT = 3000;
const EXCEL_FILE = 'rsvp_data.xlsx';

app.use(cors());
app.use(bodyParser.json());

// Ensure Excel file exists
if (!fs.existsSync(EXCEL_FILE)) {
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet([]);
    XLSX.utils.book_append_sheet(wb, ws, 'RSVPs');
    XLSX.writeFile(wb, EXCEL_FILE);
}

app.post('/submit-rsvp', (req, res) => {
    const data = req.body;

    try {
        const wb = XLSX.readFile(EXCEL_FILE);
        const ws = wb.Sheets['RSVPs'];
        const existingData = XLSX.utils.sheet_to_json(ws);

        existingData.push({
            Timestamp: new Date().toISOString(),
            Name: data.name,
            Contact: data.contact,
            Side: data.side,
            Attending: data.attending,
            Guests: data.guests,
            Events: data.events.join(', ')
        });

        const newWs = XLSX.utils.json_to_sheet(existingData);
        wb.Sheets['RSVPs'] = newWs;
        XLSX.writeFile(wb, EXCEL_FILE);

        res.status(200).json({ status: 'Success', message: 'RSVP saved!' });
    } catch (error) {
        res.status(500).json({ status: 'Error', message: 'Failed to save RSVP.' });
    }
});

const path = require("path");
const EXCEL_FILE = path.join(__dirname, "tmp_data.xlsx");

app.get("/download-rsvp", (req, res) => {
  res.download(EXCEL_FILE, "RSVP_List.xlsx", (err) => {
    if (err) {
      console.error("Download error:", err);
      res.status(500).send("Could not download the file.");
    }
  });
});

app.listen(PORT, () => {
    console.log(`Server running on http://localhost:${PORT}`);
});
