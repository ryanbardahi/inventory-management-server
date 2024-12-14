require('dotenv').config();
const express = require('express');
const cors = require('cors');
const { google } = require('googleapis');

const inventoryRoutes = require('./routes/inventory');

const app = express();
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

const corsOptions = {
    origin: [
        'http://localhost:3000'
    ],
    credentials: true,
    optionsSuccessStatus: 200
};
app.use(cors(corsOptions));

const { GS_CRED, SH_ID } = process.env;

let sheetsClient;
async function initGoogleSheets() {
    try {
        const auth = new google.auth.GoogleAuth({
            keyFile: GS_CRED,
            scopes: ['https://www.googleapis.com/auth/spreadsheets']
        });

        const authClient = await auth.getClient();
        sheetsClient = google.sheets({ version: 'v4', auth: authClient });

        console.log('Connected to Google Sheets successfully');
    } catch (err) {
        console.error('Failed to connect to Google Sheets:', err.message);
    }
}
initGoogleSheets();

app.use((req, res, next) => {
    if (!sheetsClient) {
        return res.status(500).json({ error: 'Google Sheets not initialized' });
    }
    req.sheetsClient = sheetsClient;
    req.sheetId = SH_ID;
    next();
});

app.use('/inventory', inventoryRoutes);

if (require.main === module) {
    const PORT = 4000;
    app.listen(PORT, () => {
        console.log(`API is now online on port ${PORT}`);
    });
}

module.exports = { app };