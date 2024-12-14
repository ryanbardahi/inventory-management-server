const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');

const drive = new google.drive({
    version: 'v3',
    auth: new google.auth.GoogleAuth({
        keyFile: process.env.GS_CRED,
        scopes: ['https://www.googleapis.com/auth/drive.file']
    })
});

module.exports.addInventoryWithImage = async function (req, res) {
    try {
        if (!req.file) {
            return res.status(400).json({ error: 'No image file uploaded' });
        }

        const folderId = process.env.DRIVE_FOLDER_ID;
        if (!folderId) {
            return res.status(500).json({ error: 'Drive folder ID is not configured' });
        }

        const filePath = req.file.path;

        const fileMetadata = {
            name: req.file.originalname,
            parents: [folderId]
        };

        const media = {
            mimeType: req.file.mimetype,
            body: fs.createReadStream(filePath)
        };

        const imageResponse = await drive.files.create({
            resource: fileMetadata,
            media,
            fields: 'id, name, webViewLink, webContentLink'
        });

        fs.unlinkSync(filePath);

        const imageLink = imageResponse.data.webViewLink;

        const entry = JSON.parse(req.body.entry);
        const currentDate = new Date().toISOString().split('T')[0];

        const values = [[
            entry.Location,
            entry['Item Code'],
            entry.Description,
            entry.UOM,
            entry.Qty,
            entry.Condition,
            entry['Returnable Item'],
            entry.Category,
            currentDate,
            imageLink
        ]];

        await req.sheetsClient.spreadsheets.values.append({
            spreadsheetId: req.sheetId,
            range: 'Inventory!A:J',
            valueInputOption: 'USER_ENTERED',
            resource: { values }
        });

        return res.status(201).json({
            message: 'Inventory entry and image uploaded successfully',
            data: {
                ...entry,
                'Date Counted': currentDate,
                Image: imageLink
            }
        });
    } catch (err) {
        console.error('Error:', err);
        return res.status(500).json({ error: 'Failed to add inventory entry and upload image' });
    }
};

module.exports.searchInventory = async function (req, res) {
    try {
        const { keyword } = req.query;
        if (!keyword) {
            return res.status(400).json({ error: 'Please provide a keyword to search' });
        }

        const response = await req.sheetsClient.spreadsheets.values.get({
            spreadsheetId: req.sheetId,
            range: 'Inventory!A:J'
        });

        const rows = response.data.values;
        if (!rows || rows.length === 0) {
            return res.status(404).json({ error: 'No inventory data found' });
        }

        const headers = rows[0];
        const dataRows = rows.slice(1);

        const filtered = dataRows.filter(row => {
            return row.some(field => 
                field && field.toString().toLowerCase().includes(keyword.toLowerCase())
            );
        });

        const formattedResults = filtered.map(entry => {
            const obj = {};
            headers.forEach((header, i) => {
                obj[header] = entry[i] || '';
            });
            return obj;
        });

        return res.status(200).json({ results: formattedResults });
    } catch (err) {
        console.error('Error searching inventory:', err);
        return res.status(500).json({ error: 'Failed to search inventory' });
    }
};