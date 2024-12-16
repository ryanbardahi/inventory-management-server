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

// Utility function to convert column index to letter (supports multiple letters)
function getColumnLetter(colIndex) {
    let letter = '';
    while (colIndex >= 0) {
        letter = String.fromCharCode((colIndex % 26) + 65) + letter;
        colIndex = Math.floor(colIndex / 26) - 1;
    }
    return letter;
}

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
            insertDataOption: 'INSERT_ROWS',
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

async function findFirstBlankRowAtoJ(sheetsClient, spreadsheetId) {
    // Fetch all data in columns A:J of 'Form Responses'
    const response = await sheetsClient.spreadsheets.values.get({
        spreadsheetId,
        range: 'Form Responses!A:J',
    });

    const rows = response.data.values || [];
    // Assuming row 1 has headers
    for (let i = 0; i < rows.length; i++) {
        const row = rows[i];
        // Check if all cells in A:J are blank
        const isBlank = row.slice(0, 10).every(cell => !cell || cell.toString().trim() === '');
        if (isBlank) {
            return i + 1; // Sheets are 1-indexed
        }
    }

    // If no blank row is found within existing data, append to the next row
    return rows.length + 1;
}

module.exports.issueInventory = async function (req, res) {
    try {
        const {
            itemCode,
            issuanceQty,
            issuedBy,
            activity,
            notes,
            location
        } = req.body;

        // Validate required fields
        if (!itemCode || !issuanceQty || !issuedBy || !activity || !location) {
            return res.status(400).json({ error: 'Missing required fields' });
        }

        // Fetch all inventory data
        const inventoryResponse = await req.sheetsClient.spreadsheets.values.get({
            spreadsheetId: req.sheetId,
            range: 'Inventory!A:J'
        });

        const rows = inventoryResponse.data.values;
        if (!rows || rows.length < 2) { // Assuming row 1 has headers
            return res.status(404).json({ error: 'No inventory data found' });
        }

        const headers = rows[0];
        const dataRows = rows.slice(1);

        // Identify column indices
        const colLocation = headers.indexOf('Location');
        const colItemCode = headers.indexOf('Item Code');
        const colDescription = headers.indexOf('Description');
        const colQty = headers.indexOf('Qty');
        const colReturnable = headers.indexOf('Returnable Item');
        const colImageLink = headers.indexOf('Image Link');

        if (
            colLocation === -1 || colItemCode === -1 ||
            colDescription === -1 || colQty === -1 ||
            colReturnable === -1 || colImageLink === -1
        ) {
            return res.status(500).json({ error: 'Inventory sheet headers missing required columns' });
        }

        let targetRowIndex = -1;
        let currentQty = 0;
        let description = '';
        let returnableItem = '';
        let imageLink = '';

        // Locate the inventory item
        for (let i = 0; i < dataRows.length; i++) {
            const row = dataRows[i];
            if (row[colItemCode] === itemCode && row[colLocation] === location) {
                currentQty = parseFloat(row[colQty]) || 0;
                description = row[colDescription] || '';
                returnableItem = row[colReturnable] || '';
                imageLink = row[colImageLink] || '';
                targetRowIndex = i + 2; // +2 because dataRows starts from row 2 (1-indexed)
                break;
            }
        }

        if (targetRowIndex === -1) {
            return res.status(404).json({ error: 'Item not found in inventory for the specified location' });
        }

        const requestQty = parseFloat(issuanceQty);
        if (isNaN(requestQty) || requestQty <= 0) {
            return res.status(400).json({ error: 'Invalid issuance quantity' });
        }

        if (currentQty < requestQty) {
            return res.status(400).json({ error: 'Not enough quantity in inventory to fulfill issuance' });
        }

        const newQty = currentQty - requestQty;

        // Construct the timestamp in the desired format: YYYY-MM-DD, HH:MM:SS AM/PM
        const now = new Date();
        const year = now.getFullYear();
        const month = String(now.getMonth() + 1).padStart(2, '0');
        const day = String(now.getDate()).padStart(2, '0');

        let hours = now.getHours();
        const minutes = String(now.getMinutes()).padStart(2, '0');
        const seconds = String(now.getSeconds()).padStart(2, '0');

        const ampm = hours >= 12 ? 'PM' : 'AM';
        hours = hours % 12;
        hours = hours || 12; // if hours is 0, set it to 12

        const timestamp = `${year}-${month}-${day}, ${hours}:${minutes}:${seconds} ${ampm}`;

        const formValues = [[
            timestamp,
            itemCode,
            requestQty,
            location,
            description,
            returnableItem,
            issuedBy,
            activity,
            notes || '',
            imageLink
        ]];

        // Find the first blank row in 'Form Responses!A:J'
        const blankRow = await findFirstBlankRowAtoJ(req.sheetsClient, req.sheetId);
        console.log(`Blank Row Found in Form Responses: ${blankRow}`);

        // Determine the range to update in 'Form Responses!A:J'
        const updateRange = `Form Responses!A${blankRow}:J${blankRow}`;

        // Update the identified row with formValues
        await req.sheetsClient.spreadsheets.values.update({
            spreadsheetId: req.sheetId,
            range: updateRange,
            valueInputOption: 'USER_ENTERED',
            resource: { values: formValues }
        });

        // Update the inventory quantity in 'Inventory' tab
        const qtyColumnLetter = getColumnLetter(colQty);
        const inventoryUpdateRange = `Inventory!${qtyColumnLetter}${targetRowIndex}`;

        await req.sheetsClient.spreadsheets.values.update({
            spreadsheetId: req.sheetId,
            range: inventoryUpdateRange,
            valueInputOption: 'USER_ENTERED',
            resource: { values: [[newQty]] }
        });

        return res.status(200).json({
            message: 'Inventory issued successfully',
            data: {
                Timestamp: timestamp,
                'Item Code': itemCode,
                'Issuance Qty': requestQty,
                Location: location,
                Description: description,
                'Returnable Item': returnableItem,
                'Issued by': issuedBy,
                Activity: activity,
                'Notes/Comments': notes || '',
                'Image Link': imageLink
            }
        });
    } catch (err) {
        console.error('Error issuing inventory:', err);
        return res.status(500).json({ error: 'Failed to issue inventory' });
    }
};

module.exports.receiveInventory = async function (req, res) {
    try {
        const {
            itemCode,
            receiptQty,
            receivedBy,
            notes,
            location
        } = req.body;

        // Validate required fields
        if (!itemCode || !receiptQty || !receivedBy || !location) {
            return res.status(400).json({ error: 'Missing required fields' });
        }

        // Fetch all inventory data
        const inventoryResponse = await req.sheetsClient.spreadsheets.values.get({
            spreadsheetId: req.sheetId,
            range: 'Inventory!A:J'
        });

        const rows = inventoryResponse.data.values;
        if (!rows || rows.length < 2) {
            return res.status(404).json({ error: 'No inventory data found' });
        }

        const headers = rows[0];
        const dataRows = rows.slice(1);

        // Identify column indices
        const colLocation = headers.indexOf('Location');
        const colItemCode = headers.indexOf('Item Code');
        const colDescription = headers.indexOf('Description');
        const colQty = headers.indexOf('Qty');
        const colReturnable = headers.indexOf('Returnable Item');
        const colImageLink = headers.indexOf('Image Link');

        // Check if all required columns are present
        if (
            colLocation === -1 || colItemCode === -1 ||
            colDescription === -1 || colQty === -1 ||
            colReturnable === -1 || colImageLink === -1
        ) {
            return res.status(500).json({ error: 'Inventory sheet headers missing required columns' });
        }

        // Find the target inventory row based on Item Code and Location
        let targetRowIndex = -1;
        let currentQty = 0;
        let description = '';
        let returnableItem = '';
        let imageLink = '';

        for (let i = 0; i < dataRows.length; i++) {
            const row = dataRows[i];
            if (row[colItemCode] === itemCode && row[colLocation] === location) {
                currentQty = parseFloat(row[colQty]) || 0;
                description = row[colDescription] || '';
                returnableItem = row[colReturnable] || '';
                imageLink = row[colImageLink] || '';
                targetRowIndex = i + 1; // +1 because dataRows starts from row 2
                break;
            }
        }

        if (targetRowIndex === -1) {
            return res.status(404).json({ error: 'Item not found in inventory for the specified location' });
        }

        // Validate receipt quantity
        const requestQty = parseFloat(receiptQty);
        if (isNaN(requestQty) || requestQty <= 0) {
            return res.status(400).json({ error: 'Invalid receipt quantity' });
        }

        const newQty = currentQty + requestQty;

        // Construct the timestamp in the desired format: YYYY-MM-DD, HH:MM:SS AM/PM
        const now = new Date();
        const year = now.getFullYear();
        const month = String(now.getMonth() + 1).padStart(2, '0');
        const day = String(now.getDate()).padStart(2, '0');

        let hours = now.getHours();
        const minutes = String(now.getMinutes()).padStart(2, '0');
        const seconds = String(now.getSeconds()).padStart(2, '0');

        const ampm = hours >= 12 ? 'PM' : 'AM';
        hours = hours % 12;
        hours = hours || 12; // if hours is 0, set it to 12

        const timestamp = `${year}-${month}-${day}, ${hours}:${minutes}:${seconds} ${ampm}`;

        // Prepare the data to append to Form Responses, including Image Link
        const formValues = [
            [
                timestamp,        // K: Timestamp
                itemCode,         // L: Item Code
                requestQty,       // M: Receipt Qty
                location,         // N: Location
                description,      // O: Description
                returnableItem,   // P: Returnable Item
                receivedBy,       // Q: Received by
                notes || '',      // R: Notes/Comments
                imageLink         // S: Image Link
            ]
        ];

        // Fetch existing data in K:S to find the last row with data
        const receiptRange = 'Form Responses!K3:S';
        const receiptResponse = await req.sheetsClient.spreadsheets.values.get({
            spreadsheetId: req.sheetId,
            range: receiptRange
        });

        const receiptRows = receiptResponse.data.values || [];
        console.log('Receipt Rows:', receiptRows);

        // Headers are in K2:S2, data starts from K3:S3
        let lastRow = 2; // headers are in row 2
        for (let i = 0; i < receiptRows.length; i++) {
            const row = receiptRows[i];
            if (row && row.some(cell => cell.trim() !== '')) {
                lastRow = i + 3; // +3 because rows are 1-indexed and data starts at row3
            }
        }
        const nextRow = lastRow + 1;
        console.log(`Last Row with data: ${lastRow}`);
        console.log(`Next Row to append: ${nextRow}`);

        // Define the range to update
        const updateRange = `Form Responses!K${nextRow}:S${nextRow}`;
        console.log(`Update Range: ${updateRange}`);

        // Update the range with formValues
        const updateResponse = await req.sheetsClient.spreadsheets.values.update({
            spreadsheetId: req.sheetId,
            range: updateRange,
            valueInputOption: 'USER_ENTERED',
            resource: { values: formValues }
        });

        // Log the update response to verify
        console.log('Update Response:', updateResponse.data);

        // Update the inventory quantity
        const qtyColumnLetter = getColumnLetter(colQty);
        const inventoryUpdateRange = `Inventory!${qtyColumnLetter}${targetRowIndex + 1}`; // Correct row number

        await req.sheetsClient.spreadsheets.values.update({
            spreadsheetId: req.sheetId,
            range: inventoryUpdateRange,
            valueInputOption: 'USER_ENTERED',
            resource: { values: [[newQty]] }
        });

        // Respond with success and receipt details, including Image Link
        return res.status(200).json({
            message: 'Inventory received successfully',
            data: {
                Timestamp: timestamp,
                'Item Code': itemCode,
                'Receipt Qty': requestQty,
                Location: location,
                Description: description,
                'Returnable Item': returnableItem,
                'Received by': receivedBy,
                'Notes/Comments': notes || '',
                'Image Link': imageLink
            }
        });
    } catch (err) {
        console.error('Error receiving inventory:', err);
        return res.status(500).json({ error: 'Failed to receive inventory' });
    }
};

async function findFirstBlankRowTtoAA(sheetsClient, spreadsheetId) {
    try {
        // Define the range to fetch (starting from row 2 to skip headers)
        const range = 'Form Responses!T2:AA';
        const response = await sheetsClient.spreadsheets.values.get({
            spreadsheetId,
            range,
            majorDimension: 'ROWS',
        });

        const rows = response.data.values || [];
        console.log(`Total rows fetched in T:AA: ${rows.length}`);

        for (let i = 0; i < rows.length; i++) {
            const row = rows[i];
            // Ensure the row has at least 8 cells (T:AA)
            const cells = row.length >= 8 ? row.slice(0, 8) : row.concat(Array(8 - row.length).fill(''));
            // Check if all cells in T:AA are blank
            const isBlank = cells.every(cell => !cell || cell.toString().trim() === '');
            if (isBlank) {
                const blankRowNumber = i + 2; // Adding 2: 1 for zero-based index and 1 for header row
                console.log(`First blank row found at T:AA row ${blankRowNumber}`);
                return blankRowNumber;
            }
        }

        // If no blank row is found within existing data, append to the next row
        const blankRow = rows.length + 2; // Adding 2: 1 for zero-based index and 1 for header row
        console.log(`No blank row found in existing data. Appending to row ${blankRow}`);
        return blankRow;
    } catch (error) {
        console.error('Error finding first blank row in T:AA:', error);
        throw new Error('Failed to find the first blank row in T:AA');
    }
}

module.exports.addNewItemWithoutCode = async function (req, res) {
    try {
        // Extract fields from the request body
        const {
            receiptQty,
            location,
            description,
            returnableItem,
            receivedBy,
            notes
        } = req.body;

        // Validate required fields
        if (!receiptQty || !location || !description || !returnableItem || !receivedBy) {
            return res.status(400).json({ error: 'Missing required fields' });
        }

        // Handle image upload if a file is provided
        let imageLink = '';
        if (req.file) {
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

            // Remove the file from the server after upload
            fs.unlinkSync(filePath);

            imageLink = imageResponse.data.webViewLink;
        }

        // Construct the timestamp in the desired format: YYYY-MM-DD, HH:MM:SS AM/PM
        const now = new Date();
        const year = now.getFullYear();
        const month = String(now.getMonth() + 1).padStart(2, '0');
        const day = String(now.getDate()).padStart(2, '0');

        let hours = now.getHours();
        const minutes = String(now.getMinutes()).padStart(2, '0');
        const seconds = String(now.getSeconds()).padStart(2, '0');

        const ampm = hours >= 12 ? 'PM' : 'AM';
        hours = hours % 12;
        hours = hours || 12; // if hours is 0, set it to 12

        const timestamp = `${year}-${month}-${day}, ${hours}:${minutes}:${seconds} ${ampm}`;

        // Prepare the data to insert into Form Responses!T:AA
        const formValues = [[
            timestamp,        // T: Timestamp
            receiptQty,       // U: Receipt Qty
            location,         // V: Location
            description,      // W: Description
            returnableItem,   // X: Returnable Item
            receivedBy,       // Y: Received by
            notes || '',      // Z: Notes/Comments
            imageLink         // AA: Image Link
        ]];

        // Find the first blank row in 'Form Responses!T:AA'
        const blankRow = await findFirstBlankRowTtoAA(req.sheetsClient, req.sheetId);
        console.log(`Blank Row Found in Form Responses!T:AA: ${blankRow}`);

        // Define the range to update in 'Form Responses!T:AA'
        const updateRange = `Form Responses!T${blankRow}:AA${blankRow}`;
        console.log(`Update Range: ${updateRange}`);

        // Insert the data into the identified row
        await req.sheetsClient.spreadsheets.values.update({
            spreadsheetId: req.sheetId,
            range: updateRange,
            valueInputOption: 'USER_ENTERED',
            resource: { values: formValues }
        });

        return res.status(201).json({
            message: 'New inventory item added successfully',
            data: {
                Timestamp: timestamp,
                'Receipt Qty': receiptQty,
                Location: location,
                Description: description,
                'Returnable Item': returnableItem,
                'Received by': receivedBy,
                'Notes/Comments': notes || '',
                'Image Link': imageLink
            }
        });
    } catch (err) {
        console.error('Error adding new inventory item:', err);
        return res.status(500).json({ error: 'Failed to add new inventory item' });
    }
};