const express = require('express');
const bodyParser = require('body-parser');
const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');
const tmp = require('tmp');

const app = express();
const port = 3000;

// Middleware to parse JSON bodies
app.use(bodyParser.json());

// Serve static files from the public directory
app.use(express.static('public'));

// File path for the Excel file
const filePath = path.join(__dirname, 'orders.xlsx');

// Function to read existing orders from the Excel file
function readOrdersFromFile() {
    if (fs.existsSync(filePath)) {
        const workbook = xlsx.readFile(filePath);
        const worksheet = workbook.Sheets['Orders'];
        return xlsx.utils.sheet_to_json(worksheet);
    }
    return [];
}

// Endpoint to handle order submissions
app.post('/api/orders', (req, res) => {
    const { roomNumber, items } = req.body;

    // Log the received order data
    console.log('Order received:');
    console.log('Room Number:', roomNumber);
    console.log('Items:', items);

    // Read existing orders
    let orders = readOrdersFromFile();

    // Add the new order to the orders array
    orders.push({
        RoomNumber: roomNumber,
        Items: items.map(item => `${item.name} (x${item.quantity})`).join(', '),
        Total: items.reduce((sum, item) => sum + item.price * item.quantity, 0).toFixed(2) + ' JD'
    });

    // Convert updated orders to a worksheet
    const worksheet = xlsx.utils.json_to_sheet(orders);

    // Create a new workbook and append the worksheet
    const workbook = xlsx.utils.book_new();
    xlsx.utils.book_append_sheet(workbook, worksheet, 'Orders');

    // Write to a temporary file first
    const tempFile = tmp.fileSync({ postfix: '.xlsx' });
    try {
        xlsx.writeFile(workbook, tempFile.name);

        // Replace the existing file with the new one
        fs.renameSync(tempFile.name, filePath);
    } catch (err) {
        console.error('Error writing Excel file:', err);
    } finally {
        tempFile.removeCallback(); // Clean up temporary file
    }

    res.json({ message: 'Order received successfully' });
});

// Start the server
app.listen(port, () => {
    console.log(`Server listening at http://localhost:${port}`);
});
