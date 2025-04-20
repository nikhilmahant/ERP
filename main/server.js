const express = require('express');
const cors = require('cors');
const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');

const app = express();
const port = 4000;

// Middleware
app.use(cors());
app.use(express.json());
app.use(express.static('public'));

// Excel file paths
const EXCEL_DIR = path.join(__dirname, 'excel_data');
if (!fs.existsSync(EXCEL_DIR)) {
    fs.mkdirSync(EXCEL_DIR, { recursive: true });
}

// Get today's Excel file path
function getTodayExcelPath() {
    const today = new Date();
    const dateStr = today.toISOString().split('T')[0];
    return path.join(EXCEL_DIR, `invoice_${dateStr}.xlsx`);
}

// Initialize Excel workbook with headers
async function initializeWorkbook() {
    const workbook = new ExcelJS.Workbook();
    
    // Initialize Patti sheet
    const pattiSheet = workbook.addWorksheet('Patti');
    pattiSheet.columns = [
        { header: 'Date', key: 'date', width: 12 },
        { header: 'Customer Name', key: 'customerName', width: 20 },
        { header: 'Item', key: 'item', width: 15 },
        { header: 'Packet', key: 'packet', width: 10 },
        { header: 'Quantity', key: 'quantity', width: 10 },
        { header: 'Rate', key: 'rate', width: 10 },
        { header: 'Hamali', key: 'hamali', width: 10 },
        { header: 'Amount', key: 'amount', width: 12 }
    ];

    // Initialize Kata sheet
    const kataSheet = workbook.addWorksheet('Kata');
    kataSheet.columns = [
        { header: 'Date', key: 'date', width: 12 },
        { header: 'Customer Name', key: 'customerName', width: 20 },
        { header: 'Item', key: 'item', width: 15 },
        { header: 'Net Weight', key: 'netWeight', width: 12 },
        { header: 'Less %', key: 'lessPercent', width: 10 },
        { header: 'Final Weight', key: 'finalWeight', width: 12 },
        { header: 'Rate', key: 'rate', width: 10 },
        { header: 'Packets', key: 'packets', width: 10 },
        { header: 'Hamali Rate', key: 'hamaliRate', width: 12 },
        { header: 'Amount', key: 'amount', width: 12 },
        { header: 'Kata Amount', key: 'kataAmount', width: 12 },
        { header: 'Total', key: 'total', width: 12 }
    ];

    // Initialize Barthe sheet
    const bartheSheet = workbook.addWorksheet('Barthe');
    bartheSheet.columns = [
        { header: 'Date', key: 'date', width: 12 },
        { header: 'Customer Name', key: 'customerName', width: 20 },
        { header: 'Item', key: 'item', width: 15 },
        { header: 'Packet', key: 'packet', width: 10 },
        { header: 'Weight', key: 'weight', width: 10 },
        { header: 'Adjustment', key: 'adjustment', width: 10 },
        { header: 'Quantity', key: 'quantity', width: 12 },
        { header: 'Rate', key: 'rate', width: 10 },
        { header: 'Hamali Rate', key: 'hamaliRate', width: 12 },
        { header: 'Amount', key: 'amount', width: 12 }
    ];

    return workbook;
}

// Save invoice data to Excel
async function saveToExcel(page, invoice) {
    const excelPath = getTodayExcelPath();
    let workbook;

    try {
        console.log(`Saving invoice to ${excelPath}`);
        console.log(`Page: ${page}`);
        console.log('Invoice data:', JSON.stringify(invoice, null, 2));
        
        // Try to load existing workbook
        if (fs.existsSync(excelPath)) {
            console.log('Loading existing workbook');
            workbook = new ExcelJS.Workbook();
            await workbook.xlsx.readFile(excelPath);
        } else {
            console.log('Creating new workbook');
            workbook = await initializeWorkbook();
        }

        const sheet = workbook.getWorksheet(page);
        if (!sheet) {
            throw new Error(`Sheet ${page} not found`);
        }

        const date = new Date(invoice.date).toLocaleDateString('en-IN');
        console.log(`Processing invoice for date: ${date}`);
        
        // Validate invoice data
        if (!Array.isArray(invoice.items)) {
            throw new Error('Invoice items must be an array');
        }

        console.log(`Number of items to process: ${invoice.items.length}`);

        if (!invoice.customerName) {
            console.warn('Warning: Customer name is empty');
        }

        // Add rows based on invoice type
        for (let i = 0; i < invoice.items.length; i++) {
            const item = invoice.items[i];
            if (!item) continue;
            
            console.log(`Processing item ${i + 1}:`, item);
            
            try {
                let rowData = {
                    date,
                    customerName: invoice.customerName || '',
                    item: item.item || ''
                };

                if (page === 'Patti') {
                    Object.assign(rowData, {
                        packet: item.packet || 0,
                        quantity: item.quantity || 0,
                        rate: item.rate || 0,
                        hamali: item.hamali || 0,
                        amount: item.amount || 0
                    });
                } else if (page === 'Kata') {
                    Object.assign(rowData, {
                        netWeight: item.netWeight || 0,
                        lessPercent: item.lessPercent || 0,
                        finalWeight: item.finalWeight || 0,
                        rate: item.rate || 0,
                        packets: item.packets || 0,
                        hamaliRate: item.hamaliRate || 0,
                        amount: item.amount || 0,
                        kataAmount: i === 0 ? (invoice.additionalAmount || 0) : 0,
                        total: i === 0 ? (invoice.grandTotal || 0) : 0
                    });
                } else if (page === 'Barthe') {
                    Object.assign(rowData, {
                        packet: item.packet || 0,
                        weight: item.weight || 0,
                        adjustment: item.adjustment || 0,
                        quantity: item.quantity || 0,
                        rate: item.rate || 0,
                        hamaliRate: item.hamaliRate || 0,
                        amount: item.amount || 0
                    });
                }

                const row = sheet.addRow(rowData);

                // Apply number format to numeric cells
                sheet.columns.forEach(col => {
                    if (col.key && typeof rowData[col.key] === 'number') {
                        const cell = row.getCell(col.key);
                        cell.numFmt = '#,##0.00';
                    }
                });

                console.log(`Added row for item ${i + 1}`);
            } catch (itemError) {
                console.error(`Error processing item ${i + 1}:`, itemError);
                throw new Error(`Failed to process item ${i + 1}: ${itemError.message}`);
            }
        }

        console.log('Saving workbook');
        await workbook.xlsx.writeFile(excelPath);
        console.log('Save completed successfully');
        return { success: true };
    } catch (error) {
        console.error('Error saving to Excel:', error);
        return { 
            error: error.message,
            details: error.stack
        };
    }
}

// API Endpoints
app.post('/api/save', async (req, res) => {
    try {
        console.log('Received save request');
        const { page, invoice } = req.body;
        
        // Validate request data
        if (!page || !invoice) {
            console.error('Missing required data:', { page, hasInvoice: !!invoice });
            return res.status(400).json({ 
                error: 'Missing required data',
                details: {
                    page: !page ? 'Missing page' : undefined,
                    invoice: !invoice ? 'Missing invoice data' : undefined
                }
            });
        }

        // Validate page value
        if (!['Patti', 'Kata', 'Barthe'].includes(page)) {
            console.error('Invalid page value:', page);
            return res.status(400).json({ 
                error: 'Invalid page value. Must be one of: Patti, Kata, Barthe' 
            });
        }

        const result = await saveToExcel(page, invoice);
        if (result.error) {
            console.error('Save to Excel failed:', result.error);
            return res.status(500).json(result);
        }
        res.json(result);
    } catch (error) {
        console.error('API Error:', error);
        res.status(500).json({ 
            error: error.message,
            details: error.stack
        });
    }
});

// Instead of app.listen, export the app
module.exports = app.listen(port, () => {
    console.log(`Server running at http://localhost:${port}`);
}); 