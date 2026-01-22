const express = require('express');
const multer = require('multer');
const cors = require('cors');
const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');

const app = express();
const upload = multer({ dest: 'uploads/' });

// Enable CORS and JSON parsing
app.use(cors());
app.use(express.json());
app.use(express.static('public')); // Serve frontend files

// SIE File Generator Class
class SIEBuilder {
    constructor(sieType = 'sie4', encoding = 'ISO-8859-1', companyName = "My Company") {
        this.lines = [];
        this.sieType = sieType;
        this.encoding = encoding;
        this.companyName = companyName;
        this.addHeader();
    }
    
    addHeader() {
        const today = new Date().toISOString().split('T')[0];
        const year = new Date().getFullYear();
        
        this.lines.push('#FLAGGA 0');
        this.lines.push('#PROGRAM "SIE Converter" 1.0');
        this.lines.push('#FORMAT PC8');
        this.lines.push(`#GEN ${today} "System"`);
        this.lines.push(`#SIETYP ${this.sieType.replace('sie', '')}`);
        this.lines.push(`#FNAMN "${this.companyName}"`);
        this.lines.push(`#RAR 0 ${year}0101 ${year}1231`);
        this.lines.push('#VALUTA SEK');
    }
    
    addAccount(accountNumber, accountName) {
        this.lines.push(`#KONTO ${accountNumber} "${accountName}"`);
    }
    
    addTransaction(verNum, date, accountNumber, amount, description = '') {
        this.lines.push(`#VER ${verNum} ${date} "${description}"`);
        this.lines.push('{');
        this.lines.push(`#TRANS ${accountNumber} {} ${amount.toFixed(2)}`);
        this.lines.push('}');
    }
    
    build() {
        this.lines.push('#END');
        return this.lines.join('\n');
    }
}

// Helper function to parse CSV
function parseCSV(content) {
    const lines = content.split('\n');
    const result = [];
    
    lines.forEach(line => {
        if (line.trim() === '') return;
        
        // Handle both comma and semicolon separated values
        const delimiter = line.includes(';') ? ';' : ',';
        const columns = line.split(delimiter);
        
        // Expected format: Date,Account,Amount,Description
        if (columns.length >= 3) {
            result.push({
                date: columns[0].trim(),
                account: columns[1].trim(),
                amount: parseFloat(columns[2].trim()),
                description: columns[3] ? columns[3].trim() : ''
            });
        }
    });
    
    return result;
}

// Convert CSV to SIE
function convertCSVToSIE(csvContent, sieType, companyName) {
    const transactions = parseCSV(csvContent);
    const sie = new SIEBuilder(sieType, 'ISO-8859-1', companyName);
    
    // Add sample chart of accounts (you should customize this)
    sie.addAccount('1910', 'Kassa');
    sie.addAccount('1930', 'Bank');
    sie.addAccount('3011', 'Försäljning');
    sie.addAccount('4010', 'Lokalhyra');
    
    // Add transactions
    transactions.forEach((transaction, index) => {
        sie.addTransaction(
            index + 1,
            transaction.date,
            transaction.account,
            transaction.amount,
            transaction.description
        );
    });
    
    return sie.build();
}

// Convert Excel to SIE
function convertExcelToSIE(filePath, sieType, companyName) {
    const workbook = XLSX.readFile(filePath);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const data = XLSX.utils.sheet_to_json(sheet);
    
    const sie = new SIEBuilder(sieType, 'ISO-8859-1', companyName);
    
    // Add sample accounts
    sie.addAccount('1910', 'Kassa');
    sie.addAccount('1930', 'Bank');
    sie.addAccount('3011', 'Försäljning');
    sie.addAccount('4010', 'Lokalhyra');
    
    // Process data - adjust based on your Excel structure
    data.forEach((row, index) => {
        if (row.Date && row.Account && row.Amount) {
            sie.addTransaction(
                index + 1,
                row.Date,
                row.Account,
                row.Amount,
                row.Description || ''
            );
        }
    });
    
    return sie.build();
}

// API Endpoint: Convert file
app.post('/api/convert', upload.single('file'), async (req, res) => {
    try {
        const { sieType, encoding, companyName = 'My Company' } = req.body;
        const file = req.file;
        
        if (!file) {
            return res.status(400).json({ error: 'Ingen fil har laddats upp' });
        }
        
        let sieContent;
        const fileExt = path.extname(file.originalname).toLowerCase();
        
        switch (fileExt) {
            case '.csv':
                const csvData = fs.readFileSync(file.path, 'utf8');
                sieContent = convertCSVToSIE(csvData, sieType, companyName);
                break;
                
            case '.xlsx':
            case '.xls':
                sieContent = convertExcelToSIE(file.path, sieType, companyName);
                break;
                
            default:
                throw new Error(`Filtypen ${fileExt} stöds inte. Använd CSV eller Excel.`);
        }
        
        // Clean up uploaded file
        fs.unlinkSync(file.path);
        
        // Return SIE file
        res.setHeader('Content-Type', 'text/plain; charset=' + encoding);
        res.setHeader('Content-Disposition', `attachment; filename="output_${sieType}.sie"`);
        res.send(sieContent);
        
    } catch (error) {
        console.error('Conversion error:', error);
        res.status(500).json({ error: error.message });
    }
});

// API Endpoint: Get SIE specifications
app.get('/api/sie-spec', (req, res) => {
    const specifications = {
        sie1: { 
            name: "SIE 1", 
            description: "Utvecklingsformat (ANSI)",
            encoding: "ISO-8859-1"
        },
        sie2: { 
            name: "SIE 2", 
            description: "Intern kontroll av bokföringsprogram",
            encoding: "ISO-8859-1"
        },
        sie3: { 
            name: "SIE 3", 
            description: "För överföring till revisionsprogram",
            encoding: "ISO-8859-1"
        },
        sie4: { 
            name: "SIE 4", 
            description: "För överföring mellan bokföringsprogram",
            encoding: "ISO-8859-1"
        },
        sie5: { 
            name: "SIE 5/EU", 
            description: "EU-kompatibelt format",
            encoding: "UTF-8"
        }
    };
    res.json(specifications);
});

// Serve main page
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// Start server
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`✅ SIE Converter is running!`);
    console.log(`📁 Open in browser: http://localhost:${PORT}`);
    console.log(`📤 Upload folder: ${path.join(__dirname, 'uploads')}`);
});
