import express from 'express';
import path from 'path';
import { fileURLToPath } from 'url';
import ejs from 'ejs';
import bodyParser from 'body-parser';
import fs from 'fs';
import { processFrontEndData } from './services/dataProcessor.js'; // Import function
import { createExcelReports } from './routes/report.js';

const app = express();
let frontEndData = [];

// Get the directory name
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// Set EJS as the view engine
app.set('view engine', 'ejs');

// Middleware to parse form data
app.use(bodyParser.urlencoded({ extended: true }));
app.use(express.static('public'));

// Function to get the number of days in a given month
function getDaysInMonth(year, month) {
    return new Date(year, month, 0).getDate(); // Last day of the month
}

// Route for rendering the form
app.get('/', (req, res) => {
    const today = new Date();
    let year = today.getFullYear();
    let month = today.getMonth() + 1; // Get current month (1-12)

    // If user selects a different month, update it
    if (req.query.month) {
        [year, month] = req.query.month.split('-').map(Number);
    }

    const totalDays = getDaysInMonth(year, month);
    const currentMonth = `${year}-${('0' + month).slice(-2)}`; // Format YYYY-MM

    res.render('index', { currentMonth, totalDays });
});

// Route for handling form submission
app.post('/excel', (req, res) => {
    const rawFrontEndData = processFrontEndData(req.body);
    const directory = req.body.directory || __dirname; // Default to current directory if not specified

    // Ensure the directory exists
    if (!fs.existsSync(directory)) {
        return res.status(400).send("指定されたディレクトリが存在しません。");
    }

    // Filter out days where onTime or offTime is empty
    const filteredDays = rawFrontEndData.days.filter(day => day.onTime !== "" && day.offTime !== "");

    const processedData = { 
        month: rawFrontEndData.month, 
        days: filteredDays 
    };

    console.log(processedData); // Logs filtered data
    frontEndData = processedData;

    if (frontEndData.days.length > 0) {
        const filePath = path.join(directory, 'Work_Report.xlsx');
        createExcelReports(frontEndData.days, filePath);
        res.send(`日報が生成されました。ファイルは ${filePath} に保存されました。`);
    } else {
        res.send("データがありません");
    }
});

// Start the server
const PORT = 3000;
app.listen(PORT, () => {
    console.log(`Server is running on http://localhost:${PORT}`);
});

