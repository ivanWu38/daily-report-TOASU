import express from 'express';
import path from 'path';
import ejs from 'ejs';
import bodyParser from 'body-parser';
import { processFrontEndData } from './services/dataProcessor.js'; // Import function
import { createExcelReport } from './routes/excel.js';

const app = express();
let frontEndData = [];

// Set EJS as the view engine
app.set('view engine', 'ejs');

// Middleware to parse form data
app.use(bodyParser.urlencoded({ extended: true }));

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

    // Filter out days where onTime or offTime is empty
    const filteredDays = rawFrontEndData.days.filter(day => day.onTime !== "" && day.offTime !== "");

    const processedData = { 
        month: rawFrontEndData.month, 
        days: filteredDays 
    };

    console.log(processedData); // Logs filtered data
    res.json(processedData); // Send only valid records
    frontEndData = processedData;

    console.log(frontEndData.days.length);

});







// Start the server
const PORT = 3000;
app.listen(PORT, () => {
    console.log(`Server is running on http://localhost:${PORT}`);
});

