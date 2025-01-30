import ExcelJS from 'exceljs';

export async function createExcelReport(date, onTime, offTime) {
    // 1️⃣ Create a new workbook & worksheet
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet(date);

    // -------------------------------------------------------
    // EXAMPLE: Set specific column widths (instead of all).
    // 
    // Columns are 1-based: 
    //    Column 1 = A, 2 = B, 3 = C, ..., 15 = O, etc.
    // 
    // You can also remove or modify these lines if you want 
    // different widths for each column.
    // -------------------------------------------------------
    worksheet.getColumn(1).width = 19.17; // Column A
    for (let col = 2; col <= 21; col++) {
        worksheet.getColumn(col).width = 3.83;
    }

    // -------------------------------------------------------
    // EXAMPLE: Set row height for specific rows (instead of all).
    // 
    // Rows are 1-based: 
    //    Row 1 = first row, Row 2 = second row, etc.
    // 
    // For instance, here we set row 6 to a height of 25 points.
    // You can set different heights for different rows similarly.
    // -------------------------------------------------------
    // worksheet.getRow(6).height = 24.75;   // Make Row 6 taller
    for (let row = 2; row <= 6; row++) {
        worksheet.getRow(row).height = 24.75;
    }

    worksheet.getRow(7).height = 9; 
    
    for (let row = 8; row <= 10; row++) {
        worksheet.getRow(row).height = 19.5;
    }

    worksheet.getRow(11).height = 60; 

    for (let row = 12; row <= 14; row++) {
        worksheet.getRow(row).height = 19.5;
    }

    // 2️⃣ Add Metadata Section (業務日, 所属, 氏名)
    worksheet.mergeCells('O2:P2');
    worksheet.getCell('O2').value = '業務日';
    worksheet.getCell('O2').alignment = { horizontal: 'center', vertical: 'middle' };
    worksheet.getCell('O2').border = {
        top:    { style: 'thin' },
        left:   { style: 'thin' },
        right:  { style: 'thin' },
        bottom: { style: 'thin' }
    };

    worksheet.mergeCells('Q2:U2');
    worksheet.getCell('Q2').value = '2024/10/15'; // Dynamic Date
    worksheet.getCell('Q2').alignment = { horizontal: 'center', vertical: 'middle' };
    worksheet.getCell('Q2').border = {
        top:    { style: 'thin' },
        left:   { style: 'thin' },
        right:  { style: 'thin' },
        bottom: { style: 'thin' }
    };

    worksheet.mergeCells('O3:P3');
    worksheet.getCell('O3').value = '所属';
    worksheet.getCell('O3').alignment = { horizontal: 'center', vertical: 'middle' };
    worksheet.getCell('O3').border = {
        top:    { style: 'thin' },
        left:   { style: 'thin' },
        right:  { style: 'thin' },
        bottom: { style: 'thin' }
    };

    worksheet.mergeCells('Q3:U3');
    worksheet.getCell('Q3').value = '事業開発部'; // Dynamic Department
    worksheet.getCell('Q3').alignment = { horizontal: 'center', vertical: 'middle' };
    worksheet.getCell('Q3').border = {
        top:    { style: 'thin' },
        left:   { style: 'thin' },
        right:  { style: 'thin' },
        bottom: { style: 'thin' }
    };

    worksheet.mergeCells('O4:P4');
    worksheet.getCell('O4').value = '氏名';
    worksheet.getCell('O4').alignment = { horizontal: 'center', vertical: 'middle' };
    worksheet.getCell('O4').border = {
        top:    { style: 'thin' },
        left:   { style: 'thin' },
        right:  { style: 'thin' },
        bottom: { style: 'thin' }
    };

    worksheet.mergeCells('Q4:U4');
    worksheet.getCell('Q4').value = '呉育平'; // Dynamic Employee Name
    worksheet.getCell('Q4').alignment = { horizontal: 'center', vertical: 'middle' };
    worksheet.getCell('Q4').border = {
        top:    { style: 'thin' },
        left:   { style: 'thin' },
        right:  { style: 'thin' },
        bottom: { style: 'thin' }
    };

    // 3️⃣ Create Header Row
    worksheet.mergeCells('A6');
    worksheet.getCell('A6').value = '時刻';
    worksheet.getCell('A6').font = { bold: true };
    worksheet.getCell('A6').alignment = { horizontal: 'center', vertical: 'middle' };
    worksheet.getCell('A6').border = {
        top:    { style: 'thin' },
        left:   { style: 'thin' },
        right:  { style: 'thin' },
        bottom: { style: 'thin' }
    };

    worksheet.mergeCells('B6:Q6');
    worksheet.getCell('B6').value = '業務内容';
    worksheet.getCell('B6').font = { bold: true };
    worksheet.getCell('B6').alignment = { horizontal: 'center', vertical: 'middle' };
    worksheet.getCell('B6').border = {
        top:    { style: 'thin' },
        left:   { style: 'thin' },
        right:  { style: 'thin' },
        bottom: { style: 'thin' }
    };

    worksheet.mergeCells('R6:U6');
    worksheet.getCell('R6').value = '備考';
    worksheet.getCell('R6').font = { bold: true };
    worksheet.getCell('R6').alignment = { horizontal: 'center', vertical: 'middle' };
    worksheet.getCell('R6').border = {
        top:    { style: 'thin' },
        left:   { style: 'thin' },
        right:  { style: 'thin' },
        bottom: { style: 'thin' }
    };

    // 4️⃣ Add Work Schedule
    const schedule = [
        { time: onTime + '-12:00', tasks: ['AIオープン研修企画',"nothing","gaming"] },
        { time: '12:00-13:00', tasks: ['休憩'] },
        { time: '13:00-'+offTime, tasks: ['AIオープン研修企画', '社内研修を受講', '部署定例会'] }
    ];

    let rowIndex = 8;
    schedule.forEach((entry) => {
        // Merge cells for multi-row time slots
        worksheet.mergeCells(`A${rowIndex}:A${rowIndex + entry.tasks.length - 1}`);
        worksheet.getCell(`A${rowIndex}`).value = entry.time;
        worksheet.getCell(`A${rowIndex}`).alignment = { horizontal: 'center', vertical: 'middle' };
        worksheet.getCell(`A${rowIndex}`).border = {
            top:    { style: 'thin' },
            left:   { style: 'thin' },
            right:  { style: 'thin' },
            bottom: { style: 'thin' }
        };

        entry.tasks.forEach((task, i) => {
            worksheet.mergeCells(`B${rowIndex + i}:Q${rowIndex + i}`);
            worksheet.getCell(`B${rowIndex + i}`).value = task;
            worksheet.getCell(`B${rowIndex + i}`).border = {
                top:    { style: 'thin' },
                left:   { style: 'thin' },
                right:  { style: 'thin' },
                bottom: { style: 'thin' }
            };

            worksheet.mergeCells(`R${rowIndex + i}:U${rowIndex + i}`);
            worksheet.getCell(`R${rowIndex + i}`).border = {
                top:    { style: 'thin' },
                left:   { style: 'thin' },
                right:  { style: 'thin' },
                bottom: { style: 'thin' }
            };
        });

        rowIndex += entry.tasks.length;
    });

    // -------------------------------------------------------
    // EXAMPLE: Set a completely custom border style on a single cell.
    // 
    
    // Let's say you want a thick red border around cell B8:
    // -------------------------------------------------------
    // worksheet.getCell('B8').border = {
    //     top:    { style: 'thick', color: { argb: 'FFFF0000' } },
    //     left:   { style: 'thick', color: { argb: 'FFFF0000' } },
    //     // bottom: { style: 'thick', color: { argb: 'FFFF0000' } },
    //     right:  { style: 'thick', color: { argb: 'FFFF0000' } }
    // };

    
    // Set background color of cells from A11 to R11 to light grey
    for (let col = 1; col <= 18; col++) { // Columns A to R are 1 to 18
        const cell = worksheet.getCell(11, col);
        cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFD3D3D3' } // Light grey color
        };
    }

    // Set text alignment for cells from B8 to B14
    for (let row = 8; row <= 14; row++) {
        const cell = worksheet.getCell(row, 2); // Column B is 2
        cell.alignment = {
            horizontal: 'center',
            vertical: 'middle'
        };
    }

    // 5️⃣ Save the file
    await workbook.xlsx.writeFile('Work_Report.xlsx');
    console.log('Excel file created successfully!');
}

// Run function