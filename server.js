import ExcelJS from 'exceljs';

async function createExcelReport() {
    // 1️⃣ Create a new workbook & worksheet
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('10月15日');

    // 2️⃣ Set Column Widths (Based on your template)
    worksheet.getColumn(1).width = 15;  // 時刻 (Time)
    worksheet.getColumn(2).width = 40;  // 業務内容 (Work Description)
    worksheet.getColumn(3).width = 20;  // 備考 (Remarks)

    // 3️⃣ Add Metadata Section (業務日, 所属, 氏名)
    worksheet.mergeCells('O2:P2');
    worksheet.getCell('O2').value = '業務日';
    worksheet.getCell('O2').alignment = { horizontal: 'center', vertical: 'middle' };
    worksheet.getCell('O2').border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' }, bottom: { style: 'thin' } };

    worksheet.mergeCells('Q2:U2');
    worksheet.getCell('Q2').value = '2024/10/15'; // Dynamic Date
    worksheet.getCell('Q2').alignment = { horizontal: 'center', vertical: 'middle' };
    worksheet.getCell('Q2').border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' }, bottom: { style: 'thin' } };

    worksheet.mergeCells('O3:P3');
    worksheet.getCell('O3').value = '所属';
    worksheet.getCell('O3').alignment = { horizontal: 'center', vertical: 'middle' };
    worksheet.getCell('O3').border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' }, bottom: { style: 'thin' } };

    worksheet.mergeCells('Q3:U3');
    worksheet.getCell('Q3').value = '事業開発部'; // Dynamic Department
    worksheet.getCell('Q3').alignment = { horizontal: 'center', vertical: 'middle' };
    worksheet.getCell('Q3').border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' }, bottom: { style: 'thin' } };

    worksheet.mergeCells('O4:P4');
    worksheet.getCell('O4').value = '氏名';
    worksheet.getCell('O4').alignment = { horizontal: 'center', vertical: 'middle' };
    worksheet.getCell('O4').border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' }, bottom: { style: 'thin' } };

    worksheet.mergeCells('Q4:U4');
    worksheet.getCell('Q4').value = '呉昊平'; // Dynamic Employee Name
    worksheet.getCell('Q4').alignment = { horizontal: 'center', vertical: 'middle' };
    worksheet.getCell('Q4').border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' }, bottom: { style: 'thin' } };

    // 4️⃣ Create Header Row
    worksheet.mergeCells('A6:A7');
    worksheet.getCell('A6').value = '時刻';
    worksheet.getCell('A6').font = { bold: true };
    worksheet.getCell('A6').alignment = { horizontal: 'center', vertical: 'middle' };
    worksheet.getCell('A6').border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' }, bottom: { style: 'thin' } };

    worksheet.mergeCells('B6:Q6');
    worksheet.getCell('B6').value = '業務内容';
    worksheet.getCell('B6').font = { bold: true };
    worksheet.getCell('B6').alignment = { horizontal: 'center', vertical: 'middle' };
    worksheet.getCell('B6').border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' }, bottom: { style: 'thin' } };

    worksheet.mergeCells('R6:U6');
    worksheet.getCell('R6').value = '備考';
    worksheet.getCell('R6').font = { bold: true };
    worksheet.getCell('R6').alignment = { horizontal: 'center', vertical: 'middle' };
    worksheet.getCell('R6').border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' }, bottom: { style: 'thin' } };

    // 5️⃣ Add Work Schedule
    const schedule = [
        { time: '8:00-12:00', tasks: ['AIオープン研修企画'] },
        { time: '12:00-13:00', tasks: ['休憩'] },
        { time: '13:00-18:30', tasks: ['AIオープン研修企画', '社内研修を受講', '部署定例会'] }
    ];

    let rowIndex = 8;
    schedule.forEach((entry) => {
        // Merge cells for multi-row time slots
        worksheet.mergeCells(`A${rowIndex}:A${rowIndex + entry.tasks.length - 1}`);
        worksheet.getCell(`A${rowIndex}`).value = entry.time;
        worksheet.getCell(`A${rowIndex}`).alignment = { horizontal: 'center', vertical: 'middle' };
        worksheet.getCell(`A${rowIndex}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' }, bottom: { style: 'thin' } };

        entry.tasks.forEach((task, i) => {
            worksheet.mergeCells(`B${rowIndex + i}:Q${rowIndex + i}`);
            worksheet.getCell(`B${rowIndex + i}`).value = task;
            worksheet.getCell(`B${rowIndex + i}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' }, bottom: { style: 'thin' } };

            worksheet.mergeCells(`R${rowIndex + i}:U${rowIndex + i}`);
            worksheet.getCell(`R${rowIndex + i}`).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' }, bottom: { style: 'thin' } };
        });

        rowIndex += entry.tasks.length;
    });

    // 6️⃣ Save the file
    await workbook.xlsx.writeFile('Work_Report.xlsx');
    console.log('Excel file created successfully!');
}

// Run function
createExcelReport();