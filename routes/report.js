import ExcelJS from 'exceljs';

/**
 * Function to create multiple daily reports in different sheets of the same Excel file.
 * 
 * @param {Array} reports - Array of report objects, each containing { date, onTime, offTime }.
 */
export async function createExcelReports(reports, filePathServer) {
    const workbook = new ExcelJS.Workbook(); // Create a single workbook

    for (const report of reports) {
        const { date, onTime, offTime } = report;
        const worksheet = workbook.addWorksheet(date); // Create a new worksheet (tab) for each date

        // 1️⃣ Set Column Widths
        worksheet.getColumn(1).width = 19.17; // Column A
        for (let col = 2; col <= 21; col++) {
            worksheet.getColumn(col).width = 3.83;
        }

        // 2️⃣ Set Row Heights
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

        // 3️⃣ Add Metadata Section (業務日, 所属, 氏名)
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
        worksheet.getCell('Q2').value = date; // Dynamic Date
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
      

        const workList = [
            "メール確認","競合調査","企画立案","ブレスト用の情報収集",
            "資料整理","新規企画の方向性整理","ベンチマーク企業の事例研究",
            "定例会","MTG","資料修正","打ち合わせ","事例調査"
        ]

        const workList2 = [
            "MTG","資料修正","打ち合わせ","事例調査","","","","",
            "","","","","","","","",
            "","","","","","","","",
            "","","","","","","","",
            "","","","","","","","",
            "","","","","","","","",
            "","","","","","","","",
            "","","","","","","","",
            "","","","","","","","",
            "","","","","","","","",
        ]


        // 5️⃣ Add Work Schedule
        const schedule = [
            { time: `${onTime}-12:00`, tasks: [
                workList[Math.floor(Math.random() * workList.length)],
                workList[Math.floor(Math.random() * workList.length)], 
                workList2[Math.floor(Math.random() * workList2.length)]] },

            { time: '12:00-13:00', tasks: ['休憩'] },

            { time: `13:00-${offTime}`, tasks: [
                workList[Math.floor(Math.random() * workList.length)],
                workList[Math.floor(Math.random() * workList.length)], 
                workList2[Math.floor(Math.random() * workList2.length)]]
            }
        ];

        let rowIndex = 8;
        schedule.forEach((entry) => {
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

        // 6️⃣ Set Background Colors
        for (let col = 1; col <= 18; col++) {
            worksheet.getCell(11, col).fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFD3D3D3' } // Light Grey
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

    }
    

    // 7️⃣ Save the file after all sheets are added

    await workbook.xlsx.writeFile(filePathServer);
    console.log(filePathServer);
    console.log(`Excel file created successfully with ${reports.length} sheets!`);
}