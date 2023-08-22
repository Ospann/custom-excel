import ExcelJS from 'exceljs';

const generateStyledExcel = async (data) => {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Sheet1');

    // Skip first row
    worksheet.addRow([]);

    // Add "Счет-фактура" row
    const invoiceRow = worksheet.addRow(['']); // Add an empty cell
    worksheet.mergeCells('B2:G2'); // Merge cells B2 to G2 (6 cells)
    invoiceRow.getCell(2).value = 'Образец платежного поручения'; // Set value to 'Счет-фактура'
    invoiceRow.font = { bold: true, size: 16 }; // Bold font, larger size
    invoiceRow.alignment = { vertical: 'middle', horizontal: 'left' };
    // Skip one more row
    worksheet.addRow([]);
    // Add columns with no header text to match your requirement
    worksheet.columns = [{}, ...Object.keys(data[0]).map(() => ({}))];

    // Apply styles to cells (table)
    const headerRow = worksheet.addRow(['', ...Object.keys(data[0])]); // Add an empty cell in the beginning
    headerRow.eachCell((cell) => {
        cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFFFFFFF' } // White background
        };
        cell.font = { color: { argb: 'FF000000' }, bold: true }; // Black text
        cell.border = {
            top: { style: 'thin', color: { argb: 'FFFF0000' } }, // Red top border
            bottom: { style: 'thin', color: { argb: 'FFFF0000' } }, // Red bottom border
            left: { style: 'thin', color: { argb: 'FFFF0000' } }, // Red left border
            right: { style: 'thin', color: { argb: 'FFFF0000' } }, // Red right border
        };
        cell.alignment = { horizontal: 'center' }; // Center text horizontally
    });

    data.forEach((rowData) => {
        const row = worksheet.addRow(['', ...Object.values(rowData)]);
        row.eachCell((cell) => {
            cell.font = { color: { argb: 'FF000000' } }; // Black text
            cell.border = {
                top: { style: 'thin', color: { argb: 'FFFF0000' } }, // Red top border
                bottom: { style: 'thin', color: { argb: 'FFFF0000' } }, // Red bottom border
                left: { style: 'thin', color: { argb: 'FFFF0000' } }, // Red left border
                right: { style: 'thin', color: { argb: 'FFFF0000' } }, // Red right border
            };
            cell.alignment = { horizontal: 'center' }; // Center text horizontally
        });
    });

    // Add left and right indentation
    worksheet.getColumn(1).width = 10; // Set width of the first column to create left indentation
    worksheet.getColumn(headerRow.cellCount + 1).width = 10; // Set width of the last column to create right indentation

    const excelBuffer = await workbook.xlsx.writeBuffer();
    return excelBuffer;
};

export default generateStyledExcel;
