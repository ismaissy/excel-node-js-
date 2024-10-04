const ExcelJS = require('exceljs');
const fs = require('fs').promises;

//const ee = require("../json/wedemost_obj1.json")

const HEAD = {
    head: "HASAPLAŞYK-TÖLEG WEDOMOSTY",
    date: "2024-nji ýylyn Aprel aýy üçin",
    director: {
        label: "Direktor",
        value: "M.A.Kakageldiýewa"
    },
    accountant: {
        label: "Baş hasapçy",
        value: "A.Arazowa"
    }
}


const initialRow = 5;

const H1 = [
    { code: "2000", name: "№" },
    { code: "2001", name: "Işgär ID" },
    { code: "2002", name: "Ady Familiýasy" },
    { code: "2003", name: "Wezipe" },
    { code: "2004", name: "Asyl aýlygy" },
]

const H2 = [
    { code: "2005", name: "Işlan günleri" },
    { code: "2006", name: "Aýlyk" },
    { code: "2007", name: "Gijeki" },
    { code: "2008", name: "Baýramçylyk" },
    { code: "2009", name: "Utgişdyrma" },
]

const H4 = [
    { code: "2010", name: "Gaýtadan hasaplama" },
    { code: "2011", name: "Zyýanlyk" },
    { code: "2012", name: "Jemi hasaplanan" },
    { code: "2013", name: "Hak ujy" },
]

const H6 = [
    { code: "2014", name: "Jemi tutulan" },
    { code: "2015", name: "Eline" }
]


async function readAndProcessFile() {
    try {
        const data = await fs.readFile("./json/wedemost_obj1.json", 'utf8');
        const jsonData = JSON.parse(data);
        return jsonData;
    } catch (err) {
        console.error('Error read file:', err);
    }
}

readAndProcessFile().then((json) => {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Sheet1');

    const salaryPayments = json.salaryPayments;
    const salaryPaymentTotals = json.salaryPaymentTotals;

    const H3 = salaryPaymentTotals.addPaidsTotals;
    const H5 = salaryPaymentTotals.taxesTotals;



    // Create Key and Header
    let arr = [];
    const combinedData = H1.concat(H2, H3, H4, H5, H6);
    combinedData.map((elem) => {
        arr.push({
            header: "",
            key: elem.code,
        })
        // if (elem.code === "2002" || elem.code === "2003") {
        //     arr.push({
        //         header: "",
        //         key: elem.code,
        //         width: 100
        //     })
        // } else {
        //     arr.push({
        //         header: "",
        //         key: elem.code,
        //     })
        // }
    })

    // First Header
    const headerOneEndColumLength = H1.length + H2.length + H3.length + H4.length - 1;


    // HEADER 4 zagalowk
    let cell1 = sheet.getCell(2, 5);
    cell1.value = HEAD.head;
    cell1.alignment = { vertical: 'middle', horizontal: 'center' };
    cell1.font = { name: 'Times New Roman', size: 12, bold: true };

    let cell2 = sheet.getCell(3, 5);
    cell2.value = HEAD.date;
    cell2.alignment = { vertical: 'middle', horizontal: 'center' };
    cell2.font = { name: 'Times New Roman', size: 10, bold: true };

    let cell3 = sheet.getCell(2, headerOneEndColumLength + 1);
    cell3.value = `${HEAD.director.label}_______________________${HEAD.director.value}`;
    cell3.font = { name: 'Times New Roman', size: 10, bold: true };

    let cell4 = sheet.getCell(3, headerOneEndColumLength + 1);
    cell4.value = `${HEAD.accountant.label}_______________________${HEAD.accountant.value}`;
    cell4.font = { name: 'Times New Roman', size: 10, bold: true };

    // Initial Headers
    H1.map((item, index) => {
        const position = index + 1;
        sheet.getCell(initialRow, position).value = item.name;
        if (item.code === "2000" || item.code === "2002" || item.code === "2003") {
            let cell = sheet.getCell(initialRow, position);
            cell.alignment = { vertical: 'middle', horizontal: 'center' };
            cell.font = { name: 'Times New Roman', size: 11, bold: true };
        } else {
            let cell = sheet.getCell(initialRow, position);
            cell.alignment = { textRotation: 90, vertical: 'middle', horizontal: 'center' };
            cell.font = { name: 'Times New Roman', bold: true };
        }

    })


    const rowPositionTwo = initialRow + 1;
    sheet.mergeCells(`A${initialRow}:A${rowPositionTwo}`); // Initial Headers degishli H1
    sheet.mergeCells(`B${initialRow}:B${rowPositionTwo}`); // Initial Headers degishli H1
    sheet.mergeCells(`C${initialRow}:C${rowPositionTwo}`); // Initial Headers degishli H1
    sheet.mergeCells(`D${initialRow}:D${rowPositionTwo}`); // Initial Headers degishli H1
    sheet.mergeCells(`E${initialRow}:E${rowPositionTwo}`); // Initial Headers degishli H1

    H2.map((item, index) => {
        const position = H1.length + index + 1;
        sheet.getCell(initialRow + 1, position).value = item.name;
        let cell = sheet.getCell(initialRow + 1, position);
        cell.alignment = { textRotation: 90, vertical: 'middle', horizontal: 'center' };
        cell.font = { name: 'Times New Roman', bold: true };
    })

    const H3_STARTR = H1.length + H2.length
    H3.map((item, index) => {
        const position = H3_STARTR + index + 1;
        sheet.getCell(initialRow + 1, position).value = item.name;
        let cell = sheet.getCell(initialRow + 1, position);
        cell.alignment = { textRotation: 90, vertical: 'middle', horizontal: 'center' };
        cell.font = { name: 'Times New Roman', bold: true };
    })

    const H4_START = H3_STARTR + H3.length;
    H4.map((item, index) => {
        const position = H4_START + index + 1;
        sheet.getCell(initialRow + 1, position).value = item.name;
        let cell = sheet.getCell(initialRow + 1, position);
        cell.alignment = { textRotation: 90, vertical: 'middle', horizontal: 'center' };
        cell.font = { name: 'Times New Roman', bold: true };
    })

    const H5_START = H4_START + H4.length;
    H5.map((item, index) => {
        const position = H5_START + index + 1;
        sheet.getCell(initialRow + 1, position).value = item.name;
        let cell = sheet.getCell(initialRow + 1, position);
        cell.alignment = { textRotation: 90, vertical: 'middle', horizontal: 'center' };
        cell.font = { name: 'Times New Roman', bold: true };
    })

    const H6_START = H5_START + H5.length;
    H6.map((item, index) => {
        const position = H6_START + index + 1;
        sheet.getCell(initialRow + 1, position).value = item.name;
        let cell = sheet.getCell(initialRow + 1, position);
        cell.alignment = { textRotation: 90, vertical: 'middle', horizontal: 'center' };
        cell.font = { name: 'Times New Roman', bold: true };
    })

    // Maglumatyny (salgytlaryny) goshyars her useryn
    sheet.columns = arr;
    salaryPayments.forEach((elem, i) => {
        let row = sheet.addRow({
            [parseInt(arr[0].key, 10)]: i + 1,
            [parseInt(arr[1].key, 10)]: elem.employeeId,
            [parseInt(arr[2].key, 10)]: `${elem.employee.firstName.charAt(0)}. ${elem.employee.middleName === "" ? "" : elem.employee.middleName.charAt(0) + ". "}${elem.employee.lastName}`,
            [parseInt(arr[3].key, 10)]: elem.employee.position.name,
            [parseInt(arr[4].key, 10)]: checkValue(elem.employee.position.salary),
            [parseInt(arr[5].key, 10)]: elem.workedDays,
            [parseInt(arr[6].key, 10)]: checkValue(elem.totalSalary),
            [parseInt(arr[7].key, 10)]: checkValue(elem.night),
            [parseInt(arr[8].key, 10)]: checkValue(elem.holiday),
            [parseInt(arr[9].key, 10)]: checkValue(elem.totalAddPositionPaid), // utgeshdirme
            ["2010"]: checkValue(elem.recalculated), // gaytadan hasaplama
            ["2011"]: checkValue(elem.harmfulness), // zyyanlyk
            ["2012"]: checkValue(elem.totalAccrued), //  jemi hasaplama
            ["2013"]: checkValue(elem.advance), //   hak ujy 
            ["2014"]: checkValue(elem.totalWithheld),//  jemi tutylan totalWithheld
            ["2015"]: checkValue(elem.totalOnHand),//  jemi tutylan totalOnHand
        });
        elem.salaryPaymentAddPaids.map((paid) => row.getCell(paid.additionalPaid.code).value = checkValue(paid.amount))
        elem.salaryPaymentTaxes.map((taxe) => row.getCell(taxe.tax.code).value = checkValue(taxe.amount))
    });

    // JEMI goshyars
    let row = sheet.addRow({
        [parseInt(arr[6].key, 10)]: checkValue(salaryPaymentTotals.totalSalary),
        [parseInt(arr[7].key, 10)]: checkValue(salaryPaymentTotals.night),
        [parseInt(arr[8].key, 10)]: checkValue(salaryPaymentTotals.holiday),
        [parseInt(arr[9].key, 10)]: checkValue(salaryPaymentTotals.totalAddPositionPaid), // utgeshdirme
        ["2010"]: checkValue(salaryPaymentTotals.recalculated), // gaytadan hasaplama
        ["2011"]: checkValue(salaryPaymentTotals.harmfulness), // zyyanlyk
        ["2012"]: checkValue(salaryPaymentTotals.totalAccrued), // jemi hasaplama
        ["2013"]: checkValue(salaryPaymentTotals.advance), // hak ujy 
        ["2014"]: checkValue(salaryPaymentTotals.totalWithheld), // jemi tutylan totalWithheld
        ["2015"]: checkValue(salaryPaymentTotals.totalOnHand), // jemi tutylan totalOnHand
    });
    salaryPaymentTotals.addPaidsTotals.map((t) => row.getCell(t.code).value = checkValue(t.total))
    salaryPaymentTotals.taxesTotals.map((taxe) => row.getCell(taxe.code).value = checkValue(taxe.total))

    row.eachCell((cell) => {
        cell.font = { bold: true };
    });

    const rowLengthJemi = salaryPayments.length + 2 + initialRow; //startColumn, startRow, endColumn, endRow
    sheet.mergeCells(getCellRange(1, rowLengthJemi, 6, rowLengthJemi));
    sheet.getCell(rowLengthJemi, 1).value = 'Jemi';
    sheet.getCell(rowLengthJemi, 1).style = {
        font: {
            name: 'Times New Roman',
            bold: true,
            color: { argb: 'FF000000' }
        },
    }

    sheet.mergeCells(getCellRange(H1.length + 1, initialRow, headerOneEndColumLength, initialRow));
    sheet.getCell(initialRow, H1.length + 1).value = 'Häzirki aý üçin zahmet tölegleri gaznasyndan hasaplanany';
    sheet.getCell(initialRow, H1.length + 1).alignment = { vertical: 'middle', horizontal: 'center' };
    sheet.getCell(initialRow, H1.length + 1).style = {
        font: {
            size: 11,
            bold: true,
            color: { argb: 'FF000000' },
            name: 'Times New Roman'
        },
        alignment: {
            vertical: 'middle',
            horizontal: 'center'
        }
    }

    // Second Header
    const headerTwoStartColumLength = H1.length + H2.length + H3.length + H4.length;
    const headerTwoEndColumLength = H1.length + H2.length + H3.length + H4.length + H5.length + H6.length;

    sheet.mergeCells(getCellRange(headerTwoStartColumLength, initialRow, headerTwoEndColumLength, initialRow));
    sheet.getCell(initialRow, headerTwoStartColumLength).value = 'Tutylan we bergiň ujyndan hasaplanany';
    sheet.getCell(initialRow, headerTwoStartColumLength).alignment = { vertical: 'middle', horizontal: 'center' };
    sheet.getCell(initialRow, headerTwoStartColumLength).style = {
        font: {
            size: 11,
            bold: true,
            color: { argb: 'FF000000' },
            name: 'Times New Roman',
        },
        alignment: {
            vertical: 'middle',
            horizontal: 'center'
        }
    }


    const setBordersForActiveCells = (sheet) => {
        const range = sheet.getCell('A1').address;
        const rowCount = sheet.rowCount;
        const colCount = sheet.columnCount;
        let minRow = Infinity, maxRow = -Infinity, minCol = Infinity, maxCol = -Infinity;

        for (let row = initialRow; row <= rowCount; row++) {
            for (let col = 1; col <= colCount; col++) {
                const cell = sheet.getCell(row, col);
                if (cell.value !== undefined && cell.value !== null) {
                    minRow = Math.min(minRow, row);
                    maxRow = Math.max(maxRow, row);
                    minCol = Math.min(minCol, col);
                    maxCol = Math.max(maxCol, col);
                }
            }
        }

        for (let row = minRow; row <= maxRow; row++) {
            for (let col = minCol; col <= maxCol; col++) {
                const cell = sheet.getCell(row, col);
                cell.border = {
                    top: { style: 'thin', color: { argb: 'FF000000' } },
                    left: { style: 'thin', color: { argb: 'FF000000' } },
                    bottom: { style: 'thin', color: { argb: 'FF000000' } },
                    right: { style: 'thin', color: { argb: 'FF000000' } }
                };
            }
        }
    };
    setBordersForActiveCells(sheet)

    // Padding start and end columns
    // autoFitColumns(sheet, (initialRow + 1), (initialRow + 2));
    autoFitColumns(sheet, (initialRow + 2), (initialRow + 2 + salaryPayments.length));


    const date = new Date()
    workbook.xlsx.writeFile(`termination_4${date.getFullYear()}.xlsx`).then(() => console.log('File saved'));
})

function getColumnLetter(index) {
    let letter = '';
    while (index > 0) {
        const modulo = (index - 1) % 26;
        letter = String.fromCharCode(65 + modulo) + letter;
        index = Math.floor((index - modulo) / 26);
    }
    return letter;
}

function getCellRange(startColumn, startRow, endColumn, endRow) {
    return `${getColumnLetter(startColumn)}${startRow}:${getColumnLetter(endColumn)}${endRow}`;
}

function checkValue(value) {
    return (value === 0 || !value) ? "" : value.toFixed(2);
}


function autoFitColumns(worksheet, startRow, endRow) {
    //const startRow = 8; // Начальная строка
    //const endRow = 14;  // Конечная строка
    // Собираем максимальную длину для каждого столбца в указанном диапазоне строк
    const columnLengths = {};
    for (let row = startRow; row <= endRow; row++) {
        const rowData = worksheet.getRow(row);
        rowData.eachCell({ includeEmpty: true }, (cell, colNumber) => {
            const cellValue = cell.value ? cell.value.toString() : '';
            if (!columnLengths[colNumber]) {
                columnLengths[colNumber] = 0;
            }

            if (cellValue) {
                columnLengths[colNumber] = Math.max(columnLengths[colNumber], cellValue.length);
            } else {
                columnLengths[colNumber] = Math.max(columnLengths[colNumber], 3);
            }

        });
    }
    // Устанавливаем ширину столбцов в зависимости от максимальной длины содержимого
    worksheet.columns.forEach((column, colNumber) => {
        if (columnLengths[colNumber + 1] !== undefined) { // Индекс столбца начинается с 1
            column.width = columnLengths[colNumber + 1] + 1; // Добавляем 2 для отступа
        }
    });
}


module.exports = readAndProcessFile;