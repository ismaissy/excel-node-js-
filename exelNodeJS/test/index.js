const ExcelJS = require('exceljs');
const fs = require('fs').promises;
const { getCellRange } = require("./mergeByIndex")


const headerCode = [
    { code: "2000", name: "№" },
    { code: "2001", name: "Işgär ID" },
    { code: "2002", name: "Ady Familiýasy" },
    { code: "2003", name: "Wezipe" },
    { code: "2004", name: "Asyl aýlygy" },
    { code: "2005", name: "Işlan günleri" },
    { code: "2006", name: "Aýlyk" },
    { code: "2007", name: "Gijeki" },
    { code: "2008", name: "Baýramçylyk" },
    { code: "2009", name: "Zähmet rugsady" },
]

const fistMonthData = [
    { name: "Gaýtadan hasaplama" },
    { name: "Zyýanlyk" },
    { name: "Jemi hasaplanan" },
]

const advance = { name: "Hak ujy", cell: 1 }


const firstMonthStartCount = 6;
const firstMonthEndCount = 10;






async function readAndProcessFile() {
    try {
        const data = await fs.readFile("./wedemost_obj.json", 'utf8');
        const jsonData = JSON.parse(data);
        return jsonData;
    } catch (err) {
        console.error('Ошибка при чтении или разборе файла:', err);
    }
}
readAndProcessFile().then((json) => {


    const salaryPayments = json.salaryPayments;
    const salaryPaymentTotals = json.salaryPaymentTotals;


    console.log(salaryPaymentTotals.addPaidsTotals.length)
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Sheet1');

    salaryPaymentTotals.addPaidsTotals.map((item, index) => {
        const position = firstMonthEndCount + index + 1;
        sheet.getCell(2, position).value = item.name;
        sheet.getCell(2, position).alignment = { textRotation: 90, vertical: 'middle', horizontal: 'center' };
    })

    const firstMonthStartCountContinue = 10 + salaryPaymentTotals.addPaidsTotals.length;
    const firstMonthEndCountContinue = 10 + salaryPaymentTotals.addPaidsTotals.length + fistMonthData.length;

    fistMonthData.map((item, index) => {
        const position = firstMonthStartCountContinue + index + 1;
        sheet.getCell(2, position).value = item.name;
        sheet.getCell(2, position).alignment = { textRotation: 90, vertical: 'middle', horizontal: 'center' };
    })





    // Задаем заголовки и данные
    sheet.columns = [
        { header: '№', key: 'index' },  // Увеличил ширину для лучшего отображения
        { header: 'Işgär ID', key: 'userId', },
        { header: 'Ady Familiýasy', key: 'fio', },
        { header: 'Wezipe', key: 'wezipe' },
        { header: 'Asyl aýlygy', key: 'asylAylygy' },
        { header: 'Häzirki aý üçin zahmet tölegleri gaznasyndan hasaplanany', key: 'h1' },
        { header: 'Tutylan we bergiň ujyndan hasaplanany', key: 'h2' },
        { header: 'QQ4', key: 'qq4', width: 20 },
    ];

    sheet.getCell(1, 6).value = 'Häzirki aý üçin zahmet tölegleri gaznasyndan hasaplanany'; // Устанавливаем текст в ячейке G12
    sheet.getCell(1, 6).style = {
        font: {
            size: 11,
            bold: true,
            color: { argb: 'FF000000' } // Черный цвет текста
        },
        fill: {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFC0C0C0' } // Серый цвет фона
        },
        alignment: {
            vertical: 'middle',
            horizontal: 'center'
        }
    };

    // Добавляем примерные данные
    let row = sheet.addRow({
        index: 1,
        userId: 1001,
        fio: 'John Doe',
        wezipe: 'Manager',
    });

    // Позже добавляем дополнительные данные в ту же строку
    row.getCell('asylAylygy').value = 5000;
    row.getCell('h1').value = 100;
    row.getCell('h2').value = 200;
    row.getCell('qq4').value = 'Data for QQ4';


    // Добавляем данные
    // sheet.addRow({ index: 4, userId: 30, fio: 'QWWQDWQD', wezipe: 'IT', asylAylygy: '365.3' });
    // sheet.addRow({ index: 5, userId: 25, fio: 'CCCCCCC', wezipe: 'Buh', asylAylygy: '9963.9' });
    // sheet.addRow({ index: 6, userId: 33, fio: 'CCCCCCC', wezipe: 'Buh', asylAylygy: '3333.9' });
    // sheet.addRow({ index: 7, userId: 33, fio: 'CCCCCCC', wezipe: 'Buh', asylAylygy: '3333.9' });
    sheet.getCell(3, 2).value = 'Isma'

    // sheet.getRow(2).height = 100;
    // sheet.getRow(5).height = 3; //*


    // sheet.getColumn('A').alignment = { horizontal: 'center' };
    // sheet.getColumn('B').alignment = { horizontal: 'center' };
    // // Сохраняем файл


    // Применяем вертикальное выравнивание текста для заголовков
    sheet.mergeCells('A1:A2');
    sheet.getCell('A1').alignment = { vertical: 'middle', horizontal: 'center' };

    sheet.mergeCells('B1:B2');
    sheet.getCell('B1').alignment = { textRotation: 90, vertical: 'middle', horizontal: 'center' };

    sheet.mergeCells('C1:C2');
    sheet.getCell('C1').alignment = { vertical: 'middle', horizontal: 'center' };

    sheet.mergeCells('D1:D2');
    sheet.getCell('D1').alignment = { vertical: 'middle', horizontal: 'center' };

    sheet.mergeCells('E1:E2');
    sheet.getCell('E1').alignment = { textRotation: 90, vertical: 'middle', horizontal: 'center' };



    sheet.getCell(1, firstMonthStartCount).alignment = { vertical: 'middle', horizontal: 'center' };
    sheet.mergeCells(
        getCellRange(firstMonthStartCount, 1, firstMonthEndCountContinue, 1)
    );

    sheet.getCell(2, 6).value = 'Işlan günleri'
    sheet.getCell(2, 6).alignment = { textRotation: 90, vertical: 'middle', horizontal: 'center' };

    sheet.getCell(2, 7).value = 'Aýlyk'
    sheet.getCell(2, 7).alignment = { textRotation: 90, vertical: 'middle', horizontal: 'center' };

    sheet.getCell(2, 8).value = 'Gijeki'
    sheet.getCell(2, 8).alignment = { textRotation: 90, vertical: 'middle', horizontal: 'center' };

    sheet.getCell(2, 9).value = 'Baýramçylyk'
    sheet.getCell(2, 9).alignment = { textRotation: 90, vertical: 'middle', horizontal: 'center' };

    sheet.getCell(2, 10).value = 'Utgişdyrma'
    sheet.getCell(2, 10).alignment = { textRotation: 90, vertical: 'middle', horizontal: 'center' };

    // sheet.mergeCells("A10:F10");
    // sheet.getCell(10, 1).alignment = { vertical: 'middle' };
    // sheet.getCell(10, 1).value = ' Jemi'
    // sheet.getCell(10, 1).border = {
    //     top: { style: 'thin', color: { argb: 'FF000000' } }, // Верхняя граница
    //     left: { style: 'thin', color: { argb: 'FF000000' } }, // Левая граница
    //     bottom: { style: 'thin', color: { argb: 'FF000000' } }, // Нижняя граница
    //     right: { style: 'thin', color: { argb: 'FF000000' } }  // Правая граница
    // };


    sheet.mergeCells(getCellRange(
        firstMonthEndCountContinue + advance.cell,
        1,
        firstMonthEndCountContinue + advance.cell + salaryPaymentTotals.taxesTotals.length,
        1));

    sheet.getCell(1, firstMonthEndCountContinue + advance.cell).value = 'Tutylan we bergiň ujyndan hasaplanany'
    sheet.getCell(1, firstMonthEndCountContinue + advance.cell).alignment = { vertical: 'middle', horizontal: 'center' };
    sheet.getCell(1, firstMonthEndCountContinue + advance.cell).style = {
        font: {
            size: 11,
            bold: true,
            color: { argb: 'FF000000' } // Черный цвет текста
        },
        fill: {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFC0C0C0' } // Серый цвет фона
        },
        alignment: {
            vertical: 'middle',
            horizontal: 'center'
        }
    };


    sheet.getCell(2, firstMonthEndCountContinue + advance.cell).value = advance.name;
    sheet.getCell(2, firstMonthEndCountContinue + advance.cell).alignment = { textRotation: 90, vertical: 'middle', horizontal: 'center' };

    salaryPaymentTotals.taxesTotals.map((item, index) => {
        const position = firstMonthEndCountContinue + advance.cell + index + 1;
        sheet.getCell(2, position).value = item.name;
        sheet.getCell(2, position).alignment = { textRotation: 90, vertical: 'middle', horizontal: 'center' };

    })



    // Функция для установки контуров для активных ячеек
    const setBordersForActiveCells = (sheet) => {
        const range = sheet.getCell('A1').address;  // Инициализируем диапазон
        const rowCount = sheet.rowCount;
        const colCount = sheet.columnCount;

        let minRow = Infinity, maxRow = -Infinity, minCol = Infinity, maxCol = -Infinity;

        // Определяем диапазон активных ячеек
        for (let row = 1; row <= rowCount; row++) {
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

        // Применяем контуры к активным ячейкам
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

    workbook.xlsx.writeFile('example.xlsx').then(() => console.log('File saved'));




});

