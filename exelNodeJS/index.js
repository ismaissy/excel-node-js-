// const ExcelJS = require('exceljs');

// // Создание нового Excel файла
// const workbook = new ExcelJS.Workbook();
// const sheet = workbook.addWorksheet('My Sheet');

// // Добавление данных
// sheet.columns = [
//   { header: 'Name', key: 'name', width: 30 },
//   { header: 'Age', key: 'age', width: 15 }
// ];
// sheet.addRow({ name: 'John', age: 30 });
// sheet.addRow({ name: 'Jane', age: 25 });
// // Выровнять по центру содержимое столбца age
// sheet.getColumn('age').alignment = { horizontal: 'center' };

// sheet.getRow(1).eachCell((cell) => {
//     cell.fill = {
//       type: 'pattern',
//       pattern: 'solid',
//       fgColor: { argb: 'FFFF0000' } // Красный цвет
//     };

//     // Дополнительно можно задать стиль шрифта
//     cell.font = {
//       bold: true,
//       color: { argb: 'FFFFFFFF' } // Белый цвет текста
//     };
//   });
// // Сохранение файла
// workbook.xlsx.writeFile('example.xlsx')
//   .then(() => {
//     console.log('File saved');
//   });

// // Чтение файла
// // const workbookRead = new ExcelJS.Workbook();
// // workbookRead.xlsx.readFile('example.xlsx')
// //   .then(() => {
// //     const worksheet = workbookRead.getWorksheet('My Sheet');
// //     worksheet.eachRow((row, rowNumber) => {
// //       console.log(`Row ${rowNumber}:`, row.values);
// //     });
// //   });
const ExcelJS = require('exceljs');

// Создаем новую книгу (workbook)
const workbook = new ExcelJS.Workbook();
const sheet = workbook.addWorksheet('Sheet1');

// Задаем заголовки и данные
sheet.columns = [
    { header: '№', key: 'index' },  // Увеличил ширину для лучшего отображения
    { header: 'Işgär ID', key: 'userId', width: 10, },
    { header: 'Ady Familiýasy', key: 'fio', width: 50 },
    { header: 'Wezipe', key: 'wezipe', width: 50 },
    { header: 'Asyl aýlygy', key: 'asylAylygy', width: 10 },
    { header: 'Häzirki aý üçin zahmet tölegleri gaznasyndan hasaplanany', key: 'h1', width: 20 },
];

// Добавляем данные
sheet.addRow({ index: 1, userId: 30, fio: 'QWWQDWQD', wezipe: 'IT', asylAylygy: '365.3' });
sheet.addRow({ index: 2, userId: 25, fio: 'CCCCCCC', wezipe: 'Buh', asylAylygy: '9963.9' });

sheet.getRow(2).height = 100;
sheet.getRow(5).height = 3;


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

sheet.mergeCells('F1:J1');

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

sheet.mergeCells("A7:F7");
sheet.getCell(7, 1).alignment = { vertical: 'middle' };
sheet.getCell(7, 1).value = ' Jemi'
sheet.getCell(7, 1).border = {
    top: { style: 'thin', color: { argb: 'FF000000' } }, // Верхняя граница
    left: { style: 'thin', color: { argb: 'FF000000' } }, // Левая граница
    bottom: { style: 'thin', color: { argb: 'FF000000' } }, // Нижняя граница
    right: { style: 'thin', color: { argb: 'FF000000' } }  // Правая граница
};


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

// Применяем контуры ко всем активным ячейкам
setBordersForActiveCells(sheet);



workbook.xlsx.writeFile('example.xlsx').then(() => console.log('File saved'));


// const ExcelJS = require('exceljs');

// // Создаем новую книгу (workbook)
// const workbook = new ExcelJS.Workbook();
// const sheet = workbook.addWorksheet('Sheet1');

// // Заполняем заголовки
// sheet.getCell('A1').value = 'Header 1';
// sheet.getCell('B1').value = 'Header 2';

// // Заполняем данные для столбцов
// const columnData = {
//     A: ['Data 1', 'New Data 1', 'New Data 2', 'New Data 3'],
//     B: ['Data 2', 'More Data 1', 'More Data 2', 'More Data 3']
// };

// // Добавляем данные в столбцы
// Object.keys(columnData).forEach(col => {
//     columnData[col].forEach((value, index) => {
//         sheet.getCell(`${col}${index + 2}`).value = value;
//     });
// });

// // Сохраняем файл
// workbook.xlsx.writeFile('example.xlsx')
//     .then(() => {
//         console.log('File saved');
//     });
