// const XLSX = require('xlsx');




// // Создаем новую книгу (workbook)
// const workbook = XLSX.utils.book_new();

// // Данные для листа
// const data = [
//     ['Name', 'Age', 'surname'],
//     ['John', 30, 'QWWQDWQD'],
//     ['Jane', 25, "CCCCCCC"]
// ];

// // Преобразуем массив в лист
// const worksheet = XLSX.utils.aoa_to_sheet(data);

// // Устанавливаем ширину колонок
// worksheet['!cols'] = [
//     { wch: 30 }, // ширина для первой колонки (Name)
//     { wch: 10 },  // ширина для второй колонки (Age)
//     { wch: 50 }  // ширина для второй колонки (Age)
// ];

// // Добавляем лист в книгу
// XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');

// // Сохраняем файл
// XLSX.writeFile(workbook, 'example.xlsx');

const XLSX = require('xlsx');

// Создаем новую книгу (workbook)
const workbook = XLSX.utils.book_new();

// Данные для листа
const data = [
    ['Header 1', 'Header 2'],
    ['Data 1', 'Data 2'],
    ['Data 3', 'Data 4']
];

// Преобразуем массив в лист
const worksheet = XLSX.utils.aoa_to_sheet(data);

// Устанавливаем объединение ячеек
worksheet['!merges'] = [
    { s: { r: 0, c: 0 }, e: { r: 0, c: 1 } }, // Объединяет ячейки A1 и B1
    { s: { r: 1, c: 0 }, e: { r: 2, c: 0 } }  // Объединяет ячейки A2 и A3
];

// Сохраняем файл
XLSX.writeFile(workbook, 'example.xlsx');
