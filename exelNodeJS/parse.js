// const fs = require('fs').promises;

// async function readAndParseJsonFile(filePath) {
//     try {
//         const data = await fs.readFile(filePath, 'utf8');
//         const jsonData = JSON.parse(data);
//         return jsonData;
//     } catch (err) {
//         console.error('Error reading or parsing file:', err);
//         return null;
//     }
// }

// async function main() {
//     const filePath = './wedemost_obj.json'; // Path to your JSON file
//     const data = await readAndParseJsonFile(filePath);
//     console.log(data)
//     // if (data) {
//     //     // Log the whole object to check its structure
//     //     console.log('JSON Data:', data);

//     //     // Check if the `employee` property exists
//     //     if (data.employee) {
//     //         console.log('Employee Name:', data.employee.name); // Example: John Doe

//     //         if (data.employee.details) {
//     //             console.log('Employee Address City:', data.employee.details.address?.city); // Example: Springfield
//     //         } else {
//     //             console.log('Employee details not found');
//     //         }

//     //         if (data.employee.salary && Array.isArray(data.employee.salary.bonuses)) {
//     //             console.log('First Bonus Amount:', data.employee.salary.bonuses[0]?.amount); // Example: 5000
//     //         } else {
//     //             console.log('Employee salary or bonuses not found');
//     //         }
//     //     } else {
//     //         console.log('Employee data not found');
//     //     }
//     // } else {
//     //     console.log('No data found');
//     // }
// }

// main();

const fs = require('fs').promises;

// Функция для чтения и разбора JSON файла
async function readAndProcessFile(filePath) {
    try {
        const data = await fs.readFile(filePath, 'utf8');
        return JSON.parse(data);
    } catch (err) {
        console.error('Ошибка при чтении или разборе файла:', err);
    }
}

// Функция для обработки данных по выплатам
function processSalaryPayments(salaryPayments) {
    return salaryPayments.map(payment => ({
        id: payment.id,
        employeeId: payment.employeeId,
        registrationDate: payment.registrationDate,
        workUnitId: payment.workUnitId,
        workedDays: payment.workedDays,
        workedDutyHours: payment.workedDutyHours,
        workedNightHours: payment.workedNightHours,
        workedHolidayHours: payment.workedHolidayHours,
        night: payment.night,
        holiday: payment.holiday,
        harmfulness: payment.harmfulness,
        advance: payment.advance,
        alimony: payment.alimony,
        dependentPay: payment.dependentPay,
        totalSalary: payment.totalSalary,
        totalWithheld: payment.totalWithheld,
        totalAddPaid: payment.totalAddPaid,
        totalAddPositionPaid: payment.totalAddPositionPaid,
        totalAccrued: payment.totalAccrued,
        totalOnHand: payment.totalOnHand,
        userId: payment.userId,
        createdAt: payment.createdAt,
        updatedAt: payment.updatedAt,
        version: payment.version,
        salaryPaymentAddPaids: payment.salaryPaymentAddPaids.map(addPaid => ({
            code: addPaid.code,
            name: addPaid.name,
            total: addPaid.total
        })),
        salaryPaymentTaxes: payment.salaryPaymentTaxes.map(tax => ({
            code: tax.code,
            name: tax.name,
            total: tax.total
        }))
    }));
}

// Функция для обработки сводных данных по зарплате
function processSalaryTotals(totals) {
    return {
        totalSalary: totals.totalSalary,
        night: totals.night,
        holiday: totals.holiday,
        totalAddPositionPaid: totals.totalAddPositionPaid,
        harmfulness: totals.harmfulness,
        advance: totals.advance,
        totalWithheld: totals.totalWithheld,
        totalAccrued: totals.totalAccrued,
        totalOnHand: totals.totalOnHand,
        addPaidsTotals: totals.addPaidsTotals.map(item => ({
            code: item.code,
            name: item.name,
            total: item.total
        })),
        taxesTotals: totals.taxesTotals.map(item => ({
            code: item.code,
            name: item.name,
            total: item.total
        }))
    };
}

// Основной блок
(async () => {
    const jsonData = await readAndProcessFile("./wedemost_obj.json");
    console.log(jsonData.salaryPayments[0].employee.firstName)
    // if (jsonData) {
    //     const processedSalaryPayments = processSalaryPayments(jsonData.salaryPayments);
    //     const processedSalaryTotals = processSalaryTotals(jsonData.salaryPaymentTotals);

    //     const result = {
    //         salaryPayments: processedSalaryPayments,
    //         salaryPaymentTotals: processedSalaryTotals
    //     };

    //     console.log('Processed Data:', result.salaryPayments[1].employeeId);
    // }
})();