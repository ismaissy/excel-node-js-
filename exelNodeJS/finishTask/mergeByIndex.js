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

module.exports = {
    getCellRange
}