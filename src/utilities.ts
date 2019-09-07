/**
 * Get the last row number that contains data for a specific column.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {number} column Column number.
 *
 * @return {number} Last row number.
 */
function getLastRowForColumn(sheet: GoogleAppsScript.Spreadsheet.Sheet, column: number): number {
    // Get the last row with data for the whole sheet.
    let numRows = sheet.getLastRow();

    // Get all data for the given column
    let data = sheet.getRange(1, column, numRows).getValues();

    // Iterate backwards and find first non empty cell
    for (let i = data.length - 1; i >= 0; i--) {
        if (
            data[i][0] !== null &&
            "string" === typeof data[i][0] && "" !== data[i][0].trim()
        ) {
            return i + 1;
        }
    }

    return -1;
}

/**
 * Converts a number to a letter.
 *
 * @param {number} column Number to convert to alpha.
 * @returns {string} Letter
 */
function columnToLetter(column: number): string {
    let temp, letter = '';
    while (column > 0) {
        temp = (column - 1) % 26;
        letter = String.fromCharCode(temp + 65) + letter;
        column = (column - temp - 1) / 26;
    }
    return letter;
}

/**
 * Converts a letter to a number.
 *
 * @param {string} letter Letter to convert.
 * @returns {number}
 */
function letterToColumn(letter: string) {
    let column = 0, length = letter.length;
    for (let i = 0; i < length; i++) {
        column += (letter.charCodeAt(i) - 64) * Math.pow(26, length - i - 1);
    }
    return column;
}

/**
 * Creates an A1 Notation string.
 *
 * @param {string} col1 Starting column letter.
 * @param {number} row1 Starting row number.
 * @param {string} col2 Ending column letter.
 * @param {number} row2 Ending row number.
 */
function getA1fromColRows(col1: string, row1: number, col2: string, row2: number): string {
    col2 = col2 || col1;
    row2 = row2 || row1;

    return col1 + row1 + ":" + col2 + row2;
}

