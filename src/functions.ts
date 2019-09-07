/**
 * Gets the current user's email.
 *
 * @returns {string}
 */
function getUserEmail(): string {
    return Session.getActiveUser().getEmail();
}

/**
 * Gets the battleplan state
 *
 * @return {number} State ENUM.
 */
function getBattlePlanState(): number {
    return getDataSheet().getRange(getData("state")).getValue();
}

/**
 * Sets the battle plan state.
 *
 * @param {number} state State of sheet.
 */
function setBattlePlanState(state: number): void {
    getDataSheet().getRange(getData("state")).setValue(state);
}

/**
 * Gets the battleplan state
 *
 * @return {number} State ENUM.
 */
function getBattlePlanActiveGoalsState(team: string): number {
    return getDataSheet().getRange(getData(team + "ActualGoals")).getValue();
}

/**
 * Sets the battle plan state.
 *
 * @param {string} team The team to set actual goals state.
 * @param {number} state State of sheet.
 */
function setBattlePlanActiveGoalsState(team: string, state: number): void {
    getDataSheet().getRange(getData(team + "ActualGoals")).setValue(state);
}

/**
 * Sorts a range ascending or descending on a particular sheet by column.
 *
 * @param {number} col Column integer.
 * @param {boolean} asc Whether the column should sort ascending or not.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet Sheet to target the sorting.
 * @param {string} range Range to sort.
 */
function sortCol(col: number, asc: boolean, sheet: GoogleAppsScript.Spreadsheet.Sheet, range: string) {
    sheet.getRange(range).activate();
    let currentCell = sheet.getCurrentCell();
    sheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
    currentCell.activateAsCurrentCell();
    sheet.getActiveRange().sort({column: col, ascending: asc});
}

/**
 * Sorts a range for the Active Team sheet.
 *
 * @param {number} col Column integer.
 * @param {boolean} asc Whether the column should sort ascending or not.
 * @param {string} cell Cell location to make active when done.
 */
function sortActiveTeam(col: number, asc?: boolean, cell?: string) {
    cell = cell || "A1";
    asc = asc || false;
    let sheet = getActiveTeamSheet();

    sortCol(col, asc, getActiveTeamSheet(), getActiveTeamData("rankingRange"));

    // Set selection.
    sheet.setActiveSelection(cell);
}

/**
 * Displays a YES/NO prompt.
 *
 * @param {string} prompt Prompt message.
 * @param {function} callback Callback function for the prompt.
 * @param {object} data Data to be passed.
 */
function showYesNoAlert(prompt: string, callback: (response: Response) => void, data?: any): void {
    let ui = SpreadsheetApp.getUi(),
        uiResponse = ui.alert(
            prompt,
            ui.ButtonSet.YES_NO
        );
    callback(new Response(uiResponse, data));

}

/**
 * Displays a text input prompt.
 *
 * @param {string} prompt Prompt message.
 * @param {function} callback Callback function for the prompt.
 * @param {object} data Data to be passed.
 */
function showInputPrompt(prompt: string, callback: (response: Response) => void, data?: any): void {
    let ui = SpreadsheetApp.getUi(),
        uiResponse = ui.prompt(
            prompt,
            ui.ButtonSet.OK_CANCEL
        );
    callback(new Response(uiResponse, data));
}

/**
 * Determines whether the response is a NO or a variation of NO.
 *
 * @param {Button} response The response of the dialog prompt.
 * @param {boolean} strict Whether to check for strictly NO.
 * @return {boolean} Returns true for NO, CLOSE, or CANCEL or a strict NO.
 */
function isResponseNo(response: Button, strict?: boolean): boolean {
    let ui = SpreadsheetApp.getUi();
    if (strict) {
        return (ui.Button.NO === response);
    }
    return (ui.Button.NO === response || isResponseCancelClose(response));
}

/**
 * Determines whether the response is a YES.
 *
 * @param {Button} response The response of the dialog prompt.
 * @return {boolean} Returns true for YES.
 */
function isResponseYES(response: Button): boolean {
    return (SpreadsheetApp.getUi().Button.YES === response);
}

/**
 * Determines whether the response is a variation of CANCEL or CLOSE.
 *
 * @param {Button} response The response of the dialog prompt.
 * @return {boolean} Returns true for CLOSE, or CANCEL.
 */
function isResponseCancelClose(response: Button): boolean {
    const ui = SpreadsheetApp.getUi();

    return (ui.Button.CLOSE === response || ui.Button.CANCEL === response);
}

/**
 * Renumbers a specific column.
 *
 * @param {Sheet} sheet Active sheet.
 * @param {string} col Column letter to renumber.
 */
function renumberCol(sheet: Sheet, col?: string) {
    col = col || "A";
    let counter = 1,
        range = sheet.getRange(col + getActiveTeamStartRow() + ":" + col),
        count = [];

    while (count.length < range.getNumRows()) {
        count.push([counter++]);
    }

    range.setValues(count);
}

/**
 * Inserts a row at the bottom of a sheet copying the last row's formulas and inserting data.
 *
 * @param {Sheet} sheet Active sheet.
 * @param {number} startingCol Starting column to insert data or measure column.
 * @param {object} data Data to be inserted.
 */
function insertRow(sheet: Sheet, startingCol: number, data?: any[]) {
    let lastRow = getLastRowForColumn(sheet, startingCol),
        nextRow = lastRow + 1,
        startingColLetter = columnToLetter(startingCol),
        lastCol = columnToLetter(sheet.getLastColumn());

    // Insert row.
    sheet.insertRowAfter(lastRow);

    copyFormatAndFormulas(
        sheet,
        startingColLetter + lastRow + ":" + lastCol + lastRow,
        startingColLetter + nextRow + ":" + lastCol + nextRow
    );

    // Add data
    if (data) {
        sheet.getRange(startingColLetter + nextRow + ":" + columnToLetter(startingCol + data.length - 1) + nextRow).setValues([data]);
    }
}

/**
 * Copies format and forumalas from one A! range to A! another.
 *
 * @link https://webapps.stackexchange.com/questions/95441/script-to-insert-new-row-and-copy-down-formulas-that-auto-increase-according-to
 *
 * @param {Sheet} sheet
 * @param {string} sourceRange
 * @param {string} destinationRange
 */
function copyFormatAndFormulas(sheet: Sheet, sourceRange: string, destinationRange: string) {
    let rangeSource = sheet.getRange(sourceRange),
        rangeDestination = sheet.getRange(destinationRange);

    rangeSource.copyTo(rangeDestination, {formatOnly: true});

    let formulas = rangeSource.getFormulasR1C1();
    for (let x in formulas) {
        for (let y in formulas[x]) {
            if (formulas[x][y] == "") {
                continue;
            }
            rangeDestination.getCell(parseInt(x) + 1, parseInt(y) + 1).setFormulaR1C1(formulas[x][y]);
        }
    }
}

/**
 * Creates alert dialog.
 *
 * @param {string} msg Message to display.
 */
function alert(msg: string) {
    SpreadsheetApp.getUi().alert(msg);
}

/**
 * Outputs a message with a variable stringified.
 *
 * @param {string} msg Message to output.
 * @param {object} data Variable.
 */
function alertData(msg: string, data: any): void {
    alert(msg + ": " + JSON.stringify(data));
}

