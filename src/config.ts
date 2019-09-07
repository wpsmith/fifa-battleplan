/* global LodashGS */
// @ts-ignore
let _ = LodashGS.load();

// Global config. Not guaranteed to work per Google.
let config: Config = new Config();

/**
 * Gets the config values.
 */
function getConfigValues(): object{
    const sheetNames = getSheetKeys(),
        keys = getConfigKeys();

    // Setting sheets as BattlePlanSheet instances.
    for (let i = 0; i < sheetNames.length; i++) {
        let k = sheetNames[i],
            v = new BattlePlanSheet(
                k,
                getProperty(k),
                getProperty(k + ".startRow"),
                getProperty(k + ".startCol"),
            );
        let sheetDataKeys = keys.filter(function (value) {
            return (
                0 <= value.indexOf(sheetNames[i] + ".data.")
            )
        });

        // Set data values.
        if (sheetDataKeys.length > 0) {
            for (let _i = 0; _i < sheetDataKeys.length; _i++) {
                v.setData(sheetDataKeys[_i].split(".")[2], getProperty(sheetDataKeys[_i]));
            }
        }

        // @ts-ignore
        config[k] = v;
        config.addSheet(v);
    }

    return config;
}

/**
 * Sets global properties for access later.
 */
function setProperties(): void {
    setProperty("dev", getIsDev());
    setProperty("sheets", getSheetKeys());
    setProperty("configKeys", getConfigKeys());

    let configSheet = getActive().getSheetByName("Config"),
        values = configSheet.getDataRange().getValues();

    // Store the various config properties.
    for (let i = 0, l = values.length; i < l; i++) {
        if ("" === values[i][0]) {
            continue;
        }
        setProperty(values[i][0], values[i][1]);
    }
}

setProperties();

/**
 * Whether in development or not.
 */
function getIsDev(): boolean {
    return false;
}

/**
 * Gets an array of allowed config keys.
 *
 * @return {string[]}
 */
function getConfigKeys(): string[] {
    return [
        "activeTeam",
        "activeTeam.startRow",
        "activeTeam.startCol",
        "activeTeam.data.rankingRange",
        "activeTeam.data.missedLVLsCol",
        "activeTeamRanking",
        "activeTeamRanking.startRow",
        "activeTeamRanking.startCol",
        "retiredTeam",
        "retiredTeam.startRow",
        "retiredTeam.startCol",
        "historical",
        "historical.startRow",
        "historical.startCol",
        "battlePlan",
        "battlePlan.startRow",
        "battlePlan.startCol",
        "battlePlan.data.opponentStartCol",
        "battlePlan.data.opponentLeagueName",
        "battlePlan.data.opponentLeagueRank",
        "battlePlan.data.opponentRange",
        "battlePlan.data.opponentCols",
        "battlePlan.data.opponentActualGoalsCol",
        "battlePlan.data.teamRange",
        "battlePlan.data.teamCols",
        "battlePlan.data.teamActualGoalsCol",
        "data",
        "data.startRow",
        "data.startCol",
        "data.data.state",
        "data.data.teamActualGoals",
        "data.data.opponentActualGoals"
    ];
}

/**
 * Gets an array of allowed sheet keys.
 *
 * @return {string[]}
 */
function getSheetKeys(): string[] {
    return [
        "activeTeam",
        "activeTeamRanking",
        "retiredTeam",
        "battlePlan",
        "historical",
        "data",
    ];
}

/**
 * Gets the Spreadsheet ID.
 *
 * @returns {string}
 */
function getID(): string {
    if (getIsDev()) {
        return config.getProp("ID") || getProperty("ID");
    }

    return SpreadsheetApp.getActiveSpreadsheet().getId();
}

/* HELPERS */

/**
 * Gets the start row of a sheet.
 *
 * @param {string} sheet Sheet key.
 */
function getStartRow(sheet: string): number {
    // @ts-ignore
    return config[sheet].startRow;
}

/**
 * Gets the starting column as a number or letter.
 *
 * @param {string} sheet Sheet key.
 * @param {string} colOrLetter Whether "col" or "letter".
 * @param {number} offset Number to offset given column relatively.
 * @return {number|string} Column number or letter
 */
function getStartCol(sheet: string, colOrLetter: string, offset?: number): (number | string) {
    offset = offset || 0;
    if ("col" === colOrLetter) {
        // @ts-ignore
        return config[sheet].getStartColNumber(offset);
    }
    // @ts-ignore
    return config[sheet].getStartColLetter(offset);

}

/**
 * Gets meta data of a given sheet.
 *
 * @param {string} sheet Sheet key.
 * @param {string} prop Meta data property.
 * @private
 *
 * @return {string|number|object|array} Data property value.
 */
function _getData(sheet: string, prop: string): any {
    // @ts-ignore
    return config[sheet].getData(prop);
}

/* RETIRED TEAM */

/**
 * Gets starting row from Retired Team Sheet.
 *
 * @returns {number}
 */
function getRetiredTeamStartRow(): number {
    return getStartRow("retiredTeam");
}

/**
 * Gets starting colum from Retired Team Sheet.
 *
 * @returns {number}
 */
function getRetiredTeamStartCol(): number {
    return <number>getStartCol("retiredTeam", "col", 0);
}

/**
 * Gets starting colum from Retired Team Sheet.
 *
 * @returns {number}
 */
function getRetiredTeamStartColLetter(offset?: number): string {
    offset = offset || 0;
    return <string>getStartCol("retiredTeam", "letter", offset);
}

/* ACTIVE TEAM RANKING **/

/**
 * Gets starting row from Active Team Sheet.
 *
 * @returns {number}
 */
function getActiveTeamRankingsStartRow(): number {
    return getStartRow("activeTeamRanking");
}

/**
 * Gets starting colum from Active Team Sheet.
 *
 * @returns {number}
 */
function getActiveTeamRankingsStartCol(offset?: number): number {
    return <number>getStartCol("activeTeamRanking", "col", offset);
}

/**
 * Gets starting column from Active Team Sheet.
 *
 * @returns {string}
 */
function getActiveTeamRankingsStartColLetter(offset?: number): string {
    return <string>getStartCol("activeTeamRanking", "letter", offset);
}

/* ACTIVE TEAM **/

/**
 * Gets starting row from Active Team Sheet.
 *
 * @returns {number}
 */
function getActiveTeamStartRow(): number {
    return getStartRow("activeTeam");
}

/**
 * Gets starting colum from Active Team Sheet.
 *
 * @returns {number}
 */
function getActiveTeamStartCol(): number {
    return <number>getStartCol("activeTeam", "col");
}

/**
 * Gets data from Active Team.
 *
 * @param {string} prop Data property to get.
 *
 * @returns {any} Value of the data property.
 */
function getActiveTeamData(prop: string): any {
    switch (prop) {
        case "name":
            prop = "nameCol";
            break;
        case "missedLVLs":
            prop = "missedLVLsCol";
            break;
    }
    return _getData("activeTeam", prop);
}

/* BATTLEPLAN TEAM **/

/**
 * Gets starting row from BattlePlan Sheet.
 *
 * @returns {number}
 */
function getBattlePlanStartRow(): number {
    return getStartRow("battlePlan");
}

/**
 * Gets starting column from BattlePlan Sheet.
 *
 * @returns {number}
 */
function getBattlePlanStartCol(): number {
    return <number>getStartCol("battlePlan", "col");
}

/**
 * Gets starting column from BattlePlan Sheet.
 *
 * @returns {string}
 */
function getBattlePlanStartColLetter(): string {
    return <string>getStartCol("battlePlan", "letter");
}

/**
 * Gets starting row from Active Team Sheet.
 *
 * @returns {number}
 */
function getBattlePlanData(prop: string): any {
    return _getData("battlePlan", prop);
}

/**
 * Gets starting row from Historical Sheet.
 *
 * @returns {number}
 */
function getHistoricalStartRow(): number {
    return getStartRow("historical");
}

/**
 * Gets starting colum from Historical Sheet.
 *
 * @returns {number}
 */
function getHistoricalStartCol(): number {
    return <number>getStartCol("historical", "col");
}

/**
 * Gets starting colum from Historical Sheet.
 *
 * @returns {string}
 */
function getHistoricalStartColLetter(): string {
    return <string>getStartCol("historical", "letter");
}

/* DATA **/

/**
 * Gets starting row from Active Team Sheet.
 *
 * @returns {number}
 */
function getData(prop: string): any {
    return _getData("data", prop);
}

/* SPREADSHEETS */

/**
 * Gets the active spreadsheet.
 *
 * @global {GoogleAppsScript.Spreadsheet.Spreadsheet} SpreadsheetApp
 * @returns {GoogleAppsScript.Spreadsheet.Spreadsheet}
 */
function getActive(): GoogleAppsScript.Spreadsheet.Spreadsheet {
    if (getIsDev()) {
        return SpreadsheetApp.openById(getID())
    }

    return SpreadsheetApp.getActiveSpreadsheet();
}

/**
 * Gets a specific Sheet by name.
 *
 * @private
 * @param {string} name Name of the sheet.
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function _getSheetByName(name: string): GoogleAppsScript.Spreadsheet.Sheet {
    return getActive().getSheetByName(name);
}

/**
 * Gets the Data sheet.
 *
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function getDataSheet(): GoogleAppsScript.Spreadsheet.Sheet {
    return _getSheetByName(config.getSheet("data").name);
}

/**
 * Gets the Active Team sheet.
 *
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function getActiveTeamSheet(): GoogleAppsScript.Spreadsheet.Sheet {
    return _getSheetByName(config.getSheet("activeTeam").name);
}

/**
 * Gets the Active Team Rankings sheet.
 *
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function getActiveTeamRankingsSheet(): GoogleAppsScript.Spreadsheet.Sheet {
    return _getSheetByName(config.getSheet("activeTeamRanking").name);
}

/**
 * Gets the Retired Team sheet.
 *
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function getRetiredTeamSheet(): GoogleAppsScript.Spreadsheet.Sheet {
    return _getSheetByName(config.getSheet("retiredTeam").name);
}

/**
 * Gets the Current LVL Strategy sheet.
 *
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function getBattlePlanSheet(): GoogleAppsScript.Spreadsheet.Sheet {
    return _getSheetByName(config.getSheet("battlePlan").name);
}

/**
 * Gets the Historical LVL sheet.
 *
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function getHistoricalSheet(): GoogleAppsScript.Spreadsheet.Sheet {
    return _getSheetByName(config.getSheet("historical").name);
}

/* PROPERTIES */

/**
 * Gets a script property
 *
 * @param {string} prop Property key.
 * @return {string|number|object|array} Value of property.
 */
function getProperty(prop: string): any {
    let val = PropertiesService.getScriptProperties().getProperty(prop);
    // let val = PropertiesService.getScriptProperties().getProperty(getID() + prop);

    return JSON.parse(val);
}

/**
 * Sets a script property
 *
 * @param {string} prop Property key.
 * @param {string|number|object|array} val Value of property.
 */
function setProperty(prop: string, val: any): void {
    PropertiesService.getScriptProperties().setProperty(prop, JSON.stringify(val));
    // PropertiesService.getScriptProperties().setProperty(getID() + prop, val);
}
