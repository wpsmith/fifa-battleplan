/**
 * The event handler triggered when a user opens the spreadsheet.
 *
 * @param {Event} e The onOpen event.
 */
function onOpen(e: GoogleAppsScript.Script.Trigger) {
    // Add a custom menu to the spreadsheet.
    let ui = SpreadsheetApp.getUi(),
        battlePlanSub = ui.createMenu('Battle Plan')
            .addItem('Set Team', 'Add2Strategy')
            .addItem('Sort Team', 'SortTeam')
            .addItem('Sort Opponent', 'SortOpponent')
            .addItem('Clear Battle Plan', 'ClearBattlePlan')
            .addItem('Record Battle Plan', 'RecordBattlePlan')
            .addItem('Reset Battle Plan', 'ResetBattlePlan')
        ,
        activeTeamSub = ui.createMenu('Active Team')
            .addItem('Set Team', 'Add2Strategy')
            .addItem('Sort by LVL', 'SortByLVL')
            .addItem('Sort by Advantage', 'SortByAdv')
            .addItem('Sort by Equal', 'SortByEq')
            .addItem('Sort by Disadvantage', 'SortByDisadv')
            .addItem('Add Player', 'AddPlayer')
            .addItem('Retire Player', 'RetirePlayer')
    ;

    ui.createMenu("FIFA")
        .addSubMenu(activeTeamSub)
        .addSubMenu(battlePlanSub)
        .addToUi();

    getConfigValues();
}

/**
 * The event handler triggered when a user changes a value in a spreadsheet.
 *
 * @param {Event} e The onEdit event.
 */
function onEdit(e: GoogleAppsScript.Script.Trigger) {
    let activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

    watchActualGoals(activeSheet, "team");
    watchActualGoals(activeSheet, "opponent");

}

/**
 * Watches actual goal columns for completion and state management.
 *
 * @param {Sheet} activeSheet Active sheet.
 * @param {string} team Supports "team" or "opponent"
 */
function watchActualGoals(activeSheet: Sheet, team: string): void {
    team = ("opponent" === team) ? team : "team";

    let actualGoalsCol = getBattlePlanData(team + "ActualGoalsCol"),
        columnNumberToWatch = letterToColumn(actualGoalsCol),
        range = activeSheet.getActiveCell();

    if (
        activeSheet.getName() == getProperty("battlePlan") && //  config.battlePlan.name
        range.getColumn() == columnNumberToWatch
    ) {

        let actualGoals = activeSheet.getRange(getA1fromColRows(
            actualGoalsCol,
            getBattlePlanStartRow(),
            actualGoalsCol,
            getLastRowForColumn(activeSheet, getBattlePlanStartCol())
        )).getValues();

        for( let i = 0, l = actualGoals.length; i < l; i++) {
            if ("" === actualGoals[i][0]) {
                setBattlePlanActiveGoalsState(team, CompleteIncomplete.INCOMPLETE);
                return;
            }
        }

        setBattlePlanActiveGoalsState(team, CompleteIncomplete.COMPLETE);
    }
}
