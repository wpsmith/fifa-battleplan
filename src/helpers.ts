/**
 * Adds players from the Active Team to the Battle Plan.
 *
 * @param {Button} response Response of YES/NO/CANCEL.
 * @private
 */
function _add2Strategy(response?: Response): void {
    if (response && isResponseNo(<Button>response.response)) {
        return;
    }
    // else if (response && isResponseYES(<Button>response.response)) {
    //     // @todo Maybe ask and move BattlePlan to Historical.
    //     // RecordBattlePlan();
    // }

    // Clear the Battle Plan.
    ClearBattlePlan(false);

    // @todo Move hard coded column letter to config.
    let allPlayers = getActiveTeamSheet().getRange("B" + getActiveTeamStartRow() + ":E").getValues(),
        battlePlanPlayers = [],
        battlePlanSheet = getBattlePlanSheet();

    // Filter players who need to be excluded.
    allPlayers = allPlayers.filter(function (value) {
        return (value[3] === 0 || value[3] === "");
    });

    // Add player to battlePlanPlayers.
    for (let i = 0, l = allPlayers.length; i < l; i++) {
        allPlayers[i] &&
        allPlayers[i][0] &&
        "" !== allPlayers[i][0] &&
        battlePlanPlayers.push([allPlayers[i][0], allPlayers[i][1]]);
    }

    // Copy players over.
    // @todo Move hard coded column letter to config.
    let battlePlanRange = "B" + getBattlePlanStartRow() + ":C" + ((battlePlanPlayers.length - 1) + getBattlePlanStartRow());
    // var r = battlePlanSheet.getRange(battlePlanRange);
    battlePlanSheet
        .getRange(battlePlanRange)
        .setValues(battlePlanPlayers);

    // Reset selected cell.
    battlePlanSheet.setActiveSelection("B" + getBattlePlanStartRow());
}

/**
 * Retires a player.
 *
 * Moves a player from the Active Team to the Retired Team.
 *
 * @param {Response} response The response of the YES/NO Prompt with data.
 * @private
 */
function _retirePlayer(response?: Response): void {
    if (isResponseCancelClose(response.getSelectedButton())) {
        return;
    }

    // Get Active Player
    let activeTeamSheet = getActiveTeamSheet(),
        data = activeTeamSheet.getRange(getActiveTeamData("rankingRange")).getValues(),
        playerData,
        playerRow;

    // Find Row for active player.
    for (let i = 0; i < data.length; i++) {
        if (data[i][getActiveTeamStartCol() - 1] == response.data.player) {
            Logger.log("found player");
            playerData = data[i];
            playerRow = i + getActiveTeamStartRow();
            break;
        }
    }

    // Setup Player object.
    let player = new Player(response.data.player);
    player.setData(playerData);
    player.toRetire(playerData[getActiveTeamData("missedLVLs")], !isResponseNo(<Button>response.response, true), response.data.reason);

    // Prepare to Retire.
    let retiredTeamSheet = getRetiredTeamSheet(),
        nextRow = getLastRowForColumn(retiredTeamSheet, getRetiredTeamStartCol()) + 1;

    // Copy to Retired.
    retiredTeamSheet.setActiveSelection(getRetiredTeamStartColLetter() + nextRow + ":" + getRetiredTeamStartColLetter(3) + nextRow)
        .setValues([player.retire()]);

    // Delete from Active.
    // playerRange.clear();
    getActiveTeamSheet().deleteRow(playerRow);

    // Renumber
    renumberCol(getActiveTeamSheet(), "A");
}

/**
 * Sorts the team on the Battle Plan sheet.
 *
 * @param {string} team Supports "team" or "opponent".
 */
function sortBattlePlanTeam(team: string): void {
    let cols: string[],
        startCol: string,
        endCol: string,
        col1: number,
        col2: number,
        sheet = getBattlePlanSheet();

    if ("team" === team) {
        cols = getBattlePlanData("teamRange");
    } else if ("opponent" === team) {
        cols = getBattlePlanData("opponentRange");
    }

    startCol = cols.shift();
    endCol = cols.pop();

    if ("team" === team) {
        col2 = letterToColumn(startCol);
        col1 = letterToColumn(endCol);
    } else if ("opponent" === team) {
        col1 = letterToColumn(startCol);
        col2 = letterToColumn(endCol);
    }

    sheet
        .getRange(startCol + getBattlePlanStartRow() + ":" + endCol + sheet.getLastRow())
        .sort([
            {column: col1, ascending: false},
            {column: col2, ascending: true}
        ]);
}

/**
 * Records Battle Plan to Historical sheet.
 *
 * @param response
 * @private
 */
function _recordBattlePlan(response?: Response): void {
    if (response && isResponseNo(<Button>response.response)) {
        return;
    }
    // else if (response && isResponseYES(<Button>response.response)) {
    //     // @todo Maybe ask and move BattlePlan to Historical.
    //     // RecordBattlePlan();
    // }

    // Copy the range.
    let battlePlanSheet = getBattlePlanSheet(),
        a1Range = getBattlePlanStartColLetter() + getBattlePlanStartRow() + ":" +
            battlePlanSheet.getLastColumn() + getLastRowForColumn(getBattlePlanSheet(), getBattlePlanStartCol()),
        // a1Range = getBattlePlanStartColLetter() + getBattlePlanStartRow() + ":" + columnToLetter(getBattlePlanSheet().getLastColumn()),
        battlePlanRange = battlePlanSheet.getRange(a1Range).getValues(),
        opponentLeagueName = battlePlanSheet.getRange(getBattlePlanData("opponentLeagueName")).getValue(),
        opponentLeagueRank = battlePlanSheet.getRange(getBattlePlanData("opponentLeagueRank")).getValue();

    // Move to Historical
    let historicalSheet = getHistoricalSheet(),
        historicalLastRow = getLastRowForColumn(getHistoricalSheet(), getHistoricalStartCol()),
        historicalNextRow = historicalLastRow + 1;

    // Add OpponentLeague/Rank to data.
    battlePlanRange = battlePlanRange.filter(function (value) {
        return "" !== value[0].trim();
    });
    for (let i = 0, l = battlePlanRange.length; i < l; i++) {
        battlePlanRange[i].push(opponentLeagueName, opponentLeagueRank);
        battlePlanRange[i].unshift(battlePlanRange[i][0] + " " + battlePlanRange[i][1]);
    }

    // Insert rows.
    historicalSheet.insertRowsAfter(historicalLastRow, battlePlanRange.length);

    // Copy values.
    historicalSheet
        .getRange(getHistoricalStartColLetter() + historicalNextRow + ":" + columnToLetter(getHistoricalStartCol() + battlePlanRange[0].length - 1) + (historicalNextRow + battlePlanRange.length - 1))
        .setValues(battlePlanRange);

    ClearBattlePlan();
}
