function Add2Strategy(): void {
    if (BattlePlanState.READY !== getBattlePlanState()) {
        showYesNoAlert("BattlePlan hasn't been recorded yet and you will lose all records. Do you want to reset it and continue?", _add2Strategy);
    } else {
        _add2Strategy();
    }
}

function SortByLVL(): void {
    sortActiveTeam(8, false, "H");
}

function SortByDisadv(): void {
    sortActiveTeam(9, false, "I");
}

function SortByEq(): void {
    sortActiveTeam(10, false, "J");
}

function SortByAdv(): void {
    sortActiveTeam(11, false, "K");
}

function AddPlayer(): void {

    showInputPrompt("Player name?", function (response?: Response): void {
        if (isResponseNo(response.getSelectedButton())) {
            return;
        }

        showInputPrompt(response.getResponseText() + " OVR?", function (response?: Response): void {
            if (isResponseNo(response.getSelectedButton())) {
                return;
            }

            let activeTeamSheet = getActiveTeamSheet();

            // Insert player.
            insertRow(activeTeamSheet, getActiveTeamStartCol(), [
                response.data, //name
                response.getResponseText() // OVR
            ]);

            renumberCol(activeTeamSheet);

            // Add user to Active Team Rankings
            insertRow(getActiveTeamRankingsSheet(), getActiveTeamRankingsStartCol(), [response.data]);

        }, response.getResponseText());

    });

}

function RetirePlayer(): void {
    let activeTeamSheet = getActiveTeamSheet(),
        row = activeTeamSheet.getActiveCell().getRow(),
        player = activeTeamSheet.getRange(row, getActiveTeamStartCol()).getValue();

    showInputPrompt("What is the reason " + player + " is retiring?", function (response?: Response): void {
        if (isResponseNo(response.getSelectedButton())) {
            return;
        }

        showYesNoAlert("Can " + response.data + " return?", _retirePlayer, {
            player: response.data,
            reason: response.getResponseText()
        });
    }, player);

}

function ClearBattlePlan(clearOpponent?: boolean): void {
    let battlePlanSheet = getBattlePlanSheet();

    // Clear content.
    let opponent = getBattlePlanData("opponentCols"),
        team = getBattlePlanData("teamCols");

    if (clearOpponent) {
        for (let i = 0, l = opponent.length; i < l; i++) {
            battlePlanSheet.getRange(opponent[i] + getBattlePlanStartRow() + ":" + opponent[i]).clearContent();
        }
    }
    for (let i = 0, l = team.length; i < l; i++) {
        battlePlanSheet.getRange(team[i] + getBattlePlanStartRow() + ":" + team[i]).clearContent();
    }
    // Reset state.
    setBattlePlanState(BattlePlanState.READY);
}

function RecordBattlePlan(): void {
    if (
        CompleteIncomplete.COMPLETE !== getBattlePlanActiveGoalsState("team") ||
        CompleteIncomplete.COMPLETE !== getBattlePlanActiveGoalsState("opponent")
    ) {
        showYesNoAlert("BattlePlan isn't ready to be recorded yet. Actual goals are missing. Do you still want to record it and continue?", _recordBattlePlan);
    } else {
        _recordBattlePlan();
    }
    //
    // // Copy the range.
    // let battlePlanSheet = getBattlePlanSheet(),
    //     a1Range = getBattlePlanStartColLetter() + getBattlePlanStartRow() + ":" +
    //         battlePlanSheet.getLastColumn() + getLastRowForColumn(getBattlePlanSheet(), getBattlePlanStartCol()),
    //     // a1Range = getBattlePlanStartColLetter() + getBattlePlanStartRow() + ":" + columnToLetter(getBattlePlanSheet().getLastColumn()),
    //     battlePlanRange = battlePlanSheet.getRange(a1Range).getValues(),
    //     opponentLeagueName = battlePlanSheet.getRange(getBattlePlanData("opponentLeagueName")).getValue(),
    //     opponentLeagueRank = battlePlanSheet.getRange(getBattlePlanData("opponentLeagueRank")).getValue();
    //
    // // Move to Historical
    // let historicalSheet = getHistoricalSheet(),
    //     historicalLastRow = getLastRowForColumn(getHistoricalSheet(), getHistoricalStartCol()),
    //     historicalNextRow = historicalLastRow + 1;
    //
    // // Add OpponentLeage/Rank to data.
    // battlePlanRange = battlePlanRange.filter(function (value) {
    //     return "" !== value[0].trim();
    // });
    // for (let i = 0, l = battlePlanRange.length; i < l; i++) {
    //     battlePlanRange[i].push(opponentLeagueName, opponentLeagueRank);
    //     battlePlanRange[i].unshift(battlePlanRange[i][0] + " " + battlePlanRange[i][1]);
    // }
    //
    // // Insert rows.
    // historicalSheet.insertRowsAfter(historicalLastRow, battlePlanRange.length);
    //
    // // Copy values.
    // historicalSheet
    //     .getRange(getHistoricalStartColLetter() + historicalNextRow + ":" + columnToLetter(getHistoricalStartCol() + battlePlanRange[0].length - 1) + (historicalNextRow + battlePlanRange.length - 1))
    //     .setValues(battlePlanRange);
    //
    // ClearBattlePlan();
    // // // Clear content.
    // // let opponent = getBattlePlanData("opponentCols"),
    // //     team = getBattlePlanData("teamCols");
    // // for (let i = 0, l = opponent.length; i < l; i++) {
    // //     battlePlanSheet.getRange(opponent[i] + getBattlePlanStartRow() + ":" + opponent[i]).clearContent();
    // // }
    // // for (let i = 0, l = team.length; i < l; i++) {
    // //     battlePlanSheet.getRange(team[i] + getBattlePlanStartRow() + ":" + team[i]).clearContent();
    // // }
    // //
    // // // Reset state.
    // // setBattlePlanState(BattlePlanState.READY);
}

function ResetBattlePlan(): void {
    if (BattlePlanState.READY !== getBattlePlanState()) {
        showYesNoAlert("BattlePlan hasn't been recorded yet and you will lose all records. Do you want to reset it and continue?", function (response?: Response): void {
            if (response && isResponseYES(<Button>response.response)) {
                ClearBattlePlan();
            }
        });
    } else {
        ClearBattlePlan();
    }
}

function SortTeam(): void {
    sortBattlePlanTeam("team");
}

function SortOpponent(): void {
    sortBattlePlanTeam("opponent");
}
