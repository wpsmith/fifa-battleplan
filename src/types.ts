// Type Aliases.
type Sheet = GoogleAppsScript.Spreadsheet.Sheet
type PromptResponse = GoogleAppsScript.Base.PromptResponse
type Button = GoogleAppsScript.Base.Button;

// Battle Plan State.
enum BattlePlanState { READY, DIRTY}

// Whether Complete or Incomplete.
enum CompleteIncomplete { INCOMPLETE, COMPLETE}

/**
 * Config class
 */
class Config extends Object {
    /**
     *
     */
    static sheetIndex: string[number];
    sheets: BattlePlanSheet[];

    /**
     * Gets property value.
     *
     * @param {string} prop Property key.
     * @return {string|number|array|object} Config property value.
     */
    getProp(prop: string): any {
        // Check if object property.
        if (this.hasOwnProperty(prop)) {
            // @ts-ignore
            return this[prop];
        }
    }

    /**
     * Gets a sheet by id.
     *
     * @param {string} id
     * @return {BattlePlanSheet} Battlesheet
     */
    getSheet(id: string): BattlePlanSheet {
        // Check if object property.
        // @ts-ignore
        if (this.hasOwnProperty(id) && this[id] instanceof BattlePlanSheet) {
            // @ts-ignore
            return this[id];
        }

        // Check if we have an index.
        // @ts-ignore
        if (Config.sheetIndex[id]) {
            // @ts-ignore
            let s = this.sheets[Config.sheetIndex[id]];
            if (id === s.id) {
                return s;
            }
        }

        // Cycle through it.
        for (let i = 0, l = this.sheets.length; i < l; i++) {
            if (id === this.sheets[i].id) {
                return this.sheets[i];
            }
        }
        return null;
    }

    /**
     * Adds a Battle Plan Sheet.
     *
     * @param {BattlePlanSheet} sheet
     */
    addSheet(sheet: BattlePlanSheet): void {
        // @ts-ignore
        Config.sheetIndex[sheet.id] = this.sheets.length;
        this.sheets.push(sheet);
    }
}


/**
 * Battle Sheet Class
 */
class BattlePlanSheet {
    startRow: number;
    _startCol: string;
    id: string;
    name: string;
    data: any;

    /**
     * BattlePlanSheet constructor
     *
     * @param {string} id Sheet ID.
     * @param {string} name Sheet Name.
     * @param {number} startRow Starting row for the sheet.
     * @param {string} startCol Starting column for the sheet.
     * @param {object} data Misc data.
     */
    constructor(id: string, name: string, startRow: number, startCol: string, data?: any) {
        this.id = id;
        this.name = name;
        this.startRow = startRow;
        this._startCol = startCol;
        this.data = data || {};
    }

    /**
     * Gets data property or all data.
     *
     * @param {string} prop Property key.
     * @return {string|number|array|object} Data property value.
     */
    getData(prop: string): any {
        if (this.data.hasOwnProperty(prop)) {
            return this.data[prop];
        }

        let val = getProperty(this.id + ".data." + prop);
        if (val) {
            return val;
        }

        return this.data;
    }

    /**
     * Sets the data.
     *
     * @param {string} prop Property key.
     * @param {string|number|array|object} val Data property value.
     */
    setData(prop: string, val: any) {
        this.data[prop] = val;
        setProperty(this.id + ".data." + prop, val);
    }

    /**
     * Gets the starting column number with an optional offset.
     *
     * @param {number} offset
     * @return {number}
     */
    getStartColNumber(offset?: number): number {
        offset = offset || 0;
        if (0 === offset) {
            return letterToColumn(this._startCol);
        }
        return letterToColumn(this._startCol + offset);
    }

    /**
     * Gets the starting column letter with an optional offset.
     *
     * @param {number} offset
     * @return {string}
     */
    getStartColLetter(offset?: number): string {
        offset = offset || 0;
        if (0 === offset) {
            return this._startCol;
        }
        return columnToLetter(letterToColumn(this._startCol) + offset);
    }
}

/**
 * Response wrapper for Alerts and Prompts.
 */
class Response {
    kind: string;
    response: (PromptResponse | Button);
    data: any;

    /**
     * Response constructor.
     *
     * @param {object} response Response object.
     * @param {string|number|object|array} data Can be anything really.
     */
    constructor(response: (PromptResponse | Button), data: any) {
        let ui = SpreadsheetApp.getUi();

        if (
            ui.Button.CANCEL == response ||
            ui.Button.CLOSE == response ||
            ui.Button.NO === response ||
            ui.Button.YES === response ||
            ui.Button.OK === response
        ) {
            this.kind = "Button"
        } else {
            this.kind = "PromptResponse"
        }

        this.response = response;
        this.data = data;
    }

    /**
     * Gets response text. Implements PromptResponse interface.
     */
    getResponseText(): string {
        if ("Button" === this.kind) {
            return ""
        }

        let resp = <PromptResponse>this.response;
        return resp.getResponseText();
    }

    /**
     * Gets response button. Implements PromptResponse interface.
     */
    getSelectedButton(): Button {
        if ("Button" === this.kind) {
            return <Button>this.response
        }

        let resp = <PromptResponse>this.response;
        return resp.getSelectedButton();
    }
}

/**
 * Player
 */
class Player {
    name: string;
    ovr: number;
    missedLVLs: number;
    blacklisted: boolean;
    reason: string;
    _data: any;

    /**
     * Player constructor.
     *
     * @ts-ignore
     * @param {string|number} no1 Name or OVR.
     * @param {string|number} no2 Name or OVR.
     */
    constructor(no1: (string | number), no2?: (string | number)) {
        // @ts-ignore
        for (let i = 0, l = arguments.length; i < l; i++) {
            // @ts-ignore
            if ("number" === typeof arguments[i]) {
                // @ts-ignore
                this.ovr = arguments[i];
            } else if ("string" === typeof arguments[i]) {
                // @ts-ignore
                this.name = arguments[i];
            }
        }
    }

    /**
     * Returns a team array. Name first, OVR second.
     *
     * @ts-ignore
     */
    toTeam(): (string | number)[] {
        return [this.name, this.ovr];
    }

    /**
     * Returns a team array. OVR first, Name second.
     *
     * @ts-ignore
     */
    toOpponent(): (string | number)[] {
        return [this.ovr, this.name];
    }

    /**
     * Sets meta data.
     *
     * @param {string|number|array|object} data Anything really.
     */
    setData(data: any): void {
        this._data = data;
    }

    /* PLAYER RETIREMENT */

    /**
     * Prepares a player for retirement.
     *
     * @param {number} missedLVLs Number of missed LVLs.
     * @param {boolean} blacklisted Whether player will be blacklisted.
     * @param {string} reason Reason for retirement.
     */
    toRetire(missedLVLs: number, blacklisted: boolean, reason: string): void {
        this.missedLVLs = missedLVLs;
        this.blacklisted = blacklisted;
        this.reason = reason;
    }

    /**
     * Returns a retirement data array.
     *
     * @param {string} email User email of person processing retirement.
     */
    retire(email?: string): any[] {
        email = email || getUserEmail();
        return [
            this.name,
            (this.blacklisted ? "Yes" : "No"),
            this.missedLVLs,
            this.reason,
            email
        ]
    }
}
