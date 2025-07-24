/*
 * (c) Copyright Ascensio System SIA 2010-2025
 *
 * This program is a free software product. You can redistribute it and/or
 * modify it under the terms of the GNU Affero General Public License (AGPL)
 * version 3 as published by the Free Software Foundation. In accordance with
 * Section 7(a) of the GNU AGPL its Section 15 shall be amended to the effect
 * that Ascensio System SIA expressly excludes the warranty of non-infringement
 * of any third-party rights.
 *
 * This program is distributed WITHOUT ANY WARRANTY; without even the implied
 * warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR  PURPOSE. For
 * details, see the GNU AGPL at: http://www.gnu.org/licenses/agpl-3.0.html
 *
 * You can contact Ascensio System SIA at 20A-6 Ernesta Birznieka-Upish
 * street, Riga, Latvia, EU, LV-1050.
 *
 * The  interactive user interfaces in modified source and object code versions
 * of the Program must display Appropriate Legal Notices, as required under
 * Section 5 of the GNU AGPL version 3.
 *
 * Pursuant to Section 7(b) of the License you must retain the original Product
 * logo when distributing the program. Pursuant to Section 7(e) we decline to
 * grant you any rights under trademark law for use of our trademarks.
 *
 * All the Product's GUI elements, including illustrations and icon sets, as
 * well as technical writing content are licensed under the terms of the
 * Creative Commons Attribution-ShareAlike 4.0 International. See the License
 * terms at http://creativecommons.org/licenses/by-sa/4.0/legalcode
 *
 */

function getCellFunctions() {

    let funcs = [];

   if (true)
    {
        let func = new RegisteredFunction();
        func.name = "setAutoFilter";
        func.params = [
            "range (string, optional): cell range to apply autofilter (e.g., 'A1:D10'). If omitted, uses active/selected range",
            "field (number, optional): field number for filtering (starting from 1, left-most field)",
            "fieldName (string, optional): column name/header for filtering (e.g., 'Name', 'Age'). Will automatically find the column number",
            "criteria1 (string|array|object, optional): filter criteria - string for operators (e.g., '>10'), array for multiple values (e.g., [1,2,3]), ApiColor object for color filters, or dynamic filter constant",
            "operator (string, optional): filter operator - 'xlAnd', 'xlOr', 'xlFilterValues', 'xlTop10Items', 'xlTop10Percent', 'xlBottom10Items', 'xlBottom10Percent', 'xlFilterCellColor', 'xlFilterFontColor', 'xlFilterDynamic'",
            "criteria2 (string, optional): second criteria for compound filters (used with xlAnd/xlOr operators)",
            "visibleDropDown (boolean, optional): show/hide filter dropdown arrow (default: true)"
        ];

        func.examples = [
            "If you need to apply autofilter to range A1:D10, respond:" +
            "[functionCalling (setAutoFilter)]: {\"range\": \"A1:D10\"}",

            "To apply autofilter to active/selected range, respond:" +
            "[functionCalling (setAutoFilter)]: {}",

            "To filter column 1 for values greater than 10, respond:" +
            "[functionCalling (setAutoFilter)]: {\"range\": \"A1:D10\", \"field\": 1, \"criteria1\": \">10\"}",

            "To filter by column name 'Name' for specific values, respond:" +
            "[functionCalling (setAutoFilter)]: {\"range\": \"A1:D10\", \"fieldName\": \"Name\", \"criteria1\": [\"John\", \"Jane\"], \"operator\": \"xlFilterValues\"}",

            "To filter by column header 'Age' for values greater than 18, respond:" +
            "[functionCalling (setAutoFilter)]: {\"range\": \"A1:D10\", \"fieldName\": \"Age\", \"criteria1\": \">18\"}",

            "To filter column 2 for specific values [2,5,8], respond:" +
            "[functionCalling (setAutoFilter)]: {\"range\": \"A1:D10\", \"field\": 2, \"criteria1\": [2,5,8], \"operator\": \"xlFilterValues\"}",

            "To filter column 1 for top 10 items, respond:" +
            "[functionCalling (setAutoFilter)]: {\"range\": \"A1:D10\", \"field\": 1, \"criteria1\": \"10\", \"operator\": \"xlTop10Items\"}",

            "To create compound filter (>5 OR <2), respond:" +
            "[functionCalling (setAutoFilter)]: {\"range\": \"A1:D10\", \"field\": 1, \"criteria1\": \">5\", \"operator\": \"xlOr\", \"criteria2\": \"<2\"}",

            "To filter by cell background color (yellow), respond:" +
            "[functionCalling (setAutoFilter)]: {\"range\": \"A1:D10\", \"field\": 1, \"criteria1\": {\"r\": 255, \"g\": 255, \"b\": 0}, \"operator\": \"xlFilterCellColor\"}",

            "To filter by font color (red), respond:" +
            "[functionCalling (setAutoFilter)]: {\"range\": \"A1:D10\", \"field\": 1, \"criteria1\": {\"r\": 255, \"g\": 0, \"b\": 0}, \"operator\": \"xlFilterFontColor\"}",

            "To filter active range for values greater than 5, respond:" +
            "[functionCalling (setAutoFilter)]: {\"field\": 1, \"criteria1\": \">5\"}",

            "To filter active range by column name 'Price' for values less than 100, respond:" +
            "[functionCalling (setAutoFilter)]: {\"fieldName\": \"Price\", \"criteria1\": \"<100\"}",

            "To remove autofilter from range, respond:" +
            "[functionCalling (setAutoFilter)]: {\"range\": \"A1:D10\", \"field\": null}",

            "To clear filter from specific column, respond:" +
            "[functionCalling (setAutoFilter)]: {\"range\": \"A1:D10\", \"field\": 1, \"criteria1\": null}"
        ];

        func.call = async function(params) {
            Asc.scope.range = params.range;
            Asc.scope.field = params.field;
            Asc.scope.fieldName = params.fieldName;
            Asc.scope.criteria1 = params.criteria1;
            Asc.scope.operator = params.operator;
            Asc.scope.criteria2 = params.criteria2;
            Asc.scope.visibleDropDown = params.visibleDropDown;

            await Asc.Editor.callCommand(function(){
                let ws = Api.GetActiveSheet();
                let range;
                
                if (!Asc.scope.range) {
                    range = Api.GetSelection();
                } else {
                    range = ws.GetRange(Asc.scope.range);
                }

                if (!range) {
                    return;
                }

                let field = Asc.scope.field;
                
                if (Asc.scope.fieldName && !field) {
                    let rangeValues = range.GetValue();
                    if (Array.isArray(rangeValues) && rangeValues.length > 0) {
                        let headerRow = rangeValues[0];
                        if (Array.isArray(headerRow)) {
                            for (let i = 0; i < headerRow.length; i++) {
                                if (headerRow[i] && headerRow[i].toString().toLowerCase() === Asc.scope.fieldName.toLowerCase()) {
                                    field = i + 1;
                                    break;
                                }
                            }
                        }
                    }
                }

                if (!field) {
                    field = 1;
                }

                let criteria1 = Asc.scope.criteria1;
                if (Asc.scope.operator === "xlFilterCellColor" || Asc.scope.operator === "xlFilterFontColor") {
                    if (criteria1 && typeof criteria1 === 'object' && criteria1.r !== undefined && criteria1.g !== undefined && criteria1.b !== undefined) {
                        criteria1 = Api.CreateColorFromRGB(criteria1.r, criteria1.g, criteria1.b);
                    }
                }

                range.SetAutoFilter(
                    field,
                    criteria1,
                    Asc.scope.operator,
                    Asc.scope.criteria2,
                    Asc.scope.visibleDropDown
                );
            });
        };

        funcs.push(func);
    }

    if (true)
    {
        let func = new RegisteredFunction();
        func.name = "setSort";
        func.params = [
            "range (string, optional): cell range to sort (e.g., 'A1:D10'). If omitted, uses active/selected range",
            "key1 (string|ApiRange, optional): first sort field - range or named range reference",
            "sortOrder1 (string, optional): sort order for key1 - 'xlAscending' or 'xlDescending' (default: 'xlAscending')",
            "key2 (string|ApiRange, optional): second sort field - range or named range reference",
            "sortOrder2 (string, optional): sort order for key2 - 'xlAscending' or 'xlDescending'",
            "header (string, optional): specifies if first row contains headers - 'xlYes' or 'xlNo' (default: 'xlNo')",
            "orientation (string, optional): sort orientation - 'xlSortColumns' (by rows) or 'xlSortRows' (by columns) (default: 'xlSortColumns')"
        ];

        func.examples = [
            "To sort range A1:D10 by first column in ascending order, respond:" +
            "[functionCalling (setSort)]: {\"range\": \"A1:D10\", \"key1\": \"A1\", \"sortOrder1\": \"xlAscending\"}",

            "To sort active range by first column with headers, respond:" +
            "[functionCalling (setSort)]: {\"key1\": \"A1\", \"header\": \"xlYes\"}",

            "To sort range by multiple columns (first ascending, second descending), respond:" +
            "[functionCalling (setSort)]: {\"range\": \"A1:D10\", \"key1\": \"A1\", \"sortOrder1\": \"xlAscending\", \"key2\": \"B1\", \"sortOrder2\": \"xlDescending\"}",

            "To sort range by three columns, respond:" +
            "[functionCalling (setSort)]: {\"range\": \"A1:D10\", \"key1\": \"A1\", \"sortOrder1\": \"xlAscending\", \"key2\": \"B1\", \"sortOrder2\": \"xlDescending\", \"key3\": \"C1\", \"sortOrder3\": \"xlAscending\"}",

            "To sort range by rows instead of columns, respond:" +
            "[functionCalling (setSort)]: {\"range\": \"A1:D10\", \"key1\": \"A1\", \"orientation\": \"xlSortRows\"}",

            "To sort range with headers by second column descending, respond:" +
            "[functionCalling (setSort)]: {\"range\": \"A1:D10\", \"key1\": \"B1\", \"sortOrder1\": \"xlDescending\", \"header\": \"xlYes\"}",

            "To sort active range by named range key, respond:" +
            "[functionCalling (setSort)]: {\"key1\": \"MyRange\", \"sortOrder1\": \"xlAscending\"}"
        ];

        func.call = async function(params) {
            Asc.scope.range = params.range;
            Asc.scope.key1 = params.key1;
            Asc.scope.sortOrder1 = params.sortOrder1 || "xlAscending";
            Asc.scope.key2 = params.key2;
            Asc.scope.sortOrder2 = params.sortOrder2;
            Asc.scope.key3 = params.key3;
            Asc.scope.sortOrder3 = params.sortOrder3;
            Asc.scope.header = params.header || "xlNo";
            Asc.scope.orientation = params.orientation || "xlSortColumns";

            await Asc.Editor.callCommand(function(){
                let ws = Api.GetActiveSheet();
                let range;

                if (!Asc.scope.range) {
                    range = Api.GetSelection();
                } else {
                    range = ws.GetRange(Asc.scope.range);
                }

                if (!range) {
                    return;
                }

                let key1 = null, key2 = null, key3 = null;

                if (Asc.scope.key1) {
                    if (typeof Asc.scope.key1 === 'string') {
                        try {
                            key1 = ws.GetRange(Asc.scope.key1) || Asc.scope.key1;
                        } catch {
                            key1 = Asc.scope.key1;
                        }
                    } else {
                        key1 = Asc.scope.key1;
                    }
                }

                if (Asc.scope.key2) {
                    if (typeof Asc.scope.key2 === 'string') {
                        try {
                            key2 = ws.GetRange(Asc.scope.key2) || Asc.scope.key2;
                        } catch {
                            key2 = Asc.scope.key2;
                        }
                    } else {
                        key2 = Asc.scope.key2;
                    }
                }

                if (Asc.scope.key3) {
                    if (typeof Asc.scope.key3 === 'string') {
                        try {
                            key3 = ws.GetRange(Asc.scope.key3) || Asc.scope.key3;
                        } catch {
                            key3 = Asc.scope.key3;
                        }
                    } else {
                        key3 = Asc.scope.key3;
                    }
                }

                range.SetSort(
                    key1,
                    Asc.scope.sortOrder1,
                    key2,
                    Asc.scope.sortOrder2,
                    key3,
                    Asc.scope.sortOrder3,
                    Asc.scope.header,
                    Asc.scope.orientation
                );
            });
        };

        funcs.push(func);
    }

    return funcs;
}