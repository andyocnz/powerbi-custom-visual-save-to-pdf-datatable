/*
 *  Power BI Visualizations
 *
 *  Copyright (c) Microsoft Corporation
 *  All rights reserved.
 *  MIT License
 *
 *  Permission is hereby granted, free of charge, to any person obtaining a copy
 *  of this software and associated documentation files (the ""Software""), to deal
 *  in the Software without restriction, including without limitation the rights
 *  to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 *  copies of the Software, and to permit persons to whom the Software is
 *  furnished to do so, subject to the following conditions:
 *
 *  The above copyright notice and this permission notice shall be included in
 *  all copies or substantial portions of the Software.
 *
 *  THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 *  IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 *  FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 *  AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 *  LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 *  OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
 *  THE SOFTWARE.
 */

"use strict";

import { formattingSettings } from "powerbi-visuals-utils-formattingmodel";

import FormattingSettingsCard = formattingSettings.SimpleCard;
import FormattingSettingsSlice = formattingSettings.Slice;
import FormattingSettingsModel = formattingSettings.Model;

class PagingCardSettings extends FormattingSettingsCard {
    pageSize = new formattingSettings.NumUpDown({
        name: "pageSize",
        displayName: "Rows per page",
        value: 25
    });

    name: string = "paging";
    displayName: string = "Paging";
    slices: Array<FormattingSettingsSlice> = [this.pageSize];
}

class ExportSettingsCardSettings extends FormattingSettingsCard {
    orientation = new formattingSettings.AutoDropdown({
        name: "orientation",
        displayName: "Orientation",
        value: "auto"
    });

    paperSize = new formattingSettings.AutoDropdown({
        name: "paperSize",
        displayName: "Paper size",
        value: "a4"
    });

    pdfFontSize = new formattingSettings.NumUpDown({
        name: "pdfFontSize",
        displayName: "PDF font size",
        value: 9
    });

    headerText = new formattingSettings.TextInput({
        name: "headerText",
        displayName: "PDF title",
        value: "",
        placeholder: "e.g. Stark Billing Report"
    });

    name: string = "exportSettings";
    displayName: string = "Export";
    slices: Array<FormattingSettingsSlice> = [this.orientation, this.paperSize, this.pdfFontSize, this.headerText];
}

class TotalsCardSettings extends FormattingSettingsCard {
    showTotals = new formattingSettings.ToggleSwitch({
        name: "showTotals",
        displayName: "Show total row",
        value: false
    });

    name: string = "totals";
    displayName: string = "Totals";
    slices: Array<FormattingSettingsSlice> = [this.showTotals];
}

/**
* visual settings model class
*
*/
export class VisualFormattingSettingsModel extends FormattingSettingsModel {
    pagingCard = new PagingCardSettings();
    exportCard = new ExportSettingsCardSettings();
    totalsCard = new TotalsCardSettings();
    cards = [this.pagingCard, this.exportCard, this.totalsCard];
}
