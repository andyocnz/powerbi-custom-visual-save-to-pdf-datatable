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
    orientation = new formattingSettings.ItemDropdown({
        name: "orientation",
        displayName: "Orientation",
        value: { value: "auto", displayName: "Auto" },
        items: [
            { value: "auto", displayName: "Auto" },
            { value: "portrait", displayName: "Portrait" },
            { value: "landscape", displayName: "Landscape" }
        ]
    });

    paperSize = new formattingSettings.ItemDropdown({
        name: "paperSize",
        displayName: "Paper size",
        value: { value: "a4", displayName: "A4" },
        items: [
            { value: "a4", displayName: "A4" },
            { value: "letter", displayName: "Letter" }
        ]
    });

    pdfFontSize = new formattingSettings.NumUpDown({
        name: "pdfFontSize",
        displayName: "PDF font size",
        value: 9
    });

    headerText = new formattingSettings.TextInput({
        name: "headerText",
        displayName: "Header text",
        value: "DataTable Export",
        placeholder: "Header text"
    });

    name: string = "exportSettings";
    displayName: string = "Export";
    slices: Array<FormattingSettingsSlice> = [this.orientation, this.paperSize, this.pdfFontSize, this.headerText];
}

/**
* visual settings model class
*
*/
export class VisualFormattingSettingsModel extends FormattingSettingsModel {
    pagingCard = new PagingCardSettings();
    exportCard = new ExportSettingsCardSettings();
    cards = [this.pagingCard, this.exportCard];
}
