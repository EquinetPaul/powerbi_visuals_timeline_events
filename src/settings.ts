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

/**
 * Timeline Formatting Card
 */
class TimelineCardSettings extends FormattingSettingsCard {
    scaling = new formattingSettings.NumUpDown({
        name: "scaling",
        displayName: "Scaling",
        value: 1
    });

    symbolColor = new formattingSettings.ColorPicker({
        name: "symbolColor",
        displayName: "Symbol color",
        value: { value: "#2563eb" }
    });

    showEventLabels = new formattingSettings.ToggleSwitch({
        name: "showEventLabels",
        displayName: "Show event names",
        value: true
    });

    showDates = new formattingSettings.ToggleSwitch({
        name: "showDates",
        displayName: "Show dates",
        value: false
    });

    useEventSymbols = new formattingSettings.ToggleSwitch({
        name: "useEventSymbols",
        displayName: "Use event symbols",
        value: false
    });

    showLegend = new formattingSettings.ToggleSwitch({
        name: "showLegend",
        displayName: "Show legend",
        value: false
    });

    name: string = "timeline";
    displayName: string = "Timeline";
    slices: Array<FormattingSettingsSlice> = [
        this.scaling,
        this.symbolColor,
        this.showEventLabels,
        this.showDates,
        this.useEventSymbols,
        this.showLegend
    ];
}

/**
* visual settings model class
*
*/
export class VisualFormattingSettingsModel extends FormattingSettingsModel {
    // Create formatting settings model formatting cards
    timelineCard = new TimelineCardSettings();

    cards = [this.timelineCard];
}
