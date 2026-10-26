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

import powerbi from "powerbi-visuals-api";
import { formattingSettings } from "powerbi-visuals-utils-formattingmodel";

import CompositeCard = formattingSettings.CompositeCard;
import FormattingSettingsCard = formattingSettings.SimpleCard;
import FormattingSettingsModel = formattingSettings.Model;
import FormattingSettingsSlice = formattingSettings.Slice;
import Group = formattingSettings.Group;
import ValidatorType = powerbi.visuals.ValidatorType;

// Power BI defaults: Segoe UI, secondary text color and neutral axis color of the default theme.
export const DEFAULT_FONT_FAMILY = "\"Segoe UI\", wf_segoe-ui_normal, helvetica, arial, sans-serif";
export const DEFAULT_TEXT_COLOR = "#605E5C";
export const DEFAULT_LINE_COLOR = "#C8C6C4";
export const DEFAULT_TODAY_COLOR = "#D64550";

export type LegendPosition = "Top" | "Bottom" | "Left" | "Right" | "TopCenter" | "BottomCenter" | "LeftCenter" | "RightCenter";
export type MarkerShape = "circle" | "square" | "diamond" | "triangle" | "invertedTriangle" | "cross" | "star" | "hexagon";
export type LabelPosition = "above" | "below" | "alternate";
export type DateLabelMode = "auto" | "eventDates" | "startEnd";
export type LineStyle = "solid" | "dashed" | "dotted";
export type LaneOrder = "alphabetical" | "firstDate";

function createFontControl(defaultFontSize: number): formattingSettings.FontControl {
    return new formattingSettings.FontControl({
        name: "font",
        displayNameKey: "Prop_Font",
        fontFamily: new formattingSettings.FontPicker({ name: "fontFamily", value: DEFAULT_FONT_FAMILY }),
        fontSize: new formattingSettings.NumUpDown({
            name: "fontSize",
            value: defaultFontSize,
            options: {
                minValue: { type: ValidatorType.Min, value: 6 },
                maxValue: { type: ValidatorType.Max, value: 40 }
            }
        }),
        bold: new formattingSettings.ToggleSwitch({ name: "bold", value: false }),
        italic: new formattingSettings.ToggleSwitch({ name: "italic", value: false }),
        underline: new formattingSettings.ToggleSwitch({ name: "underline", value: false })
    });
}

function createColorPicker(name: string, displayNameKey: string, color: string): formattingSettings.ColorPicker {
    return new formattingSettings.ColorPicker({ name, displayNameKey, value: { value: color } });
}

function createNumUpDown(name: string, displayNameKey: string, value: number, min: number, max: number): formattingSettings.NumUpDown {
    return new formattingSettings.NumUpDown({
        name,
        displayNameKey,
        value,
        options: {
            minValue: { type: ValidatorType.Min, value: min },
            maxValue: { type: ValidatorType.Max, value: max }
        }
    });
}

class LegendCard extends CompositeCard {
    show = new formattingSettings.ToggleSwitch({ name: "show", displayNameKey: "Prop_Show", value: true });
    position = new formattingSettings.AutoDropdown({ name: "position", displayNameKey: "Prop_Position", value: "Top" });
    showTitle = new formattingSettings.ToggleSwitch({ name: "showTitle", displayNameKey: "Prop_Title", value: true });
    titleText = new formattingSettings.TextInput({ name: "titleText", displayNameKey: "Prop_TitleText", value: "", placeholder: "" });
    font = createFontControl(9);
    labelColor = createColorPicker("labelColor", "Prop_Color", DEFAULT_TEXT_COLOR);

    name: string = "legend";
    displayNameKey: string = "Card_Legend";
    topLevelSlice = this.show;
    groups: Group[] = [
        new Group({ name: "legendOptions", displayNameKey: "Group_Options", slices: [this.position] }),
        new Group({ name: "legendText", displayNameKey: "Group_Text", slices: [this.font, this.labelColor] }),
        new Group({ name: "legendTitle", displayNameKey: "Group_Title", topLevelSlice: this.showTitle, slices: [this.titleText] })
    ];
}

/**
 * One lane per legend value. Lanes scroll vertically when there are more than `visibleLanes`.
 */
class LanesCard extends FormattingSettingsCard {
    show = new formattingSettings.ToggleSwitch({ name: "show", displayNameKey: "Prop_Show", value: false });
    visibleLanes = createNumUpDown("visibleLanes", "Prop_VisibleLanes", 4, 1, 20);
    order = new formattingSettings.AutoDropdown({ name: "order", displayNameKey: "Prop_LaneOrder", value: "alphabetical" });
    alternateBackground = new formattingSettings.ToggleSwitch({ name: "alternateBackground", displayNameKey: "Prop_AlternateBackground", value: true });

    name: string = "lanes";
    displayNameKey: string = "Card_Lanes";
    topLevelSlice = this.show;
    slices: Array<FormattingSettingsSlice> = [this.visibleLanes, this.order, this.alternateBackground];
}

/**
 * Colors per legend value are chosen in the visual's editor (the visual header's Edit button)
 * and stored as JSON in a hidden property of this object.
 */
class ColorsCard extends FormattingSettingsCard {
    defaultColor = createColorPicker("defaultColor", "Prop_DefaultColor", "");

    name: string = "colors";
    displayNameKey: string = "Card_Colors";
    slices: Array<FormattingSettingsSlice> = [this.defaultColor];
}

class MarkersCard extends FormattingSettingsCard {
    shape = new formattingSettings.AutoDropdown({ name: "shape", displayNameKey: "Prop_Shape", value: "circle" });
    shapeByEvent = new formattingSettings.ToggleSwitch({ name: "shapeByEvent", displayNameKey: "Prop_ShapeByEvent", value: false });
    size = createNumUpDown("size", "Prop_Size", 12, 4, 40);

    name: string = "markers";
    displayNameKey: string = "Card_Markers";
    slices: Array<FormattingSettingsSlice> = [this.shape, this.shapeByEvent, this.size];
}

/**
 * Icons drawn inside the markers. Icons per legend value and per event name are chosen in the visual's
 * editor (the visual header's Edit button) and stored as JSON in hidden properties of this object.
 */
class IconsCard extends CompositeCard {
    show = new formattingSettings.ToggleSwitch({ name: "show", displayNameKey: "Prop_Show", value: true });
    defaultIcon = new formattingSettings.AutoDropdown({ name: "defaultIcon", displayNameKey: "Prop_DefaultIcon", value: "none" });

    allEvents = new Group({ name: "iconsAllEvents", displayNameKey: "Group_ForAllEvents", slices: [this.defaultIcon] });

    name: string = "icons";
    displayNameKey: string = "Card_Icons";
    topLevelSlice = this.show;
    groups: Group[] = [this.allEvents];
}

class DateAxisCard extends CompositeCard {
    show = new formattingSettings.ToggleSwitch({ name: "show", displayNameKey: "Prop_Show", value: true });
    labelMode = new formattingSettings.AutoDropdown({ name: "labelMode", displayNameKey: "Prop_LabelMode", value: "auto" });
    font = createFontControl(9);
    labelColor = createColorPicker("labelColor", "Prop_Color", DEFAULT_TEXT_COLOR);
    lineColor = createColorPicker("lineColor", "Prop_LineColor", DEFAULT_LINE_COLOR);
    lineWidth = createNumUpDown("lineWidth", "Prop_LineWidth", 2, 1, 10);

    name: string = "dateAxis";
    displayNameKey: string = "Card_DateAxis";
    topLevelSlice = this.show;
    groups: Group[] = [
        new Group({ name: "dateAxisValues", displayNameKey: "Group_Values", slices: [this.labelMode, this.font, this.labelColor] }),
        new Group({ name: "dateAxisLine", displayNameKey: "Group_Line", slices: [this.lineColor, this.lineWidth] })
    ];
}

class LabelsCard extends CompositeCard {
    show = new formattingSettings.ToggleSwitch({ name: "show", displayNameKey: "Prop_Show", value: true });
    position = new formattingSettings.AutoDropdown({ name: "position", displayNameKey: "Prop_Position", value: "above" });
    leaderLines = new formattingSettings.ToggleSwitch({ name: "leaderLines", displayNameKey: "Prop_LeaderLines", value: true });
    maxRows = createNumUpDown("maxRows", "Prop_MaxRows", 4, 1, 20);
    showDate = new formattingSettings.ToggleSwitch({ name: "showDate", displayNameKey: "Prop_ShowDate", value: false });
    font = createFontControl(9);
    labelColor = createColorPicker("labelColor", "Prop_Color", DEFAULT_TEXT_COLOR);

    name: string = "labels";
    displayNameKey: string = "Card_Labels";
    topLevelSlice = this.show;
    groups: Group[] = [
        new Group({ name: "labelsOptions", displayNameKey: "Group_Options", slices: [this.position, this.leaderLines, this.maxRows, this.showDate] }),
        new Group({ name: "labelsValues", displayNameKey: "Group_Values", slices: [this.font, this.labelColor] })
    ];
}

class TodayLineCard extends FormattingSettingsCard {
    show = new formattingSettings.ToggleSwitch({ name: "show", displayNameKey: "Prop_Show", value: false });
    lineColor = createColorPicker("lineColor", "Prop_Color", DEFAULT_TODAY_COLOR);
    lineStyle = new formattingSettings.AutoDropdown({ name: "lineStyle", displayNameKey: "Prop_LineStyle", value: "dashed" });

    name: string = "todayLine";
    displayNameKey: string = "Card_TodayLine";
    topLevelSlice = this.show;
    slices: Array<FormattingSettingsSlice> = [this.lineColor, this.lineStyle];
}

/**
 * Visual settings model class
 */
export class VisualFormattingSettingsModel extends FormattingSettingsModel {
    legend = new LegendCard();
    lanes = new LanesCard();
    colors = new ColorsCard();
    markers = new MarkersCard();
    icons = new IconsCard();
    dateAxis = new DateAxisCard();
    labels = new LabelsCard();
    todayLine = new TodayLineCard();

    cards = [this.legend, this.lanes, this.colors, this.markers, this.icons, this.dateAxis, this.labels, this.todayLine];
}
