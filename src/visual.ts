/*
*  Power BI Visual CLI
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
import { FormattingSettingsService, formattingSettings } from "powerbi-visuals-utils-formattingmodel";
import "./../style/visual.less";

import ILocalizationManager = powerbi.extensibility.ILocalizationManager;
import ISelectionId = powerbi.visuals.ISelectionId;
import ISelectionManager = powerbi.extensibility.ISelectionManager;
import IVisual = powerbi.extensibility.visual.IVisual;
import IVisualEventService = powerbi.extensibility.IVisualEventService;
import IVisualHost = powerbi.extensibility.visual.IVisualHost;
import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualTooltipDataItem = powerbi.extensibility.VisualTooltipDataItem;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;

import { buildTimelineData, getTableDataView, LegendItem, TimelineData, TimelineEvent, VALUE_SETTINGS_PROPERTIES, ValueSettings } from "./dataModel";
import { ValueEditor } from "./editor";
import { createIcon } from "./icons";
import { clamp, countRows, getAxisTicks, getLabelWidths, LabelItem, LabelLayout, layoutLabels, measureText, TextStyle } from "./layout";
import { createShape } from "./shapes";
import {
    DateLabelMode,
    DEFAULT_TEXT_COLOR,
    DEFAULT_LINE_COLOR,
    DEFAULT_TODAY_COLOR,
    LabelPosition,
    LaneOrder,
    LegendPosition,
    LineStyle,
    VisualFormattingSettingsModel
} from "./settings";

const SVG_NS = "http://www.w3.org/2000/svg";
// Space kept above and below the drawing (and inside each lane).
const MARGIN = 8;
const LANE_MARGIN = 4;
// Space between a marker and its first label row, and between labels on a row.
const LABEL_GAP = 4;
const LABEL_ROW_GAP = 6;
const TICK_LENGTH = 4;
// Labels are not shortened below this width: a label that does not fit is hidden rather than cut to a few letters.
const MIN_LABEL_WIDTH = 120;
// Markers that hold an icon are drawn at least this big so that the icon stays readable.
const ICON_MARKER_SIZE = 20;
const MIN_LANE_HEIGHT = 40;
const SCROLLBAR_WIDTH = 8;
const GRID_COLOR = "#EDEBE9";
const ALTERNATE_LANE_COLOR = "#FAF9F8";
const DIMMED_CLASS = "dimmed";

interface Theme {
    isHighContrast: boolean;
    foreground: string;
    background: string;
    foregroundSelected: string;
}

interface AxisLabel {
    /** Centre (or start/end) of the label once kept inside the visual. */
    x: number;
    /** Position of the date on the axis, where its tick is drawn. */
    tickX: number;
    text: string;
    rowIndex: number;
    anchor: "start" | "middle" | "end";
}

interface Lane {
    item: LegendItem;
    legendIndex: number;
    /** Indexes of the lane's events in TimelineData.events, in date order. */
    indexes: number[];
}

/**
 * How to draw one band: the single timeline, or one lane.
 */
interface BandOptions {
    x0: number;
    x1: number;
    xOf: (date: Date) => number;
    /** Date labels drawn under the line. Lanes share one axis at the bottom instead. */
    axisLabels: AxisLabel[] | null;
    ticks: number[];
    gridlines: boolean;
    todayX: number | null;
    todayLabel: boolean;
    margin: number;
    /** Keeps the line at the same height whatever the number of label rows, so that lanes line up. */
    fixedLine: boolean;
    background: string;
}

export class Visual implements IVisual {
    private host: IVisualHost;
    private target: HTMLElement;
    private root: HTMLDivElement;
    private legend: HTMLDivElement;
    private chart: HTMLDivElement;
    private message: HTMLDivElement;
    private editor: ValueEditor;
    private selectionManager: ISelectionManager;
    private localizationManager: ILocalizationManager;
    private eventService: IVisualEventService;
    private formattingSettings: VisualFormattingSettingsModel;
    private formattingSettingsService: FormattingSettingsService;
    private theme: Theme;
    private data: TimelineData | null = null;
    private selectedKeys = new Set<string>();
    private focusedIndex = 0;
    /** Event indexes per lane, used by the keyboard navigation. The single timeline is one lane. */
    private laneIndexes: number[][] = [];
    private laneOf: number[] = [];
    private lanesScrollTop = 0;

    constructor(options: VisualConstructorOptions) {
        this.host = options.host;
        this.target = options.element;
        this.selectionManager = this.host.createSelectionManager();
        this.localizationManager = this.host.createLocalizationManager();
        this.eventService = this.host.eventService;
        this.formattingSettingsService = new FormattingSettingsService(this.localizationManager);
        this.formattingSettings = new VisualFormattingSettingsModel();
        this.target.replaceChildren();

        this.root = document.createElement("div");
        this.root.className = "timeline-visual";

        this.legend = document.createElement("div");
        this.legend.className = "timeline-legend";
        this.root.appendChild(this.legend);

        this.chart = document.createElement("div");
        this.chart.className = "timeline-chart";
        this.chart.setAttribute("role", "listbox");
        this.chart.setAttribute("aria-multiselectable", "true");
        this.root.appendChild(this.chart);

        this.message = document.createElement("div");
        this.message.className = "timeline-message";
        this.root.appendChild(this.message);
        this.editor = new ValueEditor(this.root);

        this.target.appendChild(this.root);

        this.bindChartEvents();
        // Restores the selection when a bookmark is applied.
        this.selectionManager.registerOnSelectCallback((ids: ISelectionId[]) => this.syncSelection(ids));
    }

    public update(options: VisualUpdateOptions) {
        this.eventService.renderingStarted(options);
        try {
            this.formattingSettings = this.formattingSettingsService.populateFormattingSettingsModel(
                VisualFormattingSettingsModel,
                getTableDataView(options.dataViews)
            );
            this.theme = this.getTheme();
            this.root.style.setProperty("--timeline-focus-color", this.theme.isHighContrast ? this.theme.foregroundSelected : "");

            const result = buildTimelineData(options.dataViews, this.host, this.formattingSettings, {
                isHighContrast: this.theme.isHighContrast,
                foreground: this.theme.foreground
            });
            this.data = result.status === "ok" ? result.data : null;
            this.updateFormatPane();

            // Advanced edit mode (Edit button of the visual header): the editor takes the right part of the visual.
            const { width, height } = options.viewport;
            const editing = options.editMode === powerbi.EditMode.Advanced && result.status === "ok";
            const editorWidth = editing ? Math.round(clamp(width * 0.4, 260, 380)) : 0;
            this.root.style.paddingRight = editing ? `${editorWidth}px` : "";
            if (editing) {
                this.editor.render(result.data, editorWidth, {
                    getString: (key) => this.getString(key),
                    iconColor: this.getIconColor(),
                    onChange: (changes) => this.persistValueSettings(changes)
                });
            } else {
                this.editor.hide();
            }

            if (result.status === "landing") {
                this.renderMessage(this.getString("Landing_Text"), this.getString("Landing_Title"));
            } else if (result.status === "message") {
                this.renderMessage(this.getString(result.messageKey), null);
            } else {
                this.render(result.data, Math.max(width - editorWidth, 1), height);
                if (result.data.truncated && this.host.displayWarningIcon) {
                    this.host.displayWarningIcon(this.getString("Warning_Truncated"), this.getString("Warning_TruncatedDetail"));
                }
            }
            this.eventService.renderingFinished(options);
        } catch (error) {
            this.eventService.renderingFailed(options, error instanceof Error ? error.message : String(error));
        }
    }

    /**
     * Returns properties pane formatting model content hierarchies, properties and latest formatting values, Then populate properties pane.
     * This method is called once every time we open properties pane or when the user edit any format property.
     */
    public getFormattingModel(): powerbi.visuals.FormattingModel {
        return this.formattingSettingsService.buildFormattingModel(this.formattingSettings);
    }

    private getTheme(): Theme {
        const palette = this.host.colorPalette;
        const isHighContrast = !!palette.isHighContrast;
        return {
            isHighContrast,
            foreground: isHighContrast && palette.foreground ? palette.foreground.value : DEFAULT_TEXT_COLOR,
            background: isHighContrast && palette.background ? palette.background.value : "#FFFFFF",
            foregroundSelected: isHighContrast && palette.foregroundSelected ? palette.foregroundSelected.value : DEFAULT_TEXT_COLOR
        };
    }

    private getString(key: string): string {
        return this.localizationManager.getDisplayName(key);
    }

    /**
     * With a Legend field, colors and icons are chosen per value in the editor, so the format pane
     * only keeps the settings that apply to all events.
     */
    private updateFormatPane(): void {
        const { colors, icons } = this.formattingSettings;
        const hasLegend = !!this.data && this.data.legendColumnName !== null;
        colors.visible = !hasLegend;
        if (!colors.defaultColor.value.value && this.data && !hasLegend && !this.theme.isHighContrast) {
            colors.defaultColor.value = { value: this.data.events[0].color };
        }
        icons.allEvents.visible = !hasLegend;
    }

    /**
     * Saves the editor's choices with the report, as JSON in static format properties.
     */
    private persistValueSettings(changes: Partial<ValueSettings>): void {
        const propertiesByObject = new Map<string, Record<string, string>>();
        (Object.keys(changes) as Array<keyof ValueSettings>).forEach((key) => {
            const { objectName, propertyName } = VALUE_SETTINGS_PROPERTIES[key];
            const properties = propertiesByObject.get(objectName) || {};
            properties[propertyName] = JSON.stringify(changes[key]);
            propertiesByObject.set(objectName, properties);
        });
        this.host.persistProperties({
            merge: Array.from(propertiesByObject.entries()).map(([objectName, properties]) => ({ objectName, selector: null, properties }))
        });
    }

    private renderMessage(text: string, title: string | null): void {
        this.data = null;
        this.legend.style.display = "none";
        this.chart.style.display = "none";
        this.message.style.display = "flex";
        this.message.replaceChildren();

        if (title) {
            this.message.appendChild(this.createLandingIcon());
            const heading = document.createElement("div");
            heading.className = "timeline-message-title";
            heading.textContent = title;
            this.message.appendChild(heading);
        }
        const paragraph = document.createElement("div");
        paragraph.className = "timeline-message-text";
        paragraph.textContent = text;
        this.message.appendChild(paragraph);
    }

    private createLandingIcon(): SVGSVGElement {
        const icon = document.createElementNS(SVG_NS, "svg");
        icon.setAttribute("class", "timeline-message-icon");
        icon.setAttribute("viewBox", "0 0 64 32");
        icon.setAttribute("aria-hidden", "true");
        const line = document.createElementNS(SVG_NS, "line");
        line.setAttribute("x1", "4");
        line.setAttribute("x2", "60");
        line.setAttribute("y1", "22");
        line.setAttribute("y2", "22");
        icon.appendChild(line);
        [10, 26, 36, 54].forEach((x, index) => {
            const stem = document.createElementNS(SVG_NS, "line");
            stem.setAttribute("x1", x.toString());
            stem.setAttribute("x2", x.toString());
            stem.setAttribute("y1", "18");
            stem.setAttribute("y2", index % 2 === 0 ? "6" : "11");
            stem.setAttribute("class", "stem");
            icon.appendChild(stem);
            icon.appendChild(createShape(index % 2 === 0 ? "circle" : "diamond", x, 22, 7));
        });
        return icon;
    }

    private render(data: TimelineData, width: number, height: number): void {
        this.message.style.display = "none";
        this.chart.style.display = "block";
        this.focusedIndex = Math.min(this.focusedIndex, data.events.length - 1);

        const lanes = this.getLanes(data);
        const legendSize = this.renderLegend(data, width, height, lanes !== null);
        const isVerticalLegend = this.root.classList.contains("legend-left") || this.root.classList.contains("legend-right");
        const chartWidth = Math.max(isVerticalLegend ? width - legendSize : width, 1);
        const chartHeight = Math.max(isVerticalLegend ? height : height - legendSize, 1);

        this.chart.replaceChildren();
        this.chart.setAttribute("aria-label", this.getString("Aria_Timeline"));
        if (lanes) {
            this.renderLanes(data, lanes, chartWidth, chartHeight);
        } else {
            this.renderSingle(data, chartWidth, chartHeight);
        }

        this.syncSelection(this.selectionManager.getSelectionIds() as ISelectionId[]);
    }

    /**
     * One lane per legend value when the Lanes card is on and a Legend field is set, otherwise null.
     */
    private getLanes(data: TimelineData): Lane[] | null {
        if (!this.formattingSettings.lanes.show.value || data.legendColumnName === null || data.legendItems.length === 0) {
            return null;
        }
        const indexOf = new Map(data.events.map((event, index): [TimelineEvent, number] => [event, index]));
        const lanes = data.legendItems.map((item, legendIndex): Lane => ({
            item,
            legendIndex,
            indexes: item.events.map((event) => indexOf.get(event))
        }));
        if (this.formattingSettings.lanes.order.value as LaneOrder === "firstDate") {
            lanes.sort((a, b) => a.indexes[0] - b.indexes[0]);
        }
        return lanes;
    }

    /**
     * Renders the legend and returns the space it takes (height, or width for a left/right legend).
     * Lanes show their own names, so the legend is hidden when lanes are on.
     */
    private renderLegend(data: TimelineData, width: number, height: number, hidden: boolean): number {
        const settings = this.formattingSettings.legend;
        const position = settings.position.value as LegendPosition;
        const isVertical = position.indexOf("Left") === 0 || position.indexOf("Right") === 0;
        this.root.classList.remove("legend-top", "legend-bottom", "legend-left", "legend-right");
        this.root.classList.add(`legend-${position.replace("Center", "").toLowerCase()}`);
        this.legend.classList.toggle("centered", position.indexOf("Center") !== -1);
        this.legend.replaceChildren();

        if (hidden || !settings.show.value || data.legendItems.length === 0) {
            this.legend.style.display = "none";
            return 0;
        }

        const style = this.getTextStyle(settings.font, settings.labelColor);
        this.legend.style.display = "flex";
        this.legend.style.maxHeight = isVertical ? "" : `${Math.round(height * 0.4)}px`;
        this.legend.style.maxWidth = isVertical ? `${Math.round(width * 0.35)}px` : "";
        this.legend.setAttribute("role", "list");
        applyHtmlTextStyle(this.legend, style);

        if (settings.showTitle.value) {
            const title = document.createElement("div");
            title.className = "timeline-legend-title";
            title.textContent = settings.titleText.value || data.legendTitle;
            this.legend.appendChild(title);
        }

        data.legendItems.forEach((item, index) => {
            const element = this.createLegendEntry("timeline-legend-item", item, index, style, !data.legendByEventName);
            if (!element.hasAttribute("role")) {
                element.setAttribute("role", "listitem");
            }
            this.legend.appendChild(element);
        });

        return isVertical ? this.legend.offsetWidth : this.legend.offsetHeight;
    }

    /**
     * A legend marker and name, used by the legend and by the lane names. Clicking it selects the group's events.
     */
    private createLegendEntry(className: string, item: LegendItem, legendIndex: number, style: TextStyle, interactive: boolean): HTMLDivElement {
        const element = document.createElement("div");
        element.className = className;
        element.setAttribute("data-legend-index", legendIndex.toString());
        element.title = item.label;

        const markerSize = item.icon ? Math.round(style.fontSize * 1.2) : Math.max(Math.round(style.fontSize * 0.75), 6);
        const marker = document.createElementNS(SVG_NS, "svg");
        marker.setAttribute("width", `${markerSize + 2}`);
        marker.setAttribute("height", `${markerSize + 2}`);
        marker.setAttribute("aria-hidden", "true");
        const center = markerSize / 2 + 1;
        const shape = createShape(item.icon ? "circle" : item.shape, center, center, markerSize);
        shape.setAttribute("fill", item.color);
        marker.appendChild(shape);
        if (item.icon) {
            marker.appendChild(createIcon(item.icon, center, center, markerSize * 0.62, this.getIconColor()));
        }

        const label = document.createElement("span");
        label.textContent = item.label;
        element.appendChild(marker);
        element.appendChild(label);

        if (interactive && this.allowInteractions()) {
            element.tabIndex = 0;
            element.setAttribute("role", "button");
            element.addEventListener("click", (event) => {
                event.stopPropagation();
                this.selectLegendItem(item, event);
            });
            element.addEventListener("keydown", (event) => {
                if (event.key === "Enter" || event.key === " ") {
                    event.preventDefault();
                    this.selectLegendItem(item, event);
                }
            });
        }
        return element;
    }

    private renderSingle(data: TimelineData, width: number, height: number): void {
        const settings = this.formattingSettings;
        const indexes = data.events.map((_, index) => index);
        const svg = this.appendSvg(this.chart, width, height);
        const { x0, x1, xOf } = this.getScale(data, width);

        const axisStyle = this.getTextStyle(settings.dateAxis.font, settings.dateAxis.labelColor);
        const axisMode = settings.dateAxis.labelMode.value as DateLabelMode;
        const axisLabels = settings.dateAxis.show.value
            ? this.getAxisLabels(data, axisMode, xOf, x0, x1, width, axisStyle)
            : [];

        this.laneIndexes = [indexes];
        this.laneOf = indexes.map(() => 0);
        this.drawBand(svg, data, indexes, width, height, {
            x0,
            x1,
            xOf,
            axisLabels,
            ticks: axisMode === "startEnd" ? [] : axisLabels.map((label) => label.tickX),
            gridlines: false,
            todayX: this.getTodayX(data, xOf),
            todayLabel: true,
            margin: MARGIN,
            fixedLine: false,
            background: this.theme.background
        });
    }

    /**
     * Lanes are HTML rows inside a vertical scroll container, so that wheel and touch scrolling and
     * lane-by-lane snapping come from the browser. The date axis stays at the bottom.
     */
    private renderLanes(data: TimelineData, lanes: Lane[], width: number, height: number): void {
        const settings = this.formattingSettings;
        const nameStyle = this.getTextStyle(settings.legend.font, settings.legend.labelColor);
        const axisStyle = this.getTextStyle(settings.dateAxis.font, settings.dateAxis.labelColor);
        const axisLineHeight = axisStyle.fontSize * 1.25;
        const visibleLanes = Math.max(1, Math.round(settings.lanes.visibleLanes.value));

        const longestName = Math.max(...lanes.map((lane) => measureText(lane.item.label, { ...nameStyle, bold: true })));
        const nameWidth = Math.round(clamp(longestName + nameStyle.fontSize * 1.2 + 26, 60, width * 0.3));
        const scrolls = lanes.length > visibleLanes;
        const bandWidth = Math.max(width - nameWidth - (scrolls ? SCROLLBAR_WIDTH : 0), 1);
        const { x0, x1, xOf } = this.getScale(data, bandWidth);

        const axisMode = settings.dateAxis.labelMode.value as DateLabelMode;
        const axisLabels = settings.dateAxis.show.value
            ? this.getAxisLabels(data, axisMode, xOf, x0, x1, bandWidth, axisStyle)
            : [];
        const ticks = axisMode === "startEnd" ? [] : axisLabels.map((label) => label.tickX);
        const axisRows = axisLabels.reduce((rows, label) => Math.max(rows, label.rowIndex + 1), 0);
        const todayX = this.getTodayX(data, xOf);
        const axisHeight = TICK_LENGTH + 3 + (axisRows + (todayX !== null ? 1 : 0)) * axisLineHeight + LANE_MARGIN;
        const lanesHeight = Math.max(height - axisHeight, MIN_LANE_HEIGHT);
        const laneHeight = Math.max(MIN_LANE_HEIGHT, lanesHeight / Math.min(visibleLanes, lanes.length));

        const container = document.createElement("div");
        container.className = "timeline-lanes";
        container.style.height = `${lanesHeight}px`;
        container.addEventListener("scroll", () => { this.lanesScrollTop = container.scrollTop; });
        this.chart.appendChild(container);

        const alternate = settings.lanes.alternateBackground.value && !this.theme.isHighContrast;
        this.laneIndexes = lanes.map((lane) => lane.indexes);
        this.laneOf = new Array(data.events.length);
        lanes.forEach((lane, laneIndex) => {
            lane.indexes.forEach((index) => { this.laneOf[index] = laneIndex; });
            const isAlternate = alternate && laneIndex % 2 === 0;
            const row = document.createElement("div");
            row.className = isAlternate ? "timeline-lane alternate" : "timeline-lane";
            row.style.height = `${laneHeight}px`;
            row.setAttribute("role", "group");
            row.setAttribute("aria-label", lane.item.label);

            const name = this.createLegendEntry("timeline-lane-name", lane.item, lane.legendIndex, nameStyle, true);
            name.style.width = `${nameWidth}px`;
            applyHtmlTextStyle(name, nameStyle);
            row.appendChild(name);
            container.appendChild(row);

            const svg = this.appendSvg(row, bandWidth, laneHeight - 1);
            this.drawBand(svg, data, lane.indexes, bandWidth, laneHeight - 1, {
                x0,
                x1,
                xOf,
                axisLabels: null,
                ticks,
                gridlines: true,
                todayX,
                todayLabel: false,
                margin: LANE_MARGIN,
                fixedLine: true,
                background: isAlternate ? ALTERNATE_LANE_COLOR : this.theme.background
            });
        });
        container.scrollTop = this.lanesScrollTop;

        // Shared date axis
        const axisSvg = this.appendSvg(this.chart, width, axisHeight);
        const axisGroup = this.appendGroup(axisSvg, "timeline-axis");
        axisGroup.setAttribute("transform", `translate(${nameWidth} 0)`);
        axisGroup.setAttribute("aria-hidden", "true");
        const lineColor = this.theme.isHighContrast ? this.theme.foreground : settings.dateAxis.lineColor.value.value || DEFAULT_LINE_COLOR;
        appendLine(axisGroup, x0, 0.5, x1, 0.5, GRID_COLOR);
        ticks.forEach((x) => appendLine(axisGroup, x, 0, x, TICK_LENGTH, lineColor));
        axisLabels.forEach((label) => {
            const top = TICK_LENGTH + 3 + label.rowIndex * axisLineHeight;
            this.drawLabel(axisGroup, label.text, label.x, label.anchor, top, axisLineHeight, axisStyle, false, null);
        });
        if (todayX !== null) {
            const text = this.getString("Label_Today");
            const labelWidth = measureText(text, axisStyle);
            const color = this.getTodayColor();
            const top = TICK_LENGTH + 3 + axisRows * axisLineHeight;
            const x = clamp(todayX, labelWidth / 2 + 2, bandWidth - labelWidth / 2 - 2);
            this.drawLabel(axisGroup, text, x, "middle", top, axisLineHeight, { ...axisStyle, color }, false, null);
        }
    }

    private getScale(data: TimelineData, width: number): { x0: number; x1: number; xOf: (date: Date) => number } {
        const maxMarkerSize = Math.max(...data.events.map((event) => this.getMarkerSize(event)));
        const paddingX = Math.max(16, maxMarkerSize / 2 + 4);
        const x0 = paddingX;
        const x1 = Math.max(paddingX + 1, width - paddingX);
        const minMs = data.minDate.getTime();
        const rangeMs = data.maxDate.getTime() - minMs;
        const xOf = (date: Date) => (rangeMs > 0 ? x0 + ((date.getTime() - minMs) / rangeMs) * (x1 - x0) : (x0 + x1) / 2);
        return { x0, x1, xOf };
    }

    private getTodayX(data: TimelineData, xOf: (date: Date) => number): number | null {
        const today = new Date();
        const inRange = data.maxDate.getTime() > data.minDate.getTime()
            && today.getTime() >= data.minDate.getTime() && today.getTime() <= data.maxDate.getTime();
        return this.formattingSettings.todayLine.show.value && inRange ? xOf(today) : null;
    }

    private getMarkerSize(event: TimelineEvent): number {
        const size = this.formattingSettings.markers.size.value;
        return event.icon ? Math.max(size, ICON_MARKER_SIZE) : size;
    }

    /**
     * Draws the line, markers, event labels (and for the single timeline, date labels) of a set of events.
     */
    private drawBand(svg: SVGSVGElement, data: TimelineData, indexes: number[], width: number, height: number, band: BandOptions): void {
        const settings = this.formattingSettings;
        const events = indexes.map((index) => data.events[index]);
        const markerSize = Math.max(...events.map((event) => this.getMarkerSize(event)));
        const radius = markerSize / 2;
        const { xs, anchors } = this.getMarkerPositions(events, band.xOf, markerSize, band.x0, band.x1);

        // Date labels (single timeline only)
        const axisStyle = this.getTextStyle(settings.dateAxis.font, settings.dateAxis.labelColor);
        const axisLineHeight = axisStyle.fontSize * 1.25;
        const axisLabels = band.axisLabels || [];
        const axisRows = axisLabels.reduce((rows, label) => Math.max(rows, label.rowIndex + 1), 0);
        const axisTicks = band.axisLabels ? band.ticks : [];
        const axisTop = Math.max(radius, axisTicks.length > 0 ? TICK_LENGTH : 0) + 3;
        const axisArea = axisRows > 0 ? axisTop + axisRows * axisLineHeight : radius;
        const todayLabelHeight = band.todayX !== null && band.todayLabel ? axisLineHeight + 2 : 0;

        // Event labels: as many rows as fit in the height, the others stay in the tooltip.
        const labelStyle = this.getTextStyle(settings.labels.font, settings.labels.labelColor);
        const dateStyle: TextStyle | null = settings.labels.showDate.value
            ? { ...labelStyle, fontSize: labelStyle.fontSize * 0.9, bold: false, underline: false, opacity: 0.75 }
            : null;
        const labelLineHeight = labelStyle.fontSize * 1.25;
        const dateLineHeight = dateStyle ? dateStyle.fontSize * 1.2 : 0;
        const rowHeight = labelLineHeight + dateLineHeight;
        const position = settings.labels.position.value as LabelPosition;
        const isAbove = (index: number) => position === "above" || (position === "alternate" && index % 2 === 0);
        const positions = settings.labels.show.value ? events.map((_, index) => index) : [];
        const abovePositions = positions.filter(isAbove);
        const belowPositions = positions.filter((index) => !isAbove(index));
        const fixedAbove = radius + (abovePositions.length > 0 ? LABEL_GAP : 0);
        const fixedBelow = axisArea + (belowPositions.length > 0 ? LABEL_GAP : 0);
        const capacity = Math.max(0, Math.floor((height - 2 * band.margin - todayLabelHeight - fixedAbove - fixedBelow) / rowHeight));
        const maxRows = settings.labels.maxRows.value;
        const aboveCapacity = belowPositions.length === 0 ? capacity : abovePositions.length === 0 ? 0 : Math.ceil(capacity / 2);
        const maxAbove = Math.min(aboveCapacity, maxRows);
        const maxBelow = Math.min(capacity - aboveCapacity, maxRows);
        const layoutSide = (sidePositions: number[], rows: number) => this.layoutEventLabels(
            events, indexes, sidePositions, xs, anchors, width, labelStyle, dateStyle, rows
        );
        const aboveLabels = layoutSide(abovePositions, maxAbove);
        const belowLabels = layoutSide(belowPositions, maxBelow);

        // The drawing is centred vertically. Lanes reserve all their label rows so that their lines align.
        const usedAbove = band.fixedLine && abovePositions.length > 0 ? maxAbove : countRows(aboveLabels);
        const usedBelow = band.fixedLine && belowPositions.length > 0 ? maxBelow : countRows(belowLabels);
        const aboveHeight = fixedAbove + usedAbove * rowHeight;
        const belowHeight = fixedBelow + usedBelow * rowHeight;
        const freeHeight = height - 2 * band.margin - todayLabelHeight - aboveHeight - belowHeight;
        const lineY = band.margin + todayLabelHeight + Math.max(0, freeHeight / 2) + aboveHeight;

        if (band.gridlines && !this.theme.isHighContrast) {
            const grid = this.appendGroup(svg, "timeline-grid");
            band.ticks.forEach((x) => appendLine(grid, x, 0, x, height, GRID_COLOR));
        }
        if (band.todayX !== null) {
            this.drawTodayLine(svg, band.todayX, width, band.todayLabel ? band.margin : null, band.todayLabel ? height - band.margin : height, axisStyle);
        }
        this.drawAxisLine(svg, band.x0 - radius, band.x1 + radius, lineY, axisTicks);

        // Leader lines go under the labels, and labels get a background so that no line crosses their text.
        const leaderLines = settings.labels.leaderLines.value;
        const leaderGroup = this.appendGroup(svg, "timeline-leader-lines");
        const labelGroup = this.appendGroup(svg, "timeline-labels");
        labelGroup.setAttribute("aria-hidden", "true");
        axisLabels.forEach((label) => {
            const top = lineY + axisTop + label.rowIndex * axisLineHeight;
            this.drawLabel(labelGroup, label.text, label.x, label.anchor, top, axisLineHeight, axisStyle, leaderLines, null, band.background);
        });
        const drawEventLabel = (layout: LabelLayout, top: number, leaderFrom: number, leaderTo: number) => {
            if (leaderLines) {
                this.drawLeaderLine(leaderGroup, layout, leaderFrom, leaderTo, data.events[layout.key]);
            }
            if (leaderLines) {
                appendRect(labelGroup, layout.x - layout.width / 2 - 2, top, layout.width + 4, layout.subText ? rowHeight : labelLineHeight, band.background);
            }
            this.drawLabel(labelGroup, layout.text, layout.x, "middle", top, labelLineHeight, labelStyle, false, layout.key);
            if (layout.subText) {
                this.drawLabel(labelGroup, layout.subText, layout.x, "middle", top + labelLineHeight, dateLineHeight, dateStyle, false, layout.key);
            }
        };
        aboveLabels.forEach((layout) => {
            const bottom = lineY - radius - LABEL_GAP - layout.rowIndex * rowHeight;
            const top = bottom - (layout.subText ? rowHeight : labelLineHeight);
            drawEventLabel(layout, top, lineY - radius, top + (layout.subText ? rowHeight : labelLineHeight));
        });
        belowLabels.forEach((layout) => {
            const top = lineY + fixedBelow + layout.rowIndex * rowHeight;
            drawEventLabel(layout, top, lineY + radius, top);
        });

        this.drawMarkers(svg, data, indexes, xs, lineY);
    }

    /**
     * Same-date events are spread around their date. The whole group is shifted inward at the edges
     * (instead of clamping each marker) so that no two markers share the same x.
     */
    private getMarkerPositions(
        events: TimelineEvent[],
        xOf: (date: Date) => number,
        markerSize: number,
        x0: number,
        x1: number
    ): { xs: number[]; anchors: number[] } {
        const xs: number[] = new Array(events.length);
        const anchors: number[] = new Array(events.length);
        let start = 0;
        while (start < events.length) {
            let end = start;
            while (end + 1 < events.length && events[end + 1].date.getTime() === events[start].date.getTime()) {
                end += 1;
            }
            const count = end - start + 1;
            const step = count > 1 ? Math.min(markerSize + 2, (x1 - x0) / (count - 1)) : 0;
            const halfSpread = (step * (count - 1)) / 2;
            const center = clamp(xOf(events[start].date), x0 + halfSpread, x1 - halfSpread);
            for (let index = start; index <= end; index++) {
                xs[index] = center - halfSpread + (index - start) * step;
                anchors[index] = center;
            }
            start = end + 1;
        }
        return { xs, anchors };
    }

    private getAxisLabels(
        data: TimelineData,
        mode: DateLabelMode,
        xOf: (date: Date) => number,
        x0: number,
        x1: number,
        width: number,
        style: TextStyle
    ): AxisLabel[] {
        if (mode === "startEnd") {
            if (data.maxDate.getTime() === data.minDate.getTime()) {
                const center = (x0 + x1) / 2;
                return [{ x: center, tickX: center, text: data.formatDate(data.minDate), rowIndex: 0, anchor: "middle" }];
            }
            return [
                { x: x0, tickX: x0, text: data.formatDate(data.minDate), rowIndex: 0, anchor: "start" },
                { x: x1, tickX: x1, text: data.formatDate(data.maxDate), rowIndex: 0, anchor: "end" }
            ];
        }

        if (mode === "eventDates") {
            // One label per date, centred under that date's markers.
            const { anchors } = this.getMarkerPositions(data.events, xOf, this.formattingSettings.markers.size.value, x0, x1);
            const items: LabelItem[] = [];
            data.events.forEach((event, index) => {
                if (index === 0 || event.date.getTime() !== data.events[index - 1].date.getTime()) {
                    items.push({ x: anchors[index], maxWidth: 0, text: event.dateText, key: index });
                }
            });
            const widths = getLabelWidths(items.map((item) => item.x), MIN_LABEL_WIDTH, width - 4);
            items.forEach((item, index) => { item.maxWidth = widths[index]; });
            return layoutLabels(items, width, style, LABEL_ROW_GAP, 2)
                .map((layout) => ({ x: layout.x, tickX: layout.anchorX, text: layout.text, rowIndex: layout.rowIndex, anchor: "middle" }));
        }

        return getAxisTicks(data.minDate, data.maxDate, x0, x1, style, this.host.locale).map((tick) => {
            const labelWidth = measureText(tick.text, style);
            const x = xOf(tick.date);
            return { x: clamp(x, labelWidth / 2 + 2, width - labelWidth / 2 - 2), tickX: x, text: tick.text, rowIndex: 0, anchor: "middle" };
        });
    }

    /**
     * `positions` are positions in `events`; the label keys are the matching indexes in TimelineData.events.
     */
    private layoutEventLabels(
        events: TimelineEvent[],
        indexes: number[],
        positions: number[],
        xs: number[],
        anchors: number[],
        width: number,
        style: TextStyle,
        dateStyle: TextStyle | null,
        maxRows: number
    ): LabelLayout[] {
        if (positions.length === 0 || maxRows <= 0) {
            return [];
        }
        const widths = getLabelWidths(positions.map((position) => anchors[position]), MIN_LABEL_WIDTH, width - 4);
        const items = positions.map((position, itemIndex): LabelItem => ({
            x: xs[position],
            maxWidth: widths[itemIndex],
            text: events[position].name,
            subText: dateStyle ? events[position].dateText : undefined,
            key: indexes[position]
        }));
        return layoutLabels(items, width, style, LABEL_ROW_GAP, maxRows, dateStyle || undefined);
    }

    private appendSvg(parent: HTMLElement, width: number, height: number): SVGSVGElement {
        const svg = document.createElementNS(SVG_NS, "svg");
        svg.classList.add("timeline-svg");
        svg.setAttribute("width", width.toString());
        svg.setAttribute("height", height.toString());
        parent.appendChild(svg);
        return svg;
    }

    private appendGroup(svg: SVGElement, className: string): SVGGElement {
        const group = document.createElementNS(SVG_NS, "g");
        group.setAttribute("class", className);
        svg.appendChild(group);
        return group;
    }

    private getTodayColor(): string {
        return this.theme.isHighContrast ? this.theme.foreground : this.formattingSettings.todayLine.lineColor.value.value || DEFAULT_TODAY_COLOR;
    }

    /**
     * Draws the today line from `labelTop` (below its label) or from the top when there is no label, down to `bottom`.
     */
    private drawTodayLine(svg: SVGSVGElement, x: number, width: number, labelTop: number | null, bottom: number, style: TextStyle): void {
        const color = this.getTodayColor();
        const group = this.appendGroup(svg, "timeline-today");
        group.setAttribute("aria-hidden", "true");

        const line = appendLine(group, x, labelTop !== null ? labelTop + style.fontSize * 1.25 + 2 : 0, x, bottom, color);
        line.setAttribute("stroke-width", "1.5");
        const dashes: Record<LineStyle, string> = { solid: "", dashed: "5 4", dotted: "1 3" };
        const dash = dashes[this.formattingSettings.todayLine.lineStyle.value as LineStyle];
        if (dash) {
            line.setAttribute("stroke-dasharray", dash);
            line.setAttribute("stroke-linecap", "round");
        }

        if (labelTop !== null) {
            const text = this.getString("Label_Today");
            const labelWidth = measureText(text, style);
            const label = document.createElementNS(SVG_NS, "text");
            label.setAttribute("x", clamp(x, labelWidth / 2 + 2, width - labelWidth / 2 - 2).toString());
            label.setAttribute("y", labelTop.toString());
            label.setAttribute("text-anchor", "middle");
            label.setAttribute("dominant-baseline", "hanging");
            applySvgTextStyle(label, { ...style, color });
            label.textContent = text;
            group.appendChild(label);
        }
    }

    private drawAxisLine(svg: SVGSVGElement, startX: number, endX: number, lineY: number, ticks: number[]): void {
        const settings = this.formattingSettings.dateAxis;
        const color = this.theme.isHighContrast ? this.theme.foreground : settings.lineColor.value.value || DEFAULT_LINE_COLOR;
        const lineWidth = settings.lineWidth.value;
        const group = this.appendGroup(svg, "timeline-axis");
        group.setAttribute("aria-hidden", "true");

        const line = appendLine(group, startX, lineY, endX, lineY, color);
        line.setAttribute("stroke-width", lineWidth.toString());
        line.setAttribute("stroke-linecap", "round");
        ticks.forEach((x) => appendLine(group, x, lineY, x, lineY + lineWidth / 2 + TICK_LENGTH, color));
    }

    private drawLeaderLine(group: SVGGElement, layout: LabelLayout, fromY: number, toY: number, event: TimelineEvent): void {
        const line = appendLine(group, layout.anchorX, fromY, layout.anchorX, toY, this.theme.isHighContrast ? this.theme.foreground : event.color);
        line.setAttribute("stroke-opacity", this.theme.isHighContrast ? "1" : "0.5");
        line.setAttribute("data-index", layout.key.toString());
    }

    /**
     * Draws a one-line label vertically centred in the row that starts at `top`.
     * `eventIndex` links event labels to their event, so that they are dimmed with it.
     */
    private drawLabel(
        group: SVGGElement,
        content: string,
        x: number,
        anchor: "start" | "middle" | "end",
        top: number,
        rowHeight: number,
        style: TextStyle,
        withBackground: boolean,
        eventIndex: number | null,
        background: string = this.theme.background
    ): void {
        if (withBackground) {
            const width = measureText(content, style);
            const left = anchor === "start" ? x : anchor === "end" ? x - width : x - width / 2;
            appendRect(group, left - 2, top, width + 4, rowHeight, background);
        }
        const text = document.createElementNS(SVG_NS, "text");
        text.setAttribute("x", x.toString());
        text.setAttribute("y", (top + rowHeight / 2).toString());
        text.setAttribute("text-anchor", anchor);
        text.setAttribute("dominant-baseline", "central");
        if (eventIndex !== null) {
            text.setAttribute("data-index", eventIndex.toString());
            text.setAttribute("class", "timeline-event-label");
        }
        applySvgTextStyle(text, style);
        text.textContent = content;
        group.appendChild(text);
    }

    private getIconColor(): string {
        return this.theme.isHighContrast ? this.theme.background : "#FFFFFF";
    }

    private drawMarkers(svg: SVGSVGElement, data: TimelineData, indexes: number[], xs: number[], lineY: number): void {
        const group = this.appendGroup(svg, "timeline-markers");
        const interactive = this.allowInteractions();
        const baseSize = this.formattingSettings.markers.size.value;
        // When thousands of markers overlap, their white outlines would hide every fill but the last one.
        const crowded = xs.length > 1 && (xs[xs.length - 1] - xs[0]) / (xs.length - 1) < baseSize / 2;
        const strokeWidth = this.theme.isHighContrast ? "2" : crowded ? "0" : "1";
        indexes.forEach((index, position) => {
            const event = data.events[index];
            const size = this.getMarkerSize(event);
            const marker = document.createElementNS(SVG_NS, "g");
            marker.setAttribute("class", "timeline-marker");
            marker.setAttribute("data-index", index.toString());
            marker.setAttribute("role", "option");
            marker.setAttribute("aria-label", [event.name, event.dateText, event.legend].filter((part) => !!part).join(", "));
            if (interactive) {
                marker.setAttribute("tabindex", index === this.focusedIndex ? "0" : "-1");
            }
            const shape = createShape(event.icon ? "circle" : event.shape, xs[position], lineY, size);
            shape.setAttribute("class", "timeline-marker-shape");
            shape.setAttribute("fill", event.color);
            shape.setAttribute("stroke", this.theme.isHighContrast ? this.theme.foreground : this.theme.background);
            shape.setAttribute("stroke-width", strokeWidth);
            marker.appendChild(shape);
            if (event.icon) {
                marker.appendChild(createIcon(event.icon, xs[position], lineY, size * 0.62, this.getIconColor()));
            }
            group.appendChild(marker);
        });
    }

    // ----- Interactions -----

    private allowInteractions(): boolean {
        const capabilities = this.host.hostCapabilities;
        return !capabilities || capabilities.allowInteractions !== false;
    }

    private bindChartEvents(): void {
        this.chart.addEventListener("click", (event: MouseEvent) => {
            if (!this.data || !this.allowInteractions()) {
                return;
            }
            const index = this.getEventIndex(event.target);
            if (index === null) {
                this.selectionManager.clear().then(() => this.syncSelection([]));
                return;
            }
            this.focusMarker(index);
            this.selectEvents([this.data.events[index].selectionId], isMultiSelect(event));
        });

        this.chart.addEventListener("contextmenu", (event: MouseEvent) => {
            if (!this.data || !this.allowInteractions()) {
                return;
            }
            const index = this.getEventIndex(event.target);
            const selectionId = index !== null ? this.data.events[index].selectionId : {};
            this.selectionManager.showContextMenu(selectionId as ISelectionId, { x: event.clientX, y: event.clientY });
            event.preventDefault();
        });

        this.chart.addEventListener("pointermove", (event: PointerEvent) => {
            const index = this.getEventIndex(event.target);
            if (index === null) {
                this.hideTooltip(event.pointerType === "touch");
                return;
            }
            this.showTooltip(index, [event.clientX, event.clientY], event.pointerType === "touch");
        });
        this.chart.addEventListener("pointerleave", (event: PointerEvent) => this.hideTooltip(event.pointerType === "touch"));

        this.chart.addEventListener("keydown", (event: KeyboardEvent) => this.onMarkerKeyDown(event));
        // Keyboard users get the tooltip of the focused marker.
        this.chart.addEventListener("focusin", (event: FocusEvent) => {
            const index = this.getEventIndex(event.target);
            if (index !== null) {
                const rect = (event.target as Element).getBoundingClientRect();
                this.showTooltip(index, [rect.left + rect.width / 2, rect.top + rect.height / 2], false);
            }
        });
        this.chart.addEventListener("focusout", () => this.hideTooltip(false));
    }

    /**
     * Index of the event under a marker or an event label, or null for the background.
     */
    private getEventIndex(target: EventTarget | null): number | null {
        const element = target instanceof Element ? target.closest(".timeline-marker, .timeline-event-label") : null;
        if (!element || !this.data) {
            return null;
        }
        const index = Number(element.getAttribute("data-index"));
        return Number.isInteger(index) && index < this.data.events.length ? index : null;
    }

    /**
     * Left/Right move to the previous/next event of the lane, Up/Down to the closest event of the lane above/below.
     * With a single timeline, Up/Down behave like Left/Right.
     */
    private onMarkerKeyDown(event: KeyboardEvent): void {
        const index = this.getEventIndex(event.target);
        if (index === null || !this.allowInteractions()) {
            return;
        }
        const laneIndex = this.laneOf[index] || 0;
        const lane = this.laneIndexes[laneIndex] || [index];
        const position = lane.indexOf(index);
        const singleLane = this.laneIndexes.length <= 1;
        let next: number | null = null;
        switch (event.key) {
            case "ArrowRight":
                next = lane[Math.min(position + 1, lane.length - 1)];
                break;
            case "ArrowLeft":
                next = lane[Math.max(position - 1, 0)];
                break;
            case "ArrowDown":
                next = singleLane ? lane[Math.min(position + 1, lane.length - 1)] : this.getClosestInLane(index, laneIndex + 1);
                break;
            case "ArrowUp":
                next = singleLane ? lane[Math.max(position - 1, 0)] : this.getClosestInLane(index, laneIndex - 1);
                break;
            case "Home":
                next = lane[0];
                break;
            case "End":
                next = lane[lane.length - 1];
                break;
            case "Enter":
            case " ":
                this.selectEvents([this.data.events[index].selectionId], isMultiSelect(event));
                break;
            default:
                return;
        }
        event.preventDefault();
        if (next !== null) {
            this.focusMarker(next);
        }
    }

    private getClosestInLane(index: number, laneIndex: number): number {
        const lane = this.laneIndexes[laneIndex];
        if (!lane) {
            return index;
        }
        const time = this.data.events[index].date.getTime();
        return lane.reduce((best, candidate) =>
            Math.abs(this.data.events[candidate].date.getTime() - time) < Math.abs(this.data.events[best].date.getTime() - time) ? candidate : best);
    }

    /**
     * Roving tabindex: only one marker is in the tab order, the arrow keys move between markers.
     */
    private focusMarker(index: number): void {
        const previous = this.chart.querySelector(".timeline-marker[tabindex=\"0\"]");
        if (previous) {
            previous.setAttribute("tabindex", "-1");
        }
        this.focusedIndex = index;
        const marker = this.chart.querySelector<SVGGElement>(`.timeline-marker[data-index="${index}"]`);
        if (marker) {
            marker.setAttribute("tabindex", "0");
            marker.focus();
        }
    }

    private selectEvents(ids: ISelectionId[], multiSelect: boolean): void {
        this.selectionManager.select(ids, multiSelect).then((selected: ISelectionId[]) => this.syncSelection(selected));
    }

    /**
     * Clicking a legend item selects all its events; clicking it again clears the selection.
     */
    private selectLegendItem(item: LegendItem, event: MouseEvent | KeyboardEvent): void {
        const ids = item.events.map((timelineEvent) => timelineEvent.selectionId);
        const alreadySelected = this.selectedKeys.size === ids.length && ids.every((id) => this.selectedKeys.has(id.getKey()));
        if (alreadySelected && !isMultiSelect(event)) {
            this.selectionManager.clear().then(() => this.syncSelection([]));
            return;
        }
        this.selectEvents(ids, isMultiSelect(event));
    }

    private syncSelection(ids: ISelectionId[]): void {
        this.selectedKeys = new Set((ids || []).map((id) => id.getKey()));
        this.applySelectionStyles();
    }

    /**
     * Unselected events are dimmed, like in Power BI's built-in visuals. In high contrast mode,
     * unselected markers are drawn hollow and selected ones in the selection color.
     */
    private applySelectionStyles(): void {
        if (!this.data) {
            return;
        }
        const events = this.data.events;
        const hasSelection = this.selectedKeys.size > 0;
        const isSelected = (index: number) => !hasSelection || this.selectedKeys.has(events[index].selectionId.getKey());

        this.chart.querySelectorAll("[data-index]").forEach((element) => {
            const selected = isSelected(Number(element.getAttribute("data-index")));
            if (element.classList.contains("timeline-marker")) {
                element.setAttribute("aria-selected", String(hasSelection && selected));
                if (this.theme.isHighContrast) {
                    const fill = !hasSelection ? this.theme.foreground : selected ? this.theme.foregroundSelected : this.theme.background;
                    element.firstElementChild.setAttribute("fill", fill);
                    if (element.children.length > 1) {
                        element.lastElementChild.setAttribute("fill", fill === this.theme.background ? this.theme.foreground : this.theme.background);
                    }
                    return;
                }
            }
            element.classList.toggle(DIMMED_CLASS, !selected);
        });

        this.root.querySelectorAll(".timeline-legend-item, .timeline-lane-name").forEach((element) => {
            const item = this.data.legendItems[Number(element.getAttribute("data-legend-index"))];
            const selected = !hasSelection || item.events.some((event) => this.selectedKeys.has(event.selectionId.getKey()));
            element.classList.toggle(DIMMED_CLASS, !selected);
            if (element.getAttribute("role") === "button") {
                element.setAttribute("aria-pressed", String(hasSelection && selected));
            }
        });
    }

    private showTooltip(index: number, clientPoint: [number, number], isTouchEvent: boolean): void {
        const tooltipService = this.host.tooltipService;
        if (!this.data || !tooltipService || !tooltipService.enabled()) {
            return;
        }
        const event = this.data.events[index];
        const rect = this.target.getBoundingClientRect();
        tooltipService.show({
            coordinates: [clientPoint[0] - rect.left, clientPoint[1] - rect.top],
            isTouchEvent,
            dataItems: this.getTooltipItems(event),
            identities: [event.selectionId]
        });
    }

    private hideTooltip(isTouchEvent: boolean): void {
        if (this.host.tooltipService) {
            this.host.tooltipService.hide({ isTouchEvent, immediately: true });
        }
    }

    private getTooltipItems(event: TimelineEvent): VisualTooltipDataItem[] {
        const data = this.data;
        const items: VisualTooltipDataItem[] = [
            { displayName: data.eventColumnName, value: event.name, color: event.color },
            { displayName: data.dateColumnName, value: event.dateText }
        ];
        if (data.legendColumnName !== null && event.legend !== null) {
            items.push({ displayName: data.legendColumnName, value: event.legend });
        }
        event.tooltipFields.forEach((field) => items.push({ displayName: field.displayName, value: field.value }));
        return items;
    }

    private getTextStyle(font: formattingSettings.FontControl, color: formattingSettings.ColorPicker): TextStyle {
        return {
            fontFamily: font.fontFamily.value,
            // The format pane gives font sizes in points.
            fontSize: (font.fontSize.value * 4) / 3,
            bold: !!(font.bold && font.bold.value),
            italic: !!(font.italic && font.italic.value),
            underline: !!(font.underline && font.underline.value),
            color: this.theme.isHighContrast ? this.theme.foreground : color.value.value || DEFAULT_TEXT_COLOR
        };
    }
}

function isMultiSelect(event: MouseEvent | KeyboardEvent): boolean {
    return event.ctrlKey || event.metaKey || event.shiftKey;
}

function appendLine(parent: SVGElement, x1: number, y1: number, x2: number, y2: number, color: string): SVGLineElement {
    const line = document.createElementNS(SVG_NS, "line");
    line.setAttribute("x1", x1.toString());
    line.setAttribute("y1", y1.toString());
    line.setAttribute("x2", x2.toString());
    line.setAttribute("y2", y2.toString());
    line.setAttribute("stroke", color);
    line.setAttribute("stroke-width", "1");
    parent.appendChild(line);
    return line;
}

function appendRect(parent: SVGElement, x: number, y: number, width: number, height: number, fill: string): SVGRectElement {
    const rect = document.createElementNS(SVG_NS, "rect");
    rect.setAttribute("x", x.toString());
    rect.setAttribute("y", y.toString());
    rect.setAttribute("width", width.toString());
    rect.setAttribute("height", height.toString());
    rect.setAttribute("fill", fill);
    parent.appendChild(rect);
    return rect;
}

function applySvgTextStyle(element: SVGTextElement, style: TextStyle): void {
    element.style.fontFamily = style.fontFamily;
    element.style.fontSize = `${style.fontSize}px`;
    element.style.fontWeight = style.bold ? "bold" : "normal";
    element.style.fontStyle = style.italic ? "italic" : "normal";
    element.style.textDecoration = style.underline ? "underline" : "none";
    element.style.fill = style.color;
    element.style.fillOpacity = style.opacity !== undefined ? style.opacity.toString() : "";
}

function applyHtmlTextStyle(element: HTMLElement, style: TextStyle): void {
    element.style.fontFamily = style.fontFamily;
    element.style.fontSize = `${style.fontSize}px`;
    element.style.fontWeight = style.bold ? "bold" : "normal";
    element.style.fontStyle = style.italic ? "italic" : "normal";
    element.style.textDecoration = style.underline ? "underline" : "none";
    element.style.color = style.color;
}
