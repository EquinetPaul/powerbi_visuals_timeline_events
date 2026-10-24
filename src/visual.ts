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

import { buildTimelineData, getTableDataView, LegendItem, TimelineData, TimelineEvent } from "./dataModel";
import { clamp, countRows, getAxisTicks, getLabelWidths, LabelItem, LabelLayout, layoutLabels, measureText, TextStyle } from "./layout";
import { createShape } from "./shapes";
import {
    DateLabelMode,
    DEFAULT_TEXT_COLOR,
    DEFAULT_LINE_COLOR,
    DEFAULT_TODAY_COLOR,
    LabelPosition,
    LegendPosition,
    LineStyle,
    VisualFormattingSettingsModel
} from "./settings";

const SVG_NS = "http://www.w3.org/2000/svg";
// Space kept above and below the drawing.
const MARGIN = 8;
// Space between a marker and its first label row, and between labels on a row.
const LABEL_GAP = 4;
const LABEL_ROW_GAP = 6;
const TICK_LENGTH = 4;
// Labels are not shortened below this width: a label that does not fit is hidden rather than cut to a few letters.
const MIN_LABEL_WIDTH = 80;
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

export class Visual implements IVisual {
    private host: IVisualHost;
    private target: HTMLElement;
    private root: HTMLDivElement;
    private legend: HTMLDivElement;
    private chart: HTMLDivElement;
    private svg: SVGSVGElement;
    private message: HTMLDivElement;
    private selectionManager: ISelectionManager;
    private localizationManager: ILocalizationManager;
    private eventService: IVisualEventService;
    private formattingSettings: VisualFormattingSettingsModel;
    private formattingSettingsService: FormattingSettingsService;
    private theme: Theme;
    private data: TimelineData | null = null;
    private selectedKeys = new Set<string>();
    private focusedIndex = 0;

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
        this.svg = document.createElementNS(SVG_NS, "svg");
        this.svg.classList.add("timeline-svg");
        this.svg.setAttribute("role", "listbox");
        this.svg.setAttribute("aria-multiselectable", "true");
        this.chart.appendChild(this.svg);
        this.root.appendChild(this.chart);

        this.message = document.createElement("div");
        this.message.className = "timeline-message";
        this.root.appendChild(this.message);

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
            this.updateColorSlices();

            if (result.status === "landing") {
                this.renderMessage(this.getString("Landing_Text"), this.getString("Landing_Title"));
            } else if (result.status === "message") {
                this.renderMessage(this.getString(result.messageKey), null);
            } else {
                this.render(result.data, options.viewport.width, options.viewport.height);
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
     * Shows one color picker per legend value in the Colors card, or the default color when there is no legend.
     */
    private updateColorSlices(): void {
        const colors = this.formattingSettings.colors;
        const legendItems = this.data && this.data.legendColumnName !== null ? this.data.legendItems : [];
        colors.defaultColor.visible = legendItems.length === 0;
        if (!colors.defaultColor.value.value && this.data && legendItems.length === 0 && !this.theme.isHighContrast) {
            colors.defaultColor.value = { value: this.data.events[0].color };
        }
        colors.slices = [
            colors.defaultColor,
            ...legendItems.filter((item) => item.colorSelector !== null).map((item) => new formattingSettings.ColorPicker({
                name: "fill",
                displayName: item.label,
                value: { value: item.color },
                selector: item.colorSelector
            }))
        ];
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

        const legendSize = this.renderLegend(data, width, height);
        const isVerticalLegend = this.root.classList.contains("legend-left") || this.root.classList.contains("legend-right");
        const chartWidth = Math.max(isVerticalLegend ? width - legendSize : width, 1);
        const chartHeight = Math.max(isVerticalLegend ? height : height - legendSize, 1);
        this.renderChart(data, chartWidth, chartHeight);

        this.syncSelection(this.selectionManager.getSelectionIds() as ISelectionId[]);
    }

    /**
     * Renders the legend and returns the space it takes (height, or width for a left/right legend).
     */
    private renderLegend(data: TimelineData, width: number, height: number): number {
        const settings = this.formattingSettings.legend;
        const position = settings.position.value as LegendPosition;
        const isVertical = position.indexOf("Left") === 0 || position.indexOf("Right") === 0;
        this.root.classList.remove("legend-top", "legend-bottom", "legend-left", "legend-right");
        this.root.classList.add(`legend-${position.replace("Center", "").toLowerCase()}`);
        this.legend.classList.toggle("centered", position.indexOf("Center") !== -1);
        this.legend.replaceChildren();

        if (!settings.show.value || data.legendItems.length === 0) {
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

        const markerSize = Math.max(Math.round(style.fontSize * 0.75), 6);
        const interactive = this.allowInteractions() && !data.legendByEventName;
        data.legendItems.forEach((item, index) => {
            const element = document.createElement("div");
            element.className = "timeline-legend-item";
            element.setAttribute("role", "listitem");
            element.setAttribute("data-legend-index", index.toString());
            element.title = item.label;

            const marker = document.createElementNS(SVG_NS, "svg");
            marker.setAttribute("width", `${markerSize + 2}`);
            marker.setAttribute("height", `${markerSize + 2}`);
            marker.setAttribute("aria-hidden", "true");
            const shape = createShape(item.shape, markerSize / 2 + 1, markerSize / 2 + 1, markerSize);
            shape.setAttribute("fill", item.color);
            marker.appendChild(shape);

            const label = document.createElement("span");
            label.textContent = item.label;

            element.appendChild(marker);
            element.appendChild(label);
            if (interactive) {
                element.tabIndex = 0;
                element.setAttribute("role", "button");
                element.addEventListener("click", (event) => this.selectLegendItem(item, event));
                element.addEventListener("keydown", (event) => {
                    if (event.key === "Enter" || event.key === " ") {
                        event.preventDefault();
                        this.selectLegendItem(item, event);
                    }
                });
            }
            this.legend.appendChild(element);
        });

        return isVertical ? this.legend.offsetWidth : this.legend.offsetHeight;
    }

    private renderChart(data: TimelineData, width: number, height: number): void {
        const settings = this.formattingSettings;
        this.svg.replaceChildren();
        this.svg.setAttribute("width", width.toString());
        this.svg.setAttribute("height", height.toString());
        this.svg.setAttribute("aria-label", this.getString("Aria_Timeline"));

        const markerSize = settings.markers.size.value;
        const radius = markerSize / 2;
        const paddingX = Math.max(16, radius + 4);
        const x0 = paddingX;
        const x1 = Math.max(paddingX + 1, width - paddingX);
        const minMs = data.minDate.getTime();
        const rangeMs = data.maxDate.getTime() - minMs;
        const xOf = (date: Date) => (rangeMs > 0 ? x0 + ((date.getTime() - minMs) / rangeMs) * (x1 - x0) : (x0 + x1) / 2);
        const { xs, anchors } = this.getMarkerPositions(data.events, xOf, markerSize, x0, x1);

        // Date axis labels
        const axisStyle = this.getTextStyle(settings.dateAxis.font, settings.dateAxis.labelColor);
        const axisLineHeight = axisStyle.fontSize * 1.25;
        const axisMode = settings.dateAxis.labelMode.value as DateLabelMode;
        const axisLabels = settings.dateAxis.show.value
            ? this.getAxisLabels(data, axisMode, anchors, xOf, x0, x1, width, axisStyle)
            : [];
        const axisRows = axisLabels.reduce((rows, label) => Math.max(rows, label.rowIndex + 1), 0);
        const ticks = axisMode === "startEnd" ? [] : axisLabels.map((label) => label.tickX);
        const axisTop = Math.max(radius, ticks.length > 0 ? TICK_LENGTH : 0) + 3;
        const axisArea = axisRows > 0 ? axisTop + axisRows * axisLineHeight : radius;

        // Today line
        const today = new Date();
        const showToday = settings.todayLine.show.value && rangeMs > 0
            && today.getTime() >= minMs && today.getTime() <= data.maxDate.getTime();
        const todayLabelHeight = showToday ? axisLineHeight + 2 : 0;

        // Event labels: as many rows as fit in the height, the others stay in the tooltip.
        const labelStyle = this.getTextStyle(settings.labels.font, settings.labels.labelColor);
        const labelLineHeight = labelStyle.fontSize * 1.25;
        const position = settings.labels.position.value as LabelPosition;
        const showLabels = settings.labels.show.value;
        const isAbove = (index: number) => position === "above" || (position === "alternate" && index % 2 === 0);
        const indexes = showLabels ? data.events.map((_, index) => index) : [];
        const aboveIndexes = indexes.filter(isAbove);
        const belowIndexes = indexes.filter((index) => !isAbove(index));
        const fixedAbove = radius + (aboveIndexes.length > 0 ? LABEL_GAP : 0);
        const fixedBelow = axisArea + (belowIndexes.length > 0 ? LABEL_GAP : 0);
        const capacity = Math.max(0, Math.floor((height - 2 * MARGIN - todayLabelHeight - fixedAbove - fixedBelow) / labelLineHeight));
        const maxRows = settings.labels.maxRows.value;
        const aboveCapacity = belowIndexes.length === 0 ? capacity : aboveIndexes.length === 0 ? 0 : Math.ceil(capacity / 2);
        const maxAbove = Math.min(aboveCapacity, maxRows);
        const maxBelow = Math.min(capacity - aboveCapacity, maxRows);
        const aboveLabels = this.layoutEventLabels(data, aboveIndexes, xs, anchors, width, labelStyle, maxAbove);
        const belowLabels = this.layoutEventLabels(data, belowIndexes, xs, anchors, width, labelStyle, maxBelow);

        // The drawing is centred vertically.
        const aboveHeight = fixedAbove + countRows(aboveLabels) * labelLineHeight;
        const belowHeight = fixedBelow + countRows(belowLabels) * labelLineHeight;
        const freeHeight = height - 2 * MARGIN - todayLabelHeight - aboveHeight - belowHeight;
        const lineY = MARGIN + todayLabelHeight + Math.max(0, freeHeight / 2) + aboveHeight;

        if (showToday) {
            this.drawTodayLine(xOf(today), width, height, axisStyle);
        }
        this.drawAxisLine(x0 - radius, x1 + radius, lineY, ticks);

        // Leader lines go under the labels, and labels get a background so that no line crosses their text.
        const leaderLines = settings.labels.leaderLines.value;
        const leaderGroup = this.appendGroup("timeline-leader-lines");
        const labelGroup = this.appendGroup("timeline-labels");
        labelGroup.setAttribute("aria-hidden", "true");
        axisLabels.forEach((label) => {
            const top = lineY + axisTop + label.rowIndex * axisLineHeight;
            this.drawLabel(labelGroup, label.text, label.x, label.anchor, top, axisLineHeight, axisStyle, leaderLines, null);
        });
        aboveLabels.forEach((layout) => {
            const bottom = lineY - radius - LABEL_GAP - layout.rowIndex * labelLineHeight;
            if (leaderLines) {
                this.drawLeaderLine(leaderGroup, layout, lineY - radius, bottom, data.events[layout.key]);
            }
            this.drawLabel(labelGroup, layout.text, layout.x, "middle", bottom - labelLineHeight, labelLineHeight, labelStyle, leaderLines, layout.key);
        });
        belowLabels.forEach((layout) => {
            const top = lineY + fixedBelow + layout.rowIndex * labelLineHeight;
            if (leaderLines) {
                this.drawLeaderLine(leaderGroup, layout, lineY + radius, top, data.events[layout.key]);
            }
            this.drawLabel(labelGroup, layout.text, layout.x, "middle", top, labelLineHeight, labelStyle, leaderLines, layout.key);
        });

        this.drawMarkers(data, xs, lineY, markerSize);
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
        anchors: number[],
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

    private layoutEventLabels(
        data: TimelineData,
        indexes: number[],
        xs: number[],
        anchors: number[],
        width: number,
        style: TextStyle,
        maxRows: number
    ): LabelLayout[] {
        if (indexes.length === 0 || maxRows <= 0) {
            return [];
        }
        const widths = getLabelWidths(indexes.map((index) => anchors[index]), MIN_LABEL_WIDTH, width - 4);
        const items = indexes.map((index, position): LabelItem => ({
            x: xs[index],
            maxWidth: widths[position],
            text: data.events[index].name,
            key: index
        }));
        return layoutLabels(items, width, style, LABEL_ROW_GAP, maxRows);
    }

    private appendGroup(className: string): SVGGElement {
        const group = document.createElementNS(SVG_NS, "g");
        group.setAttribute("class", className);
        this.svg.appendChild(group);
        return group;
    }

    private drawTodayLine(x: number, width: number, height: number, style: TextStyle): void {
        const settings = this.formattingSettings.todayLine;
        const color = this.theme.isHighContrast ? this.theme.foreground : settings.lineColor.value.value || DEFAULT_TODAY_COLOR;
        const group = this.appendGroup("timeline-today");
        group.setAttribute("aria-hidden", "true");

        const line = document.createElementNS(SVG_NS, "line");
        line.setAttribute("x1", x.toString());
        line.setAttribute("x2", x.toString());
        line.setAttribute("y1", (MARGIN + style.fontSize * 1.25 + 2).toString());
        line.setAttribute("y2", (height - MARGIN).toString());
        line.setAttribute("stroke", color);
        line.setAttribute("stroke-width", "1.5");
        const dashes: Record<LineStyle, string> = { solid: "", dashed: "5 4", dotted: "1 3" };
        const dash = dashes[settings.lineStyle.value as LineStyle];
        if (dash) {
            line.setAttribute("stroke-dasharray", dash);
            line.setAttribute("stroke-linecap", "round");
        }
        group.appendChild(line);

        const text = this.getString("Label_Today");
        const labelWidth = measureText(text, style);
        const label = document.createElementNS(SVG_NS, "text");
        label.setAttribute("x", clamp(x, labelWidth / 2 + 2, width - labelWidth / 2 - 2).toString());
        label.setAttribute("y", MARGIN.toString());
        label.setAttribute("text-anchor", "middle");
        label.setAttribute("dominant-baseline", "hanging");
        applySvgTextStyle(label, { ...style, color });
        label.textContent = text;
        group.appendChild(label);
    }

    private drawAxisLine(startX: number, endX: number, lineY: number, ticks: number[]): void {
        const settings = this.formattingSettings.dateAxis;
        const color = this.theme.isHighContrast ? this.theme.foreground : settings.lineColor.value.value || DEFAULT_LINE_COLOR;
        const lineWidth = settings.lineWidth.value;
        const group = this.appendGroup("timeline-axis");
        group.setAttribute("aria-hidden", "true");

        const line = document.createElementNS(SVG_NS, "line");
        line.setAttribute("x1", startX.toString());
        line.setAttribute("x2", endX.toString());
        line.setAttribute("y1", lineY.toString());
        line.setAttribute("y2", lineY.toString());
        line.setAttribute("stroke", color);
        line.setAttribute("stroke-width", lineWidth.toString());
        line.setAttribute("stroke-linecap", "round");
        group.appendChild(line);

        ticks.forEach((x) => {
            const tick = document.createElementNS(SVG_NS, "line");
            tick.setAttribute("x1", x.toString());
            tick.setAttribute("x2", x.toString());
            tick.setAttribute("y1", lineY.toString());
            tick.setAttribute("y2", (lineY + lineWidth / 2 + TICK_LENGTH).toString());
            tick.setAttribute("stroke", color);
            tick.setAttribute("stroke-width", "1");
            group.appendChild(tick);
        });
    }

    private drawLeaderLine(group: SVGGElement, layout: LabelLayout, fromY: number, toY: number, event: TimelineEvent): void {
        const line = document.createElementNS(SVG_NS, "line");
        line.setAttribute("x1", layout.anchorX.toString());
        line.setAttribute("x2", layout.anchorX.toString());
        line.setAttribute("y1", fromY.toString());
        line.setAttribute("y2", toY.toString());
        line.setAttribute("stroke", this.theme.isHighContrast ? this.theme.foreground : event.color);
        line.setAttribute("stroke-opacity", this.theme.isHighContrast ? "1" : "0.5");
        line.setAttribute("stroke-width", "1");
        line.setAttribute("data-index", layout.key.toString());
        group.appendChild(line);
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
        eventIndex: number | null
    ): void {
        if (withBackground) {
            const width = measureText(content, style);
            const left = anchor === "start" ? x : anchor === "end" ? x - width : x - width / 2;
            const background = document.createElementNS(SVG_NS, "rect");
            background.setAttribute("x", (left - 2).toString());
            background.setAttribute("y", top.toString());
            background.setAttribute("width", (width + 4).toString());
            background.setAttribute("height", rowHeight.toString());
            background.setAttribute("fill", this.theme.background);
            group.appendChild(background);
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

    private drawMarkers(data: TimelineData, xs: number[], lineY: number, markerSize: number): void {
        const group = this.appendGroup("timeline-markers");
        this.focusedIndex = Math.min(this.focusedIndex, data.events.length - 1);
        const interactive = this.allowInteractions();
        // When thousands of markers overlap, their white outlines would hide every fill but the last one.
        const crowded = xs.length > 1 && (xs[xs.length - 1] - xs[0]) / (xs.length - 1) < markerSize / 2;
        const strokeWidth = this.theme.isHighContrast ? "2" : crowded ? "0" : "1";
        data.events.forEach((event, index) => {
            const marker = document.createElementNS(SVG_NS, "g");
            marker.setAttribute("class", "timeline-marker");
            marker.setAttribute("data-index", index.toString());
            marker.setAttribute("role", "option");
            marker.setAttribute("aria-label", [event.name, event.dateText, event.legend].filter((part) => !!part).join(", "));
            if (interactive) {
                marker.setAttribute("tabindex", index === this.focusedIndex ? "0" : "-1");
            }
            const shape = createShape(event.shape, xs[index], lineY, markerSize);
            shape.setAttribute("class", "timeline-marker-shape");
            shape.setAttribute("fill", event.color);
            shape.setAttribute("stroke", this.theme.isHighContrast ? this.theme.foreground : this.theme.background);
            shape.setAttribute("stroke-width", strokeWidth);
            marker.appendChild(shape);
            group.appendChild(marker);
        });
    }

    // ----- Interactions -----

    private allowInteractions(): boolean {
        const capabilities = this.host.hostCapabilities;
        return !capabilities || capabilities.allowInteractions !== false;
    }

    private bindChartEvents(): void {
        this.svg.addEventListener("click", (event: MouseEvent) => {
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

        this.svg.addEventListener("contextmenu", (event: MouseEvent) => {
            if (!this.data || !this.allowInteractions()) {
                return;
            }
            const index = this.getEventIndex(event.target);
            const selectionId = index !== null ? this.data.events[index].selectionId : {};
            this.selectionManager.showContextMenu(selectionId as ISelectionId, { x: event.clientX, y: event.clientY });
            event.preventDefault();
        });

        this.svg.addEventListener("pointermove", (event: PointerEvent) => {
            const index = this.getEventIndex(event.target);
            if (index === null) {
                this.hideTooltip(event.pointerType === "touch");
                return;
            }
            this.showTooltip(index, [event.clientX, event.clientY], event.pointerType === "touch");
        });
        this.svg.addEventListener("pointerleave", (event: PointerEvent) => this.hideTooltip(event.pointerType === "touch"));

        this.svg.addEventListener("keydown", (event: KeyboardEvent) => this.onMarkerKeyDown(event));
        // Keyboard users get the tooltip of the focused marker.
        this.svg.addEventListener("focusin", (event: FocusEvent) => {
            const index = this.getEventIndex(event.target);
            if (index !== null) {
                const rect = (event.target as Element).getBoundingClientRect();
                this.showTooltip(index, [rect.left + rect.width / 2, rect.top + rect.height / 2], false);
            }
        });
        this.svg.addEventListener("focusout", () => this.hideTooltip(false));
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

    private onMarkerKeyDown(event: KeyboardEvent): void {
        const index = this.getEventIndex(event.target);
        if (index === null || !this.allowInteractions()) {
            return;
        }
        const last = this.data.events.length - 1;
        let next: number | null = null;
        switch (event.key) {
            case "ArrowRight":
            case "ArrowDown":
                next = Math.min(index + 1, last);
                break;
            case "ArrowLeft":
            case "ArrowUp":
                next = Math.max(index - 1, 0);
                break;
            case "Home":
                next = 0;
                break;
            case "End":
                next = last;
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

    /**
     * Roving tabindex: only one marker is in the tab order, the arrow keys move between markers.
     */
    private focusMarker(index: number): void {
        const markers = this.svg.querySelectorAll<SVGGElement>(".timeline-marker");
        const previous = markers[this.focusedIndex];
        if (previous) {
            previous.setAttribute("tabindex", "-1");
        }
        this.focusedIndex = index;
        const marker = markers[index];
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

        this.svg.querySelectorAll("[data-index]").forEach((element) => {
            const selected = isSelected(Number(element.getAttribute("data-index")));
            if (element.classList.contains("timeline-marker")) {
                element.setAttribute("aria-selected", String(hasSelection && selected));
                if (this.theme.isHighContrast) {
                    const fill = !hasSelection ? this.theme.foreground : selected ? this.theme.foregroundSelected : this.theme.background;
                    element.firstElementChild.setAttribute("fill", fill);
                    return;
                }
            }
            element.classList.toggle(DIMMED_CLASS, !selected);
        });

        this.legend.querySelectorAll(".timeline-legend-item").forEach((element) => {
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

function applySvgTextStyle(element: SVGTextElement, style: TextStyle): void {
    element.style.fontFamily = style.fontFamily;
    element.style.fontSize = `${style.fontSize}px`;
    element.style.fontWeight = style.bold ? "bold" : "normal";
    element.style.fontStyle = style.italic ? "italic" : "normal";
    element.style.textDecoration = style.underline ? "underline" : "none";
    element.style.fill = style.color;
}

function applyHtmlTextStyle(element: HTMLElement, style: TextStyle): void {
    element.style.fontFamily = style.fontFamily;
    element.style.fontSize = `${style.fontSize}px`;
    element.style.fontWeight = style.bold ? "bold" : "normal";
    element.style.fontStyle = style.italic ? "italic" : "normal";
    element.style.textDecoration = style.underline ? "underline" : "none";
    element.style.color = style.color;
}
