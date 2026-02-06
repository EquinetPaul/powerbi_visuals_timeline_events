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
import { FormattingSettingsService } from "powerbi-visuals-utils-formattingmodel";
import "./../style/visual.less";

import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisual = powerbi.extensibility.visual.IVisual;

import { VisualFormattingSettingsModel } from "./settings";

export class Visual implements IVisual {
    private target: HTMLElement;
    private container: HTMLDivElement;
    private svg: SVGSVGElement;
    private legend: HTMLDivElement;
    private emptyState: HTMLDivElement;
    private formattingSettings: VisualFormattingSettingsModel;
    private formattingSettingsService: FormattingSettingsService;

    constructor(options: VisualConstructorOptions) {
        this.formattingSettingsService = new FormattingSettingsService();
        this.target = options.element;
        this.target.innerHTML = "";

        this.container = document.createElement("div");
        this.container.className = "timeline-visual";

        this.svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
        this.svg.classList.add("timeline-svg");
        this.container.appendChild(this.svg);

        this.legend = document.createElement("div");
        this.legend.className = "timeline-legend";
        this.container.appendChild(this.legend);

        this.emptyState = document.createElement("div");
        this.emptyState.className = "timeline-empty";
        this.container.appendChild(this.emptyState);

        this.target.appendChild(this.container);
    }

    public update(options: VisualUpdateOptions) {
        this.formattingSettings = this.formattingSettingsService.populateFormattingSettingsModel(VisualFormattingSettingsModel, options.dataViews[0]);
        const dataView = options.dataViews && options.dataViews[0];
        if (!dataView || !dataView.table || !dataView.table.rows || dataView.table.rows.length === 0) {
            this.renderEmpty("Ajoutez des champs Date, Event Name et Tooltip.");
            return;
        }

        const columns = dataView.table.columns || [];
        const dateColumnIndex = columns.findIndex((column) => column.roles && column.roles.date);
        const eventColumnIndex = columns.findIndex((column) => column.roles && column.roles.event);
        const tooltipColumnIndexes = columns
            .map((column, index) => (column.roles && column.roles.tooltip ? index : -1))
            .filter((index) => index !== -1);

        if (dateColumnIndex === -1 || eventColumnIndex === -1) {
            this.renderEmpty("Assignez au minimum Date et Event Name.");
            return;
        }

        const grouped = this.groupByDate(
            dataView.table.rows,
            columns,
            dateColumnIndex,
            eventColumnIndex,
            tooltipColumnIndexes
        );

        if (grouped.length === 0) {
            this.renderEmpty("Aucune donnée exploitable.");
            return;
        }

        const scaling = this.normalizeScale(this.formattingSettings.timelineCard.scaling.value);
        const symbolColor = this.normalizeColor(this.formattingSettings.timelineCard.symbolColor.value.value);
        const showEventLabels = this.formattingSettings.timelineCard.showEventLabels.value;
        const showDates = this.formattingSettings.timelineCard.showDates.value;
        const useEventSymbols = this.formattingSettings.timelineCard.useEventSymbols.value;
        const showLegend = this.formattingSettings.timelineCard.showLegend.value;
        this.renderTimeline(grouped, options.viewport.width, options.viewport.height, {
            scaling,
            symbolColor,
            showEventLabels,
            showDates,
            useEventSymbols,
            showLegend
        });
    }

    /**
     * Returns properties pane formatting model content hierarchies, properties and latest formatting values, Then populate properties pane.
     * This method is called once every time we open properties pane or when the user edit any format property. 
     */
    public getFormattingModel(): powerbi.visuals.FormattingModel {
        return this.formattingSettingsService.buildFormattingModel(this.formattingSettings);
    }

    private renderEmpty(message: string): void {
        this.svg.style.display = "none";
        this.legend.style.display = "none";
        this.emptyState.style.display = "block";
        this.emptyState.textContent = message;
    }

    private renderTimeline(
        groups: TimelineGroup[],
        width: number,
        height: number,
        options: TimelineRenderOptions
    ): void {
        this.emptyState.style.display = "none";
        this.svg.style.display = "block";
        this.legend.style.display = "none";

        const safeWidth = Math.max(width, 1);
        const safeHeight = Math.max(height, 1);
        const padding = 24;
        const scale = options.scaling;
        const eventNames = options.useEventSymbols && options.showLegend ? this.getUniqueEventNames(groups) : [];
        const legendHeight = eventNames.length > 0
            ? this.estimateLegendHeight(eventNames, scale, safeWidth, padding)
            : 0;
        const svgHeight = Math.max(safeHeight - legendHeight, 1);
        const minDate = groups[0].date;
        const maxDate = groups[groups.length - 1].date;
        const minMs = minDate.getTime();
        const maxMs = maxDate.getTime();
        const rangeMs = Math.max(maxMs - minMs, 1);
        const innerWidth = Math.max(safeWidth - padding * 2, 1);
        const baseFontSize = 12 * scale;
        const minLabelWidth = 40 * scale;

        const eventPoints = this.buildEventPoints(groups, minMs, rangeMs, innerWidth, padding, scale);
        const sortedByX = [...eventPoints].sort((a, b) => a.x - b.x);

        const eventLabelLayout = options.showEventLabels
            ? this.buildLabelLayouts(
                sortedByX,
                innerWidth,
                baseFontSize,
                minLabelWidth,
                scale,
                (point) => point.entry.eventName
            )
            : { layouts: [], rows: 0 };

        const dateLabelLayout = options.showDates
            ? this.buildLabelLayouts(
                sortedByX,
                innerWidth,
                baseFontSize,
                minLabelWidth,
                scale,
                (point) => this.formatDate(point.group.date)
            )
            : { layouts: [], rows: 0 };

        const lineHeight = baseFontSize * 1.2;
        const labelOffset = 14 * scale;
        const eventRows = eventLabelLayout.rows;
        const dateRows = dateLabelLayout.rows;
        const topNeeded = options.showEventLabels && eventRows > 0
            ? labelOffset + (eventRows - 1) * lineHeight + baseFontSize
            : 0;
        const bottomNeeded = options.showDates && dateRows > 0
            ? labelOffset + (dateRows - 1) * lineHeight + baseFontSize
            : 0;
        const minLineY = padding + topNeeded;
        const maxLineY = svgHeight - padding - bottomNeeded;
        let lineY = svgHeight / 2;
        if (minLineY > maxLineY) {
            lineY = minLineY;
        } else {
            lineY = this.clamp(lineY, minLineY, maxLineY);
        }

        this.svg.setAttribute("width", safeWidth.toString());
        this.svg.setAttribute("height", svgHeight.toString());
        while (this.svg.firstChild) {
            this.svg.removeChild(this.svg.firstChild);
        }

        const line = document.createElementNS("http://www.w3.org/2000/svg", "line");
        line.setAttribute("x1", padding.toString());
        line.setAttribute("y1", lineY.toString());
        line.setAttribute("x2", (safeWidth - padding).toString());
        line.setAttribute("y2", lineY.toString());
        line.setAttribute("class", "timeline-line");
        line.setAttribute("style", `stroke-width:${2 * scale}px`);
        this.svg.appendChild(line);

        if (!options.showDates) {
            const dateLabelOffset = options.showEventLabels ? labelOffset : 10 * scale;
            const dateLabelY = options.showEventLabels ? lineY + dateLabelOffset : lineY - dateLabelOffset;

            const minLabel = document.createElementNS("http://www.w3.org/2000/svg", "text");
            minLabel.setAttribute("x", padding.toString());
            minLabel.setAttribute("y", dateLabelY.toString());
            minLabel.setAttribute("class", "timeline-label");
            minLabel.setAttribute("style", `font-size:${baseFontSize}px`);
            minLabel.textContent = this.formatDate(minDate);
            this.svg.appendChild(minLabel);

            const maxLabel = document.createElementNS("http://www.w3.org/2000/svg", "text");
            maxLabel.setAttribute("x", (safeWidth - padding).toString());
            maxLabel.setAttribute("y", dateLabelY.toString());
            maxLabel.setAttribute("text-anchor", "end");
            maxLabel.setAttribute("class", "timeline-label");
            maxLabel.setAttribute("style", `font-size:${baseFontSize}px`);
            maxLabel.textContent = this.formatDate(maxDate);
            this.svg.appendChild(maxLabel);
        }

        eventPoints.forEach((eventPoint) => {
            const symbol = this.createSymbol(
                eventPoint,
                lineY,
                scale,
                options.symbolColor,
                options.useEventSymbols
            );
            const title = document.createElementNS("http://www.w3.org/2000/svg", "title");
            title.textContent = this.buildTooltip(eventPoint.group);
            symbol.appendChild(title);
            this.svg.appendChild(symbol);
        });

        if (options.showEventLabels) {
            eventLabelLayout.layouts.forEach((layout) => {
                const label = document.createElementNS("http://www.w3.org/2000/svg", "text");
                label.setAttribute("class", "timeline-label");
                label.setAttribute("text-anchor", "middle");
                label.setAttribute("x", layout.point.x.toString());
                label.setAttribute("y", (lineY - (labelOffset + layout.rowIndex * lineHeight)).toString());
                label.setAttribute("style", `font-size:${layout.fontSize}px`);
                label.textContent = layout.text;
                this.svg.appendChild(label);
            });
        }

        if (options.showDates) {
            dateLabelLayout.layouts.forEach((layout) => {
                const label = document.createElementNS("http://www.w3.org/2000/svg", "text");
                label.setAttribute("class", "timeline-label");
                label.setAttribute("text-anchor", "middle");
                label.setAttribute("x", layout.point.x.toString());
                label.setAttribute("y", (lineY + (labelOffset + layout.rowIndex * lineHeight)).toString());
                label.setAttribute("style", `font-size:${layout.fontSize}px`);
                label.textContent = layout.text;
                this.svg.appendChild(label);
            });
        }

        if (options.useEventSymbols && options.showLegend) {
            this.renderLegend(eventNames, options.symbolColor, scale);
        }
    }

    private groupByDate(
        rows: powerbi.DataViewTableRow[],
        columns: powerbi.DataViewMetadataColumn[],
        dateColumnIndex: number,
        eventColumnIndex: number,
        tooltipColumnIndexes: number[]
    ): TimelineGroup[] {
        const groups = new Map<number, TimelineGroup>();

        rows.forEach((row) => {
            const dateValue = row[dateColumnIndex];
            const date = this.coerceDate(dateValue);
            if (!date) {
                return;
            }

            const eventValue = row[eventColumnIndex];
            const eventName = eventValue === null || eventValue === undefined ? "(blank)" : String(eventValue);

            const tooltipFields: TooltipField[] = tooltipColumnIndexes
                .map((index) => {
                    const value = row[index];
                    if (value === null || value === undefined || value === "") {
                        return null;
                    }
                    return {
                        label: columns[index].displayName,
                        value: this.formatValue(value)
                    };
                })
                .filter((field): field is TooltipField => field !== null);

            const key = date.getTime();
            const existing = groups.get(key);
            const entry: TimelineEntry = {
                eventName,
                tooltipFields
            };

            if (existing) {
                existing.entries.push(entry);
            } else {
                groups.set(key, {
                    date,
                    entries: [entry]
                });
            }
        });

        return Array.from(groups.values()).sort((a, b) => a.date.getTime() - b.date.getTime());
    }

    private buildTooltip(group: TimelineGroup): string {
        const lines: string[] = [];
        lines.push(this.formatDate(group.date));

        group.entries.forEach((entry) => {
            if (entry.tooltipFields.length === 0) {
                lines.push(entry.eventName);
                return;
            }
            const tooltipValues = entry.tooltipFields
                .map((field) => `${field.label}: ${field.value}`)
                .join(" | ");
            lines.push(`${entry.eventName} — ${tooltipValues}`);
        });

        return lines.join("\n");
    }

    private buildEventPoints(
        groups: TimelineGroup[],
        minMs: number,
        rangeMs: number,
        innerWidth: number,
        padding: number,
        scale: number
    ): TimelineEventPoint[] {
        const points: TimelineEventPoint[] = [];
        const jitter = 14 * scale;
        groups.forEach((group) => {
            const baseX = padding + ((group.date.getTime() - minMs) / rangeMs) * innerWidth;
            const count = group.entries.length;
            group.entries.forEach((entry, index) => {
                const offset = count > 1 ? (index - (count - 1) / 2) * jitter : 0;
                const x = this.clamp(baseX + offset, padding, padding + innerWidth);
                points.push({
                    group,
                    entry,
                    x
                });
            });
        });
        return points;
    }

    private createSymbol(
        eventPoint: TimelineEventPoint,
        lineY: number,
        scale: number,
        color: string,
        useEventSymbols: boolean
    ): SVGElement {
        const size = 6 * scale;
        const colorClass = "timeline-event";
        if (!useEventSymbols) {
            const circle = document.createElementNS("http://www.w3.org/2000/svg", "circle");
            circle.setAttribute("cx", eventPoint.x.toString());
            circle.setAttribute("cy", lineY.toString());
            circle.setAttribute("r", size.toString());
            circle.setAttribute("class", colorClass);
            circle.setAttribute("style", `fill:${color}`);
            return circle;
        }

        const variant = this.getSymbolVariant(eventPoint.entry.eventName);
        if (variant === 1) {
            const rect = document.createElementNS("http://www.w3.org/2000/svg", "rect");
            rect.setAttribute("x", (eventPoint.x - size).toString());
            rect.setAttribute("y", (lineY - size).toString());
            rect.setAttribute("width", (size * 2).toString());
            rect.setAttribute("height", (size * 2).toString());
            rect.setAttribute("class", colorClass);
            rect.setAttribute("style", `fill:${color}`);
            return rect;
        }

        if (variant === 2) {
            const diamond = document.createElementNS("http://www.w3.org/2000/svg", "polygon");
            const points = [
                `${eventPoint.x},${lineY - size}`,
                `${eventPoint.x + size},${lineY}`,
                `${eventPoint.x},${lineY + size}`,
                `${eventPoint.x - size},${lineY}`
            ].join(" ");
            diamond.setAttribute("points", points);
            diamond.setAttribute("class", colorClass);
            diamond.setAttribute("style", `fill:${color}`);
            return diamond;
        }

        if (variant === 3) {
            const triangle = document.createElementNS("http://www.w3.org/2000/svg", "polygon");
            const points = [
                `${eventPoint.x},${lineY - size}`,
                `${eventPoint.x + size},${lineY + size}`,
                `${eventPoint.x - size},${lineY + size}`
            ].join(" ");
            triangle.setAttribute("points", points);
            triangle.setAttribute("class", colorClass);
            triangle.setAttribute("style", `fill:${color}`);
            return triangle;
        }

        const circle = document.createElementNS("http://www.w3.org/2000/svg", "circle");
        circle.setAttribute("cx", eventPoint.x.toString());
        circle.setAttribute("cy", lineY.toString());
        circle.setAttribute("r", size.toString());
        circle.setAttribute("class", colorClass);
        circle.setAttribute("style", `fill:${color}`);
        return circle;
    }

    private renderLegend(eventNames: string[], color: string, scale: number): void {
        while (this.legend.firstChild) {
            this.legend.removeChild(this.legend.firstChild);
        }

        if (eventNames.length === 0) {
            this.legend.style.display = "none";
            return;
        }

        const title = document.createElement("div");
        title.className = "timeline-legend-title";
        title.textContent = "Legend";
        title.style.fontSize = `${12 * scale}px`;
        this.legend.appendChild(title);

        const list = document.createElement("div");
        list.className = "timeline-legend-list";

        eventNames.forEach((eventName) => {
            const item = document.createElement("div");
            item.className = "timeline-legend-item";

            const svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
            svg.setAttribute("width", `${16 * scale}`);
            svg.setAttribute("height", `${16 * scale}`);
            svg.classList.add("timeline-legend-symbol");

            const variant = this.getSymbolVariant(eventName);
            const symbol = this.createLegendSymbol(variant, 6 * scale, color);
            svg.appendChild(symbol);

            const label = document.createElement("span");
            label.className = "timeline-legend-label";
            label.textContent = eventName;
            label.style.fontSize = `${12 * scale}px`;

            item.appendChild(svg);
            item.appendChild(label);
            list.appendChild(item);
        });

        this.legend.appendChild(list);
        this.legend.style.display = "block";
    }

    private estimateLegendHeight(eventNames: string[], scale: number, width: number, padding: number): number {
        const titleHeight = 16 * scale;
        const rowHeight = 18 * scale;
        const gapX = 12 * scale;
        const symbolWidth = 16 * scale;
        const labelFontSize = 12 * scale;
        const available = Math.max(width - padding * 2, 1);

        let rowWidth = 0;
        let rows = 1;

        eventNames.forEach((eventName) => {
            const labelWidth = this.estimateTextWidth(eventName, labelFontSize);
            const itemWidth = symbolWidth + (6 * scale) + labelWidth;
            const nextWidth = rowWidth === 0 ? itemWidth : rowWidth + gapX + itemWidth;
            if (nextWidth > available && rowWidth > 0) {
                rows += 1;
                rowWidth = itemWidth;
            } else {
                rowWidth = nextWidth;
            }
        });

        const paddingY = 10 * scale;
        return titleHeight + rows * rowHeight + paddingY;
    }

    private createLegendSymbol(variant: number, size: number, color: string): SVGElement {
        const center = size + 1;
        if (variant === 1) {
            const rect = document.createElementNS("http://www.w3.org/2000/svg", "rect");
            rect.setAttribute("x", (center - size).toString());
            rect.setAttribute("y", (center - size).toString());
            rect.setAttribute("width", (size * 2).toString());
            rect.setAttribute("height", (size * 2).toString());
            rect.setAttribute("class", "timeline-event");
            rect.setAttribute("style", `fill:${color}`);
            return rect;
        }

        if (variant === 2) {
            const diamond = document.createElementNS("http://www.w3.org/2000/svg", "polygon");
            const points = [
                `${center},${center - size}`,
                `${center + size},${center}`,
                `${center},${center + size}`,
                `${center - size},${center}`
            ].join(" ");
            diamond.setAttribute("points", points);
            diamond.setAttribute("class", "timeline-event");
            diamond.setAttribute("style", `fill:${color}`);
            return diamond;
        }

        if (variant === 3) {
            const triangle = document.createElementNS("http://www.w3.org/2000/svg", "polygon");
            const points = [
                `${center},${center - size}`,
                `${center + size},${center + size}`,
                `${center - size},${center + size}`
            ].join(" ");
            triangle.setAttribute("points", points);
            triangle.setAttribute("class", "timeline-event");
            triangle.setAttribute("style", `fill:${color}`);
            return triangle;
        }

        const circle = document.createElementNS("http://www.w3.org/2000/svg", "circle");
        circle.setAttribute("cx", center.toString());
        circle.setAttribute("cy", center.toString());
        circle.setAttribute("r", size.toString());
        circle.setAttribute("class", "timeline-event");
        circle.setAttribute("style", `fill:${color}`);
        return circle;
    }

    private getLabelMaxWidth(
        point: TimelineEventPoint,
        sortedPoints: TimelineEventPoint[],
        minWidth: number,
        innerWidth: number
    ): number {
        const index = sortedPoints.indexOf(point);
        if (index === -1) {
            return minWidth;
        }

        const leftGap = index > 0 ? point.x - sortedPoints[index - 1].x : innerWidth;
        const rightGap = index < sortedPoints.length - 1 ? sortedPoints[index + 1].x - point.x : innerWidth;
        const maxWidth = 2 * Math.min(leftGap, rightGap);
        return Math.max(minWidth, maxWidth);
    }

    private buildLabelLayouts(
        sortedPoints: TimelineEventPoint[],
        innerWidth: number,
        baseFontSize: number,
        minLabelWidth: number,
        scale: number,
        getText: (point: TimelineEventPoint) => string
    ): LabelLayoutResult {
        const layouts: LabelLayout[] = [];
        const rowEnds: number[] = [];
        const rowPadding = 6 * scale;

        sortedPoints.forEach((point) => {
            const rawText = getText(point);
            if (!rawText) {
                return;
            }
            const maxWidth = this.getLabelMaxWidth(point, sortedPoints, minLabelWidth, innerWidth);
            const fitted = this.fitLabelText(rawText, maxWidth, baseFontSize);
            if (!fitted.text) {
                return;
            }
            const width = this.estimateTextWidth(fitted.text, fitted.fontSize);
            const start = point.x - width / 2;
            const end = point.x + width / 2;

            let rowIndex = 0;
            while (rowIndex < rowEnds.length && start <= rowEnds[rowIndex] + rowPadding) {
                rowIndex += 1;
            }

            if (rowIndex === rowEnds.length) {
                rowEnds.push(end);
            } else {
                rowEnds[rowIndex] = end;
            }

            layouts.push({
                point,
                text: fitted.text,
                fontSize: fitted.fontSize,
                rowIndex,
                width
            });
        });

        return {
            layouts,
            rows: rowEnds.length
        };
    }

    private fitLabelText(text: string, maxWidth: number, baseFontSize: number): FittedLabel {
        if (!text || maxWidth <= 0) {
            return { text: "", fontSize: baseFontSize };
        }

        const avgCharFactor = 0.6;
        let fontSize = baseFontSize;
        let estimatedWidth = text.length * fontSize * avgCharFactor;

        if (estimatedWidth > maxWidth) {
            fontSize = Math.max(baseFontSize * 0.75, 8);
            estimatedWidth = text.length * fontSize * avgCharFactor;
        }

        if (estimatedWidth > maxWidth) {
            const maxChars = Math.max(1, Math.floor(maxWidth / (fontSize * avgCharFactor)) - 3);
            const trimmed = text.slice(0, maxChars).trimEnd();
            return { text: `${trimmed}...`, fontSize };
        }

        return { text, fontSize };
    }

    private estimateTextWidth(text: string, fontSize: number): number {
        const avgCharFactor = 0.6;
        return text.length * fontSize * avgCharFactor;
    }

    private formatValue(value: unknown): string {
        if (value instanceof Date) {
            return value.toLocaleString();
        }
        if (typeof value === "number") {
            return value.toString();
        }
        if (value === null || value === undefined) {
            return "";
        }
        return String(value);
    }

    private formatDate(date: Date): string {
        return date.toLocaleDateString();
    }

    private coerceDate(value: unknown): Date | null {
        if (value instanceof Date) {
            return value;
        }
        if (typeof value === "string" || typeof value === "number") {
            const parsed = new Date(value);
            if (!Number.isNaN(parsed.getTime())) {
                return parsed;
            }
        }
        return null;
    }

    private clamp(value: number, min: number, max: number): number {
        return Math.min(Math.max(value, min), max);
    }

    private hashString(value: string): number {
        let hash = 0;
        for (let i = 0; i < value.length; i++) {
            hash = (hash << 5) - hash + value.charCodeAt(i);
            hash |= 0;
        }
        return Math.abs(hash);
    }

    private getSymbolVariant(eventName: string): number {
        return this.hashString(eventName) % 4;
    }

    private getUniqueEventNames(groups: TimelineGroup[]): string[] {
        const seen = new Set<string>();
        const results: string[] = [];
        groups.forEach((group) => {
            group.entries.forEach((entry) => {
                if (!seen.has(entry.eventName)) {
                    seen.add(entry.eventName);
                    results.push(entry.eventName);
                }
            });
        });
        return results;
    }

    private normalizeScale(value: number): number {
        if (!Number.isFinite(value) || value <= 0) {
            return 1;
        }
        return Math.min(Math.max(value, 0.5), 3);
    }

    private normalizeColor(value: string | undefined): string {
        if (!value) {
            return "#2563eb";
        }
        return value;
    }
}

interface TooltipField {
    label: string;
    value: string;
}

interface TimelineEntry {
    eventName: string;
    tooltipFields: TooltipField[];
}

interface TimelineGroup {
    date: Date;
    entries: TimelineEntry[];
}

interface TimelineEventPoint {
    group: TimelineGroup;
    entry: TimelineEntry;
    x: number;
}

interface TimelineRenderOptions {
    scaling: number;
    symbolColor: string;
    showEventLabels: boolean;
    showDates: boolean;
    useEventSymbols: boolean;
    showLegend: boolean;
}

interface FittedLabel {
    text: string;
    fontSize: number;
}

interface LabelLayout {
    point: TimelineEventPoint;
    text: string;
    fontSize: number;
    rowIndex: number;
    width: number;
}

interface LabelLayoutResult {
    layouts: LabelLayout[];
    rows: number;
}