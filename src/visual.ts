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
import { interfaces, textMeasurementService, valueFormatter } from "powerbi-visuals-utils-formattingutils";
import "./../style/visual.less";

import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisual = powerbi.extensibility.visual.IVisual;
import IValueFormatter = valueFormatter.IValueFormatter;
import TextProperties = interfaces.TextProperties;

import { MAX_SCALING, MIN_SCALING, VisualFormattingSettingsModel } from "./settings";

// Number of distinct shapes available when "Use event symbols" is on.
const SYMBOL_SHAPE_COUNT = 8;
// Minimum distance kept between a label and the left/right edge of the visual.
const LABEL_EDGE_MARGIN = 2;

export class Visual implements IVisual {
    private target: HTMLElement;
    private container: HTMLDivElement;
    private svg: SVGSVGElement;
    private legend: HTMLDivElement;
    private emptyState: HTMLDivElement;
    private formattingSettings: VisualFormattingSettingsModel;
    private formattingSettingsService: FormattingSettingsService;
    private locale: string;
    private dateFormatter: IValueFormatter | null = null;
    private fontFamily = "";

    constructor(options: VisualConstructorOptions) {
        this.formattingSettingsService = new FormattingSettingsService();
        this.locale = options.host.locale;
        this.target = options.element;
        this.target.replaceChildren();

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
        const dataView = options.dataViews && options.dataViews[0];
        this.formattingSettings = this.formattingSettingsService.populateFormattingSettingsModel(VisualFormattingSettingsModel, dataView);
        if (!dataView || !dataView.table || !dataView.table.rows || dataView.table.rows.length === 0) {
            this.renderEmpty("Ajoutez des champs Date, Event Name et Tooltip.");
            return;
        }

        const columns = dataView.table.columns || [];
        const dateColumnIndex = this.findDateColumnIndex(columns);
        const eventColumnIndex = columns.findIndex((column) => column.roles && column.roles.event);
        const tooltipColumnIndexes = columns
            .map((column, index) => (column.roles && column.roles.tooltip ? index : -1))
            .filter((index) => index !== -1);

        if (dateColumnIndex === -1 || eventColumnIndex === -1) {
            this.renderEmpty("Assignez au minimum Date et Event Name.");
            return;
        }

        const dateColumn = columns[dateColumnIndex];
        if (!this.isUsableDateColumn(dateColumn)) {
            this.renderEmpty(
                "Le champ Date doit contenir une date. Si une hiérarchie de dates est utilisée, "
                + "cliquez sur la flèche du champ Date et choisissez le nom de la colonne au lieu de « Hiérarchie de dates »."
            );
            return;
        }

        this.dateFormatter = dateColumn.format ? this.createFormatter(dateColumn.format) : null;
        const tooltipFormatters = tooltipColumnIndexes.map((index) => this.createFormatter(columns[index].format));

        const grouped = this.groupByDate(
            dataView.table.rows,
            columns,
            dateColumnIndex,
            eventColumnIndex,
            tooltipColumnIndexes,
            tooltipFormatters
        );

        if (grouped.length === 0) {
            this.renderEmpty("Aucune date valide dans le champ Date.");
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
        this.fontFamily = window.getComputedStyle(this.svg).fontFamily;

        const safeWidth = Math.max(width, 1);
        const safeHeight = Math.max(height, 1);
        const padding = 24;
        const scale = options.scaling;
        const eventNames = this.getUniqueEventNames(groups);
        const symbolVariants = this.getSymbolVariants(eventNames);
        if (options.useEventSymbols && options.showLegend) {
            this.renderLegend(eventNames, symbolVariants, options.symbolColor, scale);
        } else {
            this.legend.style.display = "none";
        }
        // Measure the rendered legend so the SVG gets exactly the remaining height.
        const legendHeight = this.legend.style.display === "none" ? 0 : this.legend.offsetHeight;
        const svgHeight = Math.max(safeHeight - legendHeight, 1);
        const minDate = groups[0].date;
        const maxDate = groups[groups.length - 1].date;
        const minMs = minDate.getTime();
        const rangeMs = maxDate.getTime() - minMs;
        const innerWidth = Math.max(safeWidth - padding * 2, 1);
        const baseFontSize = 12 * scale;
        const minLabelWidth = 40 * scale;
        const maxLabelWidth = Math.max(safeWidth - LABEL_EDGE_MARGIN * 2, 1);

        const groupPoints = this.buildGroupPoints(groups, minMs, rangeMs, innerWidth, padding, scale);
        const labelWidths = this.getGroupLabelWidths(groupPoints, minLabelWidth, maxLabelWidth);

        const eventLabelLayout = options.showEventLabels
            ? this.buildLabelLayouts(
                groupPoints.flatMap((groupPoint, groupIndex) => groupPoint.eventPoints.map((eventPoint) => ({
                    x: eventPoint.x,
                    maxWidth: labelWidths[groupIndex],
                    text: eventPoint.entry.eventName
                }))),
                safeWidth,
                baseFontSize,
                scale
            )
            : { layouts: [], rows: 0 };

        // One date label per date, centred under that date's symbols.
        const dateLabelLayout = options.showDates
            ? this.buildLabelLayouts(
                groupPoints.map((groupPoint, groupIndex) => ({
                    x: groupPoint.x,
                    maxWidth: labelWidths[groupIndex],
                    text: this.formatDate(groupPoint.group.date)
                })),
                safeWidth,
                baseFontSize,
                scale
            )
            : { layouts: [], rows: 0 };

        const lineHeight = baseFontSize * 1.2;
        const labelOffset = 14 * scale;
        const boundsLabelOffset = options.showEventLabels ? labelOffset : 10 * scale;
        const eventRows = eventLabelLayout.rows;
        const dateRows = dateLabelLayout.rows;
        const eventLabelsTop = eventRows > 0
            ? labelOffset + (eventRows - 1) * lineHeight + baseFontSize
            : 0;
        const dateLabelsBottom = dateRows > 0
            ? labelOffset + (dateRows - 1) * lineHeight + baseFontSize
            : 0;
        // Without "Show dates", the min/max dates are drawn below the line (or above it when event names are hidden).
        const boundsLabelsSpace = options.showDates ? 0 : boundsLabelOffset + baseFontSize;
        const topNeeded = Math.max(eventLabelsTop, !options.showDates && !options.showEventLabels ? boundsLabelsSpace : 0);
        const bottomNeeded = Math.max(dateLabelsBottom, !options.showDates && options.showEventLabels ? boundsLabelsSpace : 0);
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
            const dateLabelY = options.showEventLabels ? lineY + boundsLabelOffset : lineY - boundsLabelOffset;

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

        groupPoints.forEach((groupPoint) => {
            groupPoint.eventPoints.forEach((eventPoint) => {
                const variant = options.useEventSymbols ? symbolVariants.get(eventPoint.entry.eventName) ?? 0 : 0;
                const symbol = this.createShape(variant, eventPoint.x, lineY, 6 * scale, options.symbolColor);
                const title = document.createElementNS("http://www.w3.org/2000/svg", "title");
                title.textContent = this.buildTooltip(eventPoint);
                symbol.appendChild(title);
                this.svg.appendChild(symbol);
            });
        });

        if (options.showEventLabels) {
            eventLabelLayout.layouts.forEach((layout) => {
                const label = document.createElementNS("http://www.w3.org/2000/svg", "text");
                label.setAttribute("class", "timeline-label");
                label.setAttribute("text-anchor", "middle");
                label.setAttribute("x", layout.x.toString());
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
                label.setAttribute("x", layout.x.toString());
                label.setAttribute("y", (lineY + (labelOffset + layout.rowIndex * lineHeight)).toString());
                label.setAttribute("style", `font-size:${layout.fontSize}px`);
                label.textContent = layout.text;
                this.svg.appendChild(label);
            });
        }
    }

    /**
     * Picks the column of the Date role, preferring a real date column when several are present.
     */
    private findDateColumnIndex(columns: powerbi.DataViewMetadataColumn[]): number {
        const dateColumnIndexes = columns
            .map((column, index) => (column.roles && column.roles.date ? index : -1))
            .filter((index) => index !== -1);
        const dateTimeIndex = dateColumnIndexes.find((index) => columns[index].type && columns[index].type.dateTime);
        if (dateTimeIndex !== undefined) {
            return dateTimeIndex;
        }
        return dateColumnIndexes.length > 0 ? dateColumnIndexes[0] : -1;
    }

    /**
     * Accepts date and text columns. Rejects the levels of Power BI's automatic date hierarchy
     * (Year, Quarter, Month, Day), which are numbers or month names rather than dates.
     */
    private isUsableDateColumn(column: powerbi.DataViewMetadataColumn): boolean {
        const type = column.type;
        if (type && type.dateTime) {
            return true;
        }
        if ((column.queryName || "").indexOf(".Variation.") !== -1) {
            return false;
        }
        return !type || !!type.text;
    }

    private groupByDate(
        rows: powerbi.DataViewTableRow[],
        columns: powerbi.DataViewMetadataColumn[],
        dateColumnIndex: number,
        eventColumnIndex: number,
        tooltipColumnIndexes: number[],
        tooltipFormatters: IValueFormatter[]
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
                .map((index, position) => {
                    const value = row[index];
                    if (value === null || value === undefined || value === "") {
                        return null;
                    }
                    return {
                        label: columns[index].displayName,
                        value: tooltipFormatters[position].format(value)
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

    private buildTooltip(eventPoint: TimelineEventPoint): string {
        const lines = [this.formatDate(eventPoint.group.date), eventPoint.entry.eventName];
        eventPoint.entry.tooltipFields.forEach((field) => {
            lines.push(`${field.label}: ${field.value}`);
        });
        return lines.join("\n");
    }

    private buildGroupPoints(
        groups: TimelineGroup[],
        minMs: number,
        rangeMs: number,
        innerWidth: number,
        padding: number,
        scale: number
    ): TimelineGroupPoint[] {
        const jitter = 14 * scale;
        const minX = padding;
        const maxX = padding + innerWidth;
        return groups.map((group) => {
            const baseX = rangeMs > 0
                ? padding + ((group.date.getTime() - minMs) / rangeMs) * innerWidth
                : padding + innerWidth / 2;
            const count = group.entries.length;
            // Same-date events are spread around their date. The whole cluster is shifted inward
            // at the edges (instead of clamping each symbol) so that no two symbols share the same x.
            const step = count > 1 ? Math.min(jitter, innerWidth / (count - 1)) : 0;
            const halfSpread = (step * (count - 1)) / 2;
            const x = this.clamp(baseX, minX + halfSpread, maxX - halfSpread);
            const groupPoint: TimelineGroupPoint = { group, x, eventPoints: [] };
            groupPoint.eventPoints = group.entries.map((entry, index) => ({
                group,
                entry,
                x: x - halfSpread + index * step
            }));
            return groupPoint;
        });
    }

    private createShape(variant: number, cx: number, cy: number, size: number, color: string): SVGElement {
        let shape: SVGElement;
        if (variant === 0) {
            shape = document.createElementNS("http://www.w3.org/2000/svg", "circle");
            shape.setAttribute("cx", cx.toString());
            shape.setAttribute("cy", cy.toString());
            shape.setAttribute("r", size.toString());
        } else {
            shape = document.createElementNS("http://www.w3.org/2000/svg", "polygon");
            const points = this.getShapePoints(variant, cx, cy, size)
                .map(([x, y]) => `${x},${y}`)
                .join(" ");
            shape.setAttribute("points", points);
        }
        shape.setAttribute("class", "timeline-event");
        shape.setAttribute("style", `fill:${color}`);
        return shape;
    }

    private getShapePoints(variant: number, cx: number, cy: number, size: number): Array<[number, number]> {
        switch (variant) {
            case 1: // square
                return [[cx - size, cy - size], [cx + size, cy - size], [cx + size, cy + size], [cx - size, cy + size]];
            case 2: // diamond
                return [[cx, cy - size], [cx + size, cy], [cx, cy + size], [cx - size, cy]];
            case 3: // triangle
                return [[cx, cy - size], [cx + size, cy + size], [cx - size, cy + size]];
            case 4: // inverted triangle
                return [[cx - size, cy - size], [cx + size, cy - size], [cx, cy + size]];
            case 5: { // cross
                const arm = size / 3;
                return [
                    [cx - arm, cy - size], [cx + arm, cy - size], [cx + arm, cy - arm], [cx + size, cy - arm],
                    [cx + size, cy + arm], [cx + arm, cy + arm], [cx + arm, cy + size], [cx - arm, cy + size],
                    [cx - arm, cy + arm], [cx - size, cy + arm], [cx - size, cy - arm], [cx - arm, cy - arm]
                ];
            }
            case 6: // star
                return this.getRadialPoints(cx, cy, 10, (index) => (index % 2 === 0 ? size : size * 0.45));
            default: // hexagon
                return this.getRadialPoints(cx, cy, 6, () => size);
        }
    }

    private getRadialPoints(cx: number, cy: number, count: number, radius: (index: number) => number): Array<[number, number]> {
        return Array.from({ length: count }, (_, index): [number, number] => {
            const angle = -Math.PI / 2 + (index * 2 * Math.PI) / count;
            return [cx + radius(index) * Math.cos(angle), cy + radius(index) * Math.sin(angle)];
        });
    }

    private renderLegend(eventNames: string[], symbolVariants: Map<string, number>, color: string, scale: number): void {
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

            const center = 8 * scale;
            const symbol = this.createShape(symbolVariants.get(eventName) ?? 0, center, center, 6 * scale, color);
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

    /**
     * Maximum label width for each date: twice the distance to the closest neighbouring date.
     * Events sharing a date share their date's width; overlapping labels are stacked on rows.
     */
    private getGroupLabelWidths(groupPoints: TimelineGroupPoint[], minWidth: number, maxWidth: number): number[] {
        return groupPoints.map((groupPoint, index) => {
            const leftGap = index > 0 ? groupPoint.x - groupPoints[index - 1].x : Number.POSITIVE_INFINITY;
            const rightGap = index < groupPoints.length - 1 ? groupPoints[index + 1].x - groupPoint.x : Number.POSITIVE_INFINITY;
            return this.clamp(2 * Math.min(leftGap, rightGap), minWidth, maxWidth);
        });
    }

    private buildLabelLayouts(
        items: LabelItem[],
        visualWidth: number,
        baseFontSize: number,
        scale: number
    ): LabelLayoutResult {
        const layouts: LabelLayout[] = [];
        const rowEnds: number[] = [];
        const rowPadding = 6 * scale;

        [...items].sort((a, b) => a.x - b.x).forEach((item) => {
            if (!item.text) {
                return;
            }
            const fitted = this.fitLabelText(item.text, item.maxWidth, baseFontSize);
            if (!fitted.text) {
                return;
            }
            const width = this.measureText(fitted.text, fitted.fontSize);
            // Keep the label inside the visual: labels near an edge are shifted inward.
            const x = width >= visualWidth - LABEL_EDGE_MARGIN * 2
                ? visualWidth / 2
                : this.clamp(item.x, LABEL_EDGE_MARGIN + width / 2, visualWidth - LABEL_EDGE_MARGIN - width / 2);
            const start = x - width / 2;
            const end = x + width / 2;

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
                x,
                text: fitted.text,
                fontSize: fitted.fontSize,
                rowIndex
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

        let fontSize = baseFontSize;
        if (this.measureText(text, fontSize) > maxWidth) {
            // Shrink a little before truncating, but never above the base size nor below 8px.
            fontSize = Math.min(baseFontSize, Math.max(baseFontSize * 0.75, 8));
        }

        if (this.measureText(text, fontSize) > maxWidth) {
            const tailored = textMeasurementService.getTailoredTextOrDefault(this.getTextProperties(text, fontSize), maxWidth);
            return { text: tailored, fontSize };
        }

        return { text, fontSize };
    }

    private measureText(text: string, fontSize: number): number {
        return textMeasurementService.measureSvgTextWidth(this.getTextProperties(text, fontSize));
    }

    private getTextProperties(text: string, fontSize: number): TextProperties {
        return {
            text,
            fontFamily: this.fontFamily,
            fontSize: `${fontSize}px`
        };
    }

    private createFormatter(format: string | undefined): IValueFormatter {
        return valueFormatter.create({ format, cultureSelector: this.locale });
    }

    private formatDate(date: Date): string {
        if (this.dateFormatter) {
            return this.dateFormatter.format(date);
        }
        return date.toLocaleDateString(this.locale);
    }

    private coerceDate(value: unknown): Date | null {
        if (value instanceof Date) {
            return value;
        }
        if (typeof value === "string") {
            // Date-only ISO strings ("2026-01-05") are parsed as local dates: new Date() would read
            // them as UTC midnight and display the previous day in time zones behind UTC.
            const isoDate = /^(\d{4})-(\d{2})-(\d{2})$/.exec(value.trim());
            const parsed = isoDate
                ? new Date(Number(isoDate[1]), Number(isoDate[2]) - 1, Number(isoDate[3]))
                : new Date(value);
            if (!Number.isNaN(parsed.getTime())) {
                return parsed;
            }
        }
        return null;
    }

    private clamp(value: number, min: number, max: number): number {
        return Math.min(Math.max(value, min), max);
    }

    /**
     * Shapes are assigned in alphabetical order of event names, so up to SYMBOL_SHAPE_COUNT
     * different events always get different shapes.
     */
    private getSymbolVariants(eventNames: string[]): Map<string, number> {
        const variants = new Map<string, number>();
        [...eventNames]
            .sort((a, b) => a.localeCompare(b))
            .forEach((eventName, index) => variants.set(eventName, index % SYMBOL_SHAPE_COUNT));
        return variants;
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
        if (!Number.isFinite(value)) {
            return 1;
        }
        return this.clamp(value, MIN_SCALING, MAX_SCALING);
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

interface TimelineGroupPoint {
    group: TimelineGroup;
    x: number;
    eventPoints: TimelineEventPoint[];
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

interface LabelItem {
    x: number;
    maxWidth: number;
    text: string;
}

interface LabelLayout {
    x: number;
    text: string;
    fontSize: number;
    rowIndex: number;
}

interface LabelLayoutResult {
    layouts: LabelLayout[];
    rows: number;
}
