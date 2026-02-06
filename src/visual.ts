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

        this.renderTimeline(grouped, options.viewport.width, options.viewport.height);
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
        this.emptyState.style.display = "block";
        this.emptyState.textContent = message;
    }

    private renderTimeline(groups: TimelineGroup[], width: number, height: number): void {
        this.emptyState.style.display = "none";
        this.svg.style.display = "block";

        const safeWidth = Math.max(width, 1);
        const safeHeight = Math.max(height, 1);
        const padding = 24;
        const lineY = safeHeight / 2;
        const minDate = groups[0].date;
        const maxDate = groups[groups.length - 1].date;
        const minMs = minDate.getTime();
        const maxMs = maxDate.getTime();
        const rangeMs = Math.max(maxMs - minMs, 1);
        const innerWidth = Math.max(safeWidth - padding * 2, 1);

        this.svg.setAttribute("width", safeWidth.toString());
        this.svg.setAttribute("height", safeHeight.toString());
        while (this.svg.firstChild) {
            this.svg.removeChild(this.svg.firstChild);
        }

        const line = document.createElementNS("http://www.w3.org/2000/svg", "line");
        line.setAttribute("x1", padding.toString());
        line.setAttribute("y1", lineY.toString());
        line.setAttribute("x2", (safeWidth - padding).toString());
        line.setAttribute("y2", lineY.toString());
        line.setAttribute("class", "timeline-line");
        this.svg.appendChild(line);

        const minLabel = document.createElementNS("http://www.w3.org/2000/svg", "text");
        minLabel.setAttribute("x", padding.toString());
        minLabel.setAttribute("y", (lineY - 10).toString());
        minLabel.setAttribute("class", "timeline-label");
        minLabel.textContent = this.formatDate(minDate);
        this.svg.appendChild(minLabel);

        const maxLabel = document.createElementNS("http://www.w3.org/2000/svg", "text");
        maxLabel.setAttribute("x", (safeWidth - padding).toString());
        maxLabel.setAttribute("y", (lineY - 10).toString());
        maxLabel.setAttribute("text-anchor", "end");
        maxLabel.setAttribute("class", "timeline-label");
        maxLabel.textContent = this.formatDate(maxDate);
        this.svg.appendChild(maxLabel);

        groups.forEach((group) => {
            const x =
                padding + ((group.date.getTime() - minMs) / rangeMs) * innerWidth;
            const circle = document.createElementNS("http://www.w3.org/2000/svg", "circle");
            circle.setAttribute("cx", x.toString());
            circle.setAttribute("cy", lineY.toString());
            circle.setAttribute("r", "6");
            circle.setAttribute("class", "timeline-event");

            const title = document.createElementNS("http://www.w3.org/2000/svg", "title");
            title.textContent = this.buildTooltip(group);
            circle.appendChild(title);

            this.svg.appendChild(circle);
        });
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