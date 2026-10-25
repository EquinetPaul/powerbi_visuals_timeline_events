"use strict";

import { scaleTime } from "d3-scale";
import { interfaces, textMeasurementService } from "powerbi-visuals-utils-formattingutils";

import TextProperties = interfaces.TextProperties;

// Minimum distance kept between a label and the left/right edge of the visual.
export const LABEL_EDGE_MARGIN = 2;

export interface TextStyle {
    fontFamily: string;
    /** Font size in px. */
    fontSize: number;
    bold: boolean;
    italic: boolean;
    underline: boolean;
    color: string;
    /** Opacity of the text, 1 when undefined. */
    opacity?: number;
}

export interface LabelItem {
    /** Point the label belongs to. */
    x: number;
    maxWidth: number;
    text: string;
    /** Optional second line, such as the event date. */
    subText?: string;
    key: number;
}

export interface LabelLayout {
    key: number;
    /** Centre of the label once kept inside the visual. */
    x: number;
    anchorX: number;
    text: string;
    subText?: string;
    width: number;
    rowIndex: number;
}

export interface AxisTick {
    date: Date;
    text: string;
}

export function getTextProperties(text: string, style: TextStyle): TextProperties {
    return {
        text,
        fontFamily: style.fontFamily,
        fontSize: `${style.fontSize}px`,
        fontWeight: style.bold ? "bold" : "normal",
        fontStyle: style.italic ? "italic" : "normal"
    };
}

export function measureText(text: string, style: TextStyle): number {
    return textMeasurementService.measureSvgTextWidth(getTextProperties(text, style));
}

export function fitText(text: string, maxWidth: number, style: TextStyle): string {
    if (!text || maxWidth <= 0) {
        return "";
    }
    if (measureText(text, style) <= maxWidth) {
        return text;
    }
    return textMeasurementService.getTailoredTextOrDefault(getTextProperties(text, style), maxWidth);
}

/**
 * Maximum label width of each point: twice the distance to the closest point with a different x,
 * so that labels of neighbouring dates do not run over each other more than needed.
 */
export function getLabelWidths(anchors: number[], minWidth: number, maxWidth: number): number[] {
    const distinct = Array.from(new Set(anchors)).sort((a, b) => a - b);
    const indexByAnchor = new Map(distinct.map((anchor, index) => [anchor, index]));
    return anchors.map((anchor) => {
        const index = indexByAnchor.get(anchor);
        const leftGap = index > 0 ? anchor - distinct[index - 1] : Number.POSITIVE_INFINITY;
        const rightGap = index < distinct.length - 1 ? distinct[index + 1] - anchor : Number.POSITIVE_INFINITY;
        return clamp(2 * Math.min(leftGap, rightGap), minWidth, maxWidth);
    });
}

/**
 * Places labels on rows: a label goes on the first row where it does not overlap the previous label.
 * Labels that would need more than `maxRows` rows are left out (they stay available in the tooltip).
 * `subStyle` is the style of the optional second line of the labels.
 */
export function layoutLabels(
    items: LabelItem[],
    visualWidth: number,
    style: TextStyle,
    rowGap: number,
    maxRows: number,
    subStyle?: TextStyle
): LabelLayout[] {
    const layouts: LabelLayout[] = [];
    const rowEnds: number[] = [];

    [...items].sort((a, b) => a.x - b.x).forEach((item) => {
        const text = fitText(item.text, item.maxWidth, style);
        if (!text) {
            return;
        }
        const subText = item.subText && subStyle ? fitText(item.subText, item.maxWidth, subStyle) : undefined;
        const width = Math.max(measureText(text, style), subText ? measureText(subText, subStyle) : 0);
        // Keep the label inside the visual: labels near an edge are shifted inward.
        const x = width >= visualWidth - LABEL_EDGE_MARGIN * 2
            ? visualWidth / 2
            : clamp(item.x, LABEL_EDGE_MARGIN + width / 2, visualWidth - LABEL_EDGE_MARGIN - width / 2);
        const start = x - width / 2;

        let rowIndex = 0;
        while (rowIndex < rowEnds.length && start <= rowEnds[rowIndex] + rowGap) {
            rowIndex += 1;
        }
        if (rowIndex >= maxRows) {
            return;
        }
        rowEnds[rowIndex] = x + width / 2;
        layouts.push({ key: item.key, x, anchorX: item.x, text, subText, width, rowIndex });
    });

    return layouts;
}

export function countRows(layouts: LabelLayout[]): number {
    return layouts.reduce((rows, layout) => Math.max(rows, layout.rowIndex + 1), 0);
}

/**
 * "Nice" date ticks (days, weeks, months, years...) between min and max, as many as fit without overlapping.
 */
export function getAxisTicks(
    minDate: Date,
    maxDate: Date,
    x0: number,
    x1: number,
    style: TextStyle,
    locale: string
): AxisTick[] {
    if (maxDate.getTime() <= minDate.getTime()) {
        return [{ date: minDate, text: formatTicks([minDate], locale)[0] }];
    }
    const scale = scaleTime().domain([minDate, maxDate]).range([x0, x1]);
    const minGap = 12;
    for (let count = Math.max(2, Math.floor((x1 - x0) / 70)); count >= 2; count--) {
        const dates = scale.ticks(count);
        if (dates.length === 0) {
            continue;
        }
        const texts = formatTicks(dates, locale);
        const fits = dates.every((date, index) => {
            if (index === 0) {
                return true;
            }
            const gap = scale(date) - scale(dates[index - 1]);
            return gap >= (measureText(texts[index], style) + measureText(texts[index - 1], style)) / 2 + minGap;
        });
        if (fits || count === 2) {
            return dates.map((date, index) => ({ date, text: texts[index] }));
        }
    }
    return [];
}

/**
 * Formats ticks with the report locale, showing only the parts that change: years, months or days.
 * The year is repeated on the first tick and whenever it changes.
 */
function formatTicks(dates: Date[], locale: string): string[] {
    const isYearStart = (date: Date) => date.getMonth() === 0 && date.getDate() === 1;
    const hasTime = dates.some((date) => date.getHours() !== 0 || date.getMinutes() !== 0);
    const allYears = dates.length > 1 && dates.every(isYearStart);
    const allMonths = dates.length > 1 && dates.every((date) => date.getDate() === 1);

    return dates.map((date, index) => {
        const showYear = index === 0 || date.getFullYear() !== dates[index - 1].getFullYear();
        let options: Intl.DateTimeFormatOptions;
        if (hasTime) {
            options = { day: "numeric", month: "short", hour: "2-digit", minute: "2-digit" };
        } else if (allYears) {
            options = { year: "numeric" };
        } else if (allMonths) {
            options = showYear ? { month: "short", year: "numeric" } : { month: "short" };
        } else {
            options = showYear ? { day: "numeric", month: "short", year: "numeric" } : { day: "numeric", month: "short" };
        }
        return safeFormat(date, locale, options);
    });
}

function safeFormat(date: Date, locale: string, options: Intl.DateTimeFormatOptions): string {
    try {
        return new Intl.DateTimeFormat(locale, options).format(date);
    } catch {
        return new Intl.DateTimeFormat(undefined, options).format(date);
    }
}

export function clamp(value: number, min: number, max: number): number {
    return Math.min(Math.max(value, min), max);
}
