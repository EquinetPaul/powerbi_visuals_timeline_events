"use strict";

import powerbi from "powerbi-visuals-api";
import { valueFormatter } from "powerbi-visuals-utils-formattingutils";

import DataView = powerbi.DataView;
import DataViewCategoryColumn = powerbi.DataViewCategoryColumn;
import DataViewMetadataColumn = powerbi.DataViewMetadataColumn;
import DataViewTable = powerbi.DataViewTable;
import ISelectionId = powerbi.visuals.ISelectionId;
import IVisualHost = powerbi.extensibility.visual.IVisualHost;
import IValueFormatter = valueFormatter.IValueFormatter;

import { MarkerIcon, resolveIcon } from "./icons";
import { MARKER_SHAPES } from "./shapes";
import { MarkerShape, VisualFormattingSettingsModel } from "./settings";

// Must match the dataReductionAlgorithm count in capabilities.json.
export const MAX_ROWS = 30000;

export interface TooltipField {
    displayName: string;
    value: string;
}

export interface TimelineEvent {
    rowIndex: number;
    date: Date;
    dateText: string;
    name: string;
    legend: string | null;
    color: string;
    shape: MarkerShape;
    icon: MarkerIcon | null;
    tooltipFields: TooltipField[];
    selectionId: ISelectionId;
}

export interface LegendItem {
    label: string;
    color: string;
    shape: MarkerShape;
    icon: MarkerIcon | null;
    /** Icon chosen for this legend value in the format pane ("none" when there is none). */
    iconName: string;
    events: TimelineEvent[];
    /** Selector of the legend value, used to persist its color and icon. Null when the legend data view is missing. */
    selector: powerbi.data.Selector | null;
}

/** A distinct event name, with the icon chosen for it in the format pane. */
export interface EventType {
    name: string;
    /** "auto" (Icon field, then legend), "none", or an icon name. */
    iconName: string;
    selector: powerbi.data.Selector;
}

export interface TimelineData {
    events: TimelineEvent[];
    /** Distinct event names that can get an icon in the format pane. Empty when the event data view is missing. */
    eventTypes: EventType[];
    legendItems: LegendItem[];
    legendTitle: string;
    /** True when the legend lists event names (no Legend field, "Shape by event name" on). */
    legendByEventName: boolean;
    dateColumnName: string;
    eventColumnName: string;
    legendColumnName: string | null;
    minDate: Date;
    maxDate: Date;
    truncated: boolean;
    formatDate: (date: Date) => string;
}

export type BuildResult =
    | { status: "landing" }
    | { status: "message"; messageKey: string }
    | { status: "ok"; data: TimelineData };

export interface ColorOptions {
    isHighContrast: boolean;
    foreground: string;
}

interface LegendValueInfo {
    color: string | null;
    iconName: string | null;
    selector: powerbi.data.Selector;
}

/**
 * The table data view holds the events. A second, categorical data view holds one entry per legend
 * value: its identity lets the Colors card store a color per legend value (the documented pattern).
 */
export function getTableDataView(dataViews: DataView[] | undefined): DataView | undefined {
    return (dataViews || []).find((dataView) => !!(dataView && dataView.table)) || (dataViews && dataViews[0]);
}

/**
 * The categorical data view whose category has the given role: "legend" or "event".
 */
function getRoleCategory(dataViews: DataView[] | undefined, role: string): DataViewCategoryColumn | null {
    const dataView = (dataViews || []).find((view) => {
        const categories = view && view.categorical && view.categorical.categories;
        return !!(categories && categories[0] && categories[0].source.roles && categories[0].source.roles[role]);
    });
    return dataView ? dataView.categorical.categories[0] : null;
}

/**
 * Turns the data views into timeline events, legend items and selection ids.
 */
export function buildTimelineData(
    dataViews: DataView[] | undefined,
    host: IVisualHost,
    settings: VisualFormattingSettingsModel,
    colorOptions: ColorOptions
): BuildResult {
    const dataView = getTableDataView(dataViews);
    const table = dataView && dataView.table;
    const columns = (table && table.columns) || [];
    if (columns.length === 0) {
        return { status: "landing" };
    }

    const dateColumnIndex = findDateColumnIndex(columns);
    const eventColumnIndex = columns.findIndex((column) => column.roles && column.roles.event);
    const legendColumnIndex = columns.findIndex((column) => column.roles && column.roles.legend);
    const iconColumnIndex = columns.findIndex((column) => column.roles && column.roles.icon);
    const showIcons = settings.icons.show.value;
    const tooltipColumnIndexes = columns
        .map((column, index) => (column.roles && column.roles.tooltip ? index : -1))
        .filter((index) => index !== -1);

    if (dateColumnIndex === -1 || eventColumnIndex === -1) {
        return { status: "message", messageKey: "Message_MissingFields" };
    }
    if (!isUsableDateColumn(columns[dateColumnIndex])) {
        return { status: "message", messageKey: "Message_DateHierarchy" };
    }

    const locale = host.locale;
    const dateColumn = columns[dateColumnIndex];
    const dateFormatter = dateColumn.format ? createFormatter(dateColumn.format, locale) : null;
    const formatDate = (date: Date): string => (dateFormatter ? dateFormatter.format(date) : date.toLocaleDateString(locale));
    const eventFormatter = createFormatter(columns[eventColumnIndex].format, locale);
    const legendFormatter = legendColumnIndex !== -1 ? createFormatter(columns[legendColumnIndex].format, locale) : null;
    const tooltipFormatters = tooltipColumnIndexes.map((index) => createFormatter(columns[index].format, locale));

    const rows = table.rows || [];
    const events: TimelineEvent[] = [];
    rows.forEach((row, rowIndex) => {
        const date = coerceDate(row[dateColumnIndex]);
        if (!date) {
            return;
        }
        const legendValue = legendColumnIndex !== -1 ? row[legendColumnIndex] : undefined;
        events.push({
            rowIndex,
            date,
            dateText: formatDate(date),
            name: formatText(row[eventColumnIndex], eventFormatter),
            legend: legendFormatter ? formatText(legendValue, legendFormatter) : null,
            color: "",
            shape: "circle",
            icon: showIcons && iconColumnIndex !== -1 ? resolveIcon(row[iconColumnIndex]) : null,
            tooltipFields: tooltipColumnIndexes
                .map((index, position) => ({ column: columns[index], value: row[index], formatter: tooltipFormatters[position] }))
                .filter(({ value }) => value !== null && value !== undefined && value !== "")
                .map(({ column, value, formatter }) => ({ displayName: column.displayName, value: formatter.format(value) })),
            selectionId: createRowSelectionId(host, table, rowIndex)
        });
    });

    if (events.length === 0) {
        return { status: "message", messageKey: "Message_NoValidDates" };
    }
    events.sort((a, b) => a.date.getTime() - b.date.getTime() || a.rowIndex - b.rowIndex);

    const legendItems = legendColumnIndex !== -1
        ? buildLegendItems(events, getLegendValues(getRoleCategory(dataViews, "legend"), legendFormatter, host), host, settings, colorOptions)
        : [];
    const eventTypes = getEventTypes(getRoleCategory(dataViews, "event"), eventFormatter, host);
    applyEventStyles(events, legendItems, eventTypes, host, settings, colorOptions);

    const legendByEventName = legendColumnIndex === -1 && settings.markers.shapeByEvent.value;
    return {
        status: "ok",
        data: {
            events,
            eventTypes,
            legendItems: legendByEventName ? buildEventNameLegendItems(events) : legendItems,
            legendTitle: legendColumnIndex !== -1 ? columns[legendColumnIndex].displayName : columns[eventColumnIndex].displayName,
            legendByEventName,
            dateColumnName: dateColumn.displayName,
            eventColumnName: columns[eventColumnIndex].displayName,
            legendColumnName: legendColumnIndex !== -1 ? columns[legendColumnIndex].displayName : null,
            minDate: events[0].date,
            maxDate: events[events.length - 1].date,
            truncated: rows.length >= MAX_ROWS,
            formatDate
        }
    };
}

/**
 * Colors and icons picked per legend value are stored against the legend value's identity and come back
 * in the categorical data view's category objects.
 */
function getLegendValues(category: DataViewCategoryColumn | null, formatter: IValueFormatter, host: IVisualHost): Map<string, LegendValueInfo> {
    const values = new Map<string, LegendValueInfo>();
    if (!category) {
        return values;
    }
    category.values.forEach((value, index) => {
        const objects = category.objects && category.objects[index];
        values.set(formatText(value, formatter), {
            color: getFillColor(objects && objects["colors"] && objects["colors"]["fill"]),
            iconName: objects && objects["icons"] && typeof objects["icons"]["icon"] === "string" ? objects["icons"]["icon"] as string : null,
            selector: host.createSelectionIdBuilder().withCategory(category, index).createSelectionId().getSelector()
        });
    });
    return values;
}

/**
 * One legend item per legend value, in alphabetical order so that theme colors are assigned in a stable order.
 */
function buildLegendItems(
    events: TimelineEvent[],
    legendValues: Map<string, LegendValueInfo>,
    host: IVisualHost,
    settings: VisualFormattingSettingsModel,
    colorOptions: ColorOptions
): LegendItem[] {
    const itemsByLabel = new Map<string, LegendItem>();
    events.forEach((event) => {
        const label = event.legend || "";
        let item = itemsByLabel.get(label);
        if (!item) {
            item = { label, color: "", shape: "circle", icon: null, iconName: "none", events: [], selector: null };
            itemsByLabel.set(label, item);
        }
        item.events.push(event);
    });

    const items = Array.from(itemsByLabel.values()).sort((a, b) => a.label.localeCompare(b.label));
    items.forEach((item, index) => {
        const legendValue = legendValues.get(item.label);
        item.color = colorOptions.isHighContrast
            ? colorOptions.foreground
            : (legendValue && legendValue.color) || host.colorPalette.getColor(item.label).value;
        // In high contrast mode every group has the same color, so groups are told apart by their shape.
        item.shape = colorOptions.isHighContrast ? MARKER_SHAPES[index % MARKER_SHAPES.length] : settings.markers.shape.value as MarkerShape;
        item.iconName = (legendValue && legendValue.iconName) || "none";
        item.icon = settings.icons.show.value ? resolveIcon(item.iconName) : null;
        item.selector = legendValue ? legendValue.selector : null;
    });
    return items;
}

/**
 * Icons picked per event name are stored against the event name's identity, like the legend colors.
 */
function getEventTypes(category: DataViewCategoryColumn | null, formatter: IValueFormatter, host: IVisualHost): EventType[] {
    if (!category) {
        return [];
    }
    return category.values.map((value, index): EventType => {
        const objects = category.objects && category.objects[index];
        const iconName = objects && objects["icons"] && objects["icons"]["eventIcon"];
        return {
            name: formatText(value, formatter),
            iconName: typeof iconName === "string" ? iconName : "auto",
            selector: host.createSelectionIdBuilder().withCategory(category, index).createSelectionId().getSelector()
        };
    });
}

function getFillColor(value: powerbi.DataViewPropertyValue): string | null {
    const fill = value as powerbi.Fill;
    return fill && fill.solid && fill.solid.color ? fill.solid.color : null;
}

function applyEventStyles(
    events: TimelineEvent[],
    legendItems: LegendItem[],
    eventTypes: EventType[],
    host: IVisualHost,
    settings: VisualFormattingSettingsModel,
    colorOptions: ColorOptions
): void {
    const legendByLabel = new Map(legendItems.map((item) => [item.label, item]));
    const shapesByEventName = settings.markers.shapeByEvent.value ? getShapesByName(events) : null;
    const defaultColor = colorOptions.isHighContrast
        ? colorOptions.foreground
        : settings.colors.defaultColor.value.value || host.colorPalette.getColor("defaultColor").value;
    const showIcons = settings.icons.show.value;
    const defaultIcon = showIcons ? resolveIcon(settings.icons.defaultIcon.value) : null;
    const iconByEventName = new Map(eventTypes.map((eventType) => [eventType.name, eventType.iconName]));
    events.forEach((event) => {
        const legendItem = event.legend !== null ? legendByLabel.get(event.legend || "") : undefined;
        event.color = legendItem ? legendItem.color : defaultColor;
        // An icon chosen for the event name in the format pane wins; on "auto", the Icon field value wins
        // over the icon of the legend value, which wins over the icon for all events.
        const eventTypeIcon = iconByEventName.get(event.name);
        event.icon = showIcons && eventTypeIcon && eventTypeIcon !== "auto"
            ? resolveIcon(eventTypeIcon)
            : event.icon || (legendItem ? legendItem.icon : defaultIcon);
        event.shape = shapesByEventName
            ? shapesByEventName.get(event.name)
            : legendItem ? legendItem.shape : settings.markers.shape.value as MarkerShape;
    });
}

/**
 * Shapes are assigned in alphabetical order of event names, so up to 8 different names always get different shapes.
 */
function getShapesByName(events: TimelineEvent[]): Map<string, MarkerShape> {
    const names = Array.from(new Set(events.map((event) => event.name))).sort((a, b) => a.localeCompare(b));
    return new Map(names.map((name, index): [string, MarkerShape] => [name, MARKER_SHAPES[index % MARKER_SHAPES.length]]));
}

function buildEventNameLegendItems(events: TimelineEvent[]): LegendItem[] {
    const items = new Map<string, LegendItem>();
    events.forEach((event) => {
        let item = items.get(event.name);
        if (!item) {
            item = { label: event.name, color: event.color, shape: event.shape, icon: event.icon, iconName: "none", events: [], selector: null };
            items.set(event.name, item);
        }
        item.events.push(event);
    });
    return Array.from(items.values());
}

function createRowSelectionId(host: IVisualHost, table: DataViewTable, rowIndex: number): ISelectionId {
    return host.createSelectionIdBuilder()
        .withTable(table, rowIndex)
        .createSelectionId();
}

function createFormatter(format: string | undefined, locale: string): IValueFormatter {
    return valueFormatter.create({ format, cultureSelector: locale });
}

function formatText(value: powerbi.PrimitiveValue, formatter: IValueFormatter): string {
    if (value === null || value === undefined || value === "") {
        return "(Blank)";
    }
    return formatter.format(value);
}

/**
 * Picks the column of the Date role, preferring a real date column when several are present.
 */
function findDateColumnIndex(columns: DataViewMetadataColumn[]): number {
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
function isUsableDateColumn(column: DataViewMetadataColumn): boolean {
    const type = column.type;
    if (type && type.dateTime) {
        return true;
    }
    if ((column.queryName || "").indexOf(".Variation.") !== -1) {
        return false;
    }
    return !type || !!type.text;
}

function coerceDate(value: unknown): Date | null {
    if (value instanceof Date) {
        return Number.isNaN(value.getTime()) ? null : value;
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
