"use strict";

import powerbi from "powerbi-visuals-api";
import { valueFormatter } from "powerbi-visuals-utils-formattingutils";

import DataView = powerbi.DataView;
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
    /** Value of the Lane field, or null without one. */
    lane: string | null;
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
    /** Icon chosen for this legend value in the editor ("none" when there is none). */
    iconName: string;
    /** True when the color was chosen in the editor rather than taken from the theme. */
    customColor: boolean;
    events: TimelineEvent[];
}

/** A value of the Lane field and its events, in date order. */
export interface LaneItem {
    label: string;
    events: TimelineEvent[];
}

/** A distinct event name, with the icon chosen for it in the editor. */
export interface EventType {
    name: string;
    /** "auto" (Icon field, then legend), "none", or an icon name. */
    iconName: string;
}

/**
 * Colors and icons chosen per value in the editor, keyed by legend value or event name.
 * They are stored as JSON in static format properties: a visual gets a single data view (one query),
 * so per-value settings cannot be stored against a separate legend or event identity.
 */
export interface ValueSettings {
    legendColors: Record<string, string>;
    legendIcons: Record<string, string>;
    eventIcons: Record<string, string>;
}

export const VALUE_SETTINGS_PROPERTIES: Record<keyof ValueSettings, { objectName: string; propertyName: string }> = {
    legendColors: { objectName: "colors", propertyName: "legendColors" },
    legendIcons: { objectName: "icons", propertyName: "legendIcons" },
    eventIcons: { objectName: "icons", propertyName: "eventIcons" }
};

export interface TimelineData {
    events: TimelineEvent[];
    /** Distinct event names, alphabetically, with their icon setting. */
    eventTypes: EventType[];
    valueSettings: ValueSettings;
    legendItems: LegendItem[];
    legendTitle: string;
    /** True when the legend lists event names (no Legend field, "Shape by event name" on). */
    legendByEventName: boolean;
    dateColumnName: string;
    eventColumnName: string;
    legendColumnName: string | null;
    laneColumnName: string | null;
    /** One item per Lane field value, alphabetically. Empty without a Lane field. */
    laneItems: LaneItem[];
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

/**
 * The visual has a single table data mapping, so there is one data view.
 */
export function getTableDataView(dataViews: DataView[] | undefined): DataView | undefined {
    return (dataViews || []).find((dataView) => !!(dataView && dataView.table)) || (dataViews && dataViews[0]);
}

export function readValueSettings(dataView: DataView | undefined): ValueSettings {
    const objects = dataView && dataView.metadata && dataView.metadata.objects;
    const read = (key: keyof ValueSettings): Record<string, string> => {
        const { objectName, propertyName } = VALUE_SETTINGS_PROPERTIES[key];
        const value = objects && objects[objectName] && objects[objectName][propertyName];
        if (typeof value !== "string" || !value) {
            return {};
        }
        try {
            const parsed: unknown = JSON.parse(value);
            if (!parsed || typeof parsed !== "object" || Array.isArray(parsed)) {
                return {};
            }
            return Object.fromEntries(Object.entries(parsed).filter(([, entry]) => typeof entry === "string")) as Record<string, string>;
        } catch {
            return {};
        }
    };
    return { legendColors: read("legendColors"), legendIcons: read("legendIcons"), eventIcons: read("eventIcons") };
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
    const laneColumnIndex = columns.findIndex((column) => column.roles && column.roles.lane);
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
    const laneFormatter = laneColumnIndex !== -1 ? createFormatter(columns[laneColumnIndex].format, locale) : null;
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
            lane: laneFormatter ? formatText(row[laneColumnIndex], laneFormatter) : null,
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

    const valueSettings = readValueSettings(dataView);
    const legendItems = legendColumnIndex !== -1
        ? buildLegendItems(events, valueSettings, host, settings, colorOptions)
        : [];
    const eventTypes = Array.from(new Set(events.map((event) => event.name)))
        .sort((a, b) => a.localeCompare(b))
        .map((name): EventType => ({ name, iconName: valueSettings.eventIcons[name] || "auto" }));
    applyEventStyles(events, legendItems, eventTypes, host, settings, colorOptions);

    const legendByEventName = legendColumnIndex === -1 && settings.markers.shapeByEvent.value;
    return {
        status: "ok",
        data: {
            events,
            eventTypes,
            valueSettings,
            legendItems: legendByEventName ? buildEventNameLegendItems(events) : legendItems,
            legendTitle: legendColumnIndex !== -1 ? columns[legendColumnIndex].displayName : columns[eventColumnIndex].displayName,
            legendByEventName,
            dateColumnName: dateColumn.displayName,
            eventColumnName: columns[eventColumnIndex].displayName,
            legendColumnName: legendColumnIndex !== -1 ? columns[legendColumnIndex].displayName : null,
            laneColumnName: laneColumnIndex !== -1 ? columns[laneColumnIndex].displayName : null,
            laneItems: laneColumnIndex !== -1 ? buildLaneItems(events) : [],
            minDate: events[0].date,
            maxDate: events[events.length - 1].date,
            truncated: rows.length >= MAX_ROWS,
            formatDate
        }
    };
}

function buildLaneItems(events: TimelineEvent[]): LaneItem[] {
    const lanes = new Map<string, LaneItem>();
    events.forEach((event) => {
        const label = event.lane || "";
        if (!lanes.has(label)) {
            lanes.set(label, { label, events: [] });
        }
        lanes.get(label).events.push(event);
    });
    return Array.from(lanes.values()).sort((a, b) => a.label.localeCompare(b.label));
}

/**
 * One legend item per legend value, in alphabetical order so that theme colors are assigned in a stable order.
 */
function buildLegendItems(
    events: TimelineEvent[],
    valueSettings: ValueSettings,
    host: IVisualHost,
    settings: VisualFormattingSettingsModel,
    colorOptions: ColorOptions
): LegendItem[] {
    const itemsByLabel = new Map<string, LegendItem>();
    events.forEach((event) => {
        const label = event.legend || "";
        let item = itemsByLabel.get(label);
        if (!item) {
            item = { label, color: "", shape: "circle", icon: null, iconName: "none", customColor: false, events: [] };
            itemsByLabel.set(label, item);
        }
        item.events.push(event);
    });

    const items = Array.from(itemsByLabel.values()).sort((a, b) => a.label.localeCompare(b.label));
    items.forEach((item, index) => {
        const customColor = valueSettings.legendColors[item.label];
        item.customColor = !!customColor;
        item.color = colorOptions.isHighContrast
            ? colorOptions.foreground
            : customColor || host.colorPalette.getColor(item.label).value;
        // In high contrast mode every group has the same color, so groups are told apart by their shape.
        item.shape = colorOptions.isHighContrast ? MARKER_SHAPES[index % MARKER_SHAPES.length] : settings.markers.shape.value as MarkerShape;
        item.iconName = valueSettings.legendIcons[item.label] || "none";
        item.icon = settings.icons.show.value ? resolveIcon(item.iconName) : null;
    });
    return items;
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
        // An icon chosen for the event name in the editor wins; on "auto", the Icon field value wins
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
            item = { label: event.name, color: event.color, shape: event.shape, icon: event.icon, iconName: "none", customColor: false, events: [] };
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
