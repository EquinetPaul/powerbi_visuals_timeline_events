"use strict";

import { TimelineData, ValueSettings } from "./dataModel";
import { createIcon, ICON_NAMES, MarkerIcon, resolveIcon } from "./icons";

const SVG_NS = "http://www.w3.org/2000/svg";
// A filter box is shown above the event names when there are more than this many.
const FILTER_THRESHOLD = 12;

export interface EditorOptions {
    getString: (key: string) => string;
    iconColor: string;
    /** Called with the maps to save; each map replaces the stored one. */
    onChange: (changes: Partial<ValueSettings>) => void;
}

/**
 * Panel shown in Power BI's advanced edit mode (the Edit button of the visual header): a color and an icon
 * per legend value, and an icon per event name. Choices are saved with the report through the visual's
 * format properties.
 */
export class ValueEditor {
    private element: HTMLDivElement;
    private filter = "";
    private scrollTop = 0;

    constructor(parent: HTMLElement) {
        this.element = document.createElement("div");
        this.element.className = "timeline-editor";
        this.element.addEventListener("scroll", () => { this.scrollTop = this.element.scrollTop; });
        parent.appendChild(this.element);
    }

    public hide(): void {
        this.element.style.display = "none";
        this.element.replaceChildren();
    }

    public render(data: TimelineData, width: number, options: EditorOptions): void {
        const { getString, onChange } = options;
        const settings = data.valueSettings;
        this.element.style.display = "block";
        this.element.style.width = `${width}px`;
        this.element.replaceChildren();

        const title = appendElement(this.element, "div", "timeline-editor-title");
        title.textContent = getString("Editor_Title");
        appendElement(this.element, "div", "timeline-editor-hint").textContent = getString("Editor_Hint");

        // Legend values: color and icon
        if (data.legendColumnName !== null && !data.legendByEventName) {
            appendElement(this.element, "div", "timeline-editor-section").textContent = data.legendColumnName;
            data.legendItems.forEach((item) => {
                const row = appendElement(this.element, "div", "timeline-editor-row");
                const color = document.createElement("input");
                color.type = "color";
                color.value = toHexColor(item.color);
                color.title = item.label;
                color.addEventListener("change", () => onChange({ legendColors: { ...settings.legendColors, [item.label]: color.value } }));
                row.appendChild(color);
                row.appendChild(createPreview(item.color, item.icon, options.iconColor));
                appendLabel(row, item.label);
                row.appendChild(createIconSelect(["none", ...ICON_NAMES], item.iconName, getString, item.label, (value) =>
                    onChange({ legendIcons: withValue(settings.legendIcons, item.label, value, "none") })));
                const isCustom = item.customColor || item.iconName !== "none";
                row.appendChild(createResetButton(getString("Editor_Reset"), !isCustom, () => onChange({
                    legendColors: withValue(settings.legendColors, item.label, null, ""),
                    legendIcons: withValue(settings.legendIcons, item.label, null, "none")
                })));
            });
        } else {
            appendElement(this.element, "div", "timeline-editor-note").textContent = getString("Editor_NoLegend");
        }

        // Event names: icon
        appendElement(this.element, "div", "timeline-editor-section").textContent = data.eventColumnName;
        const rows: Array<{ row: HTMLElement; name: string }> = [];
        if (data.eventTypes.length > FILTER_THRESHOLD) {
            const filter = document.createElement("input");
            filter.type = "search";
            filter.className = "timeline-editor-filter";
            filter.placeholder = getString("Editor_Filter");
            filter.value = this.filter;
            filter.addEventListener("input", () => {
                this.filter = filter.value;
                rows.forEach(({ row, name }) => { row.style.display = matches(name, this.filter) ? "" : "none"; });
            });
            this.element.appendChild(filter);
        }
        const firstEventByName = new Map<string, TimelineData["events"][number]>();
        data.events.forEach((event) => {
            if (!firstEventByName.has(event.name)) {
                firstEventByName.set(event.name, event);
            }
        });
        data.eventTypes.forEach((eventType) => {
            const event = firstEventByName.get(eventType.name);
            const row = appendElement(this.element, "div", "timeline-editor-row");
            row.appendChild(createPreview(event ? event.color : "#8A8886", event ? event.icon : null, options.iconColor));
            appendLabel(row, eventType.name);
            row.appendChild(createIconSelect(["auto", "none", ...ICON_NAMES], eventType.iconName, getString, eventType.name, (value) =>
                onChange({ eventIcons: withValue(settings.eventIcons, eventType.name, value, "auto") })));
            row.style.display = matches(eventType.name, this.filter) ? "" : "none";
            rows.push({ row, name: eventType.name });
        });

        const footer = appendElement(this.element, "div", "timeline-editor-footer");
        const hasChoices = [settings.legendColors, settings.legendIcons, settings.eventIcons].some((map) => Object.keys(map).length > 0);
        footer.appendChild(createResetButton(getString("Editor_ResetAll"), !hasChoices, () =>
            onChange({ legendColors: {}, legendIcons: {}, eventIcons: {} })));

        this.element.scrollTop = this.scrollTop;
    }
}

/**
 * Copy of `values` where `key` is set to `value`, or removed when `value` is null or the default.
 */
function withValue(values: Record<string, string>, key: string, value: string | null, defaultValue: string): Record<string, string> {
    const copy = { ...values };
    if (value === null || value === defaultValue) {
        delete copy[key];
    } else {
        copy[key] = value;
    }
    return copy;
}

function matches(name: string, filter: string): boolean {
    return !filter || name.toLocaleLowerCase().indexOf(filter.trim().toLocaleLowerCase()) !== -1;
}

function appendElement(parent: HTMLElement, tag: string, className: string): HTMLElement {
    const element = document.createElement(tag);
    element.className = className;
    parent.appendChild(element);
    return element;
}

function appendLabel(row: HTMLElement, text: string): void {
    const label = appendElement(row, "span", "timeline-editor-label");
    label.textContent = text;
    label.title = text;
}

function createPreview(color: string, icon: MarkerIcon | null, iconColor: string): SVGSVGElement {
    const svg = document.createElementNS(SVG_NS, "svg");
    svg.setAttribute("width", "20");
    svg.setAttribute("height", "20");
    svg.setAttribute("aria-hidden", "true");
    const circle = document.createElementNS(SVG_NS, "circle");
    circle.setAttribute("cx", "10");
    circle.setAttribute("cy", "10");
    circle.setAttribute("r", icon ? "9" : "5");
    circle.setAttribute("fill", color);
    svg.appendChild(circle);
    if (icon) {
        svg.appendChild(createIcon(icon, 10, 10, 11, iconColor));
    }
    return svg;
}

function createIconSelect(
    values: string[],
    selected: string,
    getString: (key: string) => string,
    label: string,
    onSelect: (value: string) => void
): HTMLSelectElement {
    const select = document.createElement("select");
    select.setAttribute("aria-label", label);
    values.forEach((value) => {
        const option = document.createElement("option");
        option.value = value;
        option.textContent = getString(`Icon_${value}`);
        select.appendChild(option);
    });
    // An unknown stored value (an icon removed from the library) falls back to the first choice.
    select.value = values.indexOf(selected) !== -1 && (selected === "auto" || selected === "none" || resolveIcon(selected)) ? selected : values[0];
    select.addEventListener("change", () => onSelect(select.value));
    return select;
}

function createResetButton(text: string, disabled: boolean, onClick: () => void): HTMLButtonElement {
    const button = document.createElement("button");
    button.type = "button";
    button.className = "timeline-editor-reset";
    button.textContent = text;
    button.disabled = disabled;
    button.addEventListener("click", onClick);
    return button;
}

/**
 * Color inputs only accept #rrggbb.
 */
function toHexColor(color: string): string {
    if (/^#[0-9a-f]{6}$/i.test(color)) {
        return color;
    }
    if (/^#[0-9a-f]{3}$/i.test(color)) {
        return `#${color[1]}${color[1]}${color[2]}${color[2]}${color[3]}${color[3]}`;
    }
    return "#118DFF";
}
