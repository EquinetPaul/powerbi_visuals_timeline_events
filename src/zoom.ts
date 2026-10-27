"use strict";

import powerbi from "powerbi-visuals-api";
import { clamp } from "./layout";

import DataView = powerbi.DataView;

const SVG_NS = "http://www.w3.org/2000/svg";
const MINUTE = 60 * 1000;
/** Height taken by the range slider under the timeline, including the space above it. */
export const SLIDER_HEIGHT = 34;
const TRACK_TOP = 8;
const TRACK_HEIGHT = 22;
const HANDLE_WIDTH = 8;
const HANDLE_HIT_WIDTH = 16;
const RESET_SIZE = 18;
const DOT_RADIUS = 2.5;
const ACCENT_COLOR = "#118DFF";
const TRACK_COLOR = "#F3F2F1";
const LINE_COLOR = "#C8C6C4";
const HANDLE_BORDER_COLOR = "#8A8886";
const RESET_COLOR = "#605E5C";
// Wheel zoom speed: a wheel notch (100 pixels) zooms by about 20 %.
const WHEEL_ZOOM_SPEED = 0.002;

export interface TimeRange {
    start: number;
    end: number;
}

/** Smallest span that can be shown: a thousandth of the span of the data, and at least one minute. */
export function getMinSpan(min: number, max: number): number {
    return Math.max((max - min) / 1000, MINUTE);
}

export function canZoom(min: number, max: number): boolean {
    return max - min > getMinSpan(min, max);
}

/**
 * Keeps a range inside [min, max] and at least the minimum span wide. A range moved past an edge keeps its width.
 */
export function fitRange(range: TimeRange, min: number, max: number): TimeRange {
    const span = clamp(range.end - range.start, Math.min(getMinSpan(min, max), max - min), max - min);
    const start = clamp(range.start, min, max - span);
    return { start, end: start + span };
}

/**
 * Zooms by `factor` (below 1 zooms in) and keeps the time `anchor` at the same place on the screen.
 */
export function zoomRange(range: TimeRange, factor: number, anchor: number, min: number, max: number): TimeRange {
    const span = range.end - range.start;
    const newSpan = clamp(span * factor, getMinSpan(min, max), max - min);
    const start = anchor - (span > 0 ? (anchor - range.start) / span : 0.5) * newSpan;
    return fitRange({ start, end: start + newSpan }, min, max);
}

export function panRange(range: TimeRange, delta: number, min: number, max: number): TimeRange {
    return fitRange({ start: range.start + delta, end: range.end + delta }, min, max);
}

/**
 * The part of the zoom range that holds data, or null to show all the data: when there is no range, when it
 * covers all the data, or when the filters left no data inside it.
 */
export function getZoomDomain(range: TimeRange | null, min: number, max: number): TimeRange | null {
    if (!range || !canZoom(min, max)) {
        return null;
    }
    const start = Math.max(range.start, min);
    const end = Math.min(range.end, max);
    if (end - start < getMinSpan(min, max) / 2 || (start <= min && end >= max)) {
        return null;
    }
    return { start, end };
}

/**
 * Positions [from, to) of the times inside the range. `timeAt` gives the times in increasing order.
 */
export function getVisibleBounds(count: number, timeAt: (position: number) => number, range: TimeRange): [number, number] {
    const firstWhere = (test: (time: number) => boolean): number => {
        let low = 0;
        let high = count;
        while (low < high) {
            const middle = (low + high) >> 1;
            if (test(timeAt(middle))) {
                high = middle;
            } else {
                low = middle + 1;
            }
        }
        return low;
    };
    return [firstWhere((time) => time >= range.start), firstWhere((time) => time > range.end)];
}

/** The `zoom.start` and `zoom.end` properties that save a range with the report ("" for all the data). */
export function getRangeProperties(range: TimeRange | null): { start: string; end: string } {
    return range
        ? { start: new Date(range.start).toISOString(), end: new Date(range.end).toISOString() }
        : { start: "", end: "" };
}

export function getRangeKey(properties: { start: string; end: string }): string {
    return `${properties.start}|${properties.end}`;
}

/**
 * The range saved with the report, and a key that tells whether it changed since the last update.
 */
export function readSavedRange(dataView: DataView | undefined): { range: TimeRange | null; key: string } {
    const zoom = dataView && dataView.metadata && dataView.metadata.objects ? dataView.metadata.objects.zoom : undefined;
    const properties = {
        start: zoom && typeof zoom.start === "string" ? zoom.start : "",
        end: zoom && typeof zoom.end === "string" ? zoom.end : ""
    };
    const start = Date.parse(properties.start);
    const end = Date.parse(properties.end);
    const valid = Number.isFinite(start) && Number.isFinite(end) && end > start;
    return { range: valid ? { start, end } : null, key: getRangeKey(properties) };
}

/**
 * Wheel deltas in pixels, whatever the unit the browser uses.
 */
export function getWheelDelta(event: WheelEvent, pageSize: number): { dx: number; dy: number } {
    const unit = event.deltaMode === WheelEvent.DOM_DELTA_LINE ? 16 : event.deltaMode === WheelEvent.DOM_DELTA_PAGE ? pageSize : 1;
    return { dx: event.deltaX * unit, dy: event.deltaY * unit };
}

export function getWheelZoomFactor(dy: number): number {
    return Math.exp(dy * WHEEL_ZOOM_SPEED);
}

export interface SliderTheme {
    isHighContrast: boolean;
    foreground: string;
    background: string;
    foregroundSelected: string;
}

export interface SliderOptions {
    width: number;
    /** Left of the timeline in the chart (after the lane names). */
    left: number;
    /** Where the first and last dates of the data are drawn on the timeline when it is not zoomed. */
    x0: number;
    x1: number;
    min: number;
    max: number;
    dots: Array<{ time: number; color: string }>;
    theme: SliderTheme;
    getString: (key: string) => string;
    formatTime: (time: number, span: number) => string;
}

export interface SliderCallbacks {
    /** A new range, or null for all the data. `commit` is true at the end of a gesture. */
    onChange: (range: TimeRange | null, commit: boolean) => void;
    /** Shows the dates of the range while it is dragged, or hides them (null). */
    onTooltip: (range: TimeRange | null, clientPoint: [number, number], isTouchEvent: boolean) => void;
}

type SliderPart = "track" | "window" | "start" | "end" | "reset";

interface SliderDrag {
    part: "window" | "start" | "end";
    pointerId: number;
    startTime: number;
    startRange: TimeRange;
}

/**
 * Range slider under the timeline, like the zoom slider of Power BI's charts: an overview of all the events with
 * a window over the range shown. Dragging its handles zooms, dragging the window scrolls.
 */
export class ZoomSlider {
    public readonly element: HTMLDivElement;
    private svg: SVGSVGElement | null = null;
    private options: SliderOptions | null = null;
    private range: TimeRange = { start: 0, end: 0 };
    private drag: SliderDrag | null = null;
    private windowRect: SVGRectElement;
    private leftMask: SVGRectElement;
    private rightMask: SVGRectElement;
    private startHandle: SVGGElement;
    private endHandle: SVGGElement;
    private resetButton: SVGGElement;

    constructor(private callbacks: SliderCallbacks) {
        this.element = document.createElement("div");
        this.element.className = "timeline-zoom";
        this.element.addEventListener("pointerdown", (event) => this.onPointerDown(event));
        this.element.addEventListener("pointermove", (event) => this.onPointerMove(event));
        this.element.addEventListener("pointerup", (event) => this.onPointerUp(event));
        this.element.addEventListener("pointercancel", (event) => this.onPointerUp(event));
        this.element.addEventListener("click", (event) => {
            if (this.getPart(event.target) === "reset") {
                this.callbacks.onChange(null, true);
            }
        });
        // The pointer is captured during a drag, so a double-click targets the slider itself rather than a part.
        this.element.addEventListener("dblclick", (event) => {
            if (this.getPart(event.target) !== "reset") {
                this.callbacks.onChange(null, true);
            }
        });
        this.element.addEventListener("wheel", (event) => this.onWheel(event), { passive: false });
        this.element.addEventListener("keydown", (event) => this.onKeyDown(event));
    }

    public get isDragging(): boolean {
        return this.drag !== null;
    }

    public render(options: SliderOptions, range: TimeRange): void {
        const focusedPart = this.getFocusedPart();
        this.options = options;
        this.drag = null;
        this.element.classList.remove("dragging");

        const { theme, getString } = options;
        const hc = theme.isHighContrast;
        const svg = createElement("svg", { class: "timeline-zoom-svg", width: options.width, height: SLIDER_HEIGHT });
        const group = createElement("g", { transform: `translate(${options.left} 0)`, role: "group", "aria-label": getString("Zoom_Slider") });
        svg.appendChild(group);

        const trackLeft = options.x0 - HANDLE_WIDTH / 2 - 2;
        const trackRight = options.x1 + HANDLE_WIDTH / 2 + 2;
        const middle = TRACK_TOP + TRACK_HEIGHT / 2;
        const track = createElement("rect", {
            "data-part": "track",
            x: trackLeft,
            y: TRACK_TOP,
            width: trackRight - trackLeft,
            height: TRACK_HEIGHT,
            rx: 4,
            fill: hc ? theme.background : TRACK_COLOR,
            stroke: hc ? theme.foreground : "none"
        });
        const title = createElement("title", {});
        title.textContent = getString("Zoom_Slider");
        track.appendChild(title);
        group.appendChild(track);
        group.appendChild(createElement("line", {
            x1: options.x0, x2: options.x1, y1: middle, y2: middle, stroke: hc ? theme.foreground : LINE_COLOR, "pointer-events": "none"
        }));

        // One dot per pixel column and color: thousands of events stay cheap to draw.
        const dots = createElement("g", { "pointer-events": "none" });
        const drawn = new Set<string>();
        options.dots.forEach(({ time, color }) => {
            const x = Math.round(this.xOf(time));
            const key = `${x}|${hc ? "" : color}`;
            if (!drawn.has(key)) {
                drawn.add(key);
                dots.appendChild(createElement("circle", { cx: x, cy: middle, r: DOT_RADIUS, fill: hc ? theme.foreground : color }));
            }
        });
        group.appendChild(dots);

        const maskStyle = { y: TRACK_TOP, height: TRACK_HEIGHT, fill: theme.background, "fill-opacity": hc ? 0.75 : 0.6, "pointer-events": "none" };
        this.leftMask = createElement("rect", { ...maskStyle, x: trackLeft, width: 0 });
        this.rightMask = createElement("rect", { ...maskStyle, x: trackRight, width: 0 });
        group.appendChild(this.leftMask);
        group.appendChild(this.rightMask);

        const accent = hc ? theme.foregroundSelected : ACCENT_COLOR;
        this.windowRect = createElement("rect", {
            "data-part": "window",
            class: "timeline-zoom-window",
            y: TRACK_TOP + 0.75,
            height: TRACK_HEIGHT - 1.5,
            rx: 3,
            fill: accent,
            "fill-opacity": hc ? 0 : 0.08,
            stroke: accent,
            "stroke-width": 1.5,
            tabindex: 0,
            role: "slider",
            "aria-label": getString("Zoom_Window")
        });
        group.appendChild(this.windowRect);

        this.startHandle = this.createHandle("start", getString("Zoom_Start"), theme);
        this.endHandle = this.createHandle("end", getString("Zoom_End"), theme);
        group.appendChild(this.startHandle);
        group.appendChild(this.endHandle);

        this.resetButton = this.createResetButton(options.x1 + HANDLE_WIDTH / 2 + 4 + RESET_SIZE / 2, middle, getString("Zoom_Reset"), theme);
        group.appendChild(this.resetButton);

        if (this.svg) {
            this.svg.replaceWith(svg);
        } else {
            this.element.appendChild(svg);
        }
        this.svg = svg;
        this.setRange(range);

        if (focusedPart) {
            const part = svg.querySelector<SVGElement>(`[data-part="${focusedPart}"]`);
            if (part) {
                part.focus();
            }
        }
    }

    /**
     * Moves the window and the handles without rebuilding the slider, so that a drag goes on while the
     * timeline is redrawn.
     */
    public setRange(range: TimeRange): void {
        this.range = range;
        const options = this.options;
        if (!options || !this.svg) {
            return;
        }
        const start = this.xOf(range.start);
        const end = this.xOf(range.end);
        const trackLeft = options.x0 - HANDLE_WIDTH / 2 - 2;
        const trackRight = options.x1 + HANDLE_WIDTH / 2 + 2;
        setAttributes(this.windowRect, { x: start, width: Math.max(end - start, 1) });
        setAttributes(this.leftMask, { x: trackLeft, width: Math.max(start - trackLeft, 0) });
        setAttributes(this.rightMask, { x: end, width: Math.max(trackRight - end, 0) });
        this.startHandle.setAttribute("transform", `translate(${start} 0)`);
        this.endHandle.setAttribute("transform", `translate(${end} 0)`);

        const span = range.end - range.start;
        const startText = options.formatTime(range.start, span);
        const endText = options.formatTime(range.end, span);
        [this.windowRect, this.startHandle, this.endHandle].forEach((part) => setAttributes(part, {
            "aria-valuemin": options.min,
            "aria-valuemax": options.max
        }));
        setAttributes(this.windowRect, { "aria-valuenow": range.start, "aria-valuetext": `${startText} – ${endText}` });
        setAttributes(this.startHandle, { "aria-valuenow": range.start, "aria-valuetext": startText });
        setAttributes(this.endHandle, { "aria-valuenow": range.end, "aria-valuetext": endText });

        const zoomed = range.start > options.min || range.end < options.max;
        this.resetButton.style.display = zoomed ? "" : "none";
        this.windowRect.classList.toggle("zoomed", zoomed);
    }

    private createHandle(part: "start" | "end", label: string, theme: SliderTheme): SVGGElement {
        const handle = createElement("g", { "data-part": part, class: "timeline-zoom-handle", tabindex: 0, role: "slider", "aria-label": label });
        const top = TRACK_TOP + 4;
        const height = TRACK_HEIGHT - 8;
        handle.appendChild(createElement("rect", { x: -HANDLE_HIT_WIDTH / 2, y: TRACK_TOP, width: HANDLE_HIT_WIDTH, height: TRACK_HEIGHT, fill: "transparent" }));
        handle.appendChild(createElement("rect", {
            class: "timeline-zoom-grip",
            x: -HANDLE_WIDTH / 2,
            y: top,
            width: HANDLE_WIDTH,
            height,
            rx: 3,
            fill: theme.isHighContrast ? theme.background : "#FFFFFF",
            stroke: theme.isHighContrast ? theme.foreground : HANDLE_BORDER_COLOR
        }));
        [-1, 1].forEach((x) => handle.appendChild(createElement("line", {
            x1: x, x2: x, y1: top + 4, y2: top + height - 4, stroke: theme.isHighContrast ? theme.foreground : HANDLE_BORDER_COLOR
        })));
        return handle;
    }

    /** A button with arrows to both edges: shows all the data again. */
    private createResetButton(cx: number, cy: number, label: string, theme: SliderTheme): SVGGElement {
        const color = theme.isHighContrast ? theme.foreground : RESET_COLOR;
        const button = createElement("g", { "data-part": "reset", class: "timeline-zoom-reset", tabindex: 0, role: "button", "aria-label": label });
        const title = createElement("title", {});
        title.textContent = label;
        button.appendChild(title);
        button.appendChild(createElement("rect", {
            x: cx - RESET_SIZE / 2,
            y: cy - RESET_SIZE / 2,
            width: RESET_SIZE,
            height: RESET_SIZE,
            rx: 3,
            fill: theme.background,
            stroke: theme.isHighContrast ? theme.foreground : LINE_COLOR
        }));
        const path = `M${cx - 5} ${cy}H${cx + 5}M${cx - 2.5} ${cy - 2.5}L${cx - 5} ${cy}L${cx - 2.5} ${cy + 2.5}`
            + `M${cx + 2.5} ${cy - 2.5}L${cx + 5} ${cy}L${cx + 2.5} ${cy + 2.5}M${cx - 6.5} ${cy - 4}V${cy + 4}M${cx + 6.5} ${cy - 4}V${cy + 4}`;
        button.appendChild(createElement("path", { d: path, fill: "none", stroke: color, "stroke-width": 1.2, "stroke-linecap": "round", "stroke-linejoin": "round" }));
        return button;
    }

    private xOf(time: number): number {
        const { x0, x1, min, max } = this.options;
        return max > min ? x0 + ((time - min) / (max - min)) * (x1 - x0) : (x0 + x1) / 2;
    }

    private timeAt(clientX: number): number {
        const { left, x0, x1, min, max } = this.options;
        const x = clientX - this.svg.getBoundingClientRect().left - left;
        return min + ((x - x0) / Math.max(x1 - x0, 1)) * (max - min);
    }

    private getPart(target: EventTarget | null): SliderPart | null {
        const element = target instanceof Element ? target.closest("[data-part]") : null;
        return element ? element.getAttribute("data-part") as SliderPart : null;
    }

    private getFocusedPart(): string | null {
        const active = document.activeElement;
        return active && this.svg && this.svg.contains(active) ? active.getAttribute("data-part") : null;
    }

    private onPointerDown(event: PointerEvent): void {
        const part = this.getPart(event.target);
        if (!this.options || !part || part === "reset" || (event.pointerType === "mouse" && event.button !== 0)) {
            return;
        }
        const { min, max } = this.options;
        const time = this.timeAt(event.clientX);
        let startRange = this.range;
        if (part === "track") {
            // Clicking next to the window centres it there.
            startRange = panRange(this.range, time - (this.range.start + this.range.end) / 2, min, max);
            this.callbacks.onChange(startRange, false);
        }
        this.drag = { part: part === "track" ? "window" : part, pointerId: event.pointerId, startTime: time, startRange };
        this.element.setPointerCapture(event.pointerId);
        this.element.classList.add("dragging");
        this.callbacks.onTooltip(this.range, [event.clientX, event.clientY], event.pointerType === "touch");
    }

    private onPointerMove(event: PointerEvent): void {
        const drag = this.drag;
        if (!drag || drag.pointerId !== event.pointerId) {
            return;
        }
        const { min, max } = this.options;
        const minSpan = getMinSpan(min, max);
        const time = this.timeAt(event.clientX);
        let range: TimeRange;
        if (drag.part === "start") {
            range = { start: clamp(time, min, this.range.end - minSpan), end: this.range.end };
        } else if (drag.part === "end") {
            range = { start: this.range.start, end: clamp(time, this.range.start + minSpan, max) };
        } else {
            range = panRange(drag.startRange, time - drag.startTime, min, max);
        }
        this.callbacks.onChange(range, false);
        this.callbacks.onTooltip(this.range, [event.clientX, event.clientY], event.pointerType === "touch");
    }

    private onPointerUp(event: PointerEvent): void {
        if (!this.drag || this.drag.pointerId !== event.pointerId) {
            return;
        }
        this.drag = null;
        this.element.classList.remove("dragging");
        this.callbacks.onTooltip(null, [event.clientX, event.clientY], event.pointerType === "touch");
        this.callbacks.onChange(this.range, true);
    }

    /** The wheel zooms around the pointer; Shift + wheel and horizontal scrolling move the window. */
    private onWheel(event: WheelEvent): void {
        if (!this.options) {
            return;
        }
        event.preventDefault();
        event.stopPropagation();
        const { x0, x1, min, max } = this.options;
        const { dx, dy } = getWheelDelta(event, x1 - x0);
        const span = this.range.end - this.range.start;
        if (event.shiftKey || Math.abs(dx) > Math.abs(dy)) {
            const pixels = dx !== 0 ? dx : dy;
            this.callbacks.onChange(panRange(this.range, (pixels / Math.max(x1 - x0, 1)) * span, min, max), true);
            return;
        }
        const anchor = clamp(this.timeAt(event.clientX), this.range.start, this.range.end);
        this.callbacks.onChange(zoomRange(this.range, getWheelZoomFactor(dy), anchor, min, max), true);
    }

    /**
     * Arrow keys move a handle or the window by a small step, Page Up/Page Down by a large one, Home/End to an edge.
     * + and - zoom in and out around the middle of the range.
     */
    private onKeyDown(event: KeyboardEvent): void {
        const part = this.getPart(event.target);
        if (!this.options || !part || part === "track") {
            return;
        }
        if (part === "reset") {
            if (event.key === "Enter" || event.key === " ") {
                event.preventDefault();
                this.callbacks.onChange(null, true);
            }
            return;
        }
        const { min, max } = this.options;
        const range = this.range;
        const span = range.end - range.start;
        const minSpan = getMinSpan(min, max);
        const isWindow = part === "window";
        const smallStep = isWindow ? span / 10 : (max - min) / 100;
        const largeStep = isWindow ? span : (max - min) / 10;
        const steps: Record<string, number> = {
            ArrowLeft: -smallStep, ArrowDown: -smallStep, ArrowRight: smallStep, ArrowUp: smallStep, PageDown: -largeStep, PageUp: largeStep
        };
        let next: TimeRange | null = null;
        if (event.key === "+" || event.key === "=" || event.key === "-") {
            next = zoomRange(range, event.key === "-" ? 2 : 0.5, (range.start + range.end) / 2, min, max);
        } else if (event.key in steps) {
            const step = steps[event.key];
            next = isWindow
                ? panRange(range, step, min, max)
                : part === "start"
                    ? { start: clamp(range.start + step, min, range.end - minSpan), end: range.end }
                    : { start: range.start, end: clamp(range.end + step, range.start + minSpan, max) };
        } else if (event.key === "Home" || event.key === "End") {
            const toStart = event.key === "Home";
            next = isWindow
                ? panRange(range, toStart ? min - range.start : max - range.end, min, max)
                : part === "start"
                    ? { start: toStart ? min : range.end - minSpan, end: range.end }
                    : { start: range.start, end: toStart ? range.start + minSpan : max };
        } else {
            return;
        }
        event.preventDefault();
        this.callbacks.onChange(next, true);
    }
}

function createElement<K extends keyof SVGElementTagNameMap>(tag: K, attributes: Record<string, string | number>): SVGElementTagNameMap[K] {
    const element = document.createElementNS(SVG_NS, tag);
    setAttributes(element, attributes);
    return element;
}

function setAttributes(element: Element, attributes: Record<string, string | number>): void {
    Object.keys(attributes).forEach((name) => element.setAttribute(name, attributes[name].toString()));
}
