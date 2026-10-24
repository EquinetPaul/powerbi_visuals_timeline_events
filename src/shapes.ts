"use strict";

import { MarkerShape } from "./settings";

export const MARKER_SHAPES: MarkerShape[] = ["circle", "square", "diamond", "triangle", "invertedTriangle", "cross", "star", "hexagon"];

/**
 * Creates a marker centred on (cx, cy). `size` is the marker's width and height.
 */
export function createShape(shape: MarkerShape, cx: number, cy: number, size: number): SVGElement {
    const radius = size / 2;
    if (shape === "circle") {
        const circle = document.createElementNS("http://www.w3.org/2000/svg", "circle");
        circle.setAttribute("cx", cx.toString());
        circle.setAttribute("cy", cy.toString());
        circle.setAttribute("r", radius.toString());
        return circle;
    }
    const polygon = document.createElementNS("http://www.w3.org/2000/svg", "polygon");
    const points = getShapePoints(shape, cx, cy, radius)
        .map(([x, y]) => `${x},${y}`)
        .join(" ");
    polygon.setAttribute("points", points);
    return polygon;
}

function getShapePoints(shape: MarkerShape, cx: number, cy: number, r: number): Array<[number, number]> {
    switch (shape) {
        case "square": {
            const half = r * 0.9;
            return [[cx - half, cy - half], [cx + half, cy - half], [cx + half, cy + half], [cx - half, cy + half]];
        }
        case "diamond":
            return [[cx, cy - r], [cx + r, cy], [cx, cy + r], [cx - r, cy]];
        case "triangle":
            return [[cx, cy - r], [cx + r, cy + r * 0.8], [cx - r, cy + r * 0.8]];
        case "invertedTriangle":
            return [[cx - r, cy - r * 0.8], [cx + r, cy - r * 0.8], [cx, cy + r]];
        case "cross": {
            const arm = r / 3;
            return [
                [cx - arm, cy - r], [cx + arm, cy - r], [cx + arm, cy - arm], [cx + r, cy - arm],
                [cx + r, cy + arm], [cx + arm, cy + arm], [cx + arm, cy + r], [cx - arm, cy + r],
                [cx - arm, cy + arm], [cx - r, cy + arm], [cx - r, cy - arm], [cx - arm, cy - arm]
            ];
        }
        case "star":
            return getRadialPoints(cx, cy, 10, (index) => (index % 2 === 0 ? r : r * 0.45));
        default: // hexagon
            return getRadialPoints(cx, cy, 6, () => r);
    }
}

function getRadialPoints(cx: number, cy: number, count: number, radius: (index: number) => number): Array<[number, number]> {
    return Array.from({ length: count }, (_, index): [number, number] => {
        const angle = -Math.PI / 2 + (index * 2 * Math.PI) / count;
        return [cx + radius(index) * Math.cos(angle), cy + radius(index) * Math.sin(angle)];
    });
}
