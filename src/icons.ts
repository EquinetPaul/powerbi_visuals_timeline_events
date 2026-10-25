/*
 *  Icons from Fluent UI System Icons (https://github.com/microsoft/fluentui-system-icons),
 *  package @fluentui/svg-icons 1.1.344, "filled" variants drawn on a 20 x 20 grid.
 *
 *  MIT License
 *
 *  Copyright (c) 2020 Microsoft Corporation
 *
 *  Permission is hereby granted, free of charge, to any person obtaining a copy
 *  of this software and associated documentation files (the "Software"), to deal
 *  in the Software without restriction, including without limitation the rights
 *  to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 *  copies of the Software, and to permit persons to whom the Software is
 *  furnished to do so, subject to the following conditions:
 *
 *  The above copyright notice and this permission notice shall be included in all
 *  copies or substantial portions of the Software.
 *
 *  THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 *  IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 *  FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 *  AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 *  LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 *  OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
 *  SOFTWARE.
 */
"use strict";

const SVG_NS = "http://www.w3.org/2000/svg";
const ICON_GRID = 20;

/** Icon name -> SVG path data. The names are the values accepted in the Icon field and in the format pane. */
const ICON_PATHS: Record<string, string[]> = {
    flag: ["M5 13h11.5a.5.5 0 0 0 .42-.78L14.1 8l2.82-4.22A.5.5 0 0 0 16.5 3h-12a.5.5 0 0 0-.5.5v14a.5.5 0 0 0 1 0z"],
    check: ["M7.03 13.9 3.56 10a.75.75 0 0 0-1.12 1l4 4.5c.29.32.79.34 1.09.03l10.5-10.5a.75.75 0 0 0-1.06-1.06z"],
    warning: ["M7.37 3.56a3 3 0 0 1 5.26 0l5.5 10A3 3 0 0 1 15.5 18h-11a3 3 0 0 1-2.63-4.44zM10 12.75a.75.75 0 1 0 0 1.5.75.75 0 0 0 0-1.5m0-6.25a.5.5 0 0 0-.5.5v4a.5.5 0 0 0 1 0V7a.5.5 0 0 0-.5-.5"],
    info: ["M18 10a8 8 0 1 0-16 0 8 8 0 0 0 16 0M9.5 8.91a.5.5 0 0 1 1 0v4.69a.5.5 0 0 1-1 0V8.9m-.25-2.16a.75.75 0 1 1 1.5 0 .75.75 0 0 1-1.5 0"],
    question: ["M10 3a4 4 0 0 0-4 4 .75.75 0 0 0 1.5 0A2.5 2.5 0 0 1 10 4.5 2.5 2.5 0 0 1 12.5 7c0 .6-.16.97-.37 1.27a5 5 0 0 1-.97.92c-.4.33-.9.73-1.28 1.3-.39.58-.63 1.3-.63 2.26v.5a.75.75 0 0 0 1.5 0v-.5c0-.67.16-1.1.38-1.43.23-.35.55-.62.97-.96l.06-.05c.38-.3.84-.67 1.2-1.17C13.74 8.6 14 7.9 14 7a4 4 0 0 0-4-4m0 14a1 1 0 1 0 0-2 1 1 0 0 0 0 2"],
    bell: ["M12.45 16a2.5 2.5 0 0 1-4.9 0zM10 2a5.9 5.9 0 0 1 5.98 5.36l.02.22v3.82l.95 2.6.03.08.01.09v.08c0 .32-.19.6-.51.71l-.12.03-.11.01H3.75a.75.75 0 0 1-.74-.76v-.12l.04-.12L4 11.4V7.57A5.9 5.9 0 0 1 10 2"],
    star: ["M9.1 2.9a1 1 0 0 1 1.8 0l1.93 3.91 4.31.63a1 1 0 0 1 .56 1.7l-3.12 3.05.73 4.3a1 1 0 0 1-1.45 1.05L10 15.51l-3.86 2.03a1 1 0 0 1-1.45-1.05l.74-4.3L2.3 9.14a1 1 0 0 1 .56-1.7l4.31-.63z"],
    heart: ["M9.39 4.29a4.3 4.3 0 0 0-6.13-.02 4.4 4.4 0 0 0 .02 6.2l6.27 6.33c.26.27.69.27.95 0l6.24-6.3a4.4 4.4 0 0 0-.02-6.19 4.3 4.3 0 0 0-6.13-.01l-.6.6z"],
    target: ["M10 11.5a1.5 1.5 0 1 0 0-3 1.5 1.5 0 0 0 0 3M5 10a5 5 0 1 1 10 0 5 5 0 0 1-10 0m5-3.5a3.5 3.5 0 1 0 0 7 3.5 3.5 0 0 0 0-7M2 10a8 8 0 1 1 16 0 8 8 0 0 1-16 0m8-6.5a6.5 6.5 0 1 0 0 13 6.5 6.5 0 0 0 0-13"],
    trophy: ["M5.5 4c0-1.1.9-2 2-2h5a2 2 0 0 1 2 2h1c.82 0 1.5.67 1.5 1.5V7a3 3 0 0 1-2.6 2.97 4.5 4.5 0 0 1-3.9 3.5v1.5h2a2 2 0 0 1 2 2v.5a.5.5 0 0 1-.5.5H6a.5.5 0 0 1-.5-.5v-.5c0-1.1.9-2 2-2h2v-1.5a4.5 4.5 0 0 1-3.9-3.5A3 3 0 0 1 3 7V5.5C3 4.67 3.67 4 4.5 4zm9 4.94A2 2 0 0 0 16 7V5.5a.5.5 0 0 0-.5-.5h-1zM5.5 5h-1a.5.5 0 0 0-.5.5V7a2 2 0 0 0 1.5 1.93z"],
    rocket: ["M10.76 6.43a2 2 0 1 1 2.82 2.82 2 2 0 0 1-2.82-2.82m2.11.7a1 1 0 1 0-1.4 1.41 1 1 0 0 0 1.4-1.4M9.44 13.6l.6.6a1.5 1.5 0 0 0 1.71.3l1.12 1.12c.2.2.51.2.7 0L15 14.2c.86-.86.96-2.17.31-3.14l.85-.85a6.3 6.3 0 0 0 1.56-6.33 2.4 2.4 0 0 0-1.6-1.6A6.3 6.3 0 0 0 9.8 3.86l-.85.85A2.5 2.5 0 0 0 5.8 5L4.38 6.43a.5.5 0 0 0 0 .7L5.5 8.27a1.5 1.5 0 0 0 .3 1.7l.6.61-1.21.73a.5.5 0 0 0-.1.78l2.83 2.83a.5.5 0 0 0 .79-.1zm7.32-9.42a5.3 5.3 0 0 1-1.3 5.33l-4 4a.5.5 0 0 1-.7 0l-1.07-1.07-2.12-2.12L6.5 9.24a.5.5 0 0 1 0-.7l3.99-4a5.3 5.3 0 0 1 5.33-1.3c.44.13.8.48.93.93M5.8 14.91a.5.5 0 0 0-.7-.7l-1.77 1.76a.5.5 0 1 0 .7.7zM4.4 12.79c.2.2.2.5 0 .7l-.7.71a.5.5 0 1 1-.72-.7l.71-.71c.2-.2.51-.2.71 0m2.83 3.53a.5.5 0 1 0-.7-.7l-.72.7a.5.5 0 1 0 .71.71z"],
    flash: ["M7.21 2c-.47 0-.89.31-1.02.77l-2.25 7.87c-.2.68.32 1.36 1.02 1.36h1.27l-1.17 4.68c-.26 1.05 1.04 1.78 1.8 1l8.67-8.86c.65-.68.19-1.82-.76-1.82H12.2l1.26-3.6c.23-.69-.28-1.4-1-1.4z"],
    lightbulb: ["m13.07 15-.38 1.6a1.8 1.8 0 0 1-1.67 1.4H9.13c-.82 0-1.54-.52-1.78-1.26l-.04-.14-.38-1.6zM10 2c3.31 0 6 2.6 6 5.8q-.01 2.53-2.2 4.6a1 1 0 0 0-.15.2l-.02.09L13.3 14H6.7l-.31-1.31a.6.6 0 0 0-.17-.3A6.3 6.3 0 0 1 4 8.04L4 7.8v-.2A5.9 5.9 0 0 1 10 2"],
    search: ["M13.54 14.6a6.5 6.5 0 1 1 1.06-1.06l3.43 3.43a.75.75 0 0 1-.98 1.13l-.08-.07zm.96-5.1a5 5 0 1 0-10 0 5 5 0 0 0 10 0"],
    calendar: ["M5.5 3A2.5 2.5 0 0 0 3 5.5V6h14v-.5A2.5 2.5 0 0 0 14.5 3zM17 7H3v7.5A2.5 2.5 0 0 0 5.5 17h9a2.5 2.5 0 0 0 2.5-2.5zm-9 3a1 1 0 1 1-2 0 1 1 0 0 1 2 0m-1 4a1 1 0 1 1 0-2 1 1 0 0 1 0 2m4-4a1 1 0 1 1-2 0 1 1 0 0 1 2 0m-1 4a1 1 0 1 1 0-2 1 1 0 0 1 0 2m4-4a1 1 0 1 1-2 0 1 1 0 0 1 2 0"],
    clock: ["M10 2a8 8 0 1 1 0 16 8 8 0 0 1 0-16m-.5 3a.5.5 0 0 0-.5.5v5.09a.5.5 0 0 0 .5.41h3.09a.5.5 0 0 0-.09-1H10V5.41A.5.5 0 0 0 9.5 5"],
    location: ["M5.05 4.05a7 7 0 1 1 9.9 9.9l-1.13 1.12-2.43 2.37a2 2 0 0 1-2.64.12l-.14-.12-2.04-1.99-1.52-1.5a7 7 0 0 1 0-9.9M10 6.5a2.5 2.5 0 1 0 0 5 2.5 2.5 0 0 0 0-5"],
    people: ["M6.75 10a3.25 3.25 0 1 0 0-6.5 3.25 3.25 0 0 0 0 6.5m5.69 5.14q.78.35 2.06.36c4 0 4-3 4-3 0-.83-.67-1.5-1.5-1.5h-4.63q.62.74.63 1.75v.36l-.02.25a5 5 0 0 1-.54 1.78M17 7.5a2.5 2.5 0 1 1-5 0 2.5 2.5 0 0 1 5 0M1.5 13c0-1.1.9-2 2-2H10a2 2 0 0 1 2 2s0 4-5.25 4-5.25-4-5.25-4m11.5.1v.07Z"],
    person: ["M10 2a4 4 0 1 0 0 8 4 4 0 0 0 0-8m-5 9a2 2 0 0 0-2 2c0 1.7.83 2.97 2.13 3.8A9 9 0 0 0 10 18c1.85 0 3.58-.39 4.87-1.2A4.4 4.4 0 0 0 17 13a2 2 0 0 0-2-2z"],
    handshake: ["M9.63 3.01h2.41a4 4 0 0 1 .86 0 2.5 2.5 0 0 1 1.08.32 3.7 3.7 0 0 1 1.72 5.1L12.5 5.17a.5.5 0 0 0-.36-.15H10.8a.5.5 0 0 0-.3.1L8.35 6.76a1.12 1.12 0 0 1-1.67-.4A1.1 1.1 0 0 1 7 5zm-4 8.85-.01.02-.77.76v.01a.6.6 0 0 1-.84-.01.6.6 0 0 1 0-.84l.76-.76a.6.6 0 0 1 .85 0c.23.22.23.6.01.82m-.08 1.5a.6.6 0 0 0 .01.82.6.6 0 0 0 .85 0l.76-.76a.6.6 0 0 0-.09-.91.6.6 0 0 0-.74.06l-.01.01-.77.76zM3.75 9.8a.6.6 0 0 1 0 .84l-.77.76a.6.6 0 0 1-.84 0 .6.6 0 0 1 0-.84l.76-.76a.6.6 0 0 1 .85 0m4.98 4.33a.6.6 0 0 1 0 .84l-.76.76-.11.09a.6.6 0 0 1-.74-.1.6.6 0 0 1-.09-.72l.09-.11.76-.76a.6.6 0 0 1 .85 0M3.53 4.4a4.8 4.8 0 0 1 4.3-1.29L6.4 4.21A2.1 2.1 0 0 0 6 7.14c.7.92 2.03 1.1 2.95.4l2.02-1.53h.9l1.95 1.93 1.6 1.62.03.03.87.86a1.02 1.02 0 0 1-1.37 1.51l-.07-.07-.06-.05-.85-.85a.5.5 0 1 0-.7.71l.9.9.1.1.03.03a.7.7 0 0 1-.98.98l-.13-.14a.5.5 0 0 0-.85.37q0 .2.15.36l.17.17a.63.63 0 1 1-.9.9l-.08-.1-.1-.1-.16-.16a.5.5 0 0 0-.7.7l.11.13.06.06.04.05a.65.65 0 0 1-.96.86L9.11 16l.33-.32a1.58 1.58 0 0 0-1.09-2.72q0-.61-.47-1.08c-.3-.3-.69-.46-1.08-.47A1.6 1.6 0 0 0 4.9 9.9q-.1-.45-.44-.8A1.6 1.6 0 0 0 2.3 9 4.7 4.7 0 0 1 3.53 4.4"],
    chat: ["M10 2a8 8 0 1 1-3.61 15.14l-.12-.07-3.65.92a.5.5 0 0 1-.62-.45v-.08l.01-.08.92-3.64-.07-.12a8 8 0 0 1-.83-2.9l-.02-.37L2 10a8 8 0 0 1 8-8m.5 9H7.41a.5.5 0 0 0 0 1h3.18a.5.5 0 0 0 0-1zm2-3H7.41a.5.5 0 0 0 0 1h5.18a.5.5 0 0 0 0-1z"],
    mail: ["M18 7.37v7.13a2.5 2.5 0 0 1-2.5 2.5h-11A2.5 2.5 0 0 1 2 14.5V7.37l7.75 4.56c.15.1.35.1.5 0zM15.5 4a2.5 2.5 0 0 1 2.48 2.22L10 10.92l-7.98-4.7A2.5 2.5 0 0 1 4.5 4z"],
    megaphone: ["M16.11 3.35A1.5 1.5 0 0 1 18 4.8v10.31c0 1.03-1 1.75-1.97 1.43l-4.76-1.59A3.25 3.25 0 0 1 5 13.75v-.89l-1.97-.66A1.5 1.5 0 0 1 2 10.78V8.27c0-.68.46-1.28 1.11-1.45zM6 13.2v.56a2.25 2.25 0 0 0 4.32.88z"],
    document: ["M10 6.5V2H5.5C4.67 2 4 2.67 4 3.5v13c0 .83.67 1.5 1.5 1.5h9c.83 0 1.5-.67 1.5-1.5V8h-4.5A1.5 1.5 0 0 1 10 6.5M6.5 10h7a.5.5 0 0 1 0 1h-7a.5.5 0 0 1 0-1m0 2h7a.5.5 0 0 1 0 1h-7a.5.5 0 0 1 0-1m0 2h7a.5.5 0 0 1 0 1h-7a.5.5 0 0 1 0-1M11 6.5V2.25L15.75 7H11.5a.5.5 0 0 1-.5-.5"],
    book: ["M6 2a2 2 0 0 0-2 2v12c0 1.1.9 2 2 2h9.5a.5.5 0 0 0 0-1H6a1 1 0 0 1-1-1h10a1 1 0 0 0 1-1V4a2 2 0 0 0-2-2zm.75 2.5h6.5a.75.75 0 0 1 0 1.5h-6.5a.75.75 0 0 1 0-1.5"],
    briefcase: ["M18 14a3 3 0 0 1-3 3H5a3 3 0 0 1-3-3v-3.36c.73.83 1.8 1.36 3 1.36h4v.5c0 .28.22.5.5.5h1a.5.5 0 0 0 .5-.5V12h4c1.2 0 2.27-.53 3-1.36zM11.25 2c.97 0 1.75.78 1.75 1.75V5h2a3 3 0 1 1 0 6h-4v-.5a.5.5 0 0 0-.5-.5h-1a.5.5 0 0 0-.5.5v.5H5a3 3 0 0 1 0-6h2V3.75C7 2.78 7.78 2 8.75 2zm-2.5 1a.75.75 0 0 0-.75.75V5h4V3.75a.75.75 0 0 0-.75-.75z"],
    money: ["M3.5 4C2.67 4 2 4.67 2 5.5v7c0 .83.67 1.5 1.5 1.5h11c.83 0 1.5-.67 1.5-1.5v-7c0-.83-.67-1.5-1.5-1.5zM6 5v1a2 2 0 0 1-2 2H3V7h1a1 1 0 0 0 1-1V5zm3 5.75a1.75 1.75 0 1 1 0-3.5 1.75 1.75 0 0 1 0 3.5M3 11v-1h1a2 2 0 0 1 2 2v1H5v-1a1 1 0 0 0-1-1zm11 0a1 1 0 0 0-1 1v1h-1v-1c0-1.1.9-2 2-2h1v1zm0-4h1v1h-1a2 2 0 0 1-2-2V5h1v1a1 1 0 0 0 1 1m3 5.5a2.5 2.5 0 0 1-2.5 2.5H4.09c.2.58.76 1 1.41 1h9a3.5 3.5 0 0 0 3.5-3.5v-5c0-.65-.42-1.2-1-1.41z"],
    cart: ["M3 3.5c0-.28.22-.5.5-.5h.44c.72 0 1.14.47 1.38.94.17.32.29.72.39 1.06H16a1 1 0 0 1 .96 1.27l-1.5 5.28A2 2 0 0 1 13.55 13H8.46a2 2 0 0 1-1.93-1.47L5.9 9.17l-.01-.03-1.03-3.5-.1-.33a5 5 0 0 0-.32-.91c-.16-.31-.3-.4-.5-.4H3.5a.5.5 0 0 1-.5-.5M8.5 17a1.5 1.5 0 1 0 0-3 1.5 1.5 0 0 0 0 3m5 0a1.5 1.5 0 1 0 0-3 1.5 1.5 0 0 0 0 3"],
    truck: ["M4 2a2 2 0 0 0-2 2v11a1 1 0 0 0 1 1V9.8q0-.38.1-.72L4.16 5.6A2.25 2.25 0 0 1 6.3 4h7.4c1 0 1.87.65 2.15 1.6l1.04 3.48q.1.35.11.71V16a1 1 0 0 0 1-1V4a2 2 0 0 0-2-2zm12 7.8v6.7c0 .83-.67 1.5-1.5 1.5h-1a1.5 1.5 0 0 1-1.5-1.5V16H8v.5c0 .83-.67 1.5-1.5 1.5h-1A1.5 1.5 0 0 1 4 16.5V9.8q0-.23.06-.44L5.1 5.9c.16-.53.65-.89 1.2-.89h7.4c.55 0 1.04.36 1.2.9l1.04 3.46q.06.21.06.43M5 16.5c0 .28.22.5.5.5h1a.5.5 0 0 0 .5-.5V16H5zM14.86 9l-.66-2.62a.5.5 0 0 0-.48-.38H6.28a.5.5 0 0 0-.48.38L5.14 9zM13 16.5c0 .28.22.5.5.5h1a.5.5 0 0 0 .5-.5V16h-2zM7 13a1 1 0 1 0 0-2 1 1 0 0 0 0 2m7-1a1 1 0 1 0-2 0 1 1 0 0 0 2 0m-5.5 1a.5.5 0 0 0 0 1h3a.5.5 0 0 0 0-1z"],
    airplane: ["M9 2c.6 0 1.16.36 1.4.92l2.37 5.4 3.23-.2a1.9 1.9 0 0 1 2 1.76V10c0 1.04-.84 1.88-1.88 1.88H16l-3.23-.2-2.37 5.4c-.24.56-.8.92-1.4.92a.9.9 0 0 1-.88-.88v-.06l.7-5.63-2.46-.16-1.14 2.28a.8.8 0 0 1-.72.45.5.5 0 0 1-.5-.5v-2.38l-1.06-.06a1 1 0 0 1-.94-1v-.12a1 1 0 0 1 .94-1L4 8.88V6.5c0-.28.22-.5.5-.5.3 0 .59.17.72.45l1.14 2.28 2.46-.16L8.12 3A.9.9 0 0 1 9 2"],
    box: ["m10 7.96 3.03-1.21-7.5-3-2.59 1.03a2 2 0 0 0-.36.21zm-7.96-2.1Q2 6 2 6.17v7.64c0 .62.37 1.17.94 1.4l5.76 2.3q.4.15.8.21v-8.9zm8.46 11.87q.4-.05.8-.21l5.76-2.3c.57-.23.94-.78.94-1.4V6.18q0-.17-.04-.33L10.5 8.84zM17.42 5l-3.04 1.22-7.5-3 1.82-.73a3.5 3.5 0 0 1 2.6 0l5.76 2.3q.2.09.36.21"],
    building: ["M4 3.5C4 2.67 4.67 2 5.5 2h6c.83 0 1.5.67 1.5 1.5V8h1.5c.83 0 1.5.67 1.5 1.5v8a.5.5 0 0 1-.5.5H13v-3.5a.5.5 0 0 0-.5-.5h-5a.5.5 0 0 0-.5.5V18H4.5a.5.5 0 0 1-.5-.5zm2.75 3a.75.75 0 1 0 0-1.5.75.75 0 0 0 0 1.5m.75 2.25a.75.75 0 1 0-1.5 0 .75.75 0 0 0 1.5 0m-.75 3.75a.75.75 0 1 0 0-1.5.75.75 0 0 0 0 1.5m3.75-6.75a.75.75 0 1 0-1.5 0 .75.75 0 0 0 1.5 0M9.75 9.5a.75.75 0 1 0 0-1.5.75.75 0 0 0 0 1.5m.75 2.25a.75.75 0 1 0-1.5 0 .75.75 0 0 0 1.5 0m2.25.75a.75.75 0 1 0 0-1.5.75.75 0 0 0 0 1.5M12 15v3h-1.5v-3zm-2.5 0H8v3h1.5z"],
    home: ["M11 2.39a1.5 1.5 0 0 0-2 0L3.5 7.33c-.32.28-.5.69-.5 1.12v7.05c0 .83.67 1.5 1.5 1.5h2c.83 0 1.5-.67 1.5-1.5v-4c0-.28.22-.5.5-.5h3c.28 0 .5.22.5.5v4c0 .83.67 1.5 1.5 1.5h2c.83 0 1.5-.67 1.5-1.5V8.45c0-.43-.18-.84-.5-1.12z"],
    settings: ["M1.91 7.38A8.5 8.5 0 0 1 3.7 4.3a.5.5 0 0 1 .54-.13l1.92.68a1 1 0 0 0 1.32-.76l.36-2a.5.5 0 0 1 .4-.4 9 9 0 0 1 3.55 0q.32.08.38.4l.37 2a1 1 0 0 0 1.32.76l1.92-.68a.5.5 0 0 1 .54.13 8.5 8.5 0 0 1 1.78 3.08q.08.31-.15.54l-1.56 1.32a1 1 0 0 0 0 1.52l1.56 1.32a.5.5 0 0 1 .15.54 8.5 8.5 0 0 1-1.78 3.08.5.5 0 0 1-.54.13l-1.92-.68a1 1 0 0 0-1.32.76l-.37 2a.5.5 0 0 1-.38.4 8.5 8.5 0 0 1-3.56 0 .5.5 0 0 1-.39-.4l-.36-2a1 1 0 0 0-1.32-.76l-1.92.68a.5.5 0 0 1-.54-.13 8.5 8.5 0 0 1-1.78-3.08.5.5 0 0 1 .15-.54l1.56-1.32a1 1 0 0 0 0-1.52L2.06 7.92a.5.5 0 0 1-.15-.54M8 10a2 2 0 1 0 4 0 2 2 0 0 0-4 0"],
    wrench: ["M13.5 2a4.5 4.5 0 0 0-4.42 5.36l-6.42 6.66a2.36 2.36 0 0 0 3.37 3.3l6.37-6.46a4.5 4.5 0 0 0 5.49-5.37.5.5 0 0 0-.84-.24L14.5 7.79 12.2 5.5l2.55-2.55a.5.5 0 0 0-.24-.84 5 5 0 0 0-1-.11"],
    code: ["M11.78 1.53a.75.75 0 0 1 1.44.44l-5 16.5a.75.75 0 0 1-1.44-.44zM4.95 5a.75.75 0 1 1 1.1 1L2.57 9.83a.25.25 0 0 0 0 .34l3.48 3.82A.75.75 0 0 1 4.95 15l-3.49-3.82c-.6-.67-.6-1.7 0-2.36zm9.14-.06a.75.75 0 0 1 1.06.06l3.4 3.84c.6.66.6 1.66 0 2.32L15.16 15a.75.75 0 0 1-1.12-1l3.4-3.83a.25.25 0 0 0 0-.34L14.04 6a.75.75 0 0 1 .06-1.06"],
    bug: ["M9 2.5a.5.5 0 0 0-1 0V3q.01.6.32 1.08A3 3 0 0 0 6 7h-.5A1.5 1.5 0 0 1 4 5.5v-2a.5.5 0 0 0-1 0v2A2.5 2.5 0 0 0 5.5 8H6v1.5H2.5a.5.5 0 0 0 0 1H6V12h-.5A2.5 2.5 0 0 0 3 14.5v2a.5.5 0 0 0 1 0v-2c0-.83.67-1.5 1.5-1.5H6a4 4 0 0 0 8 0h.5c.83 0 1.5.67 1.5 1.5v2a.5.5 0 0 0 1 0v-2a2.5 2.5 0 0 0-2.5-2.5H14v-1.5h3.5a.5.5 0 0 0 0-1H14V8h.5A2.5 2.5 0 0 0 17 5.5v-2a.5.5 0 0 0-1 0v2c0 .83-.67 1.5-1.5 1.5H14a3 3 0 0 0-2.32-2.92A2 2 0 0 0 12 3v-.5a.5.5 0 0 0-1 0V3a1 1 0 1 1-2 0z"],
    database: ["M10 8c3.31 0 6-1.34 6-3s-2.69-3-6-3-6 1.34-6 3 2.69 3 6 3m4.69.02q.72-.36 1.31-.9V15c0 1.66-2.69 3-6 3s-6-1.34-6-3V7.12q.59.54 1.31.9C6.56 8.64 8.22 9 10 9s3.44-.36 4.69-.98"],
    cloud: ["M10 4c2.82 0 4.41 1.92 4.65 4.25h.07A3.33 3.33 0 0 1 18 11.62 3.33 3.33 0 0 1 14.72 15H5.28A3.33 3.33 0 0 1 2 11.62a3.33 3.33 0 0 1 3.28-3.37h.07C5.6 5.9 7.18 4 10 4"],
    lock: ["M10 1a4 4 0 0 1 4 4v2.05a2.5 2.5 0 0 1 2 2.45v6a2.5 2.5 0 0 1-2.5 2.5h-7A2.5 2.5 0 0 1 4 15.5v-6c0-1.2.86-2.22 2-2.45V5a4 4 0 0 1 4-4m0 10.5a1 1 0 1 0 0 2 1 1 0 0 0 0-2M10 2a3 3 0 0 0-3 3v2h6V5a3 3 0 0 0-3-3"],
    shield: ["M9.72 2.08a.5.5 0 0 1 .56 0q2.9 1.94 6.3 2.43A.5.5 0 0 1 17 5v4.5c0 3.9-2.3 6.73-6.82 8.47a.5.5 0 0 1-.36 0C5.31 16.23 3 13.39 3 9.5V5a.5.5 0 0 1 .43-.5 15 15 0 0 0 6.3-2.42"],
    chart: ["M2.75 2c.41 0 .75.34.75.75v12.5c0 .69.56 1.25 1.25 1.25h12.5a.75.75 0 0 1 0 1.5H4.75A2.75 2.75 0 0 1 2 15.25V2.75c0-.41.34-.75.75-.75M12 5.75c0-.41.34-.75.75-.75h3.5c.41 0 .75.34.75.75v3.5a.75.75 0 0 1-1.5 0V7.56l-3.97 3.97a.75.75 0 0 1-1.06 0L9 10.06l-2.22 2.22a.75.75 0 0 1-1.06-1.06l2.75-2.75a.75.75 0 0 1 1.06 0L11 9.94l3.44-3.44h-1.69a.75.75 0 0 1-.75-.75"],
};

/** French names accepted in the Icon field, written without accents (accents are ignored when matching). */
const ICON_ALIASES: Record<string, string> = {
    "drapeau": "flag",
    "jalon": "flag",
    "coche": "check",
    "valide": "check",
    "ok": "check",
    "termine": "check",
    "fait": "check",
    "alerte": "warning",
    "avertissement": "warning",
    "attention": "warning",
    "risque": "warning",
    "information": "info",
    "aide": "question",
    "cloche": "bell",
    "notification": "bell",
    "rappel": "bell",
    "etoile": "star",
    "favori": "star",
    "coeur": "heart",
    "cible": "target",
    "objectif": "target",
    "trophee": "trophy",
    "succes": "trophy",
    "fusee": "rocket",
    "lancement": "rocket",
    "eclair": "flash",
    "urgent": "flash",
    "ampoule": "lightbulb",
    "idee": "lightbulb",
    "loupe": "search",
    "recherche": "search",
    "revue": "search",
    "calendrier": "calendar",
    "reunion": "calendar",
    "horloge": "clock",
    "delai": "clock",
    "heure": "clock",
    "lieu": "location",
    "localisation": "location",
    "equipe": "people",
    "groupe": "people",
    "personne": "person",
    "utilisateur": "person",
    "accord": "handshake",
    "contrat": "handshake",
    "partenariat": "handshake",
    "discussion": "chat",
    "message": "chat",
    "courrier": "mail",
    "email": "mail",
    "e mail": "mail",
    "annonce": "megaphone",
    "communication": "megaphone",
    "fichier": "document",
    "livre": "book",
    "documentation": "book",
    "mallette": "briefcase",
    "projet": "briefcase",
    "argent": "money",
    "budget": "money",
    "euro": "money",
    "finance": "money",
    "panier": "cart",
    "achat": "cart",
    "commande": "cart",
    "camion": "truck",
    "livraison": "truck",
    "avion": "airplane",
    "voyage": "airplane",
    "colis": "box",
    "carton": "box",
    "produit": "box",
    "batiment": "building",
    "entreprise": "building",
    "maison": "home",
    "accueil": "home",
    "parametres": "settings",
    "reglages": "settings",
    "configuration": "settings",
    "outil": "wrench",
    "maintenance": "wrench",
    "developpement": "code",
    "bogue": "bug",
    "anomalie": "bug",
    "defaut": "bug",
    "base de donnees": "database",
    "donnees": "database",
    "nuage": "cloud",
    "cadenas": "lock",
    "securite": "lock",
    "bouclier": "shield",
    "protection": "shield",
    "graphique": "chart",
    "tendance": "chart",
    "analyse": "chart",
};

export const ICON_NAMES = Object.keys(ICON_PATHS);

export type MarkerIcon = { kind: "icon"; name: string } | { kind: "glyph"; text: string };

/**
 * Turns an Icon field value or a format pane choice into an icon: a known icon name (English or French),
 * or one or two characters (an emoji, a letter, a digit) drawn as text. Anything else gives no icon.
 */
export function resolveIcon(value: unknown): MarkerIcon | null {
    if (value === null || value === undefined) {
        return null;
    }
    const text = String(value).trim();
    if (!text || text === "none") {
        return null;
    }
    const normalized = text
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "")
        .toLowerCase()
        .replace(/[_-]+/g, " ")
        .replace(/\s+/g, " ");
    if (ICON_PATHS[normalized]) {
        return { kind: "icon", name: normalized };
    }
    if (ICON_ALIASES[normalized]) {
        return { kind: "icon", name: ICON_ALIASES[normalized] };
    }
    return countCharacters(text) <= 2 ? { kind: "glyph", text } : null;
}

function countCharacters(text: string): number {
    const Segmenter = (Intl as unknown as { Segmenter?: new () => { segment(input: string): Iterable<unknown> } }).Segmenter;
    return Segmenter ? Array.from(new Segmenter().segment(text)).length : Array.from(text).length;
}

/**
 * Draws an icon centred on (cx, cy), `size` pixels wide.
 */
export function createIcon(icon: MarkerIcon, cx: number, cy: number, size: number, color: string): SVGElement {
    if (icon.kind === "glyph") {
        const text = document.createElementNS(SVG_NS, "text");
        text.setAttribute("x", cx.toString());
        text.setAttribute("y", cy.toString());
        text.setAttribute("text-anchor", "middle");
        text.setAttribute("dominant-baseline", "central");
        text.setAttribute("fill", color);
        text.setAttribute("font-size", `${size}px`);
        text.setAttribute("font-weight", "600");
        text.setAttribute("pointer-events", "none");
        text.textContent = icon.text;
        return text;
    }
    const group = document.createElementNS(SVG_NS, "g");
    const scale = size / ICON_GRID;
    group.setAttribute("transform", `translate(${cx - size / 2} ${cy - size / 2}) scale(${scale})`);
    group.setAttribute("fill", color);
    group.setAttribute("pointer-events", "none");
    ICON_PATHS[icon.name].forEach((data) => {
        const path = document.createElementNS(SVG_NS, "path");
        path.setAttribute("d", data);
        group.appendChild(path);
    });
    return group;
}
