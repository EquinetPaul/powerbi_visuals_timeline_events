# powerbi_visuals_timeline_events

Custom Power BI visual that displays events on a horizontal timeline.

![Screenshot of the visual](assets/screenshot_sample.png)

## Features
- Horizontal timeline with automatic date ticks (days, months, years), event dates, or start and end dates
- Events colored by **Legend** group with the report theme colors, and a color picker per group
- Event labels above, below or alternating, stacked on rows with leader lines; labels that don't fit are hidden
- Same-date events spread side by side, never on top of each other
- Native Power BI tooltips, including report page tooltips
- Selection and cross-filtering: click an event or a legend item, Ctrl+click to add, click the background to clear
- Context menu, keyboard navigation (arrow keys, Enter), bookmarks
- Optional "Today" line
- Format pane in the Power BI style: Legend, Colors, Markers, Date axis, Event labels, Today line
- English and French, numbers and dates formatted with the model's format strings and the report locale
- High contrast mode, landing page, rendering events

## Data Roles
- Date (one field: a date column, or text dates such as `2026-01-05`)
- Event name (one field)
- Legend (optional, one field): groups and colors the events
- Tooltips (one or more fields, formatted with the model's format strings)

> If the Date field shows a date hierarchy (Year / Quarter / Month / Day), click the arrow next to the
> field in the Date well and select the column name instead of "Date Hierarchy".

## Sample Data (use for screenshot)

| Date       | Event Name       | Legend    | Tooltip1        | Tooltip2        |
|------------|------------------|-----------|-----------------|-----------------|
| 2026-01-05 | Project Kickoff  | Milestone | Owner: Alex     | Status: Open    |
| 2026-01-12 | Design Review    | Review    | Owner: Kim      | Status: Done    |
| 2026-01-12 | Risk Identified  | Risk      | Severity: High  | Status: Open    |
| 2026-02-03 | Prototype Ready  | Delivery  | Owner: Sam      | Status: Done    |
| 2026-02-18 | Launch           | Milestone | Region: EU      | Status: Planned |

## Development
```bash
npm install
npm start          # pbiviz start --all-locales
npm run package    # pbiviz package --all-locales -> dist/*.pbiviz
npm run lint
```

`--all-locales` keeps every culture of `powerbi-visuals-utils-formattingutils` in the bundle, so numbers and
dates follow the report's locale (e.g. `1 234,50 €` in French). Without it, `pbiviz` only keeps the locales
that have a `stringResources` folder.

### Data mappings
- A **table** mapping carries the events (one row per Date / Event name / Legend / Tooltips combination,
  up to 30,000 rows). Selection ids are built per row with `withTable`.
- A **categorical** mapping on the Legend field carries one identity per legend value. The Colors card stores
  each legend value's color against that identity (`withCategory`), as recommended by Microsoft.

### Microsoft recommendations
Implemented: rendering events, allow interactions, color palette, context menu, high contrast, keyboard focus,
landing page, localization, multi-visual selection, tooltips (default, report page and modern tooltips).

Not applicable: **Highlight data** and **conditional formatting (fx)** are not supported by Power BI for
table data mappings, so `pbiviz package` still lists "Highlight Data" as missing.
