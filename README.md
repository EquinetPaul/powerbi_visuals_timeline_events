# powerbi_visuals_timeline_events

Custom Power BI visual that displays events on a horizontal timeline.

![Screenshot of the visual](assets/screenshot_sample.png)

## Features
- Horizontal timeline with min/max date bounds
- Event symbols plotted by date (optional per-event symbol shapes)
- Tooltip per event with extra fields (Tooltip1, Tooltip2, ...)
- Optional event name labels with auto-fit/truncation
- Optional date labels below the timeline
- Adjustable scaling for line, text, and symbols
- Symbol color picker + hover highlight
- Optional legend when event symbols are enabled

## Data Roles
- Date (one field: a date column, or text dates such as `2026-01-05`)
- Event Name (one field)
- Tooltip (one or more fields, formatted with the model's format strings)

> If the Date field shows a date hierarchy (Year / Quarter / Month / Day), click the arrow next to the
> field in the Date well and select the column name instead of "Date Hierarchy".

## Sample Data (use for screenshot)

| Date       | Event Name       | Tooltip1        | Tooltip2        |
|------------|------------------|-----------------|-----------------|
| 2026-01-05 | Project Kickoff  | Owner: Alex     | Status: Open    |
| 2026-01-12 | Design Review    | Owner: Kim      | Status: Done    |
| 2026-01-12 | Risk Identified  | Severity: High  | Status: Open    |
| 2026-02-03 | Prototype Ready  | Owner: Sam      | Status: Done    |
| 2026-02-18 | Launch           | Region: EU      | Status: Planned |

## Development
```bash
npm install
npm start          # pbiviz start --all-locales
npm run package    # pbiviz package --all-locales -> dist/*.pbiviz
```

`--all-locales` keeps every culture of `powerbi-visuals-utils-formattingutils` in the bundle, so numbers and
dates follow the report's locale (e.g. `1 234,50 €` in French). Without it, `pbiviz` only keeps `en-US`.
