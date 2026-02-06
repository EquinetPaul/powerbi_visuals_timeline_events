# powerbi_visuals_timeline_events

Custom Power BI visual that displays events on a horizontal timeline.

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
- Date
- Event Name
- Tooltip (one or more fields)

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
pbiviz start
```
