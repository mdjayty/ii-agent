# Travel Agent - Dedicated Travel Planning Folder

This folder is the exclusive workspace for the II Agent Travel Agent skill. All travel planning documents, research, budgets, and itineraries are managed here.

**Access Boundary**: The Travel Agent only reads and writes within this `/travel/` folder.

## Folder Structure

```
travel/
|-- README.md                         <- You are here
|-- memory/                           <- Self-learning memory system
|   |-- approach.md                   <- Research methodology and workflow
|   |-- research_patterns.md          <- Patterns learned over time
|   |-- lessons_learned.md            <- What worked/didn't, improvements
|-- europe-2025/                      <- Trip: Europe 2025 (Sep, 21 days)
|   |-- itinerary/
|   |   |-- itinerary_v1.md           <- Full day-by-day itinerary (versioned)
|   |-- budget/
|   |   |-- budget_v1.xlsx            <- Excel budget with hyperlinks & formulas
|   |   |-- generate_budget.py        <- Python script to regenerate budget
|   |-- research/
|   |   |-- hotels/                   <- Hotel comparison research
|   |   |-- flights/                  <- Flight research
|   |   |-- activities/               <- Activity research
|   |   |-- transportation/           <- Train/transit research
|   |-- calendar/
|   |   |-- daily_calendar.md         <- Day-by-day calendar in thirds
|   |-- archive/                      <- Old versions of files
```

## Versioning Protocol

Before modifying any itinerary or budget file:
1. Copy the current file to `archive/` (e.g., `itinerary_v1.md` -> `archive/itinerary_v1.md`)
2. Create a new version (e.g., `itinerary_v2.md`) with changes
3. This ensures you can always revert to a previous version

## Key Rules

- Each day is divided into **3 thirds**: morning (8-12), afternoon (12-5), evening (5-10)
- **Max 1 major activity** per third of the day
- **At least 1 meal** per third of the day
- All research includes **source hyperlinks**
- Budget is managed in **Excel with formulas**
- Memory files are read before and updated after each research session

## Current Trips

| Trip | Dates | Status |
|------|-------|--------|
| Europe 2025 | September 2025, 21 days | Baseline plan complete |
