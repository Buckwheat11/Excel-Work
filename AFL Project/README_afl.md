# AFL Attendance Analysis Dashboard

An interactive Excel dashboard analysing AFL match attendance across 2,879 games spanning 13 seasons (2012 to 2025).

The dataset arrived clean, so this project is about deriving useful fields from the raw match data and building a dashboard that lets someone explore attendance patterns by team, venue and season rather than reading a fixed summary.

## Data source

AFL Stats on Kaggle:
https://www.kaggle.com/datasets/stoney71/aflstats

## Tools

- Microsoft Excel
- PivotTables and PivotCharts
- XLOOKUP and IF formulas
- Slicers for cross-filtering

## Derived columns

The source data has home and away scores but no result field, and attendance as a raw number with no grouping. Three calculated columns were added to make the data analysable.

**Winner (team name)**

Returns the winning team's name dynamically based on which score is higher.

```excel
=IF(P2>V2, K2, IF(P2<V2, Q2, "Draw"))
```

If the home score beats the away score it returns the home team name, if the away score is higher it returns the away team name, otherwise it returns Draw.

**Winner type**

The same logic, but returning the category rather than the team name. This is what allows analysis of home ground advantage across the whole dataset.

```excel
=IF(P2>V2, "Home Team", IF(P2<V2, "Away Team", "Draw"))
```

**Attendance category**

Buckets raw attendance figures into three bands so the data can be grouped and compared.

```excel
=IF(J2<30000, "Low", IF(J2<50000, "Medium", "High"))
```

Under 30,000 is Low, 30,000 to 50,000 is Medium, above 50,000 is High.

## Dashboard build

**PivotCharts**

Field buttons were hidden to keep the charts clean, via PivotChart Analyse, Field Buttons, then hiding the axis and value field buttons.

**KPI cards linked to PivotTables**

By default Excel inserts a `GETPIVOTDATA` formula when you reference a PivotTable cell, which locks the reference to specific field values and breaks when the PivotTable is filtered. This was turned off under PivotTable Analyse, Options, Generate GetPivotData.

To display PivotTable values in KPI cards with thousands separators:

```excel
=TEXT(A127, "#,##")
```

The `"#,##"` format string is required here; other formats did not render correctly in this setup.

## Problem solved: pivot cache conflicts

Seven interconnected PivotTables were sharing a cache, which caused cross-slicer filtering to return incorrect results while the figures still looked plausible. Diagnosing this took longer than building the original analysis. A custom sort order was also applied to correct a ranking issue where results were being ordered alphabetically rather than by value.

This is worth noting because the outputs gave no visible sign of being wrong. It is the main reason I check results that look reasonable rather than assuming they are correct.

## Findings

<!-- Add findings here: attendance trends by season, venue comparison, home ground advantage, finals vs home and away -->

## Repository contents

```
/                       Excel workbook (.xlsx)
/data                   Source dataset
README.md               This file
```

## Reference

Dashboard build approach informed by this tutorial:
https://www.youtube.com/watch?v=1ic8E58Bo2M
