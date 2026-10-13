# PMO Portfolio Performance Dashboard: From Strategy to Success

A decision-ready Power BI report that turns monthly project snapshots into a clear view of portfolio performance, delivery risk and which projects to accelerate, recover, pause or terminate.

> Built for the *PMO Portfolio: From Strategy to Success* challenge. Add a dashboard screenshot here, for example `![Dashboard](images/dashboard.png)`.

## 1. Introduction

A project portfolio can look healthy while cost, schedule, capacity and benefits quietly drift. Most projects report green, yet the numbers underneath tell a different story.

This report gives PMO leaders one place to see the state of 100 projects across 10 programmes, 8 business units and 8 locations, from October 2023 to September 2026. It is organised into three pages:

- **Portfolio Performance:** how the portfolio is doing overall.
- **Delivery Reliability:** where milestones, capacity and risk are slipping.
- **Project Summary:** a one-project drill-down for review meetings.

## 2. Problem Statement

Leadership needs to decide where to spend scarce money and people. The report answers these questions:

- How has overall portfolio performance changed over time?
- Which programmes and business units vary most in performance?
- Where are the largest schedule delays and cost overruns?
- Where do resource-capacity gaps affect progress and milestone delivery?
- Which projects carry the highest risk, and are mitigation actions done on time?
- Which projects look healthy but fail to realise their planned benefits?
- Which projects give the best expected value relative to their approved budget?
- Which projects should be accelerated, recovered, paused or terminated?

## 3. Skills Demonstrated

- Star-schema design with a monthly fact table and a one-row-per-project snapshot table.
- DAX for earned value management (EV, CPI, SPI, EAC, VAC), reliability, risk and benefit KPIs.
- Rule-based project triage logic written in DAX.
- Dynamic number formats built with user-defined functions (UDFs).
- Measure auditing: dependency checks, duplicate removal and a display-folder structure by business domain.
- Advanced visuals (ZoomCharts drill-down bar, pie, facet, scatter and map) and an HTML Content visual driven by DAX.
- Reference labels that explain each KPI in plain language and respond to filters.
- Dashboard storytelling for a business audience.

## 4. Data Sourcing

The dataset is a multi-table PMO workbook loaded into Power BI with Import mode. The model has 14 tables.

| Table | Content |
|---|---|
| `FactProjectMonthlyTable` | 1,696 monthly project snapshots (Oct 2023 – Sep 2026) with cost, schedule, milestone, resource, risk and benefit fields |
| `DimProjectTable` | Project name, ID and strategic alignment score |
| `DimProgrammeTable` | 10 programmes |
| `DimBusinessUnitTable` | 8 business units |
| `DimLocationTable` | 8 locations with coordinates |
| `DimStrategicThemeTable` | 8 strategic themes |
| `DimDeliveryModelTable` | Delivery methodology |
| `DimEmployeeTable` | 36 project managers and sponsors |
| `DimDateTable` | Calendar table |
| `CurrentSnapshotTable` | Calculated table: the latest snapshot for each project |
| `ValidationTable`, `DataDictionaryTable` | Hidden reference tables |

Key data points:

- 100 projects: 55 active, 40 completed, 3 terminated and 2 paused.
- Total approved budget of about EUR 86.4M.
- Cost, schedule and progress, milestones, resource hours, risks and mitigation actions, benefits, stakeholder confidence and a health score of 0–100.

The source workbook is included in the project folder as `ecommerce_profitability_challenge_data.xlsx`.

[ZoomCharts Power BI Challenges](https://zoomcharts.com/en/microsoft-power-bi-custom-visuals/challenges/fp20-analytics-challenge-41)

## 5. Data Transformation

Data was loaded and typed in Power Query. A calculated `CurrentSnapshotTable` keeps one latest row per project, so current-state KPIs are not inflated by monthly history. Business metrics are DAX measures, grouped into display folders by page. The core ones are below.

### Earned value

```dax
EVM EV Current =
SUMX (
    'CurrentSnapshotTable',
    'CurrentSnapshotTable'[BaselineBudgetEUR] * 'CurrentSnapshotTable'[ActualProgressPct]
)

EVM CPI Current = DIVIDE ( [EVM EV Current], [EVM AC Current] )

EVM SPI Current = DIVIDE ( [EVM EV Current], [EVM PV Current] )

EVM ETC Current =
VAR RemainingEV = [EVM BAC Current] - [EVM EV Current]
RETURN
    DIVIDE ( RemainingEV, [EVM CPI Current] )

EVM EAC Current = [EVM AC Current] + [EVM ETC Current]
```

### Cost, progress and benefits

```dax
Forecast Cost Variance EUR Current =
SUMX (
    'CurrentSnapshotTable',
    'CurrentSnapshotTable'[ForecastTotalCostEUR] - 'CurrentSnapshotTable'[BaselineBudgetEUR]
)

Portfolio Progress Gap Current =
AVERAGEX (
    'CurrentSnapshotTable',
    'CurrentSnapshotTable'[ActualProgressPct] - 'CurrentSnapshotTable'[PlannedProgressPct]
)

Forecast Benefit to Budget Ratio Current =
DIVIDE (
    SUM ( 'CurrentSnapshotTable'[ForecastBenefitEUR] ),
    SUM ( 'CurrentSnapshotTable'[BaselineBudgetEUR] )
)

Benefit Realisation Rate Current =
DIVIDE (
    SUM ( 'CurrentSnapshotTable'[RealisedBenefitToDateEUR] ),
    SUM ( 'CurrentSnapshotTable'[PlannedBenefitToDateEUR] )
)
```

### Delivery reliability and risk

```dax
Milestone On-Time Rate Current =
DIVIDE (
    SUM ( 'CurrentSnapshotTable'[MilestonesCompletedOnTime] ),
    SUM ( 'CurrentSnapshotTable'[MilestonesDue] )
)

Resource Shortfall Hours Current =
SUMX (
    'CurrentSnapshotTable',
    IF (
        'CurrentSnapshotTable'[ResourceCapacityGapHours] < 0,
        -'CurrentSnapshotTable'[ResourceCapacityGapHours],
        0
    )
)

Projects Forecast Late Current =
CALCULATE ( [Projects Current], 'CurrentSnapshotTable'[ForecastDelayDays] > 0 )

Risk Exposure Score Current =
AVERAGEX (
    VALUES ( 'DimProjectTable'[ProjectKey] ),
    CALCULATE ( MAX ( CurrentSnapshotTable[RiskExposureScore] ) )
)
```

### Project triage

```dax
Project Triage Recommendation Current =
VAR Health = SELECTEDVALUE ( 'CurrentSnapshotTable'[PortfolioHealthScore] )
VAR Alignment = SELECTEDVALUE ( 'DimProjectTable'[StrategicAlignmentScore] )
VAR BenefitRatio = [Project Benefit to Budget Ratio Current]
VAR CriticalRiskCount = SELECTEDVALUE ( 'CurrentSnapshotTable'[CriticalRisks] )
VAR DelayDays = SELECTEDVALUE ( 'CurrentSnapshotTable'[ForecastDelayDays] )
RETURN
    SWITCH (
        TRUE (),
        ISBLANK ( Health ), BLANK (),
        Health < 35 && Alignment < 45 && BenefitRatio < 0.75, "Terminate review",
        Health < 40 && CriticalRiskCount > 0, "Pause / re-scope",
        Health < 65 || DelayDays > 180 || CriticalRiskCount >= 3, "Recover",
        Health >= 80 && Alignment >= 75 && BenefitRatio >= 1.5 && CriticalRiskCount = 0, "Accelerate",
        "Monitor"
    )
```

### Additional measures behind the dashboard visuals

The measures below feed the remaining cards, charts, table and the HTML visual.

#### Portfolio counts, health and progress

```dax
Projects Current = DISTINCTCOUNT ( 'CurrentSnapshotTable'[ProjectKey] )

Portfolio Health Score Current = AVERAGE ( 'CurrentSnapshotTable'[PortfolioHealthScore] )

Portfolio Health Percent Current = DIVIDE ( [Portfolio Health Score Current], 100 )

Actual Progress Curent = AVERAGEX ( 'CurrentSnapshotTable', 'CurrentSnapshotTable'[ActualProgressPct] )

Planned Progress Curent = AVERAGEX ( 'CurrentSnapshotTable', 'CurrentSnapshotTable'[PlannedProgressPct] )

Bar 100% = 1.00   -- full-width reference bar for the progress comparison chart

Forecast Cost Overrun % =
VAR ApprovedBudget = SUM ( 'CurrentSnapshotTable'[BaselineBudgetEUR] )
RETURN
    IF (
        ApprovedBudget = 0, BLANK (),
        DIVIDE ( SUM ( 'CurrentSnapshotTable'[ForecastTotalCostEUR] ), ApprovedBudget ) - 1
    )
```

#### Earned value helpers

```dax
EVM BAC Current = SUM ( 'CurrentSnapshotTable'[BaselineBudgetEUR] )
EVM AC Current  = SUM ( 'CurrentSnapshotTable'[ActualCostToDateEUR] )
EVM PV Current  = SUM ( 'CurrentSnapshotTable'[PlannedCostToDateEUR] )
EVM CV Current  = [EVM EV Current] - [EVM AC Current]
EVM VAC Current = [EVM BAC Current] - [EVM EAC Current]

Average CPI Current = [EVM CPI Current]   -- alias bound to the portfolio CPI card
Average SPI Current = [EVM SPI Current]   -- alias bound to the portfolio SPI card
```

#### Project Summary cards

```dax
Baseline Budget = SUM ( 'CurrentSnapshotTable'[BaselineBudgetEUR] )
Planned Cost    = SUM ( 'CurrentSnapshotTable'[PlannedCostToDateEUR] )
Actual Cost     = SUM ( 'CurrentSnapshotTable'[ActualCostToDateEUR] )

Schedule Variance EUR Net Current = SUM ( 'CurrentSnapshotTable'[ScheduleVarianceEUR_Net] )
Cost Variance EUR Net Current     = SUM ( 'CurrentSnapshotTable'[CostVarianceEUR_Net] )

Stakeholders Confidence Score = AVERAGE ( 'CurrentSnapshotTable'[StakeholderConfidenceScore] )
```

#### Delivery reliability supporting measures

```dax
Average Forecast Delay Days Current = AVERAGE ( 'CurrentSnapshotTable'[ForecastDelayDays] )

Critical Risks Current              = SUM ( 'CurrentSnapshotTable'[CriticalRisks] )
Overdue Mitigation Actions Current  = SUM ( 'CurrentSnapshotTable'[OverdueMitigationActions] )

Projects With Overdue Mitigation Current =
CALCULATE ( [Projects Current], 'CurrentSnapshotTable'[OverdueMitigationActions] > 0 )

Total Planned Resource Hours = SUM ( 'CurrentSnapshotTable'[PlannedResourceHours] )
```

#### Project diagnostics table

```dax
Project Delay Days Current = SELECTEDVALUE ( 'CurrentSnapshotTable'[ForecastDelayDays] )

Project Benefit to Budget Ratio Current =
DIVIDE (
    SELECTEDVALUE ( 'CurrentSnapshotTable'[ForecastBenefitEUR] ),
    SELECTEDVALUE ( 'CurrentSnapshotTable'[BaselineBudgetEUR] )
)

Delay Color =   -- conditional-format colour for the delay column
VAR _delay = [Project Delay Days Current]
RETURN
    SWITCH (
        TRUE (),
        ISBLANK ( _delay ), BLANK (),
        _delay < 0, "#34D17A",   -- ahead of schedule
        _delay = 0, "#8FA29A",   -- on time
        _delay > 0, "#FF7D72"    -- late
    )
```

#### Health-range filter (parameter table)

```dax
Health Range Low = MIN ( HealthScore_Parameter[HealthScore_Parameter] )

Health Filter Flag =   -- 1 = project falls inside the selected health range
VAR _Health = ROUND ( MAX ( 'CurrentSnapshotTable'[PortfolioHealthScore] ), 0 )
RETURN
    IF (
        NOT ISBLANK ( _Health )
            && _Health >= [Health Range Low]
            && _Health <= [Health Range High],
        1,
        0
    )
```

#### Project-level historical lookups (HTML visual)

Each project card reads the project's latest monthly snapshot. The same pattern is repeated for the other fields: SPI, delay, status, manager, role, dates and location.

```dax
Project Snapshot Date Historical =
VAR ProjectKey = SELECTEDVALUE ( 'DimProjectTable'[ProjectKey] )
VAR SnapshotDateKey =
    CALCULATE (
        MAX ( 'FactProjectMonthlyTable'[ReportingDateKey] ),
        FILTER ( ALL ( 'FactProjectMonthlyTable' ), 'FactProjectMonthlyTable'[ProjectKey] = ProjectKey )
    )
RETURN
    IF (
        ISBLANK ( ProjectKey ), BLANK (),
        LOOKUPVALUE ( 'DimDateTable'[Date], 'DimDateTable'[DateKey], SnapshotDateKey )
    )

Project SPI Historical =
VAR ProjectKey = SELECTEDVALUE ( 'DimProjectTable'[ProjectKey] )
VAR SnapshotDateKey =
    CALCULATE (
        MAX ( 'FactProjectMonthlyTable'[ReportingDateKey] ),
        FILTER ( ALL ( 'FactProjectMonthlyTable' ), 'FactProjectMonthlyTable'[ProjectKey] = ProjectKey )
    )
RETURN
    LOOKUPVALUE (
        'FactProjectMonthlyTable'[SchedulePerformanceIndex],
        'FactProjectMonthlyTable'[ProjectKey], ProjectKey,
        'FactProjectMonthlyTable'[ReportingDateKey], SnapshotDateKey
    )
```

`Project Progress HTML Historical` then assembles these values into one inline-styled HTML card: header, lifecycle stepper, progress bar with the planned marker, KPI tiles, schedule summary, manager details and a location map. The stage, colour and percentage logic uses `SWITCH`, `FORMAT` and string concatenation.

```dax
Project Progress HTML Historical =
VAR SelectedProject = SELECTEDVALUE ( 'DimProjectTable'[ProjectKey] )
VAR StageValue  = [Project Lifecycle Stage Historical]
VAR ActualValue = [Project Actual Progress Historical]
VAR StageIndex  = SWITCH ( StageValue, "Initiate & Design", 1, "Deliver", 2, "Deploy", 3, "Completed", 4, 1 )
VAR ActualPct   = FORMAT ( MAX ( 0, MIN ( 1, ActualValue ) ) * 100, "0.0", "en-US" ) & "%"
-- ...stage steps, KPI tiles, summary and manager grid built as HTML strings...
RETURN
    IF ( ISBLANK ( SelectedProject ), "Select one project to view its latest historical delivery snapshot.", Markup )
```

#### Reference labels

Reference labels turn each KPI into a sentence that updates with the filters.

```dax
ReferenceLabel | EVM CPI =
VAR CPIValue = [EVM CPI Current]
RETURN
    IF (
        ISBLANK ( CPIValue ),
        "CPI unavailable: no valid actual cost in the selected scope.",
        "Each €1.00 of actual cost earns " & FORMAT ( CPIValue, "€0.00" ) & " of value."
    )

ReferenceLabel | Delivery Resource Shortfall =
VAR PlannedHours = [Total Planned Resource Hours]
VAR CoverageRate = DIVIDE ( PlannedHours - [Resource Shortfall Hours Current], PlannedHours )
RETURN
    IF (
        ISBLANK ( CoverageRate ),
        "Resource coverage unavailable: no planned resource demand in the selected scope.",
        FORMAT ( CoverageRate, "0.0%" ) & " of planned resource demand is currently covered."
    )

ReferenceLabel | Risk Exposure =
VAR OverdueRate = DIVIDE ( [Projects With Overdue Mitigation Current], [Projects Current] )
RETURN
    IF (
        ISBLANK ( OverdueRate ),
        "Mitigation status unavailable for the selected scope.",
        FORMAT ( OverdueRate, "0%" ) & " of projects have overdue mitigation actions."
    )
```

#### Dynamic number formats

Currency, decimal, hours and days formats are UDFs (`pbi.NumberFormatCurrency`, `pbi.NumberFormatDecimal`, `pbi.NumberFormatDecimalHours` and `pbi.NumberFormatDecimalDays`). Measures call them through `formatStringExpression`, so values scale automatically in cards and tooltips.

## 6. Modeling

The model is a star schema. Dimension tables filter two fact-level tables:

- **`FactProjectMonthlyTable`** holds the monthly history and is used for trends. It links to `DimDateTable` and `DimProjectTable`.
- **`CurrentSnapshotTable`** holds one row per project and is used for current KPIs and triage. It links to the project, programme, business unit, location, strategic theme, delivery model and date dimensions.
- All 10 relationships are one-to-many with single-direction filtering.
- Measures sit in a dedicated `_Measure` table. They are organised in display folders: Portfolio Performance, Delivery Reliability, Project Summary, Reference Labels and Parameters. Each has KPI, Supporting and diagnostic subfolders.
- Distinct project counts use `DISTINCTCOUNT` on the snapshot table, so monthly rows are never double counted.
- A health-score parameter table lets users filter the report by a health range.

## 7. Analysis & Visualization

### Visuals and purpose

- **KPI cards with reference labels:** show CPI, SPI, forecast cost variance, health and benefit-to-budget, each with a plain-language explanation.
- **Drill-down combo bar (ZoomCharts):** compares actual and planned progress by programme.
- **Pie charts (ZoomCharts):** show portfolio mix by RAG status and priority, and shortfall hours by lifecycle stage.
- **Facet chart:** shows project counts by business unit and scenario archetype.
- **Map:** shows where projects are located, by status.
- **Scatter chart:** plots risk exposure against critical risks by business unit.
- **Table with colour flags:** lists project delay, risk, benefit realisation and the triage recommendation.
- **HTML Content visual:** a project card with lifecycle stage, progress, SPI, delay and manager details.
- **Tooltip page:** shows the progress gap by programme and risk category.

### Key insights

- **Cost efficiency is weak.** CPI is 0.91, so each EUR 1 spent earns EUR 0.91 of work. The forecast overrun is EUR 11.9M, and the estimate at completion is EUR 95.1M against an EUR 86.4M budget.
- **Schedule is also behind.** SPI is 0.94 and actual progress trails plan by 5.6 points. 48 of 100 projects are forecast to finish late, with an average forecast delay of 46 days.
- **Active projects are the problem.** Active projects have CPI 0.88 and SPI 0.90. Completed projects ended at CPI 0.95 and SPI 1.02, so cost losses were locked in even where schedules recovered.
- **Programmes vary widely.** Finance Transformation is the weakest (CPI 0.85, SPI 0.83, 97 days delay, risk score 95). Regulatory Readiness has the longest delay (100 days) despite a better CPI. Data and AI Enablement is the only programme ahead of plan (SPI 1.05, delay −4 days).
- **Business units follow the same pattern.** Technology is strongest (CPI 0.94, 2 days delay). Risk & Compliance is the largest source of shortfall: 11.4K of the 33.8K shortfall hours across 24 projects.
- **Trend:** average monthly CPI fell from about 1.02 in 2023 to about 0.97 in 2026, while SPI improved from 0.90 to 0.94. The portfolio is getting faster but more expensive.
- **Delivery reliability is a concern.** The milestone on-time rate is 79%, with 55 critical milestones overdue and 33,788 hours of resource shortfall.
- **Risk follow-through is poor.** There are 54 critical risks and 107 overdue mitigation actions, spread across 65 projects.
- **Green does not mean value.** 86 projects report Green, yet benefit realisation to date is only 72%, and 47 Green projects are below 70%. Expected benefits are still strong at 1.73x the budget.
- **Triage outcome:** 9 projects qualify to accelerate, 14 to recover and 77 to monitor. None meets the pause or terminate thresholds. Eleven projects are forecast more than 180 days late.

### KPI and interaction design

- Slicers cover programme, strategic theme, delivery model and location.
- A health-score range filter helps focus on weaker projects.
- Reference labels change with filters, so every KPI explains itself.
- Dynamic format strings keep currency, hours and days consistent.

You can interact with report here [FP20 Analytics](https://zoomcharts.com/en/microsoft-power-bi-custom-visuals/challenges/submission/077143850003bbbb868c7241555ae06a?challenge=zoomcharts-power-bi-challenge-september-2026)

![](dashboard_e_commerce.png)

## 8. Conclusion

The portfolio is not failing, but it is leaking cost. Delivery is slowly getting faster, yet costs, delays and capacity gaps remain concentrated in a few programmes. A green status alone is not a reliable signal of value.

### Recommendations

1. **Recover the 14 flagged projects first.** Start with the longest delays, such as PRJ-0020, PRJ-0062 and PRJ-0076, which are 300+ days late.
2. **Accelerate the 9 qualifying projects.** They have high health, strong alignment, a benefit ratio of 1.5x or more and no critical risks.
3. **Fix Finance Transformation and Regulatory Readiness.** Re-baseline their schedules and review scope and funding.
4. **Close the capacity gap in Risk & Compliance.** Re-balance resources before more milestones slip.
5. **Enforce mitigation deadlines.** Review the 107 overdue actions at monthly governance meetings.
6. **Track benefits, not just status.** Add benefit realisation to the Green criteria and review the 47 Green projects below 70%.
7. **Review the 5 paused or terminated projects** (CPI 0.74–0.79) to capture lessons and stop further spend.

## 9. Project Contents

- `Project Portfolio Performance.pbip`: the Power BI project.
- `Project Portfolio Performance.SemanticModel/`: model definition (TMDL).
- `Project Portfolio Performance.Report/`: report definition (PBIR).
- `README.md`: this document.

## 10. Notes

- Triage is a review aid, not an automated funding decision.
- ZoomCharts visuals need the ZoomCharts custom visuals installed.
