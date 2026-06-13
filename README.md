# BTC Intraday Market Behavior Analyzer

Excel VBA dashboard analyzing BTC 5-minute intraday market behavior, volatility, volume, market sessions, candle structure and hourly activity using Binance futures data.

This project is not a trading strategy.  
The goal is to demonstrate data analysis, Excel reporting, pivot tables, dashboard design and VBA automation skills.

## Project Overview

The workbook analyzes BTC intraday market behavior using 5-minute Binance futures data from January 2025 to April 2026.

The analysis focuses on:

- intraday volatility,
- session-based market activity,
- volume distribution,
- candle structure,
- hourly behavior,
- taker buy share,
- automated dashboard generation with VBA.

## Dataset

- Instrument: BTCUSD_PERP
- Source: Binance futures data
- Timeframe: 5 minutes
- Period: 2025-01 to 2026-04
- Records: approximately 139,680 5-minute candles
- Daily observations: approximately 485 days

The dataset contains OHLCV data, trade count and taker buy volume fields.

## Workbook Structure

| Sheet | Description |
|---|---|
| `BtcData` | Cleaned 5-minute BTC data with calculated candle metrics |
| `DailyData` | Daily aggregated market behavior metrics |
| `DailyAverages` | Daily summary tables and bucket analysis |
| `DailyPivot` | Pivot tables and charts based on daily metrics |
| `SessionData` | Session-level aggregation by date and market session |
| `PivotSessions` | Pivot tables and charts for intraday/session analysis |
| `Dashboard` | Final dashboard generated automatically with VBA |

## Key Metrics

### 5-minute candle metrics

- Candle Range
- Candle Range %
- Candle Body
- Candle Body %
- Candle Close Position
- Candle Taker Buy Share
- Candle Direction
- Body-to-Range Ratio
- Candle Type
- Market Session

### Daily metrics

- Daily Range
- Daily Range %
- Daily Body %
- Close Position Daily
- Close Strength
- Range Bucket
- Volume Bucket
- Daily Quote Volume
- Daily Trade Count
- Daily Taker Buy Share
- Daily Dollar Volume

### Session metrics

Each session row represents one day and one market session.

Sessions:

- Asia: 00:00-07:59 UTC
- London: 08:00-12:59 UTC
- New York: 13:00-20:59 UTC
- AfterHours: 21:00-23:59 UTC

Session metrics include:

- Session Open
- Session High
- Session Low
- Session Close
- Session Range
- Session Range %
- Session Quote Volume
- Session Trade Count
- Session Taker Buy Quote Volume
- Session Taker Buy Share

## Dashboard

The dashboard is generated automatically using VBA.

It includes:

- KPI cards,
- session volatility chart,
- session volume chart,
- hourly volatility chart,
- hourly volume chart,
- candle type structure by market session.

Main KPIs:

- Total Days
- Total 5m Candles
- Average Daily Range %
- Average 5m Range %
- Most Volatile Session
- Highest Volume Session

## Dashboard Screenshots

![Dashboard overview top](screenshots/dashboard_overview_top.png)

![Dashboard overview bottom](screenshots/dashboard_overview_bottom.png)

## VBA Automation

The dashboard layout is created by the `CreateDashboardLayout` macro.

The macro:

- creates or clears the Dashboard sheet,
- removes old shapes,
- sets the dashboard background,
- creates the title card,
- creates KPI label cards,
- calculates KPI values using formulas,
- copies charts from `PivotSessions`,
- embeds charts inside styled dashboard cards.

The VBA code was refactored into helper procedures:

- `AddTitleCard`
- `AddKPILabelCard`
- `AddKPIValueCard`
- `AddChartCard`

The source VBA module is available here:

```text
vba/modDashboard.bas
