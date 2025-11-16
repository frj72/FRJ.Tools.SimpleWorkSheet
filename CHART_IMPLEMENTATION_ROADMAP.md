# Feature 7: Charts & Visualizations - Implementation Roadmap

**Feature Complexity:** 8/10 | **Time Estimate:** 9-13 hours  
**Status:** Not Started  
**Started:** [Date TBD]  
**Target Completion:** [Date TBD]

---

## Overview

This document breaks down Feature 7 (Charts & Visualizations) into independent, testable steps that can be implemented over multiple sessions. Each step builds on previous work while remaining independently testable.

**Goal:** Create embedded charts (Bar, Line, Pie, Scatter) in Excel worksheets with positioning, sizing, and styling support.

---

## Implementation Phases

### **Phase A: Foundation (Steps 1-2)** - ~1 hour
Building the base chart infrastructure without OpenXML complexity.

### **Phase B: Bar Chart MVP (Steps 3-7)** - ~4-5 hours
Implement one complete chart type end-to-end to validate the architecture.

### **Phase C: Additional Chart Types (Steps 8-10)** - ~2-3 hours
Extend to other chart types now that the architecture is proven.

### **Phase D: Styling & Polish (Steps 11-13)** - ~2-3 hours
Add visual customization options.

### **Phase E: Integration & Documentation (Steps 14-15)** - ~1 hour
Final validation and documentation.

---

## Detailed Step-by-Step Plan

### **PHASE A: FOUNDATION**

#### ✅ Step 1: Chart Base Infrastructure
**Time Estimate:** 30 minutes  
**Status:** ✅ Completed  
**Dependencies:** None

**Tasks:**
- [x] Create `Chart` abstract base class in `Components/Charts/Chart.cs`
- [x] Create `ChartPosition` record (FromColumn, FromRow, ToColumn, ToRow)
- [x] Create `ChartSize` record (Width, Height in EMUs)
- [x] Create `ChartType` enum (Bar, Line, Pie, Scatter)
- [x] Add abstract method `string GetChartTypeName()` to Chart

**Files to Create:**
- `FRJ.Tools.SimpleWorkSheet/Components/Charts/Chart.cs`
- `FRJ.Tools.SimpleWorkSheet/Components/Charts/ChartPosition.cs`
- `FRJ.Tools.SimpleWorkSheet/Components/Charts/ChartSize.cs`
- `FRJ.Tools.SimpleWorkSheet/Components/Charts/ChartType.cs`

**Testable Success Criteria:**
- Chart base class can be instantiated by derived classes
- ChartPosition validates coordinates (from < to)
- ChartSize validates dimensions (> 0)
- Unit tests: `ChartTests.cs` (basic instantiation tests)

**Notes:**
- EMU (English Metric Unit) = 1/914400 of an inch
- Default chart size: ~6000000 x 4000000 EMUs (~6.5" x 4.3")

---

#### ✅ Step 2: Chart Data Series
**Time Estimate:** 30 minutes  
**Status:** ✅ Completed  
**Dependencies:** Step 1

**Tasks:**
- [x] Create `ChartSeries` record (Name, DataRange)
- [x] Create `ChartDataRange` helper (validates CellRange)
- [x] Add `List<ChartSeries> Series` property to Chart base class
- [x] Add `AddSeries(string name, CellRange dataRange)` method to Chart

**Files to Create:**
- `FRJ.Tools.SimpleWorkSheet/Components/Charts/ChartSeries.cs`
- `FRJ.Tools.SimpleWorkSheet/Components/Charts/ChartDataRange.cs`

**Testable Success Criteria:**
- ChartSeries stores name and data range
- Can add multiple series to a chart
- Data range validation works (rejects invalid ranges)
- Unit tests: `ChartSeriesTests.cs`

---

### **PHASE B: BAR CHART MVP**

#### ✅ Step 3: Bar Chart Implementation
**Time Estimate:** 1 hour  
**Status:** Not Started  
**Dependencies:** Steps 1-2

**Tasks:**
- [ ] Create `BarChart` class extending Chart in `Components/Charts/BarChart.cs`
- [ ] Add `BarChartOrientation` enum (Vertical, Horizontal)
- [ ] Implement fluent builder methods:
  - [ ] `WithTitle(string title)`
  - [ ] `WithDataRange(CellRange categoriesRange, CellRange valuesRange)`
  - [ ] `WithPosition(uint fromCol, uint fromRow, uint toCol, uint toRow)`
  - [ ] `WithSize(int width, int height)`
  - [ ] `WithOrientation(BarChartOrientation orientation)`
- [ ] Add constructor and static `Create()` factory method

**Files to Create:**
- `FRJ.Tools.SimpleWorkSheet/Components/Charts/BarChart.cs`
- `FRJ.Tools.SimpleWorkSheet/Components/Charts/BarChartOrientation.cs`

**Testable Success Criteria:**
- BarChart can be created with fluent API
- All properties are settable and retrievable
- Validation works (e.g., position coordinates, size > 0)
- Unit tests: `BarChartTests.cs`

**Example API:**
```csharp
var chart = BarChart.Create()
    .WithTitle("Sales by Region")
    .WithDataRange(categoriesRange, valuesRange)
    .WithPosition(5, 0, 10, 15)
    .WithSize(6000000, 4000000);
```

---

#### ✅ Step 4: WorkSheet Chart Integration
**Time Estimate:** 30 minutes  
**Status:** Not Started  
**Dependencies:** Step 3

**Tasks:**
- [ ] Add `List<Chart> Charts` property to WorkSheet
- [ ] Add `AddChart(Chart chart)` method to WorkSheet
- [ ] Add position validation (warn if chart overlaps cells with data)

**Files to Modify:**
- `FRJ.Tools.SimpleWorkSheet/Components/Sheet/WorkSheet.cs`

**Testable Success Criteria:**
- Charts can be added to WorkSheet
- Multiple charts can be added
- Charts collection is accessible
- Unit tests: `WorkSheetTests.cs` (AddChart tests)

---

#### ✅ Step 5: SheetConverter Chart Serialization
**Time Estimate:** 3-4 hours (OpenXML complexity)  
**Status:** Not Started  
**Dependencies:** Steps 3-4

**Tasks:**
- [ ] Research OpenXML chart structure (DrawingsPart, ChartPart)
- [ ] Add DrawingsPart to worksheetPart in SheetConverter
- [ ] Add ChartPart with relationship to DrawingsPart
- [ ] Generate ChartSpace XML with language settings
- [ ] Generate BarChart XML (BarChartSeries, categories, values)
- [ ] Generate CategoryAxis and ValueAxis XML
- [ ] Generate TwoCellAnchor positioning (FromMarker, ToMarker)
- [ ] Generate GraphicFrame with chart reference
- [ ] Handle EMU coordinate conversions

**Files to Modify:**
- `FRJ.Tools.SimpleWorkSheet/LowLevel/SheetConverter.cs`

**Testable Success Criteria:**
- Excel file generates without errors
- File opens in Excel without corruption warnings
- Bar chart appears in correct position
- Chart displays data correctly
- Integration test: Generate file and manually open in Excel

**OpenXML Structure Reference:**
```
WorksheetPart
  ├── DrawingsPart
  │     └── TwoCellAnchor (positioning)
  │           └── GraphicFrame
  │                 └── ChartPart reference
  └── ChartPart
        └── ChartSpace
              ├── Chart
              │     ├── PlotArea
              │     │     ├── BarChart
              │     │     │     └── BarChartSeries
              │     │     ├── CategoryAxis
              │     │     └── ValueAxis
              │     └── Legend
              └── ExternalData
```

**Notes:**
- EMU calculations: Column/row to EMU conversion needed
- Relationship IDs must be unique
- Chart must reference worksheet data ranges

---

#### ✅ Step 6: Unit Tests for Phase B
**Time Estimate:** 30 minutes  
**Status:** Not Started  
**Dependencies:** Steps 1-5

**Tasks:**
- [ ] Complete `ChartTests.cs` - base class tests
- [ ] Complete `ChartSeriesTests.cs` - series management
- [ ] Complete `BarChartTests.cs` - builder and properties
- [ ] Complete `WorkSheetTests.cs` - AddChart method
- [ ] Add integration test: `ChartIntegrationTests.cs`
  - [ ] Test: Generate Excel file with bar chart
  - [ ] Test: Verify file structure (DrawingsPart exists)
  - [ ] Test: Verify chart relationships

**Files to Create:**
- `FRJ.Tools.SimpleWorksheetTests/ChartTests.cs`
- `FRJ.Tools.SimpleWorksheetTests/ChartSeriesTests.cs`
- `FRJ.Tools.SimpleWorksheetTests/BarChartTests.cs`
- `FRJ.Tools.SimpleWorksheetTests/ChartIntegrationTests.cs`

**Testable Success Criteria:**
- All unit tests pass
- Integration test generates valid Excel file
- Code coverage > 80% for chart code

---

#### ✅ Step 7: Bar Chart Example & Validation
**Time Estimate:** 30 minutes  
**Status:** Not Started  
**Dependencies:** Steps 1-6

**Tasks:**
- [ ] Create `BarChartExample.cs` in Examples/AdvancedExamples
- [ ] Create sample data (regions, sales figures)
- [ ] Generate bar chart from data
- [ ] Add example to Program.cs
- [ ] Generate example file: `41_BarChart.xlsx`
- [ ] Manually validate in Excel:
  - [ ] File opens without errors
  - [ ] Chart renders correctly
  - [ ] Chart position is correct
  - [ ] Data displays correctly
- [ ] Add to ExampleFileValidationTests.cs

**Files to Create:**
- `FRJ.Tools.SimpleWorkSheet.Examples/Examples/AdvancedExamples/BarChartExample.cs`

**Files to Modify:**
- `FRJ.Tools.SimpleWorkSheet.Examples/Program.cs`
- `FRJ.Tools.SimpleWorksheetTests/ExampleFileValidationTests.cs`

**Testable Success Criteria:**
- Example generates without errors
- Excel file opens and displays chart
- Validation tests pass

**🛑 CHECKPOINT:** If bar chart doesn't work properly, fix issues before proceeding to other chart types.

---

### **PHASE C: ADDITIONAL CHART TYPES**

#### ✅ Step 8: Line Chart Implementation
**Time Estimate:** 1 hour  
**Status:** Not Started  
**Dependencies:** Steps 1-7

**Tasks:**
- [ ] Create `LineChart` class extending Chart
- [ ] Add `LineChartMarkerStyle` enum (None, Circle, Square, Diamond, Triangle)
- [ ] Add `WithMarkers(LineChartMarkerStyle style)` method
- [ ] Add `WithSmoothLines(bool smooth)` method
- [ ] Update SheetConverter to handle LineChart serialization
- [ ] Create `LineChartTests.cs`
- [ ] Create `LineChartExample.cs` (trend over time)
- [ ] Generate `42_LineChart.xlsx`

**Files to Create:**
- `FRJ.Tools.SimpleWorkSheet/Components/Charts/LineChart.cs`
- `FRJ.Tools.SimpleWorkSheet/Components/Charts/LineChartMarkerStyle.cs`
- `FRJ.Tools.SimpleWorksheetTests/LineChartTests.cs`
- `FRJ.Tools.SimpleWorkSheet.Examples/Examples/AdvancedExamples/LineChartExample.cs`

**Testable Success Criteria:**
- Line chart generates and displays correctly
- Markers and smooth lines work
- Unit tests pass
- Example validates

---

#### ✅ Step 9: Pie Chart Implementation
**Time Estimate:** 1 hour  
**Status:** Not Started  
**Dependencies:** Steps 1-7

**Tasks:**
- [ ] Create `PieChart` class extending Chart
- [ ] Add `WithExplosion(uint percentage)` method (0-100, separates slices)
- [ ] Add `WithFirstSliceAngle(uint degrees)` method (0-360, rotation)
- [ ] Update SheetConverter to handle PieChart serialization
- [ ] Create `PieChartTests.cs`
- [ ] Create `PieChartExample.cs` (market share distribution)
- [ ] Generate `43_PieChart.xlsx`

**Files to Create:**
- `FRJ.Tools.SimpleWorkSheet/Components/Charts/PieChart.cs`
- `FRJ.Tools.SimpleWorksheetTests/PieChartTests.cs`
- `FRJ.Tools.SimpleWorkSheet.Examples/Examples/AdvancedExamples/PieChartExample.cs`

**Testable Success Criteria:**
- Pie chart generates and displays correctly
- Explosion and rotation work
- Unit tests pass
- Example validates

**Notes:**
- Pie charts don't have axes
- Different XML structure than bar/line charts

---

#### ✅ Step 10: Scatter Chart Implementation
**Time Estimate:** 1 hour  
**Status:** Not Started  
**Dependencies:** Steps 1-7

**Tasks:**
- [ ] Create `ScatterChart` class extending Chart
- [ ] Add `WithXYData(CellRange xRange, CellRange yRange)` method
- [ ] Add `WithTrendline(bool show)` method
- [ ] Update SheetConverter to handle ScatterChart serialization
- [ ] Create `ScatterChartTests.cs`
- [ ] Create `ScatterChartExample.cs` (correlation data)
- [ ] Generate `44_ScatterChart.xlsx`

**Files to Create:**
- `FRJ.Tools.SimpleWorkSheet/Components/Charts/ScatterChart.cs`
- `FRJ.Tools.SimpleWorksheetTests/ScatterChartTests.cs`
- `FRJ.Tools.SimpleWorkSheet.Examples/Examples/AdvancedExamples/ScatterChartExample.cs`

**Testable Success Criteria:**
- Scatter chart generates and displays correctly
- X/Y data mapping works
- Unit tests pass
- Example validates

---

### **PHASE D: STYLING & POLISH**

#### ✅ Step 11: Chart Styling
**Time Estimate:** 2-3 hours  
**Status:** Not Started  
**Dependencies:** Steps 8-10

**Tasks:**
- [ ] Create `ChartLegend` record (Position: Top/Bottom/Left/Right/None, ShowLegend)
- [ ] Create `ChartTitle` record (Text, FontSize, Bold)
- [ ] Create `ChartAxis` record (Title, MinValue, MaxValue, ShowLabels)
- [ ] Create `ChartColors` class (series color list)
- [ ] Add styling properties to Chart base class
- [ ] Add fluent methods:
  - [ ] `WithLegend(ChartLegend legend)`
  - [ ] `WithTitle(ChartTitle title)`
  - [ ] `WithXAxis(ChartAxis axis)`
  - [ ] `WithYAxis(ChartAxis axis)`
  - [ ] `WithColors(params string[] colors)`
- [ ] Update SheetConverter to serialize styling
- [ ] Update all chart types to support styling

**Files to Create:**
- `FRJ.Tools.SimpleWorkSheet/Components/Charts/ChartLegend.cs`
- `FRJ.Tools.SimpleWorkSheet/Components/Charts/ChartTitle.cs`
- `FRJ.Tools.SimpleWorkSheet/Components/Charts/ChartAxis.cs`
- `FRJ.Tools.SimpleWorkSheet/Components/Charts/ChartColors.cs`

**Files to Modify:**
- All chart classes (Bar, Line, Pie, Scatter)
- SheetConverter.cs

**Testable Success Criteria:**
- Styled charts render correctly in Excel
- Legend position works
- Axis titles and ranges work
- Custom colors apply
- Unit tests: `ChartStylingTests.cs`

---

#### ✅ Step 12: Complete Unit Tests
**Time Estimate:** 1 hour  
**Status:** Not Started  
**Dependencies:** Steps 8-11

**Tasks:**
- [ ] Review test coverage for all chart types
- [ ] Add edge case tests:
  - [ ] Empty data ranges
  - [ ] Invalid position (overlapping charts)
  - [ ] Invalid size (too small, too large)
  - [ ] Null/empty series names
- [ ] Add styling tests for each chart type
- [ ] Ensure all tests pass

**Testable Success Criteria:**
- Code coverage > 80% for all chart code
- All edge cases covered
- All tests pass

---

#### ✅ Step 13: Examples for All Chart Types
**Time Estimate:** 1 hour  
**Status:** Not Started  
**Dependencies:** Steps 8-12

**Tasks:**
- [ ] Ensure all individual chart examples exist:
  - [ ] `41_BarChart.xlsx`
  - [ ] `42_LineChart.xlsx`
  - [ ] `43_PieChart.xlsx`
  - [ ] `44_ScatterChart.xlsx`
- [ ] Create `ComprehensiveChartsExample.cs`
  - [ ] One worksheet with all 4 chart types
  - [ ] Demonstrates different styling options
  - [ ] Generate `45_ComprehensiveCharts.xlsx`
- [ ] Register all examples in Program.cs
- [ ] Update ExampleFileValidationTests.cs (expectedCount = 45)

**Files to Create:**
- `FRJ.Tools.SimpleWorkSheet.Examples/Examples/AdvancedExamples/ComprehensiveChartsExample.cs`

**Testable Success Criteria:**
- All 5 chart examples generate
- All files open in Excel
- All charts render correctly

---

### **PHASE E: INTEGRATION & DOCUMENTATION**

#### ✅ Step 14: Generate All Examples & Validation
**Time Estimate:** 30 minutes  
**Status:** Not Started  
**Dependencies:** Steps 1-13

**Tasks:**
- [ ] Run `dotnet build --configuration Release`
- [ ] Run `dotnet test --verbosity normal`
  - [ ] Verify all tests pass (should be ~330+ tests)
- [ ] Run `cd FRJ.Tools.SimpleWorkSheet.Examples && dotnet run --configuration Release -- all`
  - [ ] Verify all 45 examples generate
- [ ] Run ExampleFileValidationTests
- [ ] Run OpenXmlValidationTests
- [ ] Manually open each chart example in Excel:
  - [ ] 41_BarChart.xlsx - verify chart displays
  - [ ] 42_LineChart.xlsx - verify chart displays
  - [ ] 43_PieChart.xlsx - verify chart displays
  - [ ] 44_ScatterChart.xlsx - verify chart displays
  - [ ] 45_ComprehensiveCharts.xlsx - verify all 4 charts display

**Testable Success Criteria:**
- Build: 0 warnings, 0 errors
- Tests: All pass
- Examples: All 45 generate
- Validation: All files valid
- Manual: All charts display correctly in Excel

---

#### ✅ Step 15: Update ROADMAP.md
**Time Estimate:** 30 minutes  
**Status:** Not Started  
**Dependencies:** Step 14

**Tasks:**
- [ ] Update "Recent Updates" section
  - [ ] Add Phase 3 Visualization complete
  - [ ] Add Feature 7 completion details
- [ ] Update Feature 7 status to "✅ Completed and Verified"
- [ ] Check all implementation tasks
- [ ] Update test counts (296 → ~330+)
- [ ] Update example counts (40 → 45)
- [ ] Update "New Features" section with chart details
- [ ] Update "Examples Generated" section
- [ ] Update "Project Statistics"
  - [ ] Current Features: 16 → 17
  - [ ] Test Coverage: 296 → ~330+
  - [ ] Working Examples: 40 → 45

**Files to Modify:**
- `ROADMAP.md`

**Testable Success Criteria:**
- ROADMAP.md accurately reflects current state
- All statistics updated
- Feature 7 marked complete

---

## Testing Strategy Summary

| Phase | Test Type | Success Criteria |
|-------|-----------|------------------|
| A | Unit | Classes instantiate, properties validate |
| B | Unit + Integration | Excel file generates and opens with bar chart |
| C | Unit + Integration | All chart types work in Excel |
| D | Unit + Visual | Styled charts render correctly |
| E | End-to-end | All 330+ tests pass, all 45 examples valid |

---

## Stop Points & Risk Mitigation

### **🛑 Critical Stop Point: After Step 7**
**Do not proceed to Phase C until:**
- Bar chart displays correctly in Excel
- No corruption warnings when opening file
- Chart position and size are accurate
- Data displays correctly in chart

**If bar chart fails:**
1. Review OpenXML structure in SheetConverter
2. Check EMU calculations
3. Verify relationship IDs
4. Test with minimal data first
5. Compare generated XML with working Excel file

### **High Risk Areas:**

1. **OpenXML Complexity (Step 5)**
   - Mitigation: Research Microsoft docs, use minimal example first
   - Mitigation: Compare generated XML with manually created Excel chart

2. **Coordinate Calculations (Steps 1, 5)**
   - Mitigation: Create conversion helper methods
   - Mitigation: Test with known positions (A1:C10, etc.)

3. **Excel Compatibility (All steps)**
   - Mitigation: Test each chart type immediately after implementation
   - Mitigation: Use ExampleFileValidationTests
   - Mitigation: Test on actual Excel (not just LibreOffice)

---

## Progress Tracking

### Phase A: Foundation
- [x] Step 1: Chart Base Infrastructure
- [x] Step 2: Chart Data Series

### Phase B: Bar Chart MVP
- [ ] Step 3: Bar Chart Implementation
- [ ] Step 4: WorkSheet Chart Integration
- [ ] Step 5: SheetConverter Chart Serialization
- [ ] Step 6: Unit Tests for Phase B
- [ ] Step 7: Bar Chart Example & Validation
- [ ] **Checkpoint:** Bar chart working in Excel?

### Phase C: Additional Chart Types
- [ ] Step 8: Line Chart Implementation
- [ ] Step 9: Pie Chart Implementation
- [ ] Step 10: Scatter Chart Implementation

### Phase D: Styling & Polish
- [ ] Step 11: Chart Styling
- [ ] Step 12: Complete Unit Tests
- [ ] Step 13: Examples for All Chart Types

### Phase E: Integration & Documentation
- [ ] Step 14: Generate All Examples & Validation
- [ ] Step 15: Update ROADMAP.md

---

## Quick Reference Commands

### Build & Test
```bash
cd /Users/frj/RiderProjects/FRJ.Tools.SimpleWorkSheet

# Build
dotnet build --configuration Release

# Run all tests
dotnet test --verbosity normal

# Run specific test
dotnet test --filter "FullyQualifiedName~BarChartTests"
```

### Generate Examples
```bash
cd FRJ.Tools.SimpleWorkSheet.Examples

# Generate all examples
dotnet run --configuration Release -- all

# Generate single example (for testing)
dotnet run --configuration Release
# Then select the example number
```

### Validation
```bash
# Run file validation tests
dotnet test --filter "FullyQualifiedName~ExampleFileValidationTests"

# Run OpenXML validation
dotnet test --filter "FullyQualifiedName~OpenXmlValidationTests"
```

---

## Session Log

Use this section to track progress across sessions.

### Session 1: Nov 16, 2025
**Completed:**
- [x] Step 1: Chart Base Infrastructure
- [x] Step 2: Chart Data Series
- [x] Unit tests for both steps (25 tests added)

**Files Created:**
- `Chart.cs`, `ChartPosition.cs`, `ChartSize.cs`, `ChartType.cs`
- `ChartSeries.cs`, `ChartDataRange.cs`
- `ChartTests.cs`, `ChartPositionTests.cs`, `ChartSizeTests.cs`, `ChartSeriesTests.cs`

**Test Results:**
- Total tests: 321 (was 296, added 25)
- All passing: ✅ 321/321
- Build warnings: 0

**Notes:**
- Phase A (Foundation) complete
- All base infrastructure in place
- Chart positioning and sizing with EMU support working
- Series management and data range validation working
- Proper nullable handling (no ! operators used)

**Next session:**
- Start Phase B: Bar Chart MVP
- Begin with Step 3: Bar Chart Implementation

---

### Session 2: [Date]
**Completed:**
- [ ] Steps completed

**Notes:**
- 

**Next session:**
- 

---

## Resources

### OpenXML Documentation
- [Microsoft: Insert a chart into a spreadsheet](https://learn.microsoft.com/en-us/office/open-xml/spreadsheet/how-to-insert-a-chart-into-a-spreadsheet)
- [OpenXML SDK Documentation](https://learn.microsoft.com/en-us/office/open-xml/open-xml-sdk)

### EMU Calculations
- 1 inch = 914400 EMUs
- 1 cm = 360000 EMUs
- Default column width: ~64 pixels = ~910000 EMUs
- Default row height: ~20 pixels = ~285000 EMUs

### Chart XML Structure
- ChartSpace → Chart → PlotArea → ChartType → Series
- Drawing → TwoCellAnchor → GraphicFrame → ChartRef

---

## Estimated Timeline

**Minimum (MVP):** 6 hours (Phases A + B only)  
**Feature Complete:** 9-10 hours (Phases A + B + C)  
**Production Ready:** 12-13 hours (All phases)

Can be split across multiple sessions of 2-3 hours each.

---

**Last Updated:** Nov 16, 2025  
**Current Phase:** Phase A Complete, Ready for Phase B  
**Next Step:** Step 3 - Bar Chart Implementation
