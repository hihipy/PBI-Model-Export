# pbi-model-export

A Tabular Editor 2 C# script that exports a complete Power BI model to a structured JSON file optimized for AI consumption — includes tables, columns, measures with full recursive DAX dependency chains, relationships with fact/dimension role hints, hierarchies, partitions, RLS roles, a compact summary section with top-level KPI detection, and a **data head** (first N rows sampled directly from source files on disk).

## Why

Power BI models contain structure that's invisible outside of Power BI Desktop. This script extracts everything into a single portable JSON file so you can:

- Give an AI assistant (Claude, Copilot, etc.) full context on your model — schema and actual data — without describing it manually
- Document model architecture for handoffs and reviews
- Run impact analysis on DAX measure dependencies
- Audit relationships, RLS roles, and data sources
- Compare models over time by diffing exports

## Requirements

- **Tabular Editor 2** (tested on 2.27; should work on any 2.x with C# scripting)
- **Power BI Desktop** model open in Tabular Editor (or SSAS/AAS connection)
- Source data files accessible on disk (local or OneDrive synced)

## Installation

1. Download `PBIModelExport.csx`
2. Open your Power BI model in Tabular Editor 2
3. Go to the **Advanced Scripting** tab
4. **File → Open Script** → select the `.csx` file
5. Click **Run** (▶)
6. Find the export in your Downloads folder — named automatically after your report

**Optional — save as a Custom Action:**

- **File → Preferences → Custom Actions**
- Click **Add**, paste the script contents
- Name it "Export Full Model to JSON"
- Now available via right-click context menu on any model

## Configuration

All options are at the top of the script:

```csharp
var headRowCount            = 50;    // rows to sample per table
var includeHiddenObjects    = true;  // include hidden tables/columns/measures
var includePartitionSources = true;  // include M queries and DAX partition expressions
var includeRoles            = true;  // include RLS role definitions

var outputFolder = Path.Combine(
    Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads");
```

The output file name is generated automatically from the report name — no manual naming needed. If you have `Sales Performance Dashboard.pbix` open, the file saves as `Sales Performance Dashboard_ModelExport.json`.

## What Gets Exported

| Section | Contents |
|---------|----------|
| **Model Metadata** | Compatibility level, culture, default mode (Import/DirectQuery), object counts |
| **Tables** | Name, description, visibility, auto-generated flag, resolved source folder path, data category, table type (Regular/CalculatedTable), partition source types |
| **Columns** | Name, data type, column type (Data/Calculated/CalculatedTableColumn), key status, format string, display folder, data category, summarize-by, sort-by-column |
| **Measures** | Name, full DAX expression, description, data type, format string, display folder, visibility, top-level KPI flag, direct measure dependencies, full recursive dependency chain |
| **Partitions** | Name, source type (M/Calculated/Entity), full M query or DAX expression, storage mode |
| **Hierarchies** | Name, levels with ordinal positions and source columns |
| **Relationships** | From/to table and column, cardinality, active/inactive, cross-filter direction, security filtering, fact/dimension role hints |
| **Roles (RLS)** | Role name, description, permission level, DAX filter expressions per table |
| **Data Sources** | Name, description, type |
| **Summary** | User-facing vs auto-generated table split, top-level KPI list, measures grouped by display folder, relationship map, column inventory per table, measure counts per table |
| **Data Head** | First N rows sampled from each table's source file — column names and raw cell values |

## Data Head

The `dataHead` section reads source files directly from disk — no ADOMD connection, no ACE OLEDB driver, no Office installation required.

**Supported formats:**

| Format | Method |
|--------|--------|
| `.csv` | Native delimiter split |
| `.tsv` | Native tab split |
| `.txt` | Native delimiter split |
| `.xlsx` | Open XML ZIP parsing via `ZipArchive` + `XmlDocument` |

**OneDrive path resolution** — source paths are stored in M expressions with the original author's username and OneDrive tenant name baked in. The script resolves them dynamically on whoever's machine is running it using a three-strategy fallback:

1. Exact path match (same machine, same user)
2. Rebuild path using `OneDriveCommercial` / `OneDrive` / `OneDriveConsumer` environment variables — handles username and tenant name differences
3. Folder name search under OneDrive roots and common profile paths — last resort if directory structure has changed

If the source folder can't be resolved, the `dataHead` entry returns an `error` key with a descriptive message rather than crashing the export.

**Data head output example:**

```json
"dataHead": {
  "Sales": {
    "sourceFile": "Sales_FY2025.xlsx",
    "columns": ["Region", "Salesperson", "Product", "Quarter", "Revenue", "Units", "Cost", "Margin"],
    "rows": [
      ["Northeast", "Sarah Kim", "Enterprise License", "Q1", "142500.00", "3", "87000.00", "0.389"],
      ["Southwest", "Marcus Rivera", "Professional Tier", "Q1", "48750.00", "15", "29250.00", "0.400"]
    ]
  }
}
```

> **Note on numeric precision:** XLSX stores floating point values at full IEEE 754 precision. Decimal values may appear as scientific notation strings (e.g. `3.9E-1`). Parse as `double` and round downstream as needed.

## Output Format

```json
{
  "exportDate": "2026-03-15 09:22:41",
  "dashboardName": "Sales Performance Dashboard",
  "modelName": "a3f92c10-7d44-4e81-b312-9f0c8e45aa12",
  "modelMetadata": {
    "compatibilityLevel": 1600,
    "culture": "en-US",
    "defaultMode": "Import",
    "tableCount": 4,
    "relationshipCount": 3,
    "measureCount": 31,
    "roleCount": 1
  },
  "exportConfiguration": {
    "includeHiddenObjects": true,
    "includePartitionSources": true,
    "includeRoles": true,
    "headRowCount": 50
  },
  "tables": [ ... ],
  "relationships": [ ... ],
  "roles": [ ... ],
  "dataSources": [ ... ],
  "summary": {
    "userFacingTables": ["Sales", "Products", "Regions", "Targets"],
    "autoGeneratedTables": [],
    "topLevelKPIs": ["Revenue vs Target", "YTD Revenue", "Margin %", "Win Rate"],
    "measuresByFolder": {
      "00 - KPIs": ["Total Revenue", "Total Units", "Active Accounts"],
      "01 - Targets": ["Revenue vs Target", "Attainment %"],
      "02 - Trend": ["YTD Revenue", "3yr Avg Revenue", "Revenue Growth"]
    },
    "relationshipMap": [
      "Sales.ProductID --> Products.ProductID (Many:One)",
      "Sales.RegionID --> Regions.RegionID (Many:One)"
    ],
    "columnInventory": {
      "Sales": ["OrderID (Int64)", "Revenue (Double)", "Units (Int64)", "Quarter (String)", "RegionID (Int64)"],
      "Products": ["ProductID (Int64)", "ProductName (String)", "Category (String)", "ListPrice (Double)"]
    },
    "measuresPerTable": {
      "Sales": 31
    }
  },
  "dataHead": {
    "Sales": {
      "sourceFile": "Sales_FY2025.xlsx",
      "columns": [ ... ],
      "rows": [ ... ]
    }
  }
}
```

## Integration Examples

### Feed to an AI Assistant

Upload the JSON file directly to Claude, ChatGPT, or any LLM and ask questions like:

- "Review my DAX measures for performance issues"
- "What does the actual data look like — what are the distinct values in Region?"
- "What happens if I rename the ProductID column?"
- "Suggest measures I'm missing for a sales performance dashboard"
- "Document this model for my team"
- "Which top-level KPIs depend on the Total Revenue base measure?"

### Python Analysis

```python
import json

with open("Sales_Performance_Dashboard_ModelExport.json", "r", encoding="utf-8") as f:
    model = json.load(f)

print(f"Dashboard: {model['dashboardName']}")
print(f"Tables: {model['modelMetadata']['tableCount']}")
print(f"Measures: {model['modelMetadata']['measureCount']}")

# User-facing tables only
print("\nUser-facing tables:")
for t in model["summary"]["userFacingTables"]:
    print(f"  {t}")

# Top-level KPIs
print("\nTop-level KPIs:")
for kpi in model["summary"]["topLevelKPIs"]:
    print(f"  {kpi}")

# Data head — print first row of each table
print("\nData head sample:")
for table_name, head in model["dataHead"].items():
    if "rows" in head and head["rows"]:
        cols = head["columns"]
        row  = head["rows"][0]
        print(f"\n  {table_name}")
        for col, val in zip(cols, row):
            print(f"    {col}: {val}")

# Measures with deep dependency chains
print("\nComplex measures:")
for table in model["tables"]:
    for m in table["measures"]:
        chain = m["fullDependencyChain"]
        if len(chain) >= 3:
            print(f"  {m['name']} -> {len(chain)} upstream measures")

# Fact vs dimension tables from relationships
print("\nRelationship roles:")
for r in model["relationships"]:
    print(f"  Fact: {r['factTable']} | Dimension: {r['dimensionTable']}")
```

### Power Query Import

```powerquery
let
    Source = Json.Document(File.Contents("Sales_Performance_Dashboard_ModelExport.json")),
    Tables = Source[tables],
    TableList = Table.FromList(Tables, Splitter.SplitByNothing()),
    Expanded = Table.ExpandRecordColumn(TableList, "Column1",
        {"name", "tableType", "isHidden", "isAutoGenerated", "sourceFolder", "columns", "measures"})
in
    Expanded
```

## File Locations

The script automatically saves to your Downloads folder, named after the open report:

| OS | Path |
|----|------|
| Windows | `C:\Users\{username}\Downloads\{ReportName}_ModelExport.json` |
| Mac | `/Users/{username}/Downloads/{ReportName}_ModelExport.json` |
| Linux | `/home/{username}/Downloads/{ReportName}_ModelExport.json` |

The output folder is created automatically if it doesn't exist.

## Troubleshooting

**Script compiles but output is empty:**
Check that your model has tables. Open the Tables pane in Tabular Editor to verify the model loaded.

**Dashboard name shows as "UnknownReport":**
The script reads the Power BI Desktop window title. Make sure the report is open and fully loaded in Power BI Desktop before running the script in Tabular Editor.

**`dataHead` shows `"error": "Folder not found"`:**
The source folder path couldn't be resolved on the current machine. Make sure OneDrive is synced and the source data folder is available locally (right-click → "Always keep on this device"). The script will retry automatically using OneDrive environment variables and folder name search before surfacing this error.

**`dataHead` shows `"note": "Unsupported format"`:**
The source file is not CSV, TSV, TXT, or XLSX. The script returns a file list instead of row data for unsupported binary formats.

**XLSX numeric values in scientific notation:**
This is expected — XLSX stores full IEEE 754 precision in the XML. Parse the string as `double` and apply rounding downstream.

**"Expression" is blank on partitions:**
Some partition types (e.g. `Entity` in DirectQuery) don't expose source expressions. The script catches this gracefully — the field will be an empty string.

**`sourceFolder` is empty:**
Only tables loaded via `Folder.Files()` or `File.Contents()` in their M expression will have this populated. SQL, SharePoint, and other source types won't produce a folder path.

**Hidden tables/columns missing:**
Set `includeHiddenObjects = true` in the configuration section.

**File not appearing in Downloads:**
Check the Tabular Editor output pane for the full path. The script prints the output location on success.

## Technical Details

### Data Head — XLSX Parsing

XLSX files are ZIP archives containing XML. The script reads them directly using `ZipArchive` (part of `System.IO.Compression`) and `XmlDocument` (part of `System.Xml`) — both available natively in TE2's .NET runtime with no external references.

Two internal files are read:

- `xl/sharedStrings.xml` — Excel de-duplicates all string values into a lookup table; cell values with `type="s"` are resolved by index against this table
- `xl/worksheets/sheet1.xml` — cell data for the first worksheet; row 0 is treated as the header row

### OneDrive Path Resolution

M expressions store the absolute path of the source folder as it existed on the original author's machine. The resolver uses three strategies in order:

1. **Exact match** — try the path verbatim
2. **Environment variable rebuild** — strip the stored username and OneDrive tenant segment, then reattach the real local OneDrive root from `OneDriveCommercial`, `OneDrive`, or `OneDriveConsumer` env vars
3. **Folder name search** — recursively search for the target folder name under all OneDrive roots and common profile paths

### Measure Dependency Detection

Direct dependencies are extracted using regex bracket-token matching (`\[([^\]]+)\]`) against a pre-built `HashSet` of all measure names. Full recursive dependency chains are resolved by walking the dependency graph with circular reference protection.

### Top-Level KPI Detection

A reverse lookup is built across all measure expressions before export. Any measure whose name never appears as a dependency in another measure's expression is flagged `isTopLevelKPI: true`. These are your model's user-facing entry points.

### Auto-Generated Table Detection

Tables whose names contain a GUID pattern are flagged `isAutoGenerated: true` and surfaced separately in the summary so AI can deprioritize them.

### Relationship Role Hints

Each relationship includes `factTable` and `dimensionTable` fields derived from cardinality: the `Many` side is the fact table, the `One` side is the dimension table.

### Output Encoding

UTF-8 without BOM (`new UTF8Encoding(false)`) for maximum compatibility with JSON parsers, Python, and web tools.

## Changelog

### Current

- **Data head** — first N rows sampled directly from source files, no ADOMD or ACE OLEDB required
- **XLSX support** — Open XML parsed via `ZipArchive` + `XmlDocument`, works without Office installed
- **Dynamic OneDrive path resolution** — three-strategy resolver handles username, tenant name, and directory structure differences across machines
- `sourceFolder` field on each table (replaces `sourceFilePath` — now resolves to the folder, not individual file)
- `headRowCount` added to `exportConfiguration` section
- Assembly reference `#r "System.IO.Compression.dll"` for TE2 compatibility
- Removed version numbering from script output

### v3.0

- Auto-detects report name from Power BI Desktop window title
- `isAutoGenerated` flag on tables containing GUIDs
- `sourceFilePath` field extracted from M partition expressions
- `isTopLevelKPI` flag per measure
- `fullDependencyChain` per measure with circular reference protection
- `factTable` / `dimensionTable` role hints on relationships
- Summary expanded with `userFacingTables`, `autoGeneratedTables`, `topLevelKPIs`, `measuresByFolder`

### v2.0

- Full model export: tables, columns, relationships, partitions, hierarchies, roles, data sources
- Proper JSON escaping for all control characters
- Regex-based measure dependency detection
- Comma-safe output using list-join pattern
- UTF-8 without BOM

### v1.0

- Measures and calculated columns only
- Basic string-replace JSON escaping
- Manual comma tracking

## License

[CC BY-NC-SA 4.0](https://creativecommons.org/licenses/by-nc-sa/4.0/)

You are free to use, share, and adapt this work, including at your job, under these terms:

- **Attribution** — Credit the original author
- **NonCommercial** — No selling or commercial products
- **ShareAlike** — Derivatives must use the same license
