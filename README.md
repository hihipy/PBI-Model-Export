# PBI-Model-Export

A Tabular Editor 2 script that exports a complete Power BI model snapshot to JSON — tables, columns, relationships, measures, hierarchies, partitions, roles, data sources, and a compact summary section designed for AI consumption.

## Why

Power BI models contain structure that's invisible outside of Power BI Desktop. This script extracts everything into a single portable JSON file so you can:

- Give an AI assistant (Claude, Copilot, etc.) full context on your model without describing it manually
- Document model architecture for handoffs and reviews
- Run impact analysis on DAX measure dependencies
- Audit relationships, RLS roles, and data sources
- Compare models over time by diffing exports

## Requirements

- **Tabular Editor 2** (tested on 2.27; should work on any 2.x with C# scripting)
- **Power BI Desktop** model open in Tabular Editor (or SSAS/AAS connection)

## Installation

1. Download `PBI_Model_Export_v2.csx`
2. Open your Power BI model in Tabular Editor 2
3. Go to the **Advanced Scripting** tab
4. **File → Open Script** → select the `.csx` file
5. Click **Run** (▶)
6. Find `FullModelExport.json` in your Downloads folder

**Optional — save as a Custom Action:**

- **File → Preferences → Custom Actions**
- Click **Add**, paste the script contents
- Name it "Export Full Model to JSON"
- Now available via right-click context menu on any model

## Configuration

All options are at the top of the script:

```csharp
var outputFileName = "FullModelExport.json";    // Output file name
var outputFolder = Path.Combine(                // Defaults to Downloads
    Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads");
var includeHiddenObjects = true;                // Include hidden tables/columns/measures
var includePartitionSources = true;             // Include M queries and DAX partition expressions
var includeRoles = true;                        // Include RLS role definitions
```

## What Gets Exported

| Section | Contents |
|---------|----------|
| **Model Metadata** | Compatibility level, culture, default mode (Import/DirectQuery), object counts |
| **Tables** | Name, description, visibility, data category, table type (Regular/CalculatedTable), partition source types |
| **Columns** | Name, data type, column type (Data/Calculated/CalculatedTableColumn), key status, format string, display folder, data category, summarize-by, sort-by-column |
| **Measures** | Name, full DAX expression, description, data type, format string, display folder, visibility, referenced measures (dependency list) |
| **Partitions** | Name, source type (M/Calculated/Entity), full M query or DAX expression, storage mode |
| **Hierarchies** | Name, levels with ordinal positions and source columns |
| **Relationships** | From/to table and column, cardinality, active/inactive, cross-filter direction, security filtering |
| **Roles (RLS)** | Role name, description, permission level, DAX filter expressions per table |
| **Data Sources** | Name, description, type |
| **Summary** | Table name list, human-readable relationship map, column inventory per table, measure counts per table |

## Output Format

```json
{
  "exportDate": "2026-02-25 13:57:52",
  "modelName": "My Model",
  "scriptVersion": "2.0",
  "modelMetadata": {
    "compatibilityLevel": 1600,
    "culture": "en-US",
    "defaultMode": "Import",
    "tableCount": 5,
    "relationshipCount": 3,
    "measureCount": 94,
    "roleCount": 0
  },
  "tables": [
    {
      "name": "Sales",
      "tableType": "Regular",
      "partitionSourceTypes": ["M"],
      "partitions": [
        {
          "sourceType": "M",
          "expression": "let Source = Sql.Database(...) in ...",
          "mode": "Import"
        }
      ],
      "columns": [
        {
          "name": "Amount",
          "dataType": "Double",
          "columnType": "Data",
          "summarizeBy": "Sum"
        }
      ],
      "measures": [
        {
          "name": "Total_Revenue",
          "expression": "SUM(Sales[Amount])",
          "displayFolder": "01 - Base Metrics",
          "referencedMeasures": []
        }
      ],
      "hierarchies": []
    }
  ],
  "relationships": [
    {
      "fromTable": "Sales",
      "fromColumn": "ProductID",
      "toTable": "Products",
      "toColumn": "ProductID",
      "fromCardinality": "Many",
      "toCardinality": "One",
      "crossFilteringBehavior": "OneDirection",
      "isActive": true
    }
  ],
  "roles": [],
  "dataSources": [],
  "summary": {
    "tableNames": ["Sales", "Products"],
    "relationshipMap": [
      "Sales.ProductID --> Products.ProductID (Many:One)"
    ],
    "columnInventory": {
      "Sales": ["Amount (Double)", "ProductID (Int64)"],
      "Products": ["ProductID (Int64)", "Name (String)"]
    },
    "measuresPerTable": {
      "Sales": 12
    }
  }
}
```

The `summary` section is intentionally compact — it gives an LLM the full model shape in a few lines without parsing nested objects.

## Technical Details

### JSON Escaping

The script uses a character-level escaper that handles all JSON-unsafe characters: `\"`, `\\`, `\b`, `\f`, `\n`, `\r`, `\t`, and any control character below `U+0020` (emitted as `\uXXXX`). This prevents malformed JSON from DAX expressions, M queries, or descriptions containing special characters.

### Measure Dependency Detection

Dependencies are extracted using regex bracket-token matching (`\[([^\]]+)\]`) against a pre-built `HashSet` of all measure names. This is significantly faster than the v1 approach of looping over all measures with `.Contains()` (O(n) vs O(n²)) and eliminates false positives from column names that happen to appear in bracket notation.

### Comma Safety

All array and object sections are built as string lists and joined with commas (`string.Join`), eliminating trailing comma bugs when adding or removing optional sections.

### Output Encoding

UTF-8 without BOM (`new UTF8Encoding(false)`) for maximum compatibility with JSON parsers, Python, and web tools.

### Table Type Detection

Table type is inferred from partition source types (a partition with source type `Calculated` → `CalculatedTable`). The raw `partitionSourceTypes` array is also exported so you can verify or refine the classification downstream.

## Integration Examples

### Feed to an AI Assistant

Upload the JSON file directly to Claude, ChatGPT, or any LLM and ask questions like:

- "Review my DAX measures for performance issues"
- "What happens if I rename the Department column?"
- "Suggest measures I'm missing for this type of dashboard"
- "Document this model for my team"

### Python Analysis

```python
import json
from collections import Counter

with open("FullModelExport.json", "r", encoding="utf-8") as f:
    model = json.load(f)

print(f"Tables: {model['modelMetadata']['tableCount']}")
print(f"Measures: {model['modelMetadata']['measureCount']}")
print(f"Relationships: {model['modelMetadata']['relationshipCount']}")

# Find measures with most dependencies
for table in model["tables"]:
    for m in table["measures"]:
        deps = m["referencedMeasures"]
        if len(deps) >= 5:
            print(f"  Complex: {m['name']} → {len(deps)} deps")

# Relationship map (one-liner)
for r in model["summary"]["relationshipMap"]:
    print(f"  {r}")
```

### Power Query Import

```powerquery
let
    Source = Json.Document(File.Contents("FullModelExport.json")),
    Tables = Source[tables],
    TableList = Table.FromList(Tables, Splitter.SplitByNothing()),
    Expanded = Table.ExpandRecordColumn(TableList, "Column1",
        {"name", "tableType", "isHidden", "columns", "measures"})
in
    Expanded
```

## File Locations

The script automatically saves to your Downloads folder:

| OS | Path |
|----|------|
| Windows | `C:\Users\{username}\Downloads\FullModelExport.json` |
| Mac | `/Users/{username}/Downloads/FullModelExport.json` |
| Linux | `/home/{username}/Downloads/FullModelExport.json` |

The output folder is created automatically if it doesn't exist.

## Troubleshooting

**Script compiles but output is empty:**
Check that your model has tables. Open the Tables pane in Tabular Editor to verify the model loaded.

**"Expression" is blank on partitions:**
Some partition types (e.g., `Entity` in DirectQuery) don't expose source expressions. The script catches this gracefully — the field will be an empty string.

**Hidden tables/columns missing:**
Set `includeHiddenObjects = true` in the configuration section.

**File not appearing in Downloads:**
Check the Tabular Editor output pane for the full path. The script prints the output location on success.

## Changelog

### v2.0 (Current)

- Full model export: tables, columns, relationships, partitions, hierarchies, roles, data sources
- Proper JSON escaping for all control characters
- Regex-based measure dependency detection (faster, fewer false positives)
- Comma-safe output using list-join pattern
- Table type detection with raw `partitionSourceTypes` fallback
- UTF-8 without BOM
- Output folder auto-creation
- Compact `summary` section for AI consumption

### v1.0

- Measures and calculated columns only
- Basic string-replace JSON escaping
- O(n²) dependency detection via `.Contains()`
- Manual comma tracking

## License

[CC BY-NC-SA 4.0](https://creativecommons.org/licenses/by-nc-sa/4.0/)

You are free to use, share, and adapt this work, including at your job, under these terms:

- **Attribution** — Credit the original author
- **NonCommercial** — No selling or commercial products
- **ShareAlike** — Derivatives must use the same license
