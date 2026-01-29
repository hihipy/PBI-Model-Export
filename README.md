# Export-DAX-Lineage

A Tabular Editor macro that exports all DAX measures and calculated columns from Power BI models to JSON with complete dependency mapping, complexity analysis, and folder organization tracking. Useful for impact analysis, documentation automation, and understanding DAX architecture.

## Features

- **Complete DAX Inventory**: Exports all measures and calculated columns with expressions, descriptions, and metadata
- **Folder Structure Mapping**: Captures complete folder hierarchy with depth tracking for both measures and columns
- **Dependency Mapping**: Uses built-in dependency tracking to show which measures reference other measures
- **Complexity Analysis**: Automatically categorizes measures from Simple to Very Complex based on dependencies
- **Rich Metadata**: Captures data types, format strings, data categories, and sort columns
- **Clean JSON Output**: Properly formatted, parseable JSON with nested structure
- **Universal Compatibility**: Cross-platform file paths work on Windows, Mac, and Linux
- **Configurable Export**: Control what gets exported (measures, columns, hidden objects)

## Requirements

- **Tabular Editor** (version with C# scripting support)
- **Power BI Model** (.pbix, .pbit, or Analysis Services connection)

## Installation

1. **Download the script**: Save the code as `Export-DAX-Lineage.csx` in your preferred location

2. Option 1 - Load as Script:

   - Open Tabular Editor with your Power BI model
   - Go to **Advanced Scripting** tab
   - Click **File > Open Script**
   - Select `Export-DAX-Lineage.csx`
   - Click **Run**

3. Option 2 - Create Custom Action (Recommended):

   - In Tabular Editor: **File > Preferences > Custom Actions**
   - Click **Add** and paste the code
   - Name: "Export DAX Lineage to JSON"
   - Now available via right-click context menu

## Usage

### Quick Start

1. **Open your Power BI model** in Tabular Editor
2. **Run the macro** using either method above (no path configuration needed)
3. **Find your JSON file** in your Downloads folder as `DAX_Export.json`

### Customization

Configure the export at the top of the script:

```csharp
// Configuration - Easy to modify
var outputFileName = "DAX_Export.json";
var outputFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads");
var includeMeasures = true;              // Set to false to skip measures
var includeCalculatedColumns = true;     // Set to false to skip calculated columns
var includeHiddenObjects = true;         // Set to false to exclude hidden objects
var complexityMedium = 1;                // 1+ dependencies = Medium
var complexityComplex = 3;               // 3+ dependencies = Complex
var complexityVeryComplex = 6;           // 6+ dependencies = Very Complex
```

## Output Format

The macro generates a structured JSON file with nested objects for better organization:

```json
{
  "exportDate": "2025-11-24 11:10:32",
  "modelName": "Sales Analytics Model",
  "exportConfiguration": {
    "includedMeasures": true,
    "includedCalculatedColumns": true,
    "includedHiddenObjects": true
  },
  "measures": {
    "count": 56,
    "items": [
      {
        "name": "Total_Revenue",
        "table": "Sales",
        "expression": "SUM(Sales[Amount])",
        "description": "Total sales revenue",
        "dataType": "Currency",
        "formatString": "$#,0.00",
        "folder": {
          "main": "01 - Base Metrics",
          "sub": "Revenue",
          "full": "01 - Base Metrics/Revenue",
          "depth": 2
        },
        "complexity": {
          "level": "Simple",
          "directDependencies": 0,
          "referencedMeasures": []
        },
        "metadata": {
          "isHidden": false,
          "displayFolder": "01 - Base Metrics/Revenue",
          "dataCategory": ""
        }
      },
      {
        "name": "Revenue_Growth",
        "table": "Sales",
        "expression": "DIVIDE([Total_Revenue] - [Previous_Revenue], [Previous_Revenue])",
        "description": "YoY revenue growth percentage",
        "dataType": "Double",
        "formatString": "0.00%",
        "folder": {
          "main": "02 - KPIs",
          "sub": "Growth Metrics",
          "full": "02 - KPIs/Growth Metrics",
          "depth": 2
        },
        "complexity": {
          "level": "Medium",
          "directDependencies": 2,
          "referencedMeasures": ["Total_Revenue", "Previous_Revenue"]
        },
        "metadata": {
          "isHidden": false,
          "displayFolder": "02 - KPIs/Growth Metrics",
          "dataCategory": ""
        }
      }
    ]
  },
  "calculatedColumns": {
    "count": 13,
    "items": [
      {
        "name": "YoY_Growth_Rate",
        "table": "Sales",
        "description": "Year-over-year growth calculation",
        "dataType": "Double",
        "formatString": "0.00%",
        "folder": {
          "main": "Growth Metrics",
          "sub": "",
          "full": "Growth Metrics",
          "depth": 1
        },
        "metadata": {
          "isHidden": false,
          "displayFolder": "Growth Metrics",
          "dataCategory": "",
          "sortByColumn": ""
        }
      },
      {
        "name": "Month",
        "table": "DateTable",
        "description": "",
        "dataType": "String",
        "formatString": "",
        "folder": {
          "main": "",
          "sub": "",
          "full": "",
          "depth": 0
        },
        "metadata": {
          "isHidden": true,
          "displayFolder": "",
          "dataCategory": "Months",
          "sortByColumn": "MonthNo"
        }
      }
    ]
  }
}
```

## Key Features

### Metadata Capture

**For Measures:**
- DAX expressions with dependencies
- Data types (Currency, Double, Int64, DateTime, etc.)
- Format strings (number and date formatting)
- Folder organization with depth tracking
- Complexity classification with dependency counts
- Referenced measures list
- Visibility and data category information

**For Calculated Columns:**
- Data types and format strings
- Folder organization with depth tracking
- Sort column relationships
- Data categories (Months, Years, QuarterOfYear, etc.)
- Visibility settings
- Table context

### JSON Structure

- **Nested Objects**: Organized into logical sections (folder, complexity, metadata)
- **Count Fields**: Quick access to totals without parsing arrays
- **Export Configuration**: Documents what was included in the export
- **Model Identification**: Captures model name for multi-model environments
- **Depth Tracking**: Shows folder nesting level for organizational analysis

### Folder Structure Features

The macro captures complete folder hierarchy information:

- **`main`**: Top-level folder name (e.g., "01 - Base Metrics")
- **`sub`**: First subfolder level (e.g., "Revenue")
- **`full`**: Complete folder path (e.g., "01 - Base Metrics/Revenue")
- **`depth`**: Number of folder levels (0 = root, 1 = one level, etc.)

### Dependency Analysis

- **Accurate Tracking**: Uses Tabular Editor's built-in dependency system
- **Foundation Measures**: Identifies measures with zero dependencies
- **Derived Measures**: Shows complete dependency chain
- **Impact Analysis**: Understand what breaks if you change a measure
- **Complexity Scoring**: Automatic classification based on dependency count

### Calculated Column Support

- **Complete Inventory**: All calculated columns including auto-generated ones
- **Data Type Distribution**: Understand column types across tables
- **Format Consistency**: Document formatting patterns
- **Sorting Relationships**: Capture sort column dependencies
- **Data Categories**: Semantic information for date/time columns

## Analysis Features

### Measure Classification

- **Simple**: No measure dependencies (foundation/base measures)
- **Medium**: 1-2 measure dependencies
- **Complex**: 3-5 measure dependencies
- **Very Complex**: 6+ measure dependencies

### Organizational Analysis

- **Folder Distribution**: See how DAX objects are organized
- **Depth Analysis**: Identify deeply nested folder structures
- **Cross-Folder Dependencies**: Track relationships across organizational boundaries
- **Naming Conventions**: Analyze folder naming patterns

### Data Quality Insights

- **Hidden Object Tracking**: Identify hidden measures and columns
- **Format String Consistency**: Review formatting standards
- **Data Category Usage**: Understand semantic metadata application
- **Sort Column Relationships**: Verify proper column sorting

## Use Cases

- **Model Documentation**: Generate technical documentation for all DAX objects
- **Impact Analysis**: Understand what happens if you modify a measure
- **Architecture Planning**: Understand DAX object relationships and organization
- **Knowledge Transfer**: Onboard new team members with complete DAX inventory
- **Quality Assurance**: Find measures with excessive complexity
- **Organizational Review**: Analyze folder structure and categorization
- **Model Optimization**: Identify redundant or inefficient calculated columns
- **Data Governance**: Track data types, formats, and semantic categories
- **Dependency Mapping**: Visualize measure relationships and complexity

## Integration Examples

### Python Analysis

```python
import json
from collections import Counter

with open('DAX_Export.json', 'r', encoding='utf-8') as f:
    dax_data = json.load(f)

# Access measures and columns
measures = dax_data['measures']['items']
calc_columns = dax_data['calculatedColumns']['items']

print(f"Model: {dax_data['modelName']}")
print(f"Export Date: {dax_data['exportDate']}")
print(f"Total Measures: {dax_data['measures']['count']}")
print(f"Total Calculated Columns: {dax_data['calculatedColumns']['count']}")

# Analyze measure complexity
complexity_dist = Counter([m['complexity']['level'] for m in measures])
print(f"\nComplexity Distribution: {dict(complexity_dist)}")

# Find most complex measures
very_complex = [m for m in measures if m['complexity']['level'] == 'Very Complex']
print(f"\nVery Complex Measures ({len(very_complex)}):")
for m in very_complex:
    deps = m['complexity']['directDependencies']
    print(f"  - {m['name']}: {deps} dependencies")

# Analyze calculated columns
custom_columns = [col for col in calc_columns 
                 if not col['table'].startswith('DateTableTemplate') 
                 and not col['table'].startswith('LocalDateTable')]

data_types = Counter([col['dataType'] for col in custom_columns])
print(f"\nCustom Column Data Types: {dict(data_types)}")

# Folder depth analysis
folder_depths = Counter([m['folder']['depth'] for m in measures])
print(f"\nFolder Depth Distribution: {dict(folder_depths)}")

# Find measures with specific folder patterns
base_measures = [m for m in measures if 'Base' in m['folder']['main']]
print(f"\nBase Measures: {len(base_measures)}")

# Analyze dependencies
no_deps = [m for m in measures if m['complexity']['directDependencies'] == 0]
print(f"\nFoundation Measures (no dependencies): {len(no_deps)}")
```

### Power BI Import

```powerquery
// Import measures with full metadata
let
    Source = Json.Document(File.Contents("DAX_Export.json")),
    measures = Source[measures][items],
    MeasuresTable = Table.FromList(measures, Splitter.SplitByNothing()),
    ExpandedMeasures = Table.ExpandRecordColumn(MeasuresTable, "Column1", 
        {"name", "table", "expression", "description", "dataType", "formatString"}),
    ExpandedFolder = Table.ExpandRecordColumn(ExpandedMeasures, "folder",
        {"main", "sub", "full", "depth"},
        {"folder_main", "folder_sub", "folder_full", "folder_depth"}),
    ExpandedComplexity = Table.ExpandRecordColumn(ExpandedFolder, "complexity",
        {"level", "directDependencies"},
        {"complexity_level", "directDependencies"})
in
    ExpandedComplexity

// Import calculated columns
let
    Source = Json.Document(File.Contents("DAX_Export.json")),
    columns = Source[calculatedColumns][items],
    ColumnsTable = Table.FromList(columns, Splitter.SplitByNothing()),
    ExpandedColumns = Table.ExpandRecordColumn(ColumnsTable, "Column1", 
        {"name", "table", "dataType", "formatString"}),
    ExpandedFolder = Table.ExpandRecordColumn(ExpandedColumns, "folder",
        {"main", "full", "depth"},
        {"folder_main", "folder_full", "folder_depth"}),
    ExpandedMetadata = Table.ExpandRecordColumn(ExpandedFolder, "metadata",
        {"isHidden", "dataCategory", "sortByColumn"},
        {"isHidden", "dataCategory", "sortByColumn"})
in
    ExpandedMetadata
```

### Network Visualization (Python + NetworkX)

```python
import json
import networkx as nx
import matplotlib.pyplot as plt

with open('DAX_Export.json', 'r', encoding='utf-8') as f:
    dax_data = json.load(f)

# Create dependency graph
G = nx.DiGraph()
measures = dax_data['measures']['items']

for measure in measures:
    measure_name = measure['name']
    G.add_node(measure_name, complexity=measure['complexity']['level'])

    for dep in measure['complexity']['referencedMeasures']:
        G.add_edge(dep, measure_name)

# Visualize
pos = nx.spring_layout(G)
complexity_colors = {
    'Simple': 'lightgreen',
    'Medium': 'yellow',
    'Complex': 'orange',
    'Very Complex': 'red'
}

node_colors = [complexity_colors[G.nodes[node].get('complexity', 'Simple')]
               for node in G.nodes()]

plt.figure(figsize=(15, 10))
nx.draw(G, pos, node_color=node_colors, with_labels=True,
        node_size=1000, font_size=8, arrows=True)
plt.title("DAX Measure Dependency Network")
plt.show()
```

## Technical Details

- **Language**: C# script for Tabular Editor
- **Dependencies**: None beyond standard Tabular Editor libraries
- **Output**: UTF-8 encoded JSON file with proper escaping
- **Performance**: Processes hundreds of DAX objects in seconds
- **Memory**: Minimal memory footprint with StringBuilder optimization
- **Compatibility**: Works with all Power BI model versions
- **Cross-Platform**: Universal file paths work on Windows, Mac, and Linux
- **Folder Support**: Handles nested folder structures of any depth
- **String Safety**: Proper JSON escaping for special characters

## Troubleshooting

### Common Issues

**"Cannot convert type" errors**
- Fixed in latest version. Uses string matching for dependencies instead of type casting.

**File path issues**
- Script automatically uses cross-platform compatible Downloads folder.
- Verify write permissions to Downloads folder.

**Empty output**
- Check that your model contains DAX measures or calculated columns.
- Verify configuration settings (`includeMeasures`, `includeCalculatedColumns`).

**Missing folder information**
- Ensure objects are properly organized in folders within Power BI.
- Empty folder fields indicate objects at root level.

**Hidden objects not showing**
- Set `includeHiddenObjects = true` in configuration.
- By default, all objects including hidden ones are exported.

### File Locations

The script automatically saves to your Downloads folder:

- **Windows**: `C:\Users\[username]\Downloads\DAX_Export.json`
- **Mac**: `/Users/[username]/Downloads/DAX_Export.json`
- **Linux**: `/home/[username]/Downloads/DAX_Export.json`

### What's Exported

**Measures:**
- Name, table, DAX expression
- Data type and format string
- Complete folder hierarchy with depth
- Complexity classification
- Direct dependency count
- List of referenced measures
- Visibility and data category

**Calculated Columns:**
- Name, table, description
- Data type and format string
- Complete folder hierarchy with depth
- Sort column relationships
- Visibility and data category
- All columns including auto-generated date tables

### JSON Structure Notes

- Nested objects provide logical grouping of related metadata
- Forward slashes used in folder paths for cross-platform compatibility
- Special characters properly escaped in JSON output
- Empty strings indicate missing or not applicable values
- Count fields provide quick access to totals

## Changelog

### Version 2.0 (Current)

**Major Improvements:**
- Restructured JSON with nested objects for better organization
- Added count fields for measures and calculated columns
- Implemented export configuration tracking
- Enhanced folder structure with depth tracking
- Added data types and format strings for measures
- Improved metadata capture (data categories, sort columns)
- Better dependency tracking using built-in DependsOn property
- Cross-platform path compatibility with forward slashes
- Proper JSON string escaping
- Error handling with try-catch blocks

**Breaking Changes:**
- JSON structure changed from flat to nested format
- Folder fields now in nested `folder` object
- Complexity fields now in nested `complexity` object
- Metadata now in nested `metadata` object
- Output filename changed to `DAX_Export.json`

### Version 1.0 (Original)

- Initial release with basic measure and calculated column export
- Simple dependency tracking via string matching
- Flat JSON structure
- Basic folder organization tracking

## License

Export-DAX-Lineage © 2025

Distributed under the [Creative Commons Attribution-NonCommercial-NoDerivatives 4.0 International](https://creativecommons.org/licenses/by-nc-nd/4.0/) license.
