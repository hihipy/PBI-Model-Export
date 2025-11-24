using System.IO;
using System.Text;

// Configuration - Easy to modify
var outputFileName = "DAX_Export.json";
var outputFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads");
var includeMeasures = true;
var includeCalculatedColumns = true;
var includeHiddenObjects = true;
var complexityMedium = 1;
var complexityComplex = 3;
var complexityVeryComplex = 6;

try
{
    var sb = new StringBuilder();

    sb.AppendLine("{");
    sb.AppendLine("  \"exportDate\": \"" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "\",");
    sb.AppendLine("  \"modelName\": \"" + Model.Database.Name.Replace("\"", "'").Replace("\\", "/") + "\",");
    sb.AppendLine("  \"exportConfiguration\": {");
    sb.AppendLine("    \"includedMeasures\": " + includeMeasures.ToString().ToLower() + ",");
    sb.AppendLine("    \"includedCalculatedColumns\": " + includeCalculatedColumns.ToString().ToLower() + ",");
    sb.AppendLine("    \"includedHiddenObjects\": " + includeHiddenObjects.ToString().ToLower());
    sb.AppendLine("  },");

    var measureCount = 0;
    var columnCount = 0;

    // Export Measures
    if (includeMeasures)
    {
        sb.AppendLine("  \"measures\": {");

        var measuresList = new List<Measure>();
        foreach (var measure in Model.AllMeasures)
        {
            if (includeHiddenObjects || !measure.IsHidden)
            {
                measuresList.Add(measure);
            }
        }
        measureCount = measuresList.Count;

        sb.AppendLine("    \"count\": " + measureCount + ",");
        sb.AppendLine("    \"items\": [");

        for (int i = 0; i < measuresList.Count; i++)
        {
            var measure = measuresList[i];

            // Get dependencies - use simple string matching since DependsOn might not work as expected
            var dependencies = new List<string>();
            var depCount = 0;
            foreach (var otherMeasure in Model.AllMeasures)
            {
                if (measure != otherMeasure && measure.Expression.Contains("[" + otherMeasure.Name + "]"))
                {
                    dependencies.Add(otherMeasure.Name);
                    depCount++;
                }
            }

            var complexity = "Simple";
            if (depCount >= complexityVeryComplex)
                complexity = "Very Complex";
            else if (depCount >= complexityComplex)
                complexity = "Complex";
            else if (depCount >= complexityMedium)
                complexity = "Medium";

            // Parse folder structure
            var folderPath = measure.DisplayFolder ?? "";
            var folderParts = folderPath.Split(new char[] { '\\' }, StringSplitOptions.RemoveEmptyEntries);
            var mainFolder = folderParts.Length > 0 ? folderParts[0] : "";
            var subFolder = folderParts.Length > 1 ? folderParts[1] : "";

            // Clean strings
            var cleanName = measure.Name.Replace("\"", "'").Replace("\\", "/").Replace("\n", " ").Replace("\r", "");
            var cleanTable = measure.Table.Name.Replace("\"", "'").Replace("\\", "/");
            var cleanExpr = measure.Expression.Replace("\"", "'").Replace("\\", "/").Replace("\n", " ").Replace("\r", "");
            var cleanDesc = (measure.Description ?? "").Replace("\"", "'").Replace("\\", "/").Replace("\n", " ").Replace("\r", "");
            var cleanFormat = (measure.FormatString ?? "").Replace("\"", "'").Replace("\\", "/");
            var cleanMainFolder = mainFolder.Replace("\"", "'").Replace("\\", "/");
            var cleanSubFolder = subFolder.Replace("\"", "'").Replace("\\", "/");
            var cleanFullPath = folderPath.Replace("\"", "'").Replace("\\", "/");
            var cleanDataCat = (measure.DataCategory ?? "").Replace("\"", "'").Replace("\\", "/");
            var cleanDispFolder = (measure.DisplayFolder ?? "").Replace("\"", "'").Replace("\\", "/");

            sb.AppendLine("      {");
            sb.AppendLine("        \"name\": \"" + cleanName + "\",");
            sb.AppendLine("        \"table\": \"" + cleanTable + "\",");
            sb.AppendLine("        \"expression\": \"" + cleanExpr + "\",");
            sb.AppendLine("        \"description\": \"" + cleanDesc + "\",");
            sb.AppendLine("        \"dataType\": \"" + measure.DataType.ToString() + "\",");
            sb.AppendLine("        \"formatString\": \"" + cleanFormat + "\",");
            sb.AppendLine("        \"folder\": {");
            sb.AppendLine("          \"main\": \"" + cleanMainFolder + "\",");
            sb.AppendLine("          \"sub\": \"" + cleanSubFolder + "\",");
            sb.AppendLine("          \"full\": \"" + cleanFullPath + "\",");
            sb.AppendLine("          \"depth\": " + folderParts.Length);
            sb.AppendLine("        },");
            sb.AppendLine("        \"complexity\": {");
            sb.AppendLine("          \"level\": \"" + complexity + "\",");
            sb.AppendLine("          \"directDependencies\": " + depCount + ",");
            sb.Append("          \"referencedMeasures\": [");
            for (int j = 0; j < dependencies.Count; j++)
            {
                var cleanDep = dependencies[j].Replace("\"", "'").Replace("\\", "/");
                sb.Append("\"" + cleanDep + "\"");
                if (j < dependencies.Count - 1) sb.Append(", ");
            }
            sb.AppendLine("]");
            sb.AppendLine("        },");
            sb.AppendLine("        \"metadata\": {");
            sb.AppendLine("          \"isHidden\": " + measure.IsHidden.ToString().ToLower() + ",");
            sb.AppendLine("          \"displayFolder\": \"" + cleanDispFolder + "\",");
            sb.AppendLine("          \"dataCategory\": \"" + cleanDataCat + "\"");
            sb.AppendLine("        }");
            sb.Append("      }");
            if (i < measuresList.Count - 1) sb.AppendLine(",");
            else sb.AppendLine();
        }

        sb.Append("    ]");
        sb.AppendLine();
        sb.Append("  }");
        if (includeCalculatedColumns) sb.AppendLine(",");
        else sb.AppendLine();
    }

    // Export Calculated Columns
    if (includeCalculatedColumns)
    {
        sb.AppendLine("  \"calculatedColumns\": {");

        var columnsList = new List<Column>();
        foreach (var table in Model.Tables)
        {
            foreach (var column in table.Columns)
            {
                if (column.Type == ColumnType.Calculated)
                {
                    if (includeHiddenObjects || !column.IsHidden)
                    {
                        columnsList.Add(column);
                    }
                }
            }
        }
        columnCount = columnsList.Count;

        sb.AppendLine("    \"count\": " + columnCount + ",");
        sb.AppendLine("    \"items\": [");

        for (int i = 0; i < columnsList.Count; i++)
        {
            var column = columnsList[i];

            // Parse folder structure
            var folderPath = column.DisplayFolder ?? "";
            var folderParts = folderPath.Split(new char[] { '\\' }, StringSplitOptions.RemoveEmptyEntries);
            var mainFolder = folderParts.Length > 0 ? folderParts[0] : "";
            var subFolder = folderParts.Length > 1 ? folderParts[1] : "";

            var sortByCol = column.SortByColumn != null ? column.SortByColumn.Name : "";

            // Clean strings
            var cleanName = column.Name.Replace("\"", "'").Replace("\\", "/").Replace("\n", " ").Replace("\r", "");
            var cleanTable = column.Table.Name.Replace("\"", "'").Replace("\\", "/");
            var cleanDesc = (column.Description ?? "").Replace("\"", "'").Replace("\\", "/").Replace("\n", " ").Replace("\r", "");
            var cleanFormat = (column.FormatString ?? "").Replace("\"", "'").Replace("\\", "/");
            var cleanMainFolder = mainFolder.Replace("\"", "'").Replace("\\", "/");
            var cleanSubFolder = subFolder.Replace("\"", "'").Replace("\\", "/");
            var cleanFullPath = folderPath.Replace("\"", "'").Replace("\\", "/");
            var cleanDataCat = (column.DataCategory ?? "").Replace("\"", "'").Replace("\\", "/");
            var cleanDispFolder = (column.DisplayFolder ?? "").Replace("\"", "'").Replace("\\", "/");
            var cleanSortBy = sortByCol.Replace("\"", "'").Replace("\\", "/");

            sb.AppendLine("      {");
            sb.AppendLine("        \"name\": \"" + cleanName + "\",");
            sb.AppendLine("        \"table\": \"" + cleanTable + "\",");
            sb.AppendLine("        \"description\": \"" + cleanDesc + "\",");
            sb.AppendLine("        \"dataType\": \"" + column.DataType.ToString() + "\",");
            sb.AppendLine("        \"formatString\": \"" + cleanFormat + "\",");
            sb.AppendLine("        \"folder\": {");
            sb.AppendLine("          \"main\": \"" + cleanMainFolder + "\",");
            sb.AppendLine("          \"sub\": \"" + cleanSubFolder + "\",");
            sb.AppendLine("          \"full\": \"" + cleanFullPath + "\",");
            sb.AppendLine("          \"depth\": " + folderParts.Length);
            sb.AppendLine("        },");
            sb.AppendLine("        \"metadata\": {");
            sb.AppendLine("          \"isHidden\": " + column.IsHidden.ToString().ToLower() + ",");
            sb.AppendLine("          \"displayFolder\": \"" + cleanDispFolder + "\",");
            sb.AppendLine("          \"dataCategory\": \"" + cleanDataCat + "\",");
            sb.AppendLine("          \"sortByColumn\": \"" + cleanSortBy + "\"");
            sb.AppendLine("        }");
            sb.Append("      }");
            if (i < columnsList.Count - 1) sb.AppendLine(",");
            else sb.AppendLine();
        }

        sb.AppendLine("    ]");
        sb.AppendLine("  }");
    }

    sb.AppendLine("}");

    // Write to file
    var outputPath = Path.Combine(outputFolder, outputFileName);
    File.WriteAllText(outputPath, sb.ToString(), Encoding.UTF8);

    // Success message
    Console.WriteLine("Export completed successfully!");
    Console.WriteLine("  Model: " + Model.Database.Name);
    Console.WriteLine("  Measures: " + measureCount.ToString());
    Console.WriteLine("  Calculated Columns: " + columnCount.ToString());
    Console.WriteLine("  Location: " + outputPath);
}
catch (Exception ex)
{
    Console.WriteLine("Export failed!");
    Console.WriteLine("  Error: " + ex.Message);
    Console.WriteLine("  Type: " + ex.GetType().Name);
}