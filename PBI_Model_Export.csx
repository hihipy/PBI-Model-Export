// ============================================================================
// Full Power BI Model Export for Tabular Editor 2
// v2 - Refactored with proper JSON escaping, optimized dependency detection,
//      robust comma handling, and improved table type detection
// ============================================================================

using System;
using System.IO;
using System.Text;
using System.Linq;
using System.Collections.Generic;
using System.Text.RegularExpressions;

// --- Configuration -----------------------------------------------------------
var outputFileName = "FullModelExport.json";
var outputFolder = Path.Combine(
    Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads");
var includeHiddenObjects = true;
var includePartitionSources = true;
var includeRoles = true;
// -----------------------------------------------------------------------------

// --- JSON-safe string escaper (handles all control chars < 0x20) -------------
Func<string, string> Esc = (string s) =>
{
    if (s == null) return "";
    var esb = new StringBuilder(s.Length);
    foreach (var ch in s)
    {
        switch (ch)
        {
            case '\\': esb.Append("\\\\"); break;
            case '\"': esb.Append("\\\""); break;
            case '\b': esb.Append("\\b"); break;
            case '\f': esb.Append("\\f"); break;
            case '\n': esb.Append("\\n"); break;
            case '\r': esb.Append("\\r"); break;
            case '\t': esb.Append("\\t"); break;
            default:
                if (ch < (char)0x20)
                {
                    esb.Append("\\u");
                    esb.Append(((int)ch).ToString("x4"));
                }
                else esb.Append(ch);
                break;
        }
    }
    return esb.ToString();
};

// --- Bracket token regex for dependency detection ----------------------------
var bracketTokenRegex = new Regex(@"\[([^\]]+)\]", RegexOptions.Compiled);
var allMeasureNames = new HashSet<string>();
foreach (var m in Model.AllMeasures) allMeasureNames.Add(m.Name);

Func<string, string, List<string>> GetDeps = (string expr, string selfName) =>
{
    var deps = new HashSet<string>();
    if (string.IsNullOrEmpty(expr)) return deps.ToList();
    foreach (Match match in bracketTokenRegex.Matches(expr))
    {
        var token = match.Groups[1].Value;
        if (token != selfName && allMeasureNames.Contains(token))
            deps.Add(token);
    }
    var result = deps.ToList();
    result.Sort();
    return result;
};

// --- Helper: join a list of JSON fragments with commas -----------------------
Func<List<string>, string, string> JoinItems = (List<string> items, string indent) =>
{
    if (items.Count == 0) return "";
    return string.Join(",\n", items);
};

try
{
    var sb = new StringBuilder();

    // =========================================================================
    //  ROOT OPEN + MODEL METADATA
    // =========================================================================
    sb.AppendLine("{");
    sb.AppendLine("  \"exportDate\": \"" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "\",");
    sb.AppendLine("  \"modelName\": \"" + Esc(Model.Database.Name) + "\",");
    sb.AppendLine("  \"scriptVersion\": \"2.0\",");

    sb.AppendLine("  \"modelMetadata\": {");
    sb.AppendLine("    \"compatibilityLevel\": " + Model.Database.CompatibilityLevel + ",");
    sb.AppendLine("    \"culture\": \"" + Esc(Model.Culture) + "\",");
    sb.AppendLine("    \"defaultMode\": \"" + Model.DefaultMode.ToString() + "\",");
    sb.AppendLine("    \"description\": \"" + Esc(Model.Database.Description ?? "") + "\",");
    sb.AppendLine("    \"tableCount\": " + Model.Tables.Count + ",");
    sb.AppendLine("    \"relationshipCount\": " + Model.Relationships.Count + ",");
    sb.AppendLine("    \"measureCount\": " + Model.AllMeasures.Count() + ",");
    sb.AppendLine("    \"roleCount\": " + Model.Roles.Count);
    sb.AppendLine("  },");

    sb.AppendLine("  \"exportConfiguration\": {");
    sb.AppendLine("    \"includeHiddenObjects\": " + includeHiddenObjects.ToString().ToLower() + ",");
    sb.AppendLine("    \"includePartitionSources\": " + includePartitionSources.ToString().ToLower() + ",");
    sb.AppendLine("    \"includeRoles\": " + includeRoles.ToString().ToLower());
    sb.AppendLine("  },");

    // =========================================================================
    //  TABLES
    // =========================================================================
    var tableList = new List<Table>();
    foreach (var t in Model.Tables)
    {
        if (includeHiddenObjects || !t.IsHidden)
            tableList.Add(t);
    }

    var tableJsonItems = new List<string>();

    for (int ti = 0; ti < tableList.Count; ti++)
    {
        var table = tableList[ti];
        var tsb = new StringBuilder();

        tsb.AppendLine("    {");
        tsb.AppendLine("      \"name\": \"" + Esc(table.Name) + "\",");
        tsb.AppendLine("      \"description\": \"" + Esc(table.Description ?? "") + "\",");
        tsb.AppendLine("      \"isHidden\": " + table.IsHidden.ToString().ToLower() + ",");
        tsb.AppendLine("      \"dataCategory\": \"" + Esc(table.DataCategory ?? "") + "\",");

        // -- Table type detection (best-guess + raw partition source types) ----
        var sourceTypes = new List<string>();
        foreach (var p in table.Partitions)
            sourceTypes.Add(p.SourceType.ToString());

        var tableType = "Regular";
        if (sourceTypes.Contains("Calculated")) tableType = "CalculatedTable";

        tsb.AppendLine("      \"tableType\": \"" + tableType + "\",");
        tsb.Append("      \"partitionSourceTypes\": [");
        for (int si = 0; si < sourceTypes.Count; si++)
        {
            tsb.Append("\"" + Esc(sourceTypes[si]) + "\"");
            if (si < sourceTypes.Count - 1) tsb.Append(", ");
        }
        tsb.AppendLine("],");

        // -- Partitions -------------------------------------------------------
        if (includePartitionSources)
        {
            var partItems = new List<string>();
            foreach (var part in table.Partitions)
            {
                var psb = new StringBuilder();
                psb.AppendLine("        {");
                psb.AppendLine("          \"name\": \"" + Esc(part.Name) + "\",");
                psb.AppendLine("          \"sourceType\": \"" + part.SourceType.ToString() + "\",");

                var sourceExpr = "";
                try { sourceExpr = part.Expression ?? ""; } catch { }

                psb.AppendLine("          \"expression\": \"" + Esc(sourceExpr) + "\",");
                psb.AppendLine("          \"mode\": \"" + part.Mode.ToString() + "\"");
                psb.Append("        }");
                partItems.Add(psb.ToString());
            }
            tsb.AppendLine("      \"partitions\": [");
            tsb.AppendLine(JoinItems(partItems, "        "));
            tsb.AppendLine("      ],");
        }

        // -- Columns ----------------------------------------------------------
        var colItems = new List<string>();
        foreach (var col in table.Columns)
        {
            if (!includeHiddenObjects && col.IsHidden) continue;

            var csb = new StringBuilder();
            csb.AppendLine("        {");
            csb.AppendLine("          \"name\": \"" + Esc(col.Name) + "\",");
            csb.AppendLine("          \"dataType\": \"" + col.DataType.ToString() + "\",");
            csb.AppendLine("          \"description\": \"" + Esc(col.Description ?? "") + "\",");
            csb.AppendLine("          \"columnType\": \"" + col.Type.ToString() + "\",");
            csb.AppendLine("          \"isHidden\": " + col.IsHidden.ToString().ToLower() + ",");
            csb.AppendLine("          \"isKey\": " + col.IsKey.ToString().ToLower() + ",");
            csb.AppendLine("          \"formatString\": \"" + Esc(col.FormatString ?? "") + "\",");
            csb.AppendLine("          \"displayFolder\": \"" + Esc(col.DisplayFolder ?? "") + "\",");
            csb.AppendLine("          \"dataCategory\": \"" + Esc(col.DataCategory ?? "") + "\",");
            csb.AppendLine("          \"summarizeBy\": \"" + col.SummarizeBy.ToString() + "\",");
            csb.AppendLine("          \"sortByColumn\": \"" + Esc(col.SortByColumn != null ? col.SortByColumn.Name : "") + "\"");
            csb.Append("        }");
            colItems.Add(csb.ToString());
        }
        tsb.AppendLine("      \"columns\": [");
        tsb.AppendLine(JoinItems(colItems, "        "));
        tsb.AppendLine("      ],");

        // -- Measures ---------------------------------------------------------
        var measItems = new List<string>();
        foreach (var meas in table.Measures)
        {
            if (!includeHiddenObjects && meas.IsHidden) continue;

            var deps = GetDeps(meas.Expression ?? "", meas.Name);

            var msb = new StringBuilder();
            msb.AppendLine("        {");
            msb.AppendLine("          \"name\": \"" + Esc(meas.Name) + "\",");
            msb.AppendLine("          \"expression\": \"" + Esc(meas.Expression ?? "") + "\",");
            msb.AppendLine("          \"description\": \"" + Esc(meas.Description ?? "") + "\",");
            msb.AppendLine("          \"dataType\": \"" + meas.DataType.ToString() + "\",");
            msb.AppendLine("          \"formatString\": \"" + Esc(meas.FormatString ?? "") + "\",");
            msb.AppendLine("          \"displayFolder\": \"" + Esc(meas.DisplayFolder ?? "") + "\",");
            msb.AppendLine("          \"isHidden\": " + meas.IsHidden.ToString().ToLower() + ",");

            msb.Append("          \"referencedMeasures\": [");
            for (int di = 0; di < deps.Count; di++)
            {
                msb.Append("\"" + Esc(deps[di]) + "\"");
                if (di < deps.Count - 1) msb.Append(", ");
            }
            msb.AppendLine("]");
            msb.Append("        }");
            measItems.Add(msb.ToString());
        }
        tsb.AppendLine("      \"measures\": [");
        tsb.AppendLine(JoinItems(measItems, "        "));
        tsb.AppendLine("      ],");

        // -- Hierarchies ------------------------------------------------------
        var hierItems = new List<string>();
        foreach (var hier in table.Hierarchies)
        {
            if (!includeHiddenObjects && hier.IsHidden) continue;

            var hsb = new StringBuilder();
            hsb.AppendLine("        {");
            hsb.AppendLine("          \"name\": \"" + Esc(hier.Name) + "\",");
            hsb.AppendLine("          \"description\": \"" + Esc(hier.Description ?? "") + "\",");
            hsb.AppendLine("          \"isHidden\": " + hier.IsHidden.ToString().ToLower() + ",");
            hsb.AppendLine("          \"levels\": [");

            var lvlItems = new List<string>();
            foreach (var lv in hier.Levels)
            {
                var lsb = new StringBuilder();
                lsb.AppendLine("            {");
                lsb.AppendLine("              \"name\": \"" + Esc(lv.Name) + "\",");
                lsb.AppendLine("              \"ordinal\": " + lv.Ordinal + ",");
                lsb.AppendLine("              \"column\": \"" + Esc(lv.Column.Name) + "\"");
                lsb.Append("            }");
                lvlItems.Add(lsb.ToString());
            }
            hsb.AppendLine(JoinItems(lvlItems, "            "));
            hsb.AppendLine("          ]");
            hsb.Append("        }");
            hierItems.Add(hsb.ToString());
        }
        tsb.AppendLine("      \"hierarchies\": [");
        tsb.AppendLine(JoinItems(hierItems, "        "));
        tsb.AppendLine("      ]");

        // -- Close table
        tsb.Append("    }");
        tableJsonItems.Add(tsb.ToString());
    }

    sb.AppendLine("  \"tables\": [");
    sb.AppendLine(JoinItems(tableJsonItems, "    "));
    sb.AppendLine("  ],");

    // =========================================================================
    //  RELATIONSHIPS
    // =========================================================================
    var relItems = new List<string>();

    foreach (var r in Model.Relationships)
    {
        try
        {
            var rel = r as SingleColumnRelationship;
            if (rel == null) continue;

            var rsb = new StringBuilder();
            rsb.AppendLine("    {");
            rsb.AppendLine("      \"name\": \"" + Esc(rel.Name) + "\",");
            rsb.AppendLine("      \"isActive\": " + rel.IsActive.ToString().ToLower() + ",");
            rsb.AppendLine("      \"fromTable\": \"" + Esc(rel.FromTable.Name) + "\",");
            rsb.AppendLine("      \"fromColumn\": \"" + Esc(rel.FromColumn.Name) + "\",");
            rsb.AppendLine("      \"fromCardinality\": \"" + rel.FromCardinality.ToString() + "\",");
            rsb.AppendLine("      \"toTable\": \"" + Esc(rel.ToTable.Name) + "\",");
            rsb.AppendLine("      \"toColumn\": \"" + Esc(rel.ToColumn.Name) + "\",");
            rsb.AppendLine("      \"toCardinality\": \"" + rel.ToCardinality.ToString() + "\",");
            rsb.AppendLine("      \"crossFilteringBehavior\": \"" + rel.CrossFilteringBehavior.ToString() + "\",");
            rsb.AppendLine("      \"securityFilteringBehavior\": \"" + rel.SecurityFilteringBehavior.ToString() + "\"");
            rsb.Append("    }");
            relItems.Add(rsb.ToString());
        }
        catch { }
    }

    sb.AppendLine("  \"relationships\": [");
    sb.AppendLine(JoinItems(relItems, "    "));
    sb.AppendLine("  ],");

    // =========================================================================
    //  ROLES (Row-Level Security)
    // =========================================================================
    if (includeRoles)
    {
        var roleItems = new List<string>();
        foreach (var role in Model.Roles)
        {
            var rosb = new StringBuilder();
            rosb.AppendLine("    {");
            rosb.AppendLine("      \"name\": \"" + Esc(role.Name) + "\",");
            rosb.AppendLine("      \"description\": \"" + Esc(role.Description ?? "") + "\",");
            rosb.AppendLine("      \"modelPermission\": \"" + role.ModelPermission.ToString() + "\",");

            var tpItems = new List<string>();
            foreach (var tp in role.TablePermissions)
            {
                var tpsb = new StringBuilder();
                tpsb.AppendLine("        {");
                tpsb.AppendLine("          \"table\": \"" + Esc(tp.Table.Name) + "\",");
                tpsb.AppendLine("          \"filterExpression\": \"" + Esc(tp.FilterExpression ?? "") + "\"");
                tpsb.Append("        }");
                tpItems.Add(tpsb.ToString());
            }
            rosb.AppendLine("      \"tablePermissions\": [");
            rosb.AppendLine(JoinItems(tpItems, "        "));
            rosb.AppendLine("      ]");
            rosb.Append("    }");
            roleItems.Add(rosb.ToString());
        }

        sb.AppendLine("  \"roles\": [");
        sb.AppendLine(JoinItems(roleItems, "    "));
        sb.AppendLine("  ],");
    }

    // =========================================================================
    //  DATA SOURCES
    // =========================================================================
    var dsItems = new List<string>();
    foreach (var ds in Model.DataSources)
    {
        var dssb = new StringBuilder();
        dssb.AppendLine("    {");
        dssb.AppendLine("      \"name\": \"" + Esc(ds.Name) + "\",");
        dssb.AppendLine("      \"description\": \"" + Esc(ds.Description ?? "") + "\",");
        dssb.AppendLine("      \"type\": \"" + ds.Type.ToString() + "\"");
        dssb.Append("    }");
        dsItems.Add(dssb.ToString());
    }
    sb.AppendLine("  \"dataSources\": [");
    sb.AppendLine(JoinItems(dsItems, "    "));
    sb.AppendLine("  ],");

    // =========================================================================
    //  SUMMARY (compact AI-friendly reference)
    // =========================================================================
    sb.AppendLine("  \"summary\": {");

    // Table names
    sb.Append("    \"tableNames\": [");
    for (int ti = 0; ti < tableList.Count; ti++)
    {
        sb.Append("\"" + Esc(tableList[ti].Name) + "\"");
        if (ti < tableList.Count - 1) sb.Append(", ");
    }
    sb.AppendLine("],");

    // Relationship map
    var relMapItems = new List<string>();
    foreach (var r in Model.Relationships)
    {
        try
        {
            var rel = r as SingleColumnRelationship;
            if (rel == null) continue;
            var arrow = rel.IsActive ? " --> " : " --x ";
            var card = rel.FromCardinality.ToString() + ":" + rel.ToCardinality.ToString();
            var desc = Esc(rel.FromTable.Name) + "." + Esc(rel.FromColumn.Name)
                     + arrow
                     + Esc(rel.ToTable.Name) + "." + Esc(rel.ToColumn.Name)
                     + " (" + card + ")";
            relMapItems.Add("      \"" + desc + "\"");
        }
        catch { }
    }
    sb.AppendLine("    \"relationshipMap\": [");
    sb.AppendLine(string.Join(",\n", relMapItems));
    sb.AppendLine("    ],");

    // Column inventory per table (compact)
    sb.AppendLine("    \"columnInventory\": {");
    var invItems = new List<string>();
    for (int ti = 0; ti < tableList.Count; ti++)
    {
        var table = tableList[ti];
        var colNames = new List<string>();
        foreach (var c in table.Columns)
        {
            if (includeHiddenObjects || !c.IsHidden)
                colNames.Add("\"" + Esc(c.Name) + " (" + c.DataType.ToString() + ")\"");
        }
        invItems.Add("      \"" + Esc(table.Name) + "\": [" + string.Join(", ", colNames) + "]");
    }
    sb.AppendLine(string.Join(",\n", invItems));
    sb.AppendLine("    },");

    // Measure count per table
    sb.AppendLine("    \"measuresPerTable\": {");
    var mptItems = new List<string>();
    foreach (var t in tableList)
    {
        if (t.Measures.Count > 0)
            mptItems.Add("      \"" + Esc(t.Name) + "\": " + t.Measures.Count);
    }
    sb.AppendLine(string.Join(",\n", mptItems));
    sb.AppendLine("    }");

    sb.AppendLine("  }");

    // ---- Root Close
    sb.AppendLine("}");

    // ---- Write to disk (ensure folder exists, UTF-8 no BOM) -----------------
    Directory.CreateDirectory(outputFolder);
    var outputPath = Path.Combine(outputFolder, outputFileName);
    File.WriteAllText(outputPath, sb.ToString(), new UTF8Encoding(false));

    // ---- Console Summary ----------------------------------------------------
    Console.WriteLine("=== Full Model Export v2 Complete ===");
    Console.WriteLine("  Model:          " + Model.Database.Name);
    Console.WriteLine("  Tables:         " + tableList.Count);
    Console.WriteLine("  Relationships:  " + relItems.Count);
    Console.WriteLine("  Measures:       " + Model.AllMeasures.Count());
    Console.WriteLine("  Roles:          " + Model.Roles.Count);
    Console.WriteLine("  Data Sources:   " + dsItems.Count);
    Console.WriteLine("  Output:         " + outputPath);
}
catch (Exception ex)
{
    Console.WriteLine("Export failed!");
    Console.WriteLine("  Error: " + ex.Message);
    Console.WriteLine("  Type:  " + ex.GetType().Name);
    if (ex.InnerException != null)
        Console.WriteLine("  Inner: " + ex.InnerException.Message);
}
