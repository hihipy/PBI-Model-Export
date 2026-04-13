#r "System.IO.Compression.dll"

using System.IO;
using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;

// ============================================================================
// Power BI Model Export Script for Tabular Editor 2
//
// Exports the full tabular model schema (tables, columns, measures,
// relationships, roles) plus a data head (first N rows) read directly
// from the source files on disk. No external DLL references required.
//
// Supports CSV, TSV, TXT and XLSX source files.
// XLSX is read by parsing the Open XML ZIP format directly — no ACE OLEDB
// or Office installation required.
//
// Output: JSON file written to the current user's Downloads folder.
// ============================================================================

// ----------------------------------------------------------------------------
// Configuration
// ----------------------------------------------------------------------------
var headRowCount            = 50;    // rows to sample per table
var includeHiddenObjects    = true;  // include hidden tables, columns, measures
var includePartitionSources = true;  // include M expression source in output
var includeRoles            = true;  // include RLS role definitions

var outputFolder = Path.Combine(
    Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads");

// ----------------------------------------------------------------------------
// Regex patterns
// ----------------------------------------------------------------------------

// Detects auto-generated PBI internal tables (contain a GUID in the name)
var guidPattern = new Regex(
    @"[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}",
    RegexOptions.Compiled);

// Extracts bracket-wrapped tokens from DAX expressions e.g. [My Measure]
var bracketTokenRegex = new Regex(@"\[([^\]]+)\]", RegexOptions.Compiled);

// ----------------------------------------------------------------------------
// JSON string escaper
// Handles all control characters and special JSON characters safely.
// ----------------------------------------------------------------------------
Func<string, string> Esc = (string s) =>
{
    if (s == null) return "";
    var buf = new StringBuilder(s.Length);
    foreach (var ch in s)
    {
        switch (ch)
        {
            case '\\': buf.Append("\\\\"); break;
            case '"':  buf.Append("\\\""); break;
            case '\b': buf.Append("\\b");  break;
            case '\f': buf.Append("\\f");  break;
            case '\n': buf.Append("\\n");  break;
            case '\r': buf.Append("\\r");  break;
            case '\t': buf.Append("\\t");  break;
            default:
                if (ch < (char)0x20)
                    buf.Append("\\u" + ((int)ch).ToString("x4"));
                else
                    buf.Append(ch);
                break;
        }
    }
    return buf.ToString();
};

// ----------------------------------------------------------------------------
// JSON array joiner
// Joins a list of pre-formatted JSON fragments with comma separators.
// ----------------------------------------------------------------------------
Func<List<string>, string> JoinArray =
    (List<string> items) =>
        items.Count == 0 ? "" : string.Join(",\n", items);

// ----------------------------------------------------------------------------
// Measure dependency resolution
// Builds direct and full recursive dependency chains for each measure.
// Used to identify top-level KPIs (measures not referenced by any other).
// ----------------------------------------------------------------------------
var allMeasureNames = new HashSet<string>();
foreach (var m in Model.AllMeasures)
    allMeasureNames.Add(m.Name);

// Returns all measure names directly referenced in a given DAX expression
Func<string, string, List<string>> GetDirectDeps = (string expr, string selfName) =>
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

// Build expression lookup for recursive resolution
var measureExpressions = new Dictionary<string, string>();
foreach (var m in Model.AllMeasures)
    measureExpressions[m.Name] = m.Expression ?? "";

// Recursively resolves all upstream measure dependencies
Func<string, HashSet<string>, HashSet<string>> GetFullChain = null;
GetFullChain = (string measureName, HashSet<string> visited) =>
{
    var result = new HashSet<string>();
    if (visited.Contains(measureName)) return result; // prevent circular reference loops
    visited.Add(measureName);

    string expr;
    if (!measureExpressions.TryGetValue(measureName, out expr)) return result;

    foreach (Match match in bracketTokenRegex.Matches(expr))
    {
        var token = match.Groups[1].Value;
        if (token != measureName && allMeasureNames.Contains(token))
        {
            result.Add(token);
            foreach (var downstream in GetFullChain(token, visited))
                result.Add(downstream);
        }
    }
    return result;
};

// Track which measures are referenced by other measures.
// Measures NOT in this set are top-level KPIs (entry points).
var referencedByOthers = new HashSet<string>();
foreach (var m in Model.AllMeasures)
    foreach (var dep in GetDirectDeps(m.Expression ?? "", m.Name))
        referencedByOthers.Add(dep);

// ----------------------------------------------------------------------------
// Source folder resolver
//
// Extracts the data source folder path from the table's M expression and
// resolves it dynamically on the current machine. Handles:
//   1. Exact path match (same machine, same user)
//   2. OneDrive path rebuilt using environment variables (different user/tenant)
//   3. Folder name search under OneDrive roots and common profile paths
// ----------------------------------------------------------------------------
Func<string, string> ResolveSourceFolder = (string mExpression) =>
{
    if (string.IsNullOrEmpty(mExpression)) return "";

    // Extract raw path from M expression — try Folder.Files first, then File.Contents
    var rawPath = "";

    var folderMatch = Regex.Match(mExpression, @"Folder\.Files\(""([^""]+)""\)");
    if (folderMatch.Success)
        rawPath = folderMatch.Groups[1].Value;

    if (string.IsNullOrEmpty(rawPath))
    {
        var fileMatch = Regex.Match(mExpression, @"File\.Contents\(""([^""]+)""\)");
        if (fileMatch.Success)
            rawPath = Path.GetDirectoryName(fileMatch.Groups[1].Value) ?? "";
    }

    if (string.IsNullOrEmpty(rawPath)) return "";

    // Strategy 1: try the path exactly as stored in the M expression
    if (Directory.Exists(rawPath)) return rawPath;

    // Strategy 2: rebuild path using real OneDrive root for current user.
    // The stored path may have a different username or tenant name baked in.
    var oneDriveRoots = new List<string>();
    foreach (var envVar in new[] { "OneDriveCommercial", "OneDrive", "OneDriveConsumer" })
    {
        var val = Environment.GetEnvironmentVariable(envVar);
        if (!string.IsNullOrEmpty(val) && !oneDriveRoots.Contains(val))
            oneDriveRoots.Add(val);
    }

    var sep      = Path.DirectorySeparatorChar;
    var segments = rawPath.Split(sep);

    // Find the OneDrive segment in the stored path (e.g. "OneDrive - University of Miami")
    var oneDriveSegmentIndex = -1;
    for (int i = 0; i < segments.Length; i++)
    {
        if (segments[i].StartsWith("OneDrive", StringComparison.OrdinalIgnoreCase))
        {
            oneDriveSegmentIndex = i;
            break;
        }
    }

    if (oneDriveSegmentIndex >= 0)
    {
        // Build relative path from everything after the OneDrive root segment
        var relativeParts = new List<string>();
        for (int i = oneDriveSegmentIndex + 1; i < segments.Length; i++)
            relativeParts.Add(segments[i]);
        var relativePath = string.Join(sep.ToString(), relativeParts);

        foreach (var root in oneDriveRoots)
        {
            var candidate = Path.Combine(root, relativePath);
            if (Directory.Exists(candidate)) return candidate;
        }
    }

    // Strategy 3: search for the folder by name under OneDrive and profile roots.
    // Last resort — works even if the relative path structure has changed.
    var folderName  = segments[segments.Length - 1];
    var searchRoots = new List<string>(oneDriveRoots);
    searchRoots.Add(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile));
    searchRoots.Add(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));

    foreach (var root in searchRoots)
    {
        try
        {
            var matches = Directory.GetDirectories(root, folderName, SearchOption.AllDirectories);
            if (matches.Length > 0) return matches[0];
        }
        catch { /* skip inaccessible directories */ }
    }

    // Return original path so the error message in dataHead is still informative
    return rawPath;
};

// ----------------------------------------------------------------------------
// Data head reader
//
// Reads the first N rows from the source file in the given folder.
// Supports:
//   - CSV / TSV / TXT  : split by delimiter, no dependencies
//   - XLSX             : parsed via Open XML ZIP format using ZipArchive
//                        and XmlDocument — no ACE OLEDB needed
//
// Returns a pre-formatted JSON fragment ready for embedding.
// ----------------------------------------------------------------------------
Func<string, int, string> ReadDataHead = (string folderPath, int maxRows) =>
{
    if (string.IsNullOrEmpty(folderPath))
        return "\"error\": \"No source path found in M expression\"";

    if (!Directory.Exists(folderPath))
        return "\"error\": \"Folder not found: " + Esc(folderPath) + "\"";

    var files = Directory.GetFiles(folderPath);
    if (files.Length == 0)
        return "\"error\": \"No files found in folder\"";

    // Prefer delimited text formats; fall back to first file found
    var preferredExtensions = new HashSet<string> { ".csv", ".txt", ".tsv" };
    var filePath = files.FirstOrDefault(
        f => preferredExtensions.Contains(Path.GetExtension(f).ToLower())) ?? files[0];

    var extension = Path.GetExtension(filePath).ToLower();
    var result    = new StringBuilder();

    // -------------------------------------------------------------------------
    // CSV / TSV / TXT
    // -------------------------------------------------------------------------
    if (extension == ".csv" || extension == ".txt" || extension == ".tsv")
    {
        try
        {
            var delimiter = extension == ".tsv" ? '\t' : ',';
            var lines     = File.ReadAllLines(filePath);

            if (lines.Length == 0)
                return "\"error\": \"File is empty\"";

            var headers = lines[0].Split(delimiter);

            result.Append("\"sourceFile\": \"" + Esc(Path.GetFileName(filePath)) + "\",\n      ");
            result.Append("\"columns\": [");
            result.Append(string.Join(", ",
                headers.Select(h => "\"" + Esc(h.Trim('"').Trim()) + "\"")));
            result.AppendLine("],");

            var rows = new List<string>();
            for (int i = 1; i < Math.Min(lines.Length, maxRows + 1); i++)
            {
                if (string.IsNullOrWhiteSpace(lines[i])) continue;
                var cells = lines[i].Split(delimiter);
                rows.Add("        [" + string.Join(", ",
                    cells.Select(c => "\"" + Esc(c.Trim('"').Trim()) + "\"")) + "]");
            }

            result.AppendLine("\"rows\": [");
            result.Append(string.Join(",\n", rows));
            result.Append("\n      ]");
        }
        catch (Exception ex)
        {
            return "\"error\": \"" + Esc(ex.Message) + "\"";
        }
    }

    // -------------------------------------------------------------------------
    // XLSX — parsed directly from Open XML ZIP structure using ZipArchive.
    // No ACE OLEDB or Office installation required.
    //
    // XLSX internals used:
    //   xl/sharedStrings.xml     — lookup table for all string cell values
    //   xl/worksheets/sheet1.xml — cell data for the first worksheet
    //
    // Cell type="s" means the value is an index into sharedStrings.
    // All other cell types are read as raw text from the <v> element.
    // -------------------------------------------------------------------------
    else if (extension == ".xlsx")
    {
        try
        {
            result.Append("\"sourceFile\": \"" + Esc(Path.GetFileName(filePath)) + "\",\n      ");

            // Open the xlsx as a ZIP archive using ZipArchive (no ZipFile needed)
            using (var fileStream = File.OpenRead(filePath))
            using (var zip = new ZipArchive(fileStream, ZipArchiveMode.Read))
            {
                // Load shared strings — Excel de-duplicates all strings into this table
                var sharedStrings = new List<string>();
                var ssEntry = zip.GetEntry("xl/sharedStrings.xml");
                if (ssEntry != null)
                {
                    using (var stream = ssEntry.Open())
                    {
                        var ssDoc = new XmlDocument();
                        ssDoc.Load(stream);
                        var nsMgr = new XmlNamespaceManager(ssDoc.NameTable);
                        nsMgr.AddNamespace("x",
                            "http://schemas.openxmlformats.org/spreadsheetml/2006/main");

                        foreach (XmlNode si in ssDoc.SelectNodes("//x:si", nsMgr))
                            sharedStrings.Add(si.InnerText);
                    }
                }

                // Find the first worksheet — prefer sheet1.xml, fall back to any sheet
                var sheetEntry = zip.GetEntry("xl/worksheets/sheet1.xml")
                    ?? zip.Entries.FirstOrDefault(
                        e => e.FullName.StartsWith("xl/worksheets/sheet")
                          && e.FullName.EndsWith(".xml"));

                if (sheetEntry == null)
                    return "\"error\": \"No worksheet found in xlsx\"";

                using (var stream = sheetEntry.Open())
                {
                    var doc   = new XmlDocument();
                    doc.Load(stream);
                    var nsMgr = new XmlNamespaceManager(doc.NameTable);
                    nsMgr.AddNamespace("x",
                        "http://schemas.openxmlformats.org/spreadsheetml/2006/main");

                    var allRows = doc.SelectNodes("//x:sheetData/x:row", nsMgr);
                    if (allRows == null || allRows.Count == 0)
                        return "\"error\": \"No rows found in worksheet\"";

                    // Resolves a single cell node to its display string value
                    Func<XmlNode, string> CellValue = (XmlNode cell) =>
                    {
                        var typeAttr = cell.Attributes["t"];
                        var vNode    = cell.SelectSingleNode("x:v", nsMgr);
                        if (vNode == null) return "";
                        var raw = vNode.InnerText;

                        // type="s" — value is a shared string index
                        if (typeAttr != null && typeAttr.Value == "s")
                        {
                            int idx;
                            if (int.TryParse(raw, out idx) && idx < sharedStrings.Count)
                                return sharedStrings[idx];
                        }
                        return raw;
                    };

                    // Row 0 is the header row
                    var headerCells = allRows[0].SelectNodes("x:c", nsMgr);
                    var headers     = new List<string>();
                    foreach (XmlNode cell in headerCells)
                        headers.Add(CellValue(cell));

                    result.Append("\"columns\": [");
                    result.Append(string.Join(", ",
                        headers.Select(h => "\"" + Esc(h) + "\"")));
                    result.AppendLine("],");

                    // Rows 1..N are data rows
                    var rows = new List<string>();
                    for (int i = 1; i < Math.Min(allRows.Count, maxRows + 1); i++)
                    {
                        var rowCells = allRows[i].SelectNodes("x:c", nsMgr);
                        var cells    = new List<string>();
                        foreach (XmlNode cell in rowCells)
                            cells.Add("\"" + Esc(CellValue(cell)) + "\"");
                        rows.Add("        [" + string.Join(", ", cells) + "]");
                    }

                    result.AppendLine("\"rows\": [");
                    result.Append(string.Join(",\n", rows));
                    result.Append("\n      ]");
                }
            }
        }
        catch (Exception ex)
        {
            return "\"error\": \"" + Esc(ex.Message) + "\"";
        }
    }

    // -------------------------------------------------------------------------
    // Unsupported format — return file list only
    // -------------------------------------------------------------------------
    else
    {
        var fileNames = files.Take(maxRows).Select(f => "\"" + Esc(Path.GetFileName(f)) + "\"");
        return "\"note\": \"Unsupported format — file list only\",\n      " +
               "\"filesFound\": [" + string.Join(", ", fileNames) + "]";
    }

    return result.ToString();
};

// ----------------------------------------------------------------------------
// Dashboard name
// Reads the Power BI Desktop window title to name the output file.
// ----------------------------------------------------------------------------
var pbiProcess    = System.Diagnostics.Process.GetProcessesByName("PBIDesktop").FirstOrDefault();
var dashboardName = "UnknownReport";

if (pbiProcess != null && !string.IsNullOrEmpty(pbiProcess.MainWindowTitle))
{
    var windowTitle = pbiProcess.MainWindowTitle;
    var suffixIndex = windowTitle.LastIndexOf(" - Power BI Desktop");
    dashboardName   = suffixIndex > 0
        ? windowTitle.Substring(0, suffixIndex).Trim()
        : windowTitle.Trim();
}

var safeFileName = string.Concat(dashboardName.Split(Path.GetInvalidFileNameChars()));
if (string.IsNullOrWhiteSpace(safeFileName)) safeFileName = "UnknownReport";
var outputPath = Path.Combine(outputFolder, safeFileName + "_ModelExport.json");

// ----------------------------------------------------------------------------
// Build JSON output
// ----------------------------------------------------------------------------
Directory.CreateDirectory(outputFolder);
var sb = new StringBuilder();

sb.AppendLine("{");
sb.AppendLine("  \"exportDate\": \""    + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "\",");
sb.AppendLine("  \"dashboardName\": \"" + Esc(dashboardName) + "\",");
sb.AppendLine("  \"modelName\": \""     + Esc(Model.Database.Name) + "\",");

// --- Model metadata ----------------------------------------------------------
sb.AppendLine("  \"modelMetadata\": {");
sb.AppendLine("    \"compatibilityLevel\": " + Model.Database.CompatibilityLevel.ToString() + ",");
sb.AppendLine("    \"culture\": \""           + Esc(Model.Culture) + "\",");
sb.AppendLine("    \"defaultMode\": \""       + Model.DefaultMode.ToString() + "\",");
sb.AppendLine("    \"description\": \""       + Esc(Model.Database.Description ?? "") + "\",");
sb.AppendLine("    \"tableCount\": "          + Model.Tables.Count.ToString() + ",");
sb.AppendLine("    \"relationshipCount\": "   + Model.Relationships.Count.ToString() + ",");
sb.AppendLine("    \"measureCount\": "        + Model.AllMeasures.Count().ToString() + ",");
sb.AppendLine("    \"roleCount\": "           + Model.Roles.Count.ToString());
sb.AppendLine("  },");

// --- Export configuration ----------------------------------------------------
sb.AppendLine("  \"exportConfiguration\": {");
sb.AppendLine("    \"includeHiddenObjects\": "    + includeHiddenObjects.ToString().ToLower() + ",");
sb.AppendLine("    \"includePartitionSources\": " + includePartitionSources.ToString().ToLower() + ",");
sb.AppendLine("    \"includeRoles\": "            + includeRoles.ToString().ToLower() + ",");
sb.AppendLine("    \"headRowCount\": "            + headRowCount.ToString());
sb.AppendLine("  },");

// --- Tables ------------------------------------------------------------------
var tableList  = new List<Table>();
foreach (var t in Model.Tables)
    if (includeHiddenObjects || !t.IsHidden) tableList.Add(t);

var userTables = tableList.Where(t => !guidPattern.IsMatch(t.Name)).ToList();
var autoTables = tableList.Where(t =>  guidPattern.IsMatch(t.Name)).ToList();

// Cache resolved source folder per table for reuse in the data head section
var tableSourceFolders = new Dictionary<string, string>();
var tableJsonItems     = new List<string>();

foreach (var table in tableList)
{
    // Resolve the source folder path from the partition M expression
    var sourceFolder = "";
    foreach (var part in table.Partitions)
    {
        try
        {
            var resolved = ResolveSourceFolder(part.Expression ?? "");
            if (!string.IsNullOrEmpty(resolved)) { sourceFolder = resolved; break; }
        }
        catch { }
    }
    tableSourceFolders[table.Name] = sourceFolder;

    var sourceTypes = table.Partitions.Select(p => p.SourceType.ToString()).ToList();
    var tableType   = sourceTypes.Contains("Calculated") ? "CalculatedTable" : "Regular";

    var tsb = new StringBuilder();
    tsb.AppendLine("    {");
    tsb.AppendLine("      \"name\": \""           + Esc(table.Name) + "\",");
    tsb.AppendLine("      \"description\": \""    + Esc(table.Description ?? "") + "\",");
    tsb.AppendLine("      \"isHidden\": "         + table.IsHidden.ToString().ToLower() + ",");
    tsb.AppendLine("      \"isAutoGenerated\": "  + guidPattern.IsMatch(table.Name).ToString().ToLower() + ",");
    tsb.AppendLine("      \"dataCategory\": \""   + Esc(table.DataCategory ?? "") + "\",");
    tsb.AppendLine("      \"sourceFolder\": \""   + Esc(sourceFolder) + "\",");
    tsb.AppendLine("      \"tableType\": \""      + tableType + "\",");
    tsb.Append    ("      \"partitionSourceTypes\": [");
    tsb.Append    (string.Join(", ", sourceTypes.Select(st => "\"" + st + "\"")));
    tsb.AppendLine("],");

    if (includePartitionSources)
    {
        var partItems = new List<string>();
        foreach (var part in table.Partitions)
        {
            var srcExpr = "";
            try { srcExpr = part.Expression ?? ""; } catch { }
            partItems.Add(
                "        {\n" +
                "          \"name\": \""       + Esc(part.Name) + "\",\n" +
                "          \"sourceType\": \"" + part.SourceType.ToString() + "\",\n" +
                "          \"expression\": \"" + Esc(srcExpr) + "\",\n" +
                "          \"mode\": \""       + part.Mode.ToString() + "\"\n        }");
        }
        tsb.AppendLine("      \"partitions\": [");
        tsb.AppendLine(JoinArray(partItems));
        tsb.AppendLine("      ],");
    }

    // Columns
    var colItems = new List<string>();
    foreach (var col in table.Columns)
    {
        if (!includeHiddenObjects && col.IsHidden) continue;
        colItems.Add(
            "        {\n" +
            "          \"name\": \""           + Esc(col.Name) + "\",\n" +
            "          \"dataType\": \""       + col.DataType.ToString() + "\",\n" +
            "          \"description\": \""    + Esc(col.Description ?? "") + "\",\n" +
            "          \"columnType\": \""     + col.Type.ToString() + "\",\n" +
            "          \"isHidden\": "         + col.IsHidden.ToString().ToLower() + ",\n" +
            "          \"isKey\": "            + col.IsKey.ToString().ToLower() + ",\n" +
            "          \"formatString\": \""   + Esc(col.FormatString ?? "") + "\",\n" +
            "          \"displayFolder\": \""  + Esc(col.DisplayFolder ?? "") + "\",\n" +
            "          \"dataCategory\": \""   + Esc(col.DataCategory ?? "") + "\",\n" +
            "          \"summarizeBy\": \""    + col.SummarizeBy.ToString() + "\",\n" +
            "          \"sortByColumn\": \""   +
                Esc(col.SortByColumn != null ? col.SortByColumn.Name : "") + "\"\n        }");
    }
    tsb.AppendLine("      \"columns\": [");
    tsb.AppendLine(JoinArray(colItems));
    tsb.AppendLine("      ],");

    // Measures
    var measItems = new List<string>();
    foreach (var meas in table.Measures)
    {
        if (!includeHiddenObjects && meas.IsHidden) continue;
        var directDeps = GetDirectDeps(meas.Expression ?? "", meas.Name);
        var fullChain  = GetFullChain(meas.Name, new HashSet<string>())
                             .OrderBy(x => x).ToList();
        measItems.Add(
            "        {\n" +
            "          \"name\": \""               + Esc(meas.Name) + "\",\n" +
            "          \"expression\": \""         + Esc(meas.Expression ?? "") + "\",\n" +
            "          \"description\": \""        + Esc(meas.Description ?? "") + "\",\n" +
            "          \"dataType\": \""           + meas.DataType.ToString() + "\",\n" +
            "          \"formatString\": \""       + Esc(meas.FormatString ?? "") + "\",\n" +
            "          \"displayFolder\": \""      + Esc(meas.DisplayFolder ?? "") + "\",\n" +
            "          \"isHidden\": "             + meas.IsHidden.ToString().ToLower() + ",\n" +
            "          \"isTopLevelKPI\": "        +
                (!referencedByOthers.Contains(meas.Name)).ToString().ToLower() + ",\n" +
            "          \"referencedMeasures\": ["  +
                string.Join(", ", directDeps.Select(d => "\"" + Esc(d) + "\"")) + "],\n" +
            "          \"fullDependencyChain\": [" +
                string.Join(", ", fullChain.Select(d => "\"" + Esc(d) + "\"")) + "]\n        }");
    }
    tsb.AppendLine("      \"measures\": [");
    tsb.AppendLine(JoinArray(measItems));
    tsb.AppendLine("      ],");

    // Hierarchies
    var hierItems = new List<string>();
    foreach (var hier in table.Hierarchies)
    {
        if (!includeHiddenObjects && hier.IsHidden) continue;
        var lvlItems = hier.Levels.Select(lv =>
            "            {\n" +
            "              \"name\": \""   + Esc(lv.Name) + "\",\n" +
            "              \"ordinal\": "  + lv.Ordinal.ToString() + ",\n" +
            "              \"column\": \"" + Esc(lv.Column.Name) + "\"\n            }")
            .ToList();
        hierItems.Add(
            "        {\n" +
            "          \"name\": \""        + Esc(hier.Name) + "\",\n" +
            "          \"description\": \"" + Esc(hier.Description ?? "") + "\",\n" +
            "          \"isHidden\": "       + hier.IsHidden.ToString().ToLower() + ",\n" +
            "          \"levels\": [\n"      + JoinArray(lvlItems) + "\n          ]\n        }");
    }
    tsb.AppendLine("      \"hierarchies\": [");
    tsb.AppendLine(JoinArray(hierItems));
    tsb.AppendLine("      ]");
    tsb.Append    ("    }");
    tableJsonItems.Add(tsb.ToString());
}

sb.AppendLine("  \"tables\": [");
sb.AppendLine(JoinArray(tableJsonItems));
sb.AppendLine("  ],");

// --- Relationships -----------------------------------------------------------
var relItems = new List<string>();
foreach (var r in Model.Relationships)
{
    try
    {
        var rel        = r as SingleColumnRelationship;
        if (rel == null) continue;
        var isFromMany = rel.FromCardinality.ToString() == "Many";
        relItems.Add(
            "    {\n" +
            "      \"name\": \""                      + Esc(rel.Name) + "\",\n" +
            "      \"isActive\": "                    + rel.IsActive.ToString().ToLower() + ",\n" +
            "      \"fromTable\": \""                 + Esc(rel.FromTable.Name) + "\",\n" +
            "      \"fromColumn\": \""                + Esc(rel.FromColumn.Name) + "\",\n" +
            "      \"fromCardinality\": \""           + rel.FromCardinality.ToString() + "\",\n" +
            "      \"toTable\": \""                   + Esc(rel.ToTable.Name) + "\",\n" +
            "      \"toColumn\": \""                  + Esc(rel.ToColumn.Name) + "\",\n" +
            "      \"toCardinality\": \""             + rel.ToCardinality.ToString() + "\",\n" +
            "      \"crossFilteringBehavior\": \""    + rel.CrossFilteringBehavior.ToString() + "\",\n" +
            "      \"securityFilteringBehavior\": \"" + rel.SecurityFilteringBehavior.ToString() + "\",\n" +
            "      \"factTable\": \""      + Esc(isFromMany ? rel.FromTable.Name : rel.ToTable.Name) + "\",\n" +
            "      \"dimensionTable\": \"" + Esc(isFromMany ? rel.ToTable.Name : rel.FromTable.Name) + "\"\n    }");
    }
    catch { }
}
sb.AppendLine("  \"relationships\": [");
sb.AppendLine(JoinArray(relItems));
sb.AppendLine("  ],");

// --- Roles (Row-Level Security) ----------------------------------------------
if (includeRoles)
{
    var roleItems = new List<string>();
    foreach (var role in Model.Roles)
    {
        var permItems = role.TablePermissions.Select(tp =>
            "        {\n" +
            "          \"table\": \""            + Esc(tp.Table.Name) + "\",\n" +
            "          \"filterExpression\": \"" + Esc(tp.FilterExpression ?? "") + "\"\n        }")
            .ToList();
        roleItems.Add(
            "    {\n" +
            "      \"name\": \""            + Esc(role.Name) + "\",\n" +
            "      \"description\": \""     + Esc(role.Description ?? "") + "\",\n" +
            "      \"modelPermission\": \"" + role.ModelPermission.ToString() + "\",\n" +
            "      \"tablePermissions\": [\n" + JoinArray(permItems) + "\n      ]\n    }");
    }
    sb.AppendLine("  \"roles\": [");
    sb.AppendLine(JoinArray(roleItems));
    sb.AppendLine("  ],");
}

// --- Data sources ------------------------------------------------------------
var dsItems = Model.DataSources.Select(ds =>
    "    {\n" +
    "      \"name\": \""        + Esc(ds.Name) + "\",\n" +
    "      \"description\": \"" + Esc(ds.Description ?? "") + "\",\n" +
    "      \"type\": \""        + ds.Type.ToString() + "\"\n    }").ToList();
sb.AppendLine("  \"dataSources\": [");
sb.AppendLine(JoinArray(dsItems));
sb.AppendLine("  ],");

// --- Summary (compact AI-friendly reference) ---------------------------------
var topLevelMeasures = Model.AllMeasures
    .Where(m => !referencedByOthers.Contains(m.Name))
    .Select(m => m.Name).OrderBy(n => n).ToList();

var folderGroups = new Dictionary<string, List<string>>();
foreach (var m in Model.AllMeasures)
{
    var folder = string.IsNullOrEmpty(m.DisplayFolder) ? "(No Folder)" : m.DisplayFolder;
    if (!folderGroups.ContainsKey(folder)) folderGroups[folder] = new List<string>();
    folderGroups[folder].Add(m.Name);
}

var relMapItems = new List<string>();
foreach (var r in Model.Relationships)
{
    try
    {
        var rel = r as SingleColumnRelationship;
        if (rel == null) continue;
        relMapItems.Add(
            "      \"" +
            Esc(rel.FromTable.Name) + "." + Esc(rel.FromColumn.Name) +
            (rel.IsActive ? " --> " : " --x ") +
            Esc(rel.ToTable.Name) + "." + Esc(rel.ToColumn.Name) +
            " (" + rel.FromCardinality.ToString() + ":" + rel.ToCardinality.ToString() + ")\"");
    }
    catch { }
}

sb.AppendLine("  \"summary\": {");
sb.Append    ("    \"userFacingTables\": [");
sb.Append    (string.Join(", ", userTables.Select(t => "\"" + Esc(t.Name) + "\"")));
sb.AppendLine("],");
sb.Append    ("    \"autoGeneratedTables\": [");
sb.Append    (string.Join(", ", autoTables.Select(t => "\"" + Esc(t.Name) + "\"")));
sb.AppendLine("],");
sb.Append    ("    \"topLevelKPIs\": [");
sb.Append    (string.Join(", ", topLevelMeasures.Select(n => "\"" + Esc(n) + "\"")));
sb.AppendLine("],");
sb.AppendLine("    \"measuresByFolder\": {");
sb.AppendLine(string.Join(",\n", folderGroups.Select(kv =>
    "      \"" + Esc(kv.Key) + "\": [" +
    string.Join(", ", kv.Value.Select(n => "\"" + Esc(n) + "\"")) + "]")));
sb.AppendLine("    },");
sb.AppendLine("    \"relationshipMap\": [");
sb.AppendLine(string.Join(",\n", relMapItems));
sb.AppendLine("    ],");
sb.AppendLine("    \"columnInventory\": {");
sb.AppendLine(string.Join(",\n", tableList.Select(t =>
    "      \"" + Esc(t.Name) + "\": [" + string.Join(", ",
        t.Columns
            .Where(c => includeHiddenObjects || !c.IsHidden)
            .Select(c => "\"" + Esc(c.Name) + " (" + c.DataType.ToString() + ")\"")) + "]")));
sb.AppendLine("    },");
sb.AppendLine("    \"measuresPerTable\": {");
sb.AppendLine(string.Join(",\n", tableList
    .Where(t => t.Measures.Count > 0)
    .Select(t => "      \"" + Esc(t.Name) + "\": " + t.Measures.Count.ToString())));
sb.AppendLine("    }");
sb.AppendLine("  },");

// --- Data head ---------------------------------------------------------------
sb.AppendLine("  \"dataHead\": {");
var dataHeadItems = new List<string>();
foreach (var table in userTables)
{
    var folderPath = tableSourceFolders.ContainsKey(table.Name)
        ? tableSourceFolders[table.Name] : "";
    dataHeadItems.Add(
        "    \"" + Esc(table.Name) + "\": {\n" +
        "      " + ReadDataHead(folderPath, headRowCount) +
        "\n    }");
}
sb.AppendLine(JoinArray(dataHeadItems));
sb.AppendLine("  }");
sb.AppendLine("}");

// ----------------------------------------------------------------------------
// Write output file (UTF-8, no BOM)
// ----------------------------------------------------------------------------
File.WriteAllText(outputPath, sb.ToString(), new UTF8Encoding(false));

Console.WriteLine("=== Model export complete ===");
Console.WriteLine("  Dashboard  : " + dashboardName);
Console.WriteLine("  Model      : " + Model.Database.Name);
Console.WriteLine("  Tables     : " + tableList.Count.ToString() + " (" + autoTables.Count.ToString() + " auto-generated)");
Console.WriteLine("  Measures   : " + Model.AllMeasures.Count().ToString() + " (" + topLevelMeasures.Count.ToString() + " top-level KPIs)");
Console.WriteLine("  Relations  : " + relItems.Count.ToString());
Console.WriteLine("  Head rows  : " + headRowCount.ToString() + " per table");
Console.WriteLine("  Output     : " + outputPath);