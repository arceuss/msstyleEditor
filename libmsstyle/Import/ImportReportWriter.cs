using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace libmsstyle.Import
{
    public static class ImportReportWriter
    {
        /// <summary>writes &lt;prefix&gt;.json, &lt;prefix&gt;.md and &lt;prefix&gt;.threeway.json; returns the paths.</summary>
        public static List<string> Write(ImportReport r, string prefix)
        {
            var paths = new List<string> { prefix + ".json", prefix + ".md", prefix + ".threeway.json" };
            Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(prefix)));
            File.WriteAllText(paths[0], JsonWriter.Serialize(r), Encoding.UTF8);
            File.WriteAllText(paths[1], ToMarkdown(r), Encoding.UTF8);
            File.WriteAllText(paths[2], JsonWriter.Serialize(new { Summary = r.ThreeWaySummary, Entries = r.ThreeWayEntries }), Encoding.UTF8);
            return paths;
        }

        public static string ToMarkdown(ImportReport r)
        {
            var sb = new StringBuilder();
            sb.AppendLine("# Windows 7 style import report");
            sb.AppendLine();
            sb.AppendLine($"Status: **{r.Status}** ({(r.DryRun ? "dry run" : "write")}). {r.StatusText}");
            sb.AppendLine();
            sb.AppendLine("Offline checks only. The candidate was never loaded by uxtheme or applied; it is unverified visually until tested in a disposable VM.");
            sb.AppendLine();
            sb.AppendLine("| input | path |");
            sb.AppendLine("|---|---|");
            foreach (var kv in r.Inputs) sb.AppendLine($"| {kv.Key} | `{kv.Value}` |");
            if (r.Output != null) sb.AppendLine($"| output | `{r.Output}` |");
            sb.AppendLine();

            List("Errors", r.Errors, sb);
            List("Warnings", r.Warnings, sb);

            sb.AppendLine("## Counts");
            sb.AppendLine();
            sb.AppendLine("| what | count |");
            sb.AppendLine("|---|---:|");
            foreach (var kv in r.Counts.OrderBy(k => k.Key)) sb.AppendLine($"| {kv.Key} | {kv.Value} |");
            sb.AppendLine();

            sb.AppendLine("## Three-way comparison (keyed by class / part name / state / property)");
            sb.AppendLine();
            sb.AppendLine("| category | entries |");
            sb.AppendLine("|---|---:|");
            foreach (var kv in r.ThreeWaySummary.OrderBy(k => k.Key)) sb.AppendLine($"| {kv.Key} | {kv.Value} |");
            sb.AppendLine();
            sb.AppendLine("`source-changed`: source differs from stock win7. `template-adapted`: template differs from stock win7. `conflict`: both. `win7-only`: no such record in the template. Full list in the `.threeway.json` file.");
            sb.AppendLine();

            var d = r.Dwm;
            sb.AppendLine("## DWMWindow");
            sb.AppendLine();
            sb.AppendLine($"- template atlas is the stock win7 atlas: {d.TemplateAtlasIsStockWin7}");
            sb.AppendLine($"- source atlas differs from stock: {d.SourceAtlasChanged} ({d.SourceAtlasSize})");
            sb.AppendLine($"- atlas stream: template {d.TemplateAtlasId} -> output {d.NewAtlasId} ({d.NewAtlasSize}{(d.AtlasReencoded ? ", re-encoded with kept template sprites" : ", source bytes unchanged")})");
            sb.AppendLine($"- win7-only parts dropped: {(d.Win7OnlyParts.Count == 0 ? "none" : string.Join(", ", d.Win7OnlyParts))}");
            sb.AppendLine($"- template sprites kept (fallback): {(d.FallbackParts.Count == 0 ? "none" : string.Join(", ", d.FallbackParts.Distinct()))}");
            sb.AppendLine();
            if (d.Parts.Count > 0)
            {
                sb.AppendLine("| id | template part | from win7 part | method | template rect | new rect | note |");
                sb.AppendLine("|---:|---|---|---|---|---|---|");
                foreach (var p in d.Parts)
                    sb.AppendLine($"| {p.TargetPartId} | {p.TargetPart} | {p.SourcePart} | {p.Method} | {Rect(p.TargetRect)} | {Rect(p.NewRect)} | {p.Note} |");
                sb.AppendLine();
            }

            sb.AppendLine("## Validation");
            sb.AppendLine();
            foreach (var c in r.Validation.Checks) sb.AppendLine("- ok: " + c);
            foreach (var e in r.Validation.Errors) sb.AppendLine("- FAIL: " + e);
            foreach (var w in r.Validation.Warnings) sb.AppendLine("- warn: " + w);
            sb.AppendLine();

            if (r.Resources.Count > 0)
            {
                sb.AppendLine("## Resources copied");
                sb.AppendLine();
                sb.AppendLine($"{r.Resources.Count(x => !x.Reused)} added, {r.Resources.Count(x => x.Reused)} shared with identical bytes already in the template.");
                sb.AppendLine();
                sb.AppendLine("| type | source id | output id | reused | size | bytes |");
                sb.AppendLine("|---|---:|---:|---|---|---:|");
                foreach (var m in r.Resources) sb.AppendLine($"| {m.Type} | {m.SourceId} | {m.TargetId} | {m.Reused} | {m.Dimensions} | {m.Bytes} |");
                sb.AppendLine();
            }

            List("Inheritance differences (win7 vs template, informational)", r.InheritanceDifferences, sb);
            List("States that keep template-only images (mostly 192-384 dpi variants the win7 source never had)", r.TemplateOnlyImages, sb);
            List("Source classes missing in the template", r.SourceClassesMissingInTarget, sb);
            List("Notes", r.Notes, sb);

            var interesting = r.Decisions.Where(x => x.Action != "applied").ToList();
            sb.AppendLine("## Decisions other than plain applies");
            sb.AppendLine();
            if (interesting.Count == 0) sb.AppendLine("none");
            else
            {
                sb.AppendLine("| class | part | state | property | action | reason | win7 | source | template |");
                sb.AppendLine("|---|---|---:|---|---|---|---|---|---|");
                foreach (var x in interesting)
                    sb.AppendLine($"| {x.Class} | {x.Part} | {x.State} | {x.Property} | {x.Action} | {x.Reason} | {x.Base} | {x.Source} | {x.Target} |");
            }
            sb.AppendLine();
            sb.AppendLine("Every decision, including applies, is in the `.json` report.");
            return sb.ToString();
        }

        private static string Rect(string s)
        {
            return s?.Replace("{", "").Replace("}", "");
        }

        private static void List(string title, List<string> items, StringBuilder sb)
        {
            if (items.Count == 0) return;
            sb.AppendLine("## " + title);
            sb.AppendLine();
            foreach (var i in items) sb.AppendLine("- " + i);
            sb.AppendLine();
        }
    }
}
