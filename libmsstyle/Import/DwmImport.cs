using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;

namespace libmsstyle.Import
{
    /// <summary>
    /// DWMWindow is a single PNG atlas (Common Properties, IMAGEATLAS, DISKSTREAM) plus one
    /// ATLASRECT per part. Win7 and win10 number the parts differently, so parts are matched by
    /// name, never by id.
    ///
    /// The mapping only works because the template's atlas is the stock win7 atlas (checked, not
    /// assumed): then every template rect is a stock win7 rect and can be traced back to the win7
    /// part that owns it, even for win10-only parts that reuse a win7 sprite (the DARK glyphs, the
    /// extra caption button backgrounds 88-91). The source atlas replaces the template atlas and
    /// each rect is pointed at the source's sprite for the same win7 part. A template sprite that
    /// can't be traced or whose new geometry doesn't fit is copied over from the template atlas
    /// unchanged, so no frame/caption/button part is ever guessed.
    /// </summary>
    internal class DwmImport
    {
        private readonly Win7StyleImporter m_imp;
        private readonly DwmReport m_rep;

        private class PartData
        {
            public string Name;
            public int Id;
            public List<SemanticEntry> Entries = new List<SemanticEntry>();
            public IEnumerable<SemanticEntry> Rects => Entries.Where(e => e.NameId == (int)IDENTIFIER.ATLASRECT && e.TypeId == (int)IDENTIFIER.RECTTYPE);
            public SemanticEntry Find(int state, string prop, int typeId, int occurrence)
            {
                return Entries.FirstOrDefault(e => e.StateId == state && e.PropertyName == prop && e.TypeId == typeId && e.Occurrence == occurrence);
            }
        }

        private class RectPlan
        {
            public PartData Part;
            public SemanticEntry Entry;
            public Rectangle TargetRect;
            public string Src;
            public Rectangle SourceRect;
            public string Method;
            public string Note;
        }

        public DwmImport(Win7StyleImporter importer)
        {
            m_imp = importer;
            m_rep = importer.Report.Dwm;
        }

        private static Dictionary<string, PartData> Group(SemanticModel m)
        {
            var d = new Dictionary<string, PartData>(StringComparer.OrdinalIgnoreCase);
            foreach (var e in m.Entries.Where(e => string.Equals(e.ClassName, Win7StyleImporter.DwmClass, StringComparison.OrdinalIgnoreCase)))
            {
                if (!d.TryGetValue(e.PartName, out var p))
                {
                    p = new PartData { Name = e.PartName, Id = e.PartId };
                    d[e.PartName] = p;
                }
                p.Entries.Add(e);
            }
            return d;
        }

        internal static Rectangle RectOf(SemanticEntry e)
        {
            var m = (Margins)e.Property.GetValue();
            return Rectangle.FromLTRB(m.Left, m.Top, m.Right, m.Bottom);
        }

        private static SemanticEntry AtlasEntry(Dictionary<string, PartData> parts)
        {
            return parts.Values.Where(p => p.Id == 0)
                .SelectMany(p => p.Entries)
                .FirstOrDefault(e => e.TypeId == (int)IDENTIFIER.DISKSTREAM);
        }

        public void Run()
        {
            var report = m_imp.Report;
            var b = Group(m_imp.Base);
            var s = Group(m_imp.Source);
            var t = Group(m_imp.Target);

            if (t.Count == 0)
            {
                report.Errors.Add("template has no DWMWindow class; it can't be a working win10 style.");
                return;
            }
            var ta = AtlasEntry(t);
            var ba = AtlasEntry(b);
            var sa = AtlasEntry(s);
            if (ta == null)
            {
                report.Errors.Add("template DWMWindow has no IMAGEATLAS stream.");
                return;
            }
            m_rep.TemplateAtlasId = ta.ResourceId;
            m_rep.NewAtlasId = ta.ResourceId;
            if (KeyPattern.Matches(m_imp.Options.Exclude, Win7StyleImporter.DwmClass, ta.PartName, ta.PropertyName))
            {
                report.Notes.Add("DWMWindow atlas excluded by rule; template frames kept.");
                return;
            }
            if (ba == null || sa == null)
            {
                report.Warnings.Add("DWMWindow not imported: base or source has no DWMWindow IMAGEATLAS. template frames kept.");
                return;
            }

            var tAtlas = m_imp.Decoded(m_imp.Target, "STREAM", ta.ResourceId);
            var bAtlas = m_imp.Decoded(m_imp.Base, "STREAM", ba.ResourceId);
            var sAtlas = m_imp.Decoded(m_imp.Source, "STREAM", sa.ResourceId);
            if (tAtlas == null || bAtlas == null || sAtlas == null)
            {
                report.Errors.Add("a DWMWindow atlas is missing or isn't a decodable png; DWM data can't be checked.");
                return;
            }
            m_rep.SourceAtlasSize = $"{sAtlas.Width}x{sAtlas.Height}";
            m_rep.TemplateAtlasIsStockWin7 = Win7StyleImporter.NearlyEqual(tAtlas, bAtlas, 0);
            m_rep.SourceAtlasChanged = !Win7StyleImporter.NearlyEqual(sAtlas, bAtlas, 0);

            if (!m_rep.TemplateAtlasIsStockWin7)
            {
                report.Warnings.Add("DWMWindow not imported: the template atlas isn't the stock win7 atlas, so its rects can't be traced to win7 parts. template frames kept.");
                return;
            }

            // stock rect -> win7 part names owning it
            var baseByRect = new Dictionary<string, List<string>>();
            foreach (var p in b.Values)
                foreach (var r in p.Rects)
                {
                    string k = r.StateId + ":" + RectOf(r);
                    if (!baseByRect.TryGetValue(k, out var l)) baseByRect[k] = l = new List<string>();
                    l.Add(p.Name);
                }

            var plans = new List<RectPlan>();
            var srcOfPart = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var tp in t.Values.OrderBy(p => p.Id))
            {
                foreach (var re in tp.Rects)
                {
                    var plan = new RectPlan { Part = tp, Entry = re, TargetRect = RectOf(re) };
                    plans.Add(plan);
                    if (plan.TargetRect.Width <= 0 || plan.TargetRect.Height <= 0)
                    {
                        plan.Method = "none";
                        plan.Note = "empty rect in the template, left alone";
                        continue;
                    }
                    ResolveSource(plan, b, s, sAtlas);
                    if (plan.Src != null && !srcOfPart.ContainsKey(tp.Name)) srcOfPart[tp.Name] = plan.Src;
                }
            }

            // parts without a rect still carry colors, glow sizes, border settings...
            foreach (var tp in t.Values)
            {
                if (!srcOfPart.ContainsKey(tp.Name) && b.ContainsKey(tp.Name)) srcOfPart[tp.Name] = tp.Name;
            }

            // property merge per part, with rollback when the merged geometry doesn't fit the sprite.
            foreach (var tp in t.Values.OrderBy(p => p.Id))
            {
                if (!srcOfPart.TryGetValue(tp.Name, out var src)) continue;
                if (plans.Any(p => p.Part == tp && p.Method == "fallback"))
                {
                    // the template sprite stays, so its margins/frame counts must stay with it.
                    report.Notes.Add($"DWMWindow {tp.Name}: template sprite kept, template properties kept with it.");
                    continue;
                }
                b.TryGetValue(src, out var bp);
                s.TryGetValue(src, out var sp);
                if (bp == null || sp == null) continue;

                var partPlans = plans.Where(p => p.Part == tp && p.Method != null && p.Method != "fallback" && p.Method != "none").ToList();
                MergePart(tp, bp, sp, src, partPlans.Count > 0);

                // the sprite came over with its own frame count and margins, so whatever size it has,
                // it's described the way the source described it. an odd height is fine for the editor
                // (it slices proportionally, 129px in 2 frames is 64 + 65); how uDWM 19041 rounds that
                // isn't verified. a frame with nothing left after uDWM's 1px inset, or margins wider than
                // the frame, only gets a warning, since the source shipped the same numbers on win7.
                foreach (var plan in partPlans)
                {
                    if (GeometryFits(tp, plan.Entry.StateId, plan.SourceRect, out string why)) continue;
                    report.Warnings.Add($"DWMWindow {tp.Name}: {why}; same as the source ships it, check this part in the VM.");
                    m_imp.GeometryWaivers.Add(tp.Name);
                    report.Count("dwm-geometry-warning");
                }
            }

            BuildAtlas(plans, ta, sa, tAtlas, sAtlas);

            var usedSrc = new HashSet<string>(srcOfPart.Values, StringComparer.OrdinalIgnoreCase);
            foreach (var bp in b.Values.Where(p => p.Id != 0 && !usedSrc.Contains(p.Name)).OrderBy(p => p.Id))
            {
                m_rep.Win7OnlyParts.Add(bp.Name);
                bool changed = s.TryGetValue(bp.Name, out var sp2) && bp.Entries.Any(e =>
                {
                    var se = sp2.Find(e.StateId, e.PropertyName, e.TypeId, e.Occurrence);
                    return se != null && e.NameId != (int)IDENTIFIER.ATLASRECT && !m_imp.Same(m_imp.Base, e, m_imp.Source, se);
                });
                if (changed)
                    report.Notes.Add($"DWMWindow {bp.Name}: source customized a win7-only part (no win10 consumer); dropped.");
            }
        }

        private void ResolveSource(RectPlan plan, Dictionary<string, PartData> b, Dictionary<string, PartData> s, ArgbImage sAtlas)
        {
            int state = plan.Entry.StateId;
            if (KeyPattern.Matches(m_imp.Options.Exclude, Win7StyleImporter.DwmClass, plan.Part.Name, "ATLASRECT"))
            {
                plan.Method = "fallback";
                plan.Note = "excluded by rule, template sprite kept";
                return;
            }
            if (b.TryGetValue(plan.Part.Name, out var bp))
            {
                var br = bp.Rects.FirstOrDefault(r => r.StateId == state);
                if (br != null && RectOf(br) == plan.TargetRect)
                {
                    plan.Src = bp.Name;
                    plan.Method = "name";
                }
            }

            if (plan.Src == null)
            {
                var candidates = new List<string>();
                foreach (var bp2 in b.Values)
                    foreach (var r in bp2.Rects)
                        if (r.StateId == state && RectOf(r) == plan.TargetRect) candidates.Add(bp2.Name);

                var withSource = candidates.Where(n => s.TryGetValue(n, out var sp) && sp.Rects.Any(r => r.StateId == state)).OrderBy(n => n).ToList();
                if (withSource.Count > 0)
                {
                    string chosen = ChooseAlias(withSource, n => SafeCrop(sAtlas, RectOf(s[n].Rects.First(r => r.StateId == state))));
                    if (chosen != null)
                    {
                        plan.Src = chosen;
                        plan.Method = "alias";
                        plan.Note = (b.ContainsKey(plan.Part.Name) ? "template moved this part onto the sprite of " : "win10-only part reuses the stock sprite of ")
                            + string.Join("/", withSource);
                        if (b.ContainsKey(plan.Part.Name))
                        {
                            // e.g. aero11 points SMALLBUTTONCLOSEGLYPH192 at the restore glyph. kept, since the
                            // template is the authority, but it smells like a template bug worth a look.
                            m_imp.Report.Warnings.Add($"DWMWindow {plan.Part.Name}: template uses the stock sprite of {string.Join("/", withSource)} instead of its own; the output keeps that pairing with the source's {withSource[0]} sprite.");
                        }
                    }
                    else
                    {
                        plan.Method = "fallback";
                        plan.Note = "ambiguous: stock rect shared by " + string.Join("/", withSource) + " and the source gave them different sprites";
                        return;
                    }
                }
            }

            if (plan.Src == null)
            {
                plan.Method = "fallback";
                plan.Note = "no win7 part owns this rect in the stock atlas";
                return;
            }

            var srcRect = s.TryGetValue(plan.Src, out var spart) ? spart.Rects.FirstOrDefault(r => r.StateId == state) : null;
            if (srcRect == null)
            {
                plan.Method = "fallback";
                plan.Note = "source has no ATLASRECT for " + plan.Src;
                plan.Src = null;
                return;
            }
            plan.SourceRect = RectOf(srcRect);
            if (SafeCrop(sAtlas, plan.SourceRect) == null || plan.SourceRect.Width <= 0 || plan.SourceRect.Height <= 0)
            {
                plan.Method = "fallback";
                plan.Note = $"source rect {plan.SourceRect} is outside the source atlas";
                plan.Src = null;
            }
        }

        private static ArgbImage SafeCrop(ArgbImage img, Rectangle r)
        {
            if (r.X < 0 || r.Y < 0 || r.Right > img.Width || r.Bottom > img.Height || r.Width < 0 || r.Height < 0) return null;
            return img.Crop(r);
        }

        /// <summary>
        /// Several win7 parts can share one stock sprite. Following the shared rect is only safe when
        /// the source also gave all of them the same pixels; otherwise there's no telling which one
        /// the template meant. Returns the first name (ordinal order) or null when ambiguous.
        /// </summary>
        internal static string ChooseAlias(IList<string> candidates, Func<string, ArgbImage> sourceSprite)
        {
            var ordered = candidates.OrderBy(n => n, StringComparer.Ordinal).ToList();
            if (ordered.Count == 0) return null;
            var sprites = ordered.Select(sourceSprite).ToList();
            if (sprites.Any(x => x == null)) return null;
            return sprites.Skip(1).All(x => Win7StyleImporter.NearlyEqual(x, sprites[0], 0)) ? ordered[0] : null;
        }

        // what describes the pixels of an atlas sprite: frame count, nine-grid and content padding.
        private static readonly HashSet<int> SpriteGeometry = new HashSet<int>
        {
            (int)IDENTIFIER.IMAGECOUNT, (int)IDENTIFIER.SIZINGMARGINS, (int)IDENTIFIER.CONTENTMARGINS,
        };

        /// <param name="spriteFromSource">the part's ATLASRECT now points at the source's sprite.</param>
        private void MergePart(PartData tp, PartData bp, PartData sp, string src, bool spriteFromSource)
        {
            string label = tp.Name + (src != tp.Name ? " (from " + src + ")" : "");
            foreach (var te in tp.Entries)
            {
                if (te.NameId == (int)IDENTIFIER.ATLASRECT && te.TypeId == (int)IDENTIFIER.RECTTYPE) continue;
                if (te.TypeId == (int)IDENTIFIER.DISKSTREAM) continue;
                var be = bp.Find(te.StateId, te.PropertyName, te.TypeId, te.Occurrence);
                var se = sp.Find(te.StateId, te.PropertyName, te.TypeId, te.Occurrence);
                if (be == null && se == null) continue;

                // the sprite comes with the numbers that describe it, like replacing a part in the
                // editor: even where the template adapted them, they belonged to the template's sprite.
                if (spriteFromSource && se != null && SpriteGeometry.Contains(te.NameId) && !m_imp.Same(m_imp.Source, se, m_imp.Target, te))
                {
                    if (m_imp.ApplySourceValue(m_imp.Source, se, te.Property, out string err))
                    {
                        m_imp.MarkTouched(te.Property, $"{te.ClassName}|{te.PartName}|S{te.StateId}|{te.PropertyName}");
                        m_imp.Decide(te, "applied-with-sprite", "describes the source sprite this part now uses", be, se, te, SemanticModel.FormatValue(te.Property));
                        m_imp.Report.Count("applied-with-sprite");
                    }
                    else m_imp.Report.Warnings.Add($"DWMWindow {tp.Name}: {te.PropertyName} not applied, {err}");
                    continue;
                }
                m_imp.ThreeWay(Win7StyleImporter.DwmClass, label, te.StateId, te.PropertyName, be, se, te, m_imp.Source, true);
            }
            foreach (var se in sp.Entries)
            {
                if (se.NameId == (int)IDENTIFIER.ATLASRECT && se.TypeId == (int)IDENTIFIER.RECTTYPE) continue;
                if (se.TypeId == (int)IDENTIFIER.DISKSTREAM) continue;
                if (tp.Find(se.StateId, se.PropertyName, se.TypeId, se.Occurrence) != null) continue;
                var be = bp.Find(se.StateId, se.PropertyName, se.TypeId, se.Occurrence);
                m_imp.ThreeWay(Win7StyleImporter.DwmClass, tp.Name, se.StateId, se.PropertyName, be, se, null, m_imp.Source, true);
            }
        }

        private bool GeometryFits(PartData tp, int state, Rectangle rect, out string why)
        {
            var style = m_imp.Target.Style;
            var cls = style.Classes[tp.Entries[0].ClassId];
            return CheckGeometry(rect, EffectiveCount(style, cls, tp.Id, state), EffectiveSizing(style, cls, tp.Id, state), out why);
        }

        internal static int EffectiveCount(VisualStyle style, StyleClass cls, int part, int state)
        {
            return SemanticModel.FindEffective(style, cls, part, state, IDENTIFIER.IMAGECOUNT)?.GetValue() is int n && n > 0 ? n : 1;
        }

        internal static Margins EffectiveSizing(VisualStyle style, StyleClass cls, int part, int state)
        {
            return SemanticModel.FindEffective(style, cls, part, state, IDENTIFIER.SIZINGMARGINS)?.GetValue() as Margins;
        }

        /// <summary>
        /// uDWM takes the atlas rect, drops a 1px border on every side and cuts it into IMAGECOUNT
        /// vertical strips (4 for buttons/glyphs, 2 for frames), each nine-gridded with SIZINGMARGINS.
        /// Strips are taken as sliced proportionally, like the editor does (129px in 2 = 64 + 65; how
        /// uDWM rounds isn't verified), so only a strip left empty by the inset or margins wider than
        /// the smallest strip count as a problem.
        /// </summary>
        internal static bool CheckGeometry(Rectangle rect, int imageCount, Margins sizing, out string why)
        {
            why = null;
            if (imageCount < 1) imageCount = 1;
            int fw = rect.Width - 2, fh = rect.Height / imageCount - 2;
            if (fw < 1 || fh < 1)
            {
                why = $"rect {rect.Width}x{rect.Height} leaves empty {imageCount}-frame strips after the 1px inset";
                return false;
            }
            if (sizing != null && (sizing.Left + sizing.Right > fw || sizing.Top + sizing.Bottom > fh))
            {
                why = $"sizing margins l{sizing.Left} r{sizing.Right} t{sizing.Top} b{sizing.Bottom} don't fit a {fw}x{fh} frame";
                return false;
            }
            return true;
        }

        private void BuildAtlas(List<RectPlan> plans, SemanticEntry ta, SemanticEntry sa, ArgbImage tAtlas, ArgbImage sAtlas)
        {
            var report = m_imp.Report;
            var fallbacks = plans.Where(p => p.Method == "fallback").ToList();

            // shelf-pack the kept template sprites under the source atlas. identical rects share a slot.
            var slots = new Dictionary<Rectangle, Rectangle>();
            int width = sAtlas.Width, y = sAtlas.Height, x = 0, rowH = 0;
            foreach (var fr in fallbacks.Select(p => p.TargetRect).Distinct().OrderByDescending(r => r.Height))
            {
                width = Math.Max(width, fr.Width);
                if (x + fr.Width > width) { y += rowH; x = 0; rowH = 0; }
                slots[fr] = new Rectangle(x, y, fr.Width, fr.Height);
                x += fr.Width;
                rowH = Math.Max(rowH, fr.Height);
            }
            int height = y + rowH;

            byte[] atlasBytes;
            ArgbImage outAtlas;
            if (slots.Count == 0)
            {
                // nothing to add: take the source atlas bytes as they are, no recompression.
                atlasBytes = m_imp.Source.GetResourceBytes("STREAM", sa.ResourceId);
                outAtlas = sAtlas;
            }
            else
            {
                outAtlas = new ArgbImage(width, height);
                outAtlas.Blit(sAtlas, 0, 0);
                foreach (var kv in slots) outAtlas.Blit(tAtlas.Crop(kv.Key), kv.Value.X, kv.Value.Y);
                atlasBytes = ImageIdentity.EncodePngVerified(outAtlas);
                m_rep.AtlasReencoded = true;
            }
            m_rep.NewAtlasSize = $"{outAtlas.Width}x{outAtlas.Height}";

            bool atlasChanged = !Win7StyleImporter.NearlyEqual(outAtlas, tAtlas, 0);
            int atlasId = ta.ResourceId;
            if (atlasChanged)
            {
                var info = ImageIdentity.Analyze(atlasBytes);
                atlasId = m_imp.PlaceResource("STREAM", sa.ResourceId, atlasBytes, info);
                ta.Property.SetValue(atlasId);
                m_imp.MarkTouched(ta.Property, $"{ta.ClassName}|{ta.PartName}|S{ta.StateId}|{ta.PropertyName}");
                m_imp.Decide(ta, "applied", "DWM atlas replaced by the source atlas" + (slots.Count > 0 ? $" plus {slots.Count} kept template sprites" : ""), null, null, ta, atlasId.ToString());
            }
            m_rep.NewAtlasId = atlasId;

            foreach (var p in plans)
            {
                var map = new DwmPartMapping
                {
                    TargetPartId = p.Part.Id,
                    TargetPart = p.Part.Name + (p.Entry.StateId != 0 ? " S" + p.Entry.StateId : ""),
                    SourcePart = p.Src,
                    Method = p.Method,
                    TargetRect = p.TargetRect.ToString(),
                    Note = p.Note
                };
                m_rep.Parts.Add(map);
                if (p.Method == "none") continue;

                Rectangle newRect;
                ArgbImage expected;
                if (p.Method == "fallback")
                {
                    newRect = slots[p.TargetRect];
                    expected = tAtlas.Crop(p.TargetRect);
                    m_rep.FallbackParts.Add(p.Part.Name);
                }
                else
                {
                    newRect = p.SourceRect;
                    expected = sAtlas.Crop(p.SourceRect);
                }
                map.NewRect = newRect.ToString();
                m_imp.SpriteChecks.Add(new SpriteExpectation { Part = p.Part.Name, OutputRect = newRect, Expected = expected });

                if (newRect != p.TargetRect)
                {
                    p.Entry.Property.SetValue(new Margins(newRect.Left, newRect.Top, newRect.Right, newRect.Bottom));
                    m_imp.MarkTouched(p.Entry.Property, $"DWMWindow|{p.Part.Name}|S{p.Entry.StateId}|ATLASRECT");
                }
                report.Count("dwm-rect-" + p.Method);
            }
        }
    }
}
