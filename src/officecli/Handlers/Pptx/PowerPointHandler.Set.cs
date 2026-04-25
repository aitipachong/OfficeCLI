// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeCli.Core;
using Drawing = DocumentFormat.OpenXml.Drawing;

namespace OfficeCli.Handlers;

public partial class PowerPointHandler
{
    public List<string> Set(string path, Dictionary<string, string> properties)
    {
        path = NormalizeCellPath(path);
        path = ResolveIdPath(path);

        // Batch Set: if path looks like a selector (not starting with /), Query → Set each
        if (!string.IsNullOrEmpty(path) && !path.StartsWith("/"))
        {
            var unsupported = new List<string>();
            var targets = Query(path);
            if (targets.Count == 0)
                throw new ArgumentException($"No elements matched selector: {path}");
            foreach (var target in targets)
            {
                var targetUnsupported = Set(target.Path, properties);
                foreach (var u in targetUnsupported)
                    if (!unsupported.Contains(u)) unsupported.Add(u);
            }
            return unsupported;
        }

        if (path.Equals("/theme", StringComparison.OrdinalIgnoreCase))
            return SetThemeProperties(properties);

        // Unified find: if 'find' key is present, route to ProcessPptFind
        if (properties.TryGetValue("find", out var findText))
        {
            var replace = properties.TryGetValue("replace", out var r) ? r : null;
            var formatProps = new Dictionary<string, string>(properties, StringComparer.OrdinalIgnoreCase);
            formatProps.Remove("find");
            formatProps.Remove("replace");
            formatProps.Remove("scope");
            formatProps.Remove("regex");

            if (replace == null && formatProps.Count == 0)
                throw new ArgumentException("'find' requires either 'replace' and/or format properties (e.g. bold, color, size).");

            // Support regex=true as an alternative to r"..." prefix.
            // CONSISTENCY(find-regex): mirror of WordHandler.Set.cs:60-61. grep
            // "CONSISTENCY(find-regex)" for every project-wide call site.
            if (properties.TryGetValue("regex", out var regexFlag) && ParseHelpers.IsTruthySafe(regexFlag) && !findText.StartsWith("r\"") && !findText.StartsWith("r'"))
                findText = $"r\"{findText}\"";

            var matchCount = ProcessPptFind(path, findText, replace, formatProps);
            LastFindMatchCount = matchCount;
            return [];
        }

        // Presentation-level properties: / or /presentation
        if (path is "/" or "" or "/presentation")
        {

            var presentation = _doc.PresentationPart?.Presentation
                ?? throw new InvalidOperationException("No presentation");
            var unsupported = new List<string>();
            foreach (var (key, value) in properties)
            {
                switch (key.ToLowerInvariant())
                {
                    case "slidewidth" or "width":
                        var sldSz = presentation.GetFirstChild<SlideSize>()
                            ?? presentation.AppendChild(new SlideSize());
                        sldSz.Cx = Core.EmuConverter.ParseEmuAsInt(value);
                        sldSz.Type = SlideSizeValues.Custom;
                        break;
                    case "slideheight" or "height":
                        var sldSz2 = presentation.GetFirstChild<SlideSize>()
                            ?? presentation.AppendChild(new SlideSize());
                        sldSz2.Cy = Core.EmuConverter.ParseEmuAsInt(value);
                        sldSz2.Type = SlideSizeValues.Custom;
                        break;
                    case "slidesize":
                        var sz = presentation.GetFirstChild<SlideSize>()
                            ?? presentation.AppendChild(new SlideSize());
                        if (SlideSizeDefaults.Presets.TryGetValue(value, out var preset))
                        {
                            sz.Cx = (int)preset.Cx;
                            sz.Cy = (int)preset.Cy;
                            sz.Type = preset.Type;
                        }
                        else
                        {
                            unsupported.Add(key);
                        }
                        break;
                    // Core document properties
                    case "title":
                        _doc.PackageProperties.Title = value;
                        break;
                    case "author" or "creator":
                        _doc.PackageProperties.Creator = value;
                        break;
                    case "subject":
                        _doc.PackageProperties.Subject = value;
                        break;
                    case "description":
                        _doc.PackageProperties.Description = value;
                        break;
                    case "category":
                        _doc.PackageProperties.Category = value;
                        break;
                    case "keywords":
                        _doc.PackageProperties.Keywords = value;
                        break;
                    case "lastmodifiedby":
                        _doc.PackageProperties.LastModifiedBy = value;
                        break;
                    case "revision":
                        _doc.PackageProperties.Revision = value;
                        break;
                    case "defaultfont" or "font":
                    {
                        var masterPart = _doc.PresentationPart?.SlideMasterParts?.FirstOrDefault();
                        var theme = masterPart?.ThemePart?.Theme;
                        var fontScheme = theme?.ThemeElements?.FontScheme;
                        if (fontScheme != null)
                        {
                            if (fontScheme.MajorFont?.LatinFont != null)
                                fontScheme.MajorFont.LatinFont.Typeface = value;
                            if (fontScheme.MinorFont?.LatinFont != null)
                                fontScheme.MinorFont.LatinFont.Typeface = value;
                            masterPart!.ThemePart!.Theme!.Save();
                        }
                        break;
                    }
                    default:
                        var lowerKey = key.ToLowerInvariant();
                        if (!TrySetPresentationSetting(lowerKey, value)
                            && !Core.ThemeHandler.TrySetTheme(
                                _doc.PresentationPart?.SlideMasterParts?.FirstOrDefault()?.ThemePart, lowerKey, value)
                            && !Core.ExtendedPropertiesHandler.TrySetExtendedProperty(
                                Core.ExtendedPropertiesHandler.GetOrCreateExtendedPart(_doc), lowerKey, value))
                        {
                            if (unsupported.Count == 0)
                                unsupported.Add($"{key} (valid presentation props: slideWidth, slideHeight, slideSize, title, author, defaultFont, firstSlideNum, rtl, compatMode, print.*, show.*)");
                            else
                                unsupported.Add(key);
                        }
                        break;
                }
            }
            presentation.Save();
            return unsupported;
        }

        // Try slidemaster/slidelayout bg-aware path first (case-insensitive):
        // /slidemaster[N], /slidemaster[N]/slidelayout[M], /slidelayout[N]
        // Handles background and name props. Falls through for shape-nested paths.
        {
            var masterBgMatch = Regex.Match(path, @"^/slidemaster\[(\d+)\](?:/slidelayout\[(\d+)\])?$", RegexOptions.IgnoreCase);
            var layoutBgMatch = Regex.Match(path, @"^/slidelayout\[(\d+)\]$", RegexOptions.IgnoreCase);
            if (masterBgMatch.Success || layoutBgMatch.Success)
                return SetMasterOrLayoutBackgroundByPath(masterBgMatch, layoutBgMatch, properties);
        }

        // Try slideMaster/slideLayout shape editing: /slideMaster[N]/shape[M] or /slideLayout[N]/shape[M]
        var masterShapeMatch = Regex.Match(path, @"^/(slideMaster|slideLayout)\[(\d+)\](?:/(\w+)\[(\d+)\])?$");
        if (masterShapeMatch.Success) return SetMasterShapeByPath(masterShapeMatch, properties);

        // Try notes path: /slide[N]/notes
        var notesSetMatch = Regex.Match(path, @"^/slide\[(\d+)\]/notes$");
        if (notesSetMatch.Success) return SetNotesByPath(notesSetMatch, properties);

        // Try run-level path: /slide[N]/shape[M]/run[K]
        var runMatch = Regex.Match(path, @"^/slide\[(\d+)\]/shape\[(\d+)\]/run\[(\d+)\]$");
        if (runMatch.Success) return SetShapeRunByPath(runMatch, properties);

        // Try paragraph/run path: /slide[N]/shape[M]/paragraph[P]/run[K]
        var paraRunMatch = Regex.Match(path, @"^/slide\[(\d+)\]/shape\[(\d+)\]/paragraph\[(\d+)\]/run\[(\d+)\]$");
        if (paraRunMatch.Success) return SetParagraphRunByPath(paraRunMatch, properties);

        // Try paragraph-level path: /slide[N]/shape[M]/paragraph[P]
        var paraMatch = Regex.Match(path, @"^/slide\[(\d+)\]/shape\[(\d+)\]/paragraph\[(\d+)\]$");
        if (paraMatch.Success) return SetParagraphByPath(paraMatch, properties);

        // Try chart axis-by-role sub-path: /slide[N]/chart[M]/axis[@role=ROLE].
        // Routed separately from the chart[]/series[] path because the role capture
        // needs to drive a different forwarder (SetAxisProperties, not series-prefix).
        var chartAxisSetMatch = Regex.Match(path,
            @"^/slide\[(\d+)\]/chart\[(\d+)\]/axis\[@role=([a-zA-Z0-9_]+)\]$");
        if (chartAxisSetMatch.Success) return SetChartAxisByPath(chartAxisSetMatch, properties);

        // Try chart path: /slide[N]/chart[M] or /slide[N]/chart[M]/series[K]
        var chartSetMatch = Regex.Match(path, @"^/slide\[(\d+)\]/chart\[(\d+)\](?:/series\[(\d+)\])?$");
        if (chartSetMatch.Success) return SetChartByPath(chartSetMatch, properties);

        // Try table cell path: /slide[N]/table[M]/tr[R]/tc[C]
        var tblCellMatch = Regex.Match(path, @"^/slide\[(\d+)\]/table\[(\d+)\]/tr\[(\d+)\]/tc\[(\d+)\]$");
        if (tblCellMatch.Success) return SetTableCellByPath(tblCellMatch, properties);

        // Try table-level path: /slide[N]/table[M]
        var tblMatch = Regex.Match(path, @"^/slide\[(\d+)\]/table\[(\d+)\]$");
        if (tblMatch.Success) return SetTableByPath(tblMatch, properties);

        // Try table row path: /slide[N]/table[M]/tr[R]
        var tblRowMatch = Regex.Match(path, @"^/slide\[(\d+)\]/table\[(\d+)\]/tr\[(\d+)\]$");
        if (tblRowMatch.Success) return SetTableRowByPath(tblRowMatch, properties);

        // Try placeholder path: /slide[N]/placeholder[M] or /slide[N]/placeholder[type]
        var phMatch = Regex.Match(path, @"^/slide\[(\d+)\]/placeholder\[(\w+)\]$");
        if (phMatch.Success) return SetPlaceholderByPath(phMatch, properties);

        // Try video/audio path: /slide[N]/video[M] or /slide[N]/audio[M]
        var mediaSetMatch = Regex.Match(path, @"^/slide\[(\d+)\]/(video|audio)\[(\d+)\]$");
        if (mediaSetMatch.Success) return SetMediaByPath(mediaSetMatch, properties);

        // Try picture path: /slide[N]/picture[M] or /slide[N]/pic[M]
        // OLE set path: /slide[N]/ole[M]
        // Replace backing embedded part + refresh ProgID automatically
        // when the extension changes. Cleans up the old part to avoid
        // storage bloat (mirrors picture path clean-up).
        var oleSetMatch = Regex.Match(path, @"^/slide\[(\d+)\]/(?:ole|object|embed)\[(\d+)\]$");
        if (oleSetMatch.Success) return SetOleByPath(oleSetMatch, properties);

        var picSetMatch = Regex.Match(path, @"^/slide\[(\d+)\]/(?:picture|pic)\[(\d+)\]$");
        if (picSetMatch.Success) return SetPictureByPath(picSetMatch, properties);

        // Try slide-level path: /slide[N]
        var slideOnlyMatch = Regex.Match(path, @"^/slide\[(\d+)\]$");
        if (slideOnlyMatch.Success) return SetSlideByPath(slideOnlyMatch, properties);

        // Try model3d-level path: /slide[N]/model3d[M]
        var model3dSetMatch = Regex.Match(path, @"^/slide\[(\d+)\]/model3d\[(\d+)\]$");
        if (model3dSetMatch.Success) return SetModel3DByPath(model3dSetMatch, properties);

        // Try zoom-level path: /slide[N]/zoom[M]
        var zoomSetMatch = Regex.Match(path, @"^/slide\[(\d+)\]/zoom\[(\d+)\]$");
        if (zoomSetMatch.Success) return SetZoomByPath(zoomSetMatch, properties);

        // Try shape-level path: /slide[N]/shape[M]
        var match = Regex.Match(path, @"^/slide\[(\d+)\]/shape\[(\d+)\]$");
        if (match.Success) return SetShapeByPath(match, properties);

        // Try connector path: /slide[N]/connector[M] or /slide[N]/connection[M]
        var cxnMatch = Regex.Match(path, @"^/slide\[(\d+)\]/(?:connector|connection)\[(\d+)\]$");
        if (cxnMatch.Success) return SetConnectorByPath(cxnMatch, properties);

        // Try group path: /slide[N]/group[M]
        var grpMatch = Regex.Match(path, @"^/slide\[(\d+)\]/group\[(\d+)\]$");
        if (grpMatch.Success) return SetGroupByPath(grpMatch, properties);

        // Generic XML fallback: navigate to element and set attributes
        {
            SlidePart fbSlidePart;
            OpenXmlElement target;

            // Try logical path resolution first (table/placeholder paths)
            var logicalResult = ResolveLogicalPath(path);
            if (logicalResult.HasValue)
            {
                fbSlidePart = logicalResult.Value.slidePart;
                target = logicalResult.Value.element;
            }
            else
            {
                var allSegments = GenericXmlQuery.ParsePathSegments(path);
                if (allSegments.Count == 0 || !allSegments[0].Name.Equals("slide", StringComparison.OrdinalIgnoreCase) || !allSegments[0].Index.HasValue)
                    throw new ArgumentException($"Path must start with /slide[N]: {path}");

                var fbSlideIdx = allSegments[0].Index!.Value;
                var fbSlideParts = GetSlideParts().ToList();
                if (fbSlideIdx < 1 || fbSlideIdx > fbSlideParts.Count)
                    throw new ArgumentException($"Slide {fbSlideIdx} not found (total: {fbSlideParts.Count})");

                fbSlidePart = fbSlideParts[fbSlideIdx - 1];
                var remaining = allSegments.Skip(1).ToList();
                target = GetSlide(fbSlidePart);
                if (remaining.Count > 0)
                {
                    target = GenericXmlQuery.NavigateByPath(target, remaining)
                        ?? throw new ArgumentException($"Element not found: {path}");
                }
            }

            var unsup = new List<string>();
            foreach (var (key, value) in properties)
            {
                if (!GenericXmlQuery.SetGenericAttribute(target, key, value))
                    unsup.Add(key);
            }
            GetSlide(fbSlidePart).Save();
            return unsup;
        }
    }

    // ==================== Per-element-type Set helpers ====================
    // Mechanical extractions from the original god-method Set(); each helper
    // owns one path-pattern's full handling. Splitting was for AI-readability
    // (each helper now <100 lines, fits in one Read) — no behavior change.

    private List<string> SetShapeRunByPath(Match runMatch, Dictionary<string, string> properties)
    {
        var slideIdx = int.Parse(runMatch.Groups[1].Value);
        var shapeIdx = int.Parse(runMatch.Groups[2].Value);
        var runIdx = int.Parse(runMatch.Groups[3].Value);

        var (slidePart, shape) = ResolveShape(slideIdx, shapeIdx);
        var allRuns = GetAllRuns(shape);
        if (runIdx < 1 || runIdx > allRuns.Count)
            throw new ArgumentException($"Run {runIdx} not found (shape has {allRuns.Count} runs)");

        var targetRun = allRuns[runIdx - 1];
        var linkValRun = properties.GetValueOrDefault("link");
        var tooltipValRun = properties.GetValueOrDefault("tooltip");
        var runOnlyProps = properties
            .Where(kv => !kv.Key.Equals("link", StringComparison.OrdinalIgnoreCase)
                      && !kv.Key.Equals("tooltip", StringComparison.OrdinalIgnoreCase))
            .ToDictionary(kv => kv.Key, kv => kv.Value);
        var unsupported = SetRunOrShapeProperties(runOnlyProps, new List<Drawing.Run> { targetRun }, shape, slidePart);
        if (linkValRun != null) ApplyRunHyperlink(slidePart, targetRun, linkValRun, tooltipValRun);
        GetSlide(slidePart).Save();
        return unsupported;
    }

    private List<string> SetParagraphRunByPath(Match paraRunMatch, Dictionary<string, string> properties)
    {
        var slideIdx = int.Parse(paraRunMatch.Groups[1].Value);
        var shapeIdx = int.Parse(paraRunMatch.Groups[2].Value);
        var paraIdx = int.Parse(paraRunMatch.Groups[3].Value);
        var runIdx = int.Parse(paraRunMatch.Groups[4].Value);

        var (slidePart, shape) = ResolveShape(slideIdx, shapeIdx);
        var paragraphs = shape.TextBody?.Elements<Drawing.Paragraph>().ToList()
            ?? throw new ArgumentException("Shape has no text body");
        if (paraIdx < 1 || paraIdx > paragraphs.Count)
            throw new ArgumentException($"Paragraph {paraIdx} not found (shape has {paragraphs.Count} paragraphs)");

        var para = paragraphs[paraIdx - 1];
        var paraRuns = para.Elements<Drawing.Run>().ToList();
        if (runIdx < 1 || runIdx > paraRuns.Count)
            throw new ArgumentException($"Run {runIdx} not found (paragraph has {paraRuns.Count} runs)");

        var targetRun = paraRuns[runIdx - 1];
        var linkVal = properties.GetValueOrDefault("link");
        var tooltipVal = properties.GetValueOrDefault("tooltip");
        var runOnlyProps = properties
            .Where(kv => !kv.Key.Equals("link", StringComparison.OrdinalIgnoreCase)
                      && !kv.Key.Equals("tooltip", StringComparison.OrdinalIgnoreCase))
            .ToDictionary(kv => kv.Key, kv => kv.Value);
        var unsupported = SetRunOrShapeProperties(runOnlyProps, new List<Drawing.Run> { targetRun }, shape, slidePart);
        if (linkVal != null) ApplyRunHyperlink(slidePart, targetRun, linkVal, tooltipVal);
        GetSlide(slidePart).Save();
        return unsupported;
    }


    private List<string> SetParagraphByPath(Match paraMatch, Dictionary<string, string> properties)
    {
        var slideIdx = int.Parse(paraMatch.Groups[1].Value);
        var shapeIdx = int.Parse(paraMatch.Groups[2].Value);
        var paraIdx = int.Parse(paraMatch.Groups[3].Value);

        var (slidePart, shape) = ResolveShape(slideIdx, shapeIdx);
        var paragraphs = shape.TextBody?.Elements<Drawing.Paragraph>().ToList()
            ?? throw new ArgumentException("Shape has no text body");
        if (paraIdx < 1 || paraIdx > paragraphs.Count)
            throw new ArgumentException($"Paragraph {paraIdx} not found (shape has {paragraphs.Count} paragraphs)");

        var para = paragraphs[paraIdx - 1];
        var paraRuns = para.Elements<Drawing.Run>().ToList();
        var unsupported = new List<string>();

        foreach (var (key, value) in properties)
        {
            switch (key.ToLowerInvariant())
            {
                case "align":
                {
                    var pProps = para.ParagraphProperties ?? (para.ParagraphProperties = new Drawing.ParagraphProperties());
                    pProps.Alignment = ParseTextAlignment(value);
                    break;
                }
                case "indent":
                {
                    var pProps = para.ParagraphProperties ?? (para.ParagraphProperties = new Drawing.ParagraphProperties());
                    pProps.Indent = (int)ParseEmu(value);
                    break;
                }
                case "level":
                {
                    var pProps = para.ParagraphProperties ?? (para.ParagraphProperties = new Drawing.ParagraphProperties());
                    if (!int.TryParse(value, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out var lvl) || lvl < 0 || lvl > 8)
                        throw new ArgumentException($"Invalid 'level' value: '{value}'. Expected an integer between 0 and 8 (OOXML a:pPr/@lvl).");
                    pProps.Level = lvl;
                    break;
                }
                case "marginleft" or "marl":
                {
                    var pProps = para.ParagraphProperties ?? (para.ParagraphProperties = new Drawing.ParagraphProperties());
                    pProps.LeftMargin = (int)ParseEmu(value);
                    break;
                }
                case "marginright" or "marr":
                {
                    var pProps = para.ParagraphProperties ?? (para.ParagraphProperties = new Drawing.ParagraphProperties());
                    pProps.RightMargin = (int)ParseEmu(value);
                    break;
                }
                case "linespacing" or "line.spacing":
                {
                    var pProps = para.ParagraphProperties ?? (para.ParagraphProperties = new Drawing.ParagraphProperties());
                    pProps.RemoveAllChildren<Drawing.LineSpacing>();
                    var (lsVal2, lsIsPercent) = SpacingConverter.ParsePptLineSpacing(value);
                    if (lsIsPercent)
                        pProps.AppendChild(new Drawing.LineSpacing(
                            new Drawing.SpacingPercent { Val = lsVal2 }));
                    else
                        pProps.AppendChild(new Drawing.LineSpacing(
                            new Drawing.SpacingPoints { Val = lsVal2 }));
                    break;
                }
                case "spacebefore" or "space.before":
                {
                    var pProps = para.ParagraphProperties ?? (para.ParagraphProperties = new Drawing.ParagraphProperties());
                    pProps.RemoveAllChildren<Drawing.SpaceBefore>();
                    pProps.AppendChild(new Drawing.SpaceBefore(new Drawing.SpacingPoints { Val = SpacingConverter.ParsePptSpacing(value) }));
                    break;
                }
                case "spaceafter" or "space.after":
                {
                    var pProps = para.ParagraphProperties ?? (para.ParagraphProperties = new Drawing.ParagraphProperties());
                    pProps.RemoveAllChildren<Drawing.SpaceAfter>();
                    pProps.AppendChild(new Drawing.SpaceAfter(new Drawing.SpacingPoints { Val = SpacingConverter.ParsePptSpacing(value) }));
                    break;
                }
                case "link":
                {
                    var paraTooltip = properties.GetValueOrDefault("tooltip");
                    foreach (var r in paraRuns)
                        ApplyRunHyperlink(slidePart, r, value, paraTooltip);
                    break;
                }
                case "tooltip":
                    // handled in tandem with "link"; standalone tooltip change is not supported here
                    break;
                default:
                    // Apply run-level properties to all runs in this paragraph
                    var runUnsup = SetRunOrShapeProperties(
                        new Dictionary<string, string> { { key, value } }, paraRuns, shape, slidePart);
                    unsupported.AddRange(runUnsup);
                    break;
            }
        }

        GetSlide(slidePart).Save();
        return unsupported;
    }



    private List<string> SetPlaceholderByPath(Match phMatch, Dictionary<string, string> properties)
    {
        var slideIdx = int.Parse(phMatch.Groups[1].Value);
        var phId = phMatch.Groups[2].Value;

        var slideParts2 = GetSlideParts().ToList();
        if (slideIdx < 1 || slideIdx > slideParts2.Count)
            throw new ArgumentException($"Slide {slideIdx} not found (total: {slideParts2.Count})");
        var slidePart = slideParts2[slideIdx - 1];
        var shape = ResolvePlaceholderShape(slidePart, phId);

        var allRuns = shape.Descendants<Drawing.Run>().ToList();
        var unsupported = SetRunOrShapeProperties(properties, allRuns, shape, slidePart);
        GetSlide(slidePart).Save();
        return unsupported;
    }

    private List<string> SetGroupByPath(Match grpMatch, Dictionary<string, string> properties)
    {
        var slideIdx = int.Parse(grpMatch.Groups[1].Value);
        var grpIdx = int.Parse(grpMatch.Groups[2].Value);

        var slideParts6 = GetSlideParts().ToList();
        if (slideIdx < 1 || slideIdx > slideParts6.Count)
            throw new ArgumentException($"Slide {slideIdx} not found (total: {slideParts6.Count})");

        var slidePart = slideParts6[slideIdx - 1];
        var shapeTree = GetSlide(slidePart).CommonSlideData?.ShapeTree
            ?? throw new ArgumentException("Slide has no shape tree");
        var groups = shapeTree.Elements<GroupShape>().ToList();
        if (grpIdx < 1 || grpIdx > groups.Count)
            throw new ArgumentException($"Group {grpIdx} not found (total: {groups.Count})");

        var grp = groups[grpIdx - 1];
        var unsupported = new List<string>();
        foreach (var (key, value) in properties)
        {
            switch (key.ToLowerInvariant())
            {
                case "name":
                    var nvGrpPr = grp.NonVisualGroupShapeProperties?.NonVisualDrawingProperties;
                    if (nvGrpPr != null) nvGrpPr.Name = value;
                    break;
                case "x" or "y" or "width" or "height":
                {
                    var grpSpPr = grp.GroupShapeProperties ?? (grp.GroupShapeProperties = new GroupShapeProperties());
                    var xfrm = grpSpPr.TransformGroup ?? (grpSpPr.TransformGroup = new Drawing.TransformGroup());
                    var off = xfrm.Offset ?? (xfrm.Offset = new Drawing.Offset());
                    var ext = xfrm.Extents ?? (xfrm.Extents = new Drawing.Extents());
                    var keyLower = key.ToLowerInvariant();
                    // CONSISTENCY(group-scale-baseline): group scaling needs <a:chOff>/<a:chExt>
                    // as a child-coordinate baseline. Before we mutate ext/off, snapshot the
                    // current ext/off into chExt/chOff if they aren't already present — that
                    // way the first Set of width/height captures the "before" as the logical
                    // child coordinate space, so shrinking ext shrinks the rendered children.
                    if (keyLower is "x" or "y")
                    {
                        if (xfrm.ChildOffset == null)
                            xfrm.ChildOffset = new Drawing.ChildOffset { X = off.X ?? 0, Y = off.Y ?? 0 };
                    }
                    else // width or height
                    {
                        if (xfrm.ChildExtents == null)
                            xfrm.ChildExtents = new Drawing.ChildExtents { Cx = ext.Cx ?? 0, Cy = ext.Cy ?? 0 };
                    }
                    TryApplyPositionSize(keyLower, value, off, ext);
                    break;
                }
                case "rotation" or "rotate":
                {
                    var grpSpPr = grp.GroupShapeProperties ?? (grp.GroupShapeProperties = new GroupShapeProperties());
                    var xfrm = grpSpPr.TransformGroup ?? (grpSpPr.TransformGroup = new Drawing.TransformGroup());
                    xfrm.Rotation = (int)(ParseHelpers.SafeParseDouble(value, "rotation") * 60000);
                    break;
                }
                case "fill":
                {
                    var grpSpPr = grp.GroupShapeProperties ?? (grp.GroupShapeProperties = new GroupShapeProperties());
                    grpSpPr.RemoveAllChildren<Drawing.SolidFill>();
                    grpSpPr.RemoveAllChildren<Drawing.NoFill>();
                    grpSpPr.RemoveAllChildren<Drawing.GradientFill>();
                    if (value.Equals("none", StringComparison.OrdinalIgnoreCase))
                        grpSpPr.AppendChild(new Drawing.NoFill());
                    else
                        grpSpPr.AppendChild(BuildSolidFill(value));
                    break;
                }
                default:
                    if (!GenericXmlQuery.SetGenericAttribute(grp, key, value))
                    {
                        if (unsupported.Count == 0)
                            unsupported.Add($"{key} (valid group props: x, y, width, height, rotation, name, fill)");
                        else
                            unsupported.Add(key);
                    }
                    break;
            }
        }
        GetSlide(slidePart).Save();
        return unsupported;
    }

    private List<string> SetConnectorByPath(Match cxnMatch, Dictionary<string, string> properties)
    {
        var slideIdx = int.Parse(cxnMatch.Groups[1].Value);
        var cxnIdx = int.Parse(cxnMatch.Groups[2].Value);

        var slideParts5 = GetSlideParts().ToList();
        if (slideIdx < 1 || slideIdx > slideParts5.Count)
            throw new ArgumentException($"Slide {slideIdx} not found (total: {slideParts5.Count})");

        var slidePart = slideParts5[slideIdx - 1];
        var shapeTree = GetSlide(slidePart).CommonSlideData?.ShapeTree
            ?? throw new ArgumentException("Slide has no shape tree");
        var connectors = shapeTree.Elements<ConnectionShape>().ToList();
        if (cxnIdx < 1 || cxnIdx > connectors.Count)
            throw new ArgumentException($"Connector {cxnIdx} not found (total: {connectors.Count})");

        var cxn = connectors[cxnIdx - 1];
        var unsupported = new List<string>();
        foreach (var (key, value) in properties)
        {
            switch (key.ToLowerInvariant())
            {
                case "name":
                    var nvCxnPr = cxn.NonVisualConnectionShapeProperties?.NonVisualDrawingProperties;
                    if (nvCxnPr != null) nvCxnPr.Name = value;
                    break;
                case "x" or "y" or "width" or "height":
                {
                    var spPr = cxn.ShapeProperties ?? (cxn.ShapeProperties = new ShapeProperties());
                    var xfrm = spPr.Transform2D ?? (spPr.Transform2D = new Drawing.Transform2D());
                    TryApplyPositionSize(key.ToLowerInvariant(), value,
                        xfrm.Offset ?? (xfrm.Offset = new Drawing.Offset()),
                        xfrm.Extents ?? (xfrm.Extents = new Drawing.Extents()));
                    break;
                }
                case "linewidth" or "line.width":
                {
                    var spPr = cxn.ShapeProperties ?? (cxn.ShapeProperties = new ShapeProperties());
                    var outline = spPr.GetFirstChild<Drawing.Outline>()
                        ?? spPr.AppendChild(new Drawing.Outline());
                    outline.Width = Core.EmuConverter.ParseLineWidth(value);
                    break;
                }
                case "linecolor" or "line.color" or "line" or "color":
                {
                    var spPr = cxn.ShapeProperties ?? (cxn.ShapeProperties = new ShapeProperties());
                    var outline = spPr.GetFirstChild<Drawing.Outline>()
                        ?? spPr.AppendChild(new Drawing.Outline());
                    var (rgb, _) = ParseHelpers.SanitizeColorForOoxml(value);
                    outline.RemoveAllChildren<Drawing.SolidFill>();
                    var newFill = new Drawing.SolidFill(
                        new Drawing.RgbColorModelHex { Val = rgb });
                    // CT_LineProperties schema: fill → prstDash → ... → headEnd → tailEnd
                    var prstDash = outline.GetFirstChild<Drawing.PresetDash>();
                    if (prstDash != null)
                        outline.InsertBefore(newFill, prstDash);
                    else
                    {
                        var headEnd = outline.GetFirstChild<Drawing.HeadEnd>();
                        if (headEnd != null)
                            outline.InsertBefore(newFill, headEnd);
                        else
                        {
                            var tailEnd = outline.GetFirstChild<Drawing.TailEnd>();
                            if (tailEnd != null)
                                outline.InsertBefore(newFill, tailEnd);
                            else
                                outline.AppendChild(newFill);
                        }
                    }
                    break;
                }
                case "fill":
                {
                    var spPr = cxn.ShapeProperties ?? (cxn.ShapeProperties = new ShapeProperties());
                    ApplyShapeFill(spPr, value);
                    break;
                }
                case "linedash" or "line.dash":
                {
                    var spPr = cxn.ShapeProperties ?? (cxn.ShapeProperties = new ShapeProperties());
                    var outline = spPr.GetFirstChild<Drawing.Outline>()
                        ?? spPr.AppendChild(new Drawing.Outline());
                    outline.RemoveAllChildren<Drawing.PresetDash>();
                    var newDash = new Drawing.PresetDash { Val = value.ToLowerInvariant() switch
                    {
                        "solid" => Drawing.PresetLineDashValues.Solid,
                        "dot" => Drawing.PresetLineDashValues.Dot,
                        "dash" => Drawing.PresetLineDashValues.Dash,
                        "dashdot" or "dash_dot" => Drawing.PresetLineDashValues.DashDot,
                        "longdash" or "lgdash" or "lg_dash" => Drawing.PresetLineDashValues.LargeDash,
                        "longdashdot" or "lgdashdot" or "lg_dash_dot" => Drawing.PresetLineDashValues.LargeDashDot,
                        _ => throw new ArgumentException($"Invalid 'lineDash' value: '{value}'. Valid values: solid, dot, dash, dashdot, longdash, longdashdot.")
                    }};
                    // CT_LineProperties schema: fill → prstDash → ... → headEnd → tailEnd
                    var headEnd = outline.GetFirstChild<Drawing.HeadEnd>();
                    if (headEnd != null)
                        outline.InsertBefore(newDash, headEnd);
                    else
                    {
                        var tailEnd = outline.GetFirstChild<Drawing.TailEnd>();
                        if (tailEnd != null)
                            outline.InsertBefore(newDash, tailEnd);
                        else
                            outline.AppendChild(newDash);
                    }
                    break;
                }
                case "lineopacity" or "line.opacity":
                {
                    var spPr = cxn.ShapeProperties ?? (cxn.ShapeProperties = new ShapeProperties());
                    if (!double.TryParse(value, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var lnOpacity)
                        || double.IsNaN(lnOpacity) || double.IsInfinity(lnOpacity))
                        throw new ArgumentException($"Invalid 'lineOpacity' value: '{value}'. Expected a finite decimal 0.0-1.0.");
                    var outline = spPr.GetFirstChild<Drawing.Outline>()
                        ?? spPr.AppendChild(new Drawing.Outline());
                    var solidFill = outline.GetFirstChild<Drawing.SolidFill>();
                    if (solidFill == null)
                    {
                        // Auto-create a black line fill (matching Apache POI behavior)
                        // CT_LineProperties schema: fill → prstDash → ... → headEnd → tailEnd
                        solidFill = new Drawing.SolidFill(new Drawing.RgbColorModelHex { Val = "000000" });
                        var prstDashEl = outline.GetFirstChild<Drawing.PresetDash>();
                        if (prstDashEl != null)
                            outline.InsertBefore(solidFill, prstDashEl);
                        else
                        {
                            var headEndEl = outline.GetFirstChild<Drawing.HeadEnd>();
                            if (headEndEl != null)
                                outline.InsertBefore(solidFill, headEndEl);
                            else
                            {
                                var tailEndEl = outline.GetFirstChild<Drawing.TailEnd>();
                                if (tailEndEl != null)
                                    outline.InsertBefore(solidFill, tailEndEl);
                                else
                                    outline.AppendChild(solidFill);
                            }
                        }
                    }
                    {
                        var colorEl = solidFill.GetFirstChild<Drawing.RgbColorModelHex>() as OpenXmlElement
                            ?? solidFill.GetFirstChild<Drawing.SchemeColor>();
                        if (colorEl != null)
                        {
                            colorEl.RemoveAllChildren<Drawing.Alpha>();
                            colorEl.AppendChild(new Drawing.Alpha { Val = (int)(lnOpacity * 100000) });
                        }
                    }
                    break;
                }
                case "rotation" or "rotate":
                {
                    var spPr = cxn.ShapeProperties ?? (cxn.ShapeProperties = new ShapeProperties());
                    var xfrm = spPr.Transform2D ?? (spPr.Transform2D = new Drawing.Transform2D());
                    xfrm.Rotation = (int)(ParseHelpers.SafeParseDouble(value, "rotation") * 60000);
                    break;
                }
                case "preset" or "prstgeom" or "shape":
                {
                    // CONSISTENCY(canonical-key): schema canonical is 'shape';
                    // 'preset'/'prstgeom' retained as legacy aliases.
                    var spPr = cxn.ShapeProperties ?? (cxn.ShapeProperties = new ShapeProperties());
                    var prstGeom = spPr.GetFirstChild<Drawing.PresetGeometry>()
                        ?? spPr.AppendChild(new Drawing.PresetGeometry());
                    prstGeom.Preset = new Drawing.ShapeTypeValues(value);
                    break;
                }
                case "headend" or "headEnd":
                {
                    var spPr = cxn.ShapeProperties ?? (cxn.ShapeProperties = new ShapeProperties());
                    var outline = spPr.GetFirstChild<Drawing.Outline>()
                        ?? spPr.AppendChild(new Drawing.Outline());
                    outline.RemoveAllChildren<Drawing.HeadEnd>();
                    var newHeadEnd = new Drawing.HeadEnd { Type = ParseLineEndType(value) };
                    // CT_LineProperties: ... → headEnd → tailEnd (headEnd before tailEnd)
                    var existingTailEnd = outline.GetFirstChild<Drawing.TailEnd>();
                    if (existingTailEnd != null)
                        outline.InsertBefore(newHeadEnd, existingTailEnd);
                    else
                        outline.AppendChild(newHeadEnd);
                    break;
                }
                case "tailend" or "tailEnd":
                {
                    var spPr = cxn.ShapeProperties ?? (cxn.ShapeProperties = new ShapeProperties());
                    var outline = spPr.GetFirstChild<Drawing.Outline>()
                        ?? spPr.AppendChild(new Drawing.Outline());
                    outline.RemoveAllChildren<Drawing.TailEnd>();
                    // CT_LineProperties: tailEnd is last — always append
                    outline.AppendChild(new Drawing.TailEnd { Type = ParseLineEndType(value) });
                    break;
                }
                default:
                    if (!GenericXmlQuery.SetGenericAttribute(cxn, key, value))
                    {
                        if (unsupported.Count == 0)
                            unsupported.Add($"{key} (valid connector props: line, color, fill, x, y, width, height, rotation, name, headEnd, tailEnd, geometry)");
                        else
                            unsupported.Add(key);
                    }
                    break;
            }
        }
        GetSlide(slidePart).Save();
        return unsupported;
    }

    private List<string> SetShapeByPath(Match match, Dictionary<string, string> properties)
    {
        var slideIdx = int.Parse(match.Groups[1].Value);
        var shapeIdx = int.Parse(match.Groups[2].Value);

        var (slidePart, shape) = ResolveShape(slideIdx, shapeIdx);

        // Handle z-order first (changes shape position in tree)
        var zOrderValue = properties.GetValueOrDefault("zorder")
            ?? properties.GetValueOrDefault("z-order")
            ?? properties.GetValueOrDefault("order");
        if (zOrderValue != null)
        {
            ApplyZOrder(slidePart, shape, zOrderValue);
        }

        // Clone shape for rollback on failure (atomic: no partial modifications)
        var shapeBackup = shape.CloneNode(true);

        try
        {
            var allRuns = shape.Descendants<Drawing.Run>().ToList();

            // Separate animation, motionPath, link, and z-order from other shape properties
            var animValue = properties.GetValueOrDefault("animation")
                ?? properties.GetValueOrDefault("animate");
            var motionPathValue = properties.GetValueOrDefault("motionpath")
                ?? properties.GetValueOrDefault("motionPath");
            var linkValue = properties.GetValueOrDefault("link");
            var tooltipValue = properties.GetValueOrDefault("tooltip");
            var excludeKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
                { "animation", "animate", "motionpath", "motionPath", "link", "tooltip", "zorder", "z-order", "order" };
            var shapeProps = properties
                .Where(kv => !excludeKeys.Contains(kv.Key))
                .ToDictionary(kv => kv.Key, kv => kv.Value);

            var unsupported = SetRunOrShapeProperties(shapeProps, allRuns, shape, slidePart);

            if (animValue != null)
            {
                // Remove existing animations before applying new one (replace, not accumulate)
                var shapeId = shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id?.Value;
                if (shapeId.HasValue)
                    RemoveShapeAnimations(slidePart.Slide!, shapeId.Value);
                ApplyShapeAnimation(slidePart, shape, animValue);
            }
            if (motionPathValue != null)
                ApplyMotionPathAnimation(slidePart, shape, motionPathValue);
            if (linkValue != null)
                ApplyShapeHyperlink(slidePart, shape, linkValue, tooltipValue);

            GetSlide(slidePart).Save();
            return unsupported;
        }
        catch
        {
            // Rollback: restore shape to pre-modification state
            shape.Parent?.ReplaceChild(shapeBackup, shape);
            throw;
        }
    }
}
