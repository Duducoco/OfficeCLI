// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using M = DocumentFormat.OpenXml.Math;

namespace OfficeCli.Handlers;

public partial class WordHandler
{
    private const string DefaultMarkdownCodeFont = "Consolas";

    // Compiled regex for Markdown inline image: ![alt](src)
    // Used both in the main line scanner and in the inline segment splitter.
    private static readonly Regex s_markdownImageRegex = new(@"^!\[([^\]]*)\]\(([^)]+)\)", RegexOptions.Compiled);

    public string ImportMarkdown(string parentPath, string markdownContent, string? styleSourceFile = null)
    {
        if (!string.Equals(parentPath, "/body", StringComparison.OrdinalIgnoreCase)
            && !string.Equals(parentPath, "/", StringComparison.OrdinalIgnoreCase))
            throw new ArgumentException("Markdown import for .docx currently supports parent path /body (or / as alias) only");
        parentPath = "/body";

        var body = _doc.MainDocumentPart?.Document?.Body
            ?? throw new InvalidOperationException("Document body not found");

        var styleMap = ResolveMarkdownStyleMap(styleSourceFile);
        var lines = markdownContent.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');

        var paragraphBuffer = new List<string>();
        var inCodeFence = false;

        int addedParagraphs = 0;
        int headingCount = 0;
        int bulletCount = 0;
        int numberedCount = 0;
        int codeLineCount = 0;
        int imageCount = 0;
        int formulaCount = 0;

        void FlushParagraphBuffer()
        {
            if (paragraphBuffer.Count == 0) return;
            var text = string.Join(" ", paragraphBuffer).Trim();
            paragraphBuffer.Clear();
            if (text.Length == 0) return;
            var (imgs, fmls) = AppendMarkdownParagraph(body, text, styleMap.NormalStyleId, null, false);
            addedParagraphs++;
            imageCount += imgs;
            formulaCount += fmls;
        }

        foreach (var rawLine in lines)
        {
            var line = rawLine ?? string.Empty;

            if (Regex.IsMatch(line, @"^\s*```"))
            {
                FlushParagraphBuffer();
                inCodeFence = !inCodeFence;
                continue;
            }

            if (inCodeFence)
            {
                AppendMarkdownParagraph(body, line, styleMap.CodeStyleId ?? styleMap.NormalStyleId, null, true);
                addedParagraphs++;
                codeLineCount++;
                continue;
            }

            if (string.IsNullOrWhiteSpace(line))
            {
                FlushParagraphBuffer();
                continue;
            }

            // Block formula: $$ ... $$ on its own line (or spanning lines is not supported here)
            // Greedy quantifier ensures the full content between $$ delimiters is captured.
            var blockFormulaMatch = Regex.Match(line.Trim(), @"^\$\$(.+)\$\$$");
            if (blockFormulaMatch.Success)
            {
                FlushParagraphBuffer();
                var latex = blockFormulaMatch.Groups[1].Value.Trim();
                AppendMarkdownDisplayFormula(body, latex);
                addedParagraphs++;
                formulaCount++;
                continue;
            }

            // Standalone image line: ![alt](src)
            var standaloneLineStr = line.Trim();
            var standaloneImageMatch = standaloneLineStr.Length > 0
                ? s_markdownImageRegex.Match(standaloneLineStr)
                : Match.Empty;
            // Ensure it matches the whole trimmed line (add end-of-string anchor check)
            if (standaloneImageMatch.Success && standaloneImageMatch.Index == 0 && standaloneImageMatch.Length == standaloneLineStr.Length)
            {
                FlushParagraphBuffer();
                var alt = standaloneImageMatch.Groups[1].Value;
                var src = standaloneImageMatch.Groups[2].Value.Trim();
                AppendMarkdownImageParagraph(body, alt, src);
                addedParagraphs++;
                imageCount++;
                continue;
            }

            var headingMatch = Regex.Match(line, @"^\s{0,3}(#{1,6})\s+(.+?)\s*$");
            if (headingMatch.Success)
            {
                FlushParagraphBuffer();
                var level = headingMatch.Groups[1].Value.Length;
                var headingText = headingMatch.Groups[2].Value.Trim();
                AppendMarkdownParagraph(body, headingText, styleMap.GetHeadingStyle(level) ?? styleMap.NormalStyleId, null, false);
                addedParagraphs++;
                headingCount++;
                continue;
            }

            var bulletMatch = Regex.Match(line, @"^(\s*)[-*+]\s+(.+?)\s*$");
            if (bulletMatch.Success)
            {
                FlushParagraphBuffer();
                var level = ComputeListLevel(bulletMatch.Groups[1].Value);
                var text = bulletMatch.Groups[2].Value.Trim();
                var (bImgs, bFmls) = AppendMarkdownParagraph(body, text, styleMap.NormalStyleId, ("bullet", level), false);
                addedParagraphs++;
                bulletCount++;
                imageCount += bImgs;
                formulaCount += bFmls;
                continue;
            }

            var numberedMatch = Regex.Match(line, @"^(\s*)\d+[.)]\s+(.+?)\s*$");
            if (numberedMatch.Success)
            {
                FlushParagraphBuffer();
                var level = ComputeListLevel(numberedMatch.Groups[1].Value);
                var text = numberedMatch.Groups[2].Value.Trim();
                var (nImgs, nFmls) = AppendMarkdownParagraph(body, text, styleMap.NormalStyleId, ("number", level), false);
                addedParagraphs++;
                numberedCount++;
                imageCount += nImgs;
                formulaCount += nFmls;
                continue;
            }

            paragraphBuffer.Add(line.Trim());
        }

        FlushParagraphBuffer();
        _doc.MainDocumentPart?.Document?.Save();

        return $"Imported Markdown into /body: {addedParagraphs} paragraph(s), headings={headingCount}, bullets={bulletCount}, numbered={numberedCount}, code-lines={codeLineCount}, images={imageCount}, formulas={formulaCount}";
    }

    /// <summary>
    /// Appends a paragraph to <paramref name="body"/>, splitting the text into runs, inline images,
    /// and inline math ($...$). Returns (imageCount, formulaCount) of elements added to the paragraph.
    /// </summary>
    private (int Images, int Formulas) AppendMarkdownParagraph(Body body, string text, string? styleId,
        (string ListStyle, int Level)? listStyle, bool forceCodeFont)
    {
        var para = new Paragraph();
        AssignParaId(para);

        var pProps = para.ParagraphProperties ?? para.PrependChild(new ParagraphProperties());
        if (!string.IsNullOrWhiteSpace(styleId))
            pProps.ParagraphStyleId = new ParagraphStyleId { Val = styleId };
        if (listStyle.HasValue)
            ApplyListStyle(para, listStyle.Value.ListStyle, listLevel: listStyle.Value.Level);

        int images = 0;
        int formulas = 0;

        if (forceCodeFont)
        {
            // Code fence lines: single run with code font, no inline splitting
            var codeRun = new Run();
            codeRun.RunProperties = new RunProperties(
                new RunFonts { Ascii = DefaultMarkdownCodeFont, HighAnsi = DefaultMarkdownCodeFont, EastAsia = DefaultMarkdownCodeFont }
            );
            codeRun.Append(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
            para.Append(codeRun);
        }
        else
        {
            // Split text into segments: plain text, ![alt](src), $formula$
            // Pattern order: image > block-inline formula ($$...$$) > inline formula ($...$)
            var segments = SplitMarkdownInlineSegments(text);
            foreach (var seg in segments)
            {
                if (seg.Kind == MarkdownSegmentKind.Text)
                {
                    if (seg.Value.Length > 0)
                    {
                        var run = new Run(new Text(seg.Value) { Space = SpaceProcessingModeValues.Preserve });
                        para.Append(run);
                    }
                }
                else if (seg.Kind == MarkdownSegmentKind.Image)
                {
                    var imgRun = TryCreateMarkdownImageRun(seg.Alt ?? string.Empty, seg.Value);
                    if (imgRun != null)
                    {
                        para.Append(imgRun);
                        images++;
                    }
                    else
                    {
                        // Fallback: insert alt text as plain text if image loading fails
                        if (!string.IsNullOrWhiteSpace(seg.Alt))
                            para.Append(new Run(new Text($"[{seg.Alt}]") { Space = SpaceProcessingModeValues.Preserve }));
                    }
                }
                else if (seg.Kind == MarkdownSegmentKind.Formula)
                {
                    try
                    {
                        var mathElem = Core.FormulaParser.Parse(seg.Value);
                        M.OfficeMath oMath = mathElem is M.OfficeMath om ? om : new M.OfficeMath(mathElem.CloneNode(true));
                        para.Append(oMath);
                        formulas++;
                    }
                    catch
                    {
                        // Fallback: insert raw LaTeX wrapped in dollar signs
                        para.Append(new Run(new Text($"${seg.Value}$") { Space = SpaceProcessingModeValues.Preserve }));
                    }
                }
            }
        }

        body.AppendChild(para);
        return (images, formulas);
    }

    /// <summary>
    /// Appends a standalone image paragraph (one image per paragraph, as display blocks).
    /// Returns false and skips silently if the image cannot be loaded.
    /// </summary>
    private void AppendMarkdownImageParagraph(Body body, string alt, string src)
    {
        var imgRun = TryCreateMarkdownImageRun(alt, src);
        if (imgRun == null) return;

        var para = new Paragraph(imgRun);
        AssignParaId(para);
        para.PrependChild(new ParagraphProperties(
            new SpacingBetweenLines { Line = "240", LineRule = LineSpacingRuleValues.Auto }));
        body.AppendChild(para);
    }

    /// <summary>
    /// Appends a display-mode formula paragraph (m:oMathPara wrapped in w:p).
    /// Silently falls back to a plain-text paragraph if parsing fails.
    /// </summary>
    private void AppendMarkdownDisplayFormula(Body body, string latex)
    {
        try
        {
            var mathElem = Core.FormulaParser.Parse(latex);
            M.OfficeMath oMath = mathElem is M.OfficeMath om ? om : new M.OfficeMath(mathElem.CloneNode(true));
            var mathPara = new M.Paragraph(oMath);
            var wrapPara = new Paragraph(mathPara);
            AssignParaId(wrapPara);
            body.AppendChild(wrapPara);
        }
        catch
        {
            // Fallback: insert raw LaTeX as plain text
            var para = new Paragraph();
            AssignParaId(para);
            para.Append(new Run(new Text($"$${latex}$$") { Space = SpaceProcessingModeValues.Preserve }));
            body.AppendChild(para);
        }
    }

    /// <summary>
    /// Creates an inline image Run from an image source string. Returns null if the image cannot be loaded.
    /// </summary>
    private Run? TryCreateMarkdownImageRun(string alt, string src)
    {
        try
        {
            var (rawStream, imgPartType) = Core.ImageSource.Resolve(src);
            using var rawStreamDispose = rawStream;
            using var imgStream = new MemoryStream();
            rawStream.CopyTo(imgStream);
            imgStream.Position = 0;

            var mainPart = _doc.MainDocumentPart!;
            var imagePart = mainPart.AddImagePart(imgPartType);
            imagePart.FeedData(imgStream);
            imgStream.Position = 0;
            var relId = mainPart.GetIdOfPart(imagePart);

            // Default to 6-inch wide inline, maintaining aspect ratio
            long cxEmu = 5486400L;
            long cyEmu = 3657600L;
            var dims = Core.ImageSource.TryGetDimensions(imgStream);
            if (dims is { Width: > 0, Height: > 0 } d)
                cyEmu = (long)(cxEmu * ((double)d.Height / d.Width));

            var altText = string.IsNullOrWhiteSpace(alt) ? Path.GetFileName(src) : alt;
            return CreateImageRun(relId, cxEmu, cyEmu, altText, NextDocPropId());
        }
        catch
        {
            return null;
        }
    }

    // ==================== Inline Markdown Segment Splitter ====================

    private enum MarkdownSegmentKind { Text, Image, Formula }

    private sealed class MarkdownSegment
    {
        public MarkdownSegmentKind Kind { get; init; }
        public string Value { get; init; } = string.Empty;
        /// <summary>Alt text for images.</summary>
        public string? Alt { get; init; }
    }

    /// <summary>
    /// Splits a line of Markdown text into typed segments: plain text, images (![alt](src)),
    /// and inline formulas ($...$).
    /// </summary>
    private static List<MarkdownSegment> SplitMarkdownInlineSegments(string text)
    {
        var result = new List<MarkdownSegment>();
        int pos = 0;

        // Matches (in priority order): image ![alt](src), inline formula $...$
        // We scan left-to-right and pick the earliest match.
        while (pos < text.Length)
        {
            // Find next candidate: '!' for image, '$' for formula
            int nextBang = text.IndexOf("![", pos, StringComparison.Ordinal);
            int nextDollar = text.IndexOf('$', pos);

            // Pick the earliest candidate; images take priority when tied with a dollar sign
            // since '!' and '$' are different characters and can't share the same position.
            // Using strict less-than to give '$' priority when nextBang == nextDollar is
            // impossible in practice, but strict < is semantically clearer.
            int nextSpecial = -1;
            bool isImage = false;
            if (nextBang >= 0 && (nextDollar < 0 || nextBang < nextDollar))
            {
                nextSpecial = nextBang;
                isImage = true;
            }
            else if (nextDollar >= 0)
            {
                nextSpecial = nextDollar;
                isImage = false;
            }

            if (nextSpecial < 0)
            {
                // No more special patterns — consume remaining text
                if (pos < text.Length)
                    result.Add(new MarkdownSegment { Kind = MarkdownSegmentKind.Text, Value = text[pos..] });
                break;
            }

            // Emit leading plain text
            if (nextSpecial > pos)
                result.Add(new MarkdownSegment { Kind = MarkdownSegmentKind.Text, Value = text[pos..nextSpecial] });

            if (isImage)
            {
                // Try to parse ![alt](src) using shared compiled regex
                var imageMatch = s_markdownImageRegex.Match(text[nextSpecial..]);
                if (imageMatch.Success)
                {
                    result.Add(new MarkdownSegment
                    {
                        Kind = MarkdownSegmentKind.Image,
                        Alt = imageMatch.Groups[1].Value,
                        Value = imageMatch.Groups[2].Value.Trim()
                    });
                    pos = nextSpecial + imageMatch.Length;
                }
                else
                {
                    // Not a valid image syntax — treat '!' as plain text
                    result.Add(new MarkdownSegment { Kind = MarkdownSegmentKind.Text, Value = "!" });
                    pos = nextSpecial + 1;
                }
            }
            else
            {
                // Try to parse $...$ (inline formula)
                int closingDollar = FindClosingDollar(text, nextDollar + 1);
                if (closingDollar > nextDollar)
                {
                    var latex = text[(nextDollar + 1)..closingDollar].Trim();
                    result.Add(new MarkdownSegment { Kind = MarkdownSegmentKind.Formula, Value = latex });
                    pos = closingDollar + 1;
                }
                else
                {
                    // No closing dollar — treat '$' as plain text
                    result.Add(new MarkdownSegment { Kind = MarkdownSegmentKind.Text, Value = "$" });
                    pos = nextDollar + 1;
                }
            }
        }

        return result;
    }

    /// <summary>
    /// Finds the position of the closing '$' for an inline formula starting at <paramref name="start"/>.
    /// Returns -1 if no closing '$' is found before end-of-string or a newline.
    /// </summary>
    private static int FindClosingDollar(string text, int start)
    {
        for (int i = start; i < text.Length; i++)
        {
            if (text[i] == '\n') return -1;
            if (text[i] == '$') return i;
        }
        return -1;
    }

    private MarkdownStyleMap ResolveMarkdownStyleMap(string? styleSourceFile)
    {
        var targetMainPart = _doc.MainDocumentPart
            ?? throw new InvalidOperationException("No main document part");
        var targetStylesPart = targetMainPart.StyleDefinitionsPart
            ?? targetMainPart.AddNewPart<StyleDefinitionsPart>();
        targetStylesPart.Styles ??= new Styles();

        if (string.IsNullOrWhiteSpace(styleSourceFile))
            return BuildMarkdownStyleMap(targetStylesPart.Styles);

        using var sourceDoc = WordprocessingDocument.Open(styleSourceFile, false);
        var sourceStyles = sourceDoc.MainDocumentPart?.StyleDefinitionsPart?.Styles;
        if (sourceStyles == null)
            return BuildMarkdownStyleMap(targetStylesPart.Styles);

        var mapFromSource = BuildMarkdownStyleMap(sourceStyles);
        foreach (var styleId in mapFromSource.AllStyleIds())
            CopyStyleChainIfMissing(sourceStyles, targetStylesPart.Styles, styleId);

        targetStylesPart.Styles.Save();
        return mapFromSource;
    }

    private static MarkdownStyleMap BuildMarkdownStyleMap(Styles styles)
    {
        var map = new MarkdownStyleMap();
        var paragraphStyles = styles.Elements<Style>()
            .Where(s => s.Type?.Value == StyleValues.Paragraph)
            .ToList();

        map.NormalStyleId =
            paragraphStyles.FirstOrDefault(s => s.Default?.Value == true)?.StyleId?.Value
            ?? FindParagraphStyleId(paragraphStyles, "Normal", "normal", "正文", "body")
            ?? paragraphStyles.FirstOrDefault()?.StyleId?.Value;

        for (int level = 1; level <= 6; level++)
        {
            var headingStyleId =
                FindHeadingByOutlineLevel(paragraphStyles, level)
                ?? FindParagraphStyleId(paragraphStyles, $"Heading{level}", $"heading {level}", $"heading{level}", $"标题{level}", $"标题 {level}", $"h{level}");
            map.SetHeadingStyle(level, headingStyleId);
        }

        map.CodeStyleId = FindParagraphStyleId(paragraphStyles, "Code", "code", "代码");
        return map;
    }

    private static string? FindHeadingByOutlineLevel(List<Style> paragraphStyles, int level)
    {
        var outlineVal = level - 1;
        return paragraphStyles
            .FirstOrDefault(s => s.StyleParagraphProperties?.OutlineLevel?.Val?.Value == outlineVal)
            ?.StyleId?.Value;
    }

    private static string? FindParagraphStyleId(List<Style> paragraphStyles, params string[] keys)
    {
        foreach (var key in keys)
        {
            var byId = paragraphStyles.FirstOrDefault(s =>
                string.Equals(s.StyleId?.Value, key, StringComparison.OrdinalIgnoreCase));
            if (byId?.StyleId?.Value != null) return byId.StyleId.Value;

            var byName = paragraphStyles.FirstOrDefault(s =>
                string.Equals(s.StyleName?.Val?.Value, key, StringComparison.OrdinalIgnoreCase));
            if (byName?.StyleId?.Value != null) return byName.StyleId.Value;
        }

        foreach (var key in keys)
        {
            var byContains = paragraphStyles.FirstOrDefault(s =>
                (s.StyleName?.Val?.Value?.Contains(key, StringComparison.OrdinalIgnoreCase) ?? false));
            if (byContains?.StyleId?.Value != null) return byContains.StyleId.Value;
        }

        return null;
    }

    private static void CopyStyleChainIfMissing(Styles sourceStyles, Styles targetStyles, string? styleId, HashSet<string>? visited = null)
    {
        if (string.IsNullOrWhiteSpace(styleId)) return;
        visited ??= new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        if (!visited.Add(styleId)) return;
        if (targetStyles.Elements<Style>().Any(s => string.Equals(s.StyleId?.Value, styleId, StringComparison.OrdinalIgnoreCase)))
            return;

        var sourceStyle = sourceStyles.Elements<Style>()
            .FirstOrDefault(s => string.Equals(s.StyleId?.Value, styleId, StringComparison.OrdinalIgnoreCase));
        if (sourceStyle == null) return;

        CopyStyleChainIfMissing(sourceStyles, targetStyles, sourceStyle.BasedOn?.Val?.Value, visited);
        CopyStyleChainIfMissing(sourceStyles, targetStyles, sourceStyle.NextParagraphStyle?.Val?.Value, visited);
        CopyStyleChainIfMissing(sourceStyles, targetStyles, sourceStyle.LinkedStyle?.Val?.Value, visited);

        targetStyles.Append((Style)sourceStyle.CloneNode(true));
    }

    private static int ComputeListLevel(string leadingWhitespace)
    {
        // Markdown indentation heuristic:
        // - 2 spaces = 1 nesting level
        // - 1 tab counts as 4 spaces (so tab = level 2)
        // - clamp to Word's supported ilvl range [0..8]
        if (string.IsNullOrEmpty(leadingWhitespace))
            return 0;
        int spaces = 0;
        foreach (var ch in leadingWhitespace)
            spaces += ch == '\t' ? 4 : (ch == ' ' ? 1 : 0);
        return Math.Clamp(spaces / 2, 0, 8);
    }

    private sealed class MarkdownStyleMap
    {
        private readonly string?[] _headingStyleIds = new string?[6];
        public string? NormalStyleId { get; set; }
        public string? CodeStyleId { get; set; }

        public void SetHeadingStyle(int level, string? styleId)
        {
            if (level < 1 || level > 6) return;
            _headingStyleIds[level - 1] = styleId;
        }

        public string? GetHeadingStyle(int level)
        {
            if (level < 1 || level > 6) return null;
            return _headingStyleIds[level - 1];
        }

        public IEnumerable<string?> AllStyleIds()
        {
            yield return NormalStyleId;
            yield return CodeStyleId;
            foreach (var styleId in _headingStyleIds)
                yield return styleId;
        }
    }
}
