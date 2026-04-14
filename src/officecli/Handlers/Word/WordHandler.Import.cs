// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeCli.Handlers;

public partial class WordHandler
{
    public string ImportMarkdown(string parentPath, string markdownContent, string? styleSourceFile = null)
    {
        if (string.IsNullOrWhiteSpace(parentPath) || parentPath == "/")
            parentPath = "/body";
        if (!string.Equals(parentPath, "/body", StringComparison.OrdinalIgnoreCase))
            throw new ArgumentException("Markdown import for .docx currently supports parent path /body only");

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

        void FlushParagraphBuffer()
        {
            if (paragraphBuffer.Count == 0) return;
            var text = string.Join(" ", paragraphBuffer).Trim();
            paragraphBuffer.Clear();
            if (text.Length == 0) return;
            AppendMarkdownParagraph(body, text, styleMap.NormalStyleId, null, false);
            addedParagraphs++;
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
                AppendMarkdownParagraph(body, text, styleMap.NormalStyleId, ("bullet", level), false);
                addedParagraphs++;
                bulletCount++;
                continue;
            }

            var numberedMatch = Regex.Match(line, @"^(\s*)\d+[.)]\s+(.+?)\s*$");
            if (numberedMatch.Success)
            {
                FlushParagraphBuffer();
                var level = ComputeListLevel(numberedMatch.Groups[1].Value);
                var text = numberedMatch.Groups[2].Value.Trim();
                AppendMarkdownParagraph(body, text, styleMap.NormalStyleId, ("number", level), false);
                addedParagraphs++;
                numberedCount++;
                continue;
            }

            paragraphBuffer.Add(line.Trim());
        }

        FlushParagraphBuffer();
        _doc.MainDocumentPart?.Document?.Save();

        return $"Imported Markdown into /body: {addedParagraphs} paragraph(s), headings={headingCount}, bullets={bulletCount}, numbered={numberedCount}, code-lines={codeLineCount}";
    }

    private void AppendMarkdownParagraph(Body body, string text, string? styleId, (string ListStyle, int Level)? listStyle, bool forceCodeFont)
    {
        var para = new Paragraph();
        AssignParaId(para);

        var pProps = para.ParagraphProperties ?? para.PrependChild(new ParagraphProperties());
        if (!string.IsNullOrWhiteSpace(styleId))
            pProps.ParagraphStyleId = new ParagraphStyleId { Val = styleId };
        if (listStyle.HasValue)
            ApplyListStyle(para, listStyle.Value.ListStyle, listLevel: listStyle.Value.Level);

        var run = new Run();
        if (forceCodeFont)
        {
            run.RunProperties = new RunProperties(
                new RunFonts { Ascii = "Consolas", HighAnsi = "Consolas", EastAsia = "Consolas" }
            );
        }
        run.Append(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
        para.Append(run);
        body.AppendChild(para);
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
        if (string.IsNullOrEmpty(leadingWhitespace))
            return 0;
        int spaces = 0;
        foreach (var ch in leadingWhitespace)
            spaces += ch == '\t' ? 4 : 1;
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
