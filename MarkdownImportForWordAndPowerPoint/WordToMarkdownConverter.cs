using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace MarkdownImportForWordAndPowerPoint
{
    /// <summary>
    /// Converts a .docx file to Markdown using the Open XML SDK.
    /// Supported:
    ///  - headings (Heading 1-9, Title, or outline level) = "#" headings
    ///  - bulleted and numbered lists, with nesting (read from the numbering definitions)
    ///  - bold, italic, strikethrough, hyperlinks, line breaks
    ///  - inline code and code blocks (monospace fonts or code-like styles)
    ///  - quotes (Quote and Block Text styles)
    ///  - tables (the first row becomes the header row)
    /// Ignored: pictures, text boxes, footnotes, headers and footers, comments, deleted (tracked) text.
    /// Markdown special characters found in the text are escaped.
    /// </summary>
    internal sealed class WordToMarkdownConverter
    {
        // An underscore at the start or end of a word could be read as emphasis; intra-word underscores are safe.
        private static readonly Regex UnderscoreAtWordEdge =
            new Regex(@"(?<![\p{L}\p{N}])_|_(?![\p{L}\p{N}])", RegexOptions.Compiled);

        // Constructs that Markdown would interpret at the beginning of a line (heading, quote, list, rule).
        private static readonly Regex LineStart = new Regex(
            @"^(?:(?<num>\d+)(?<dot>[.)])(?=\s)|(?<hash>#{1,6})(?=\s|$)|(?<quote>>)|(?<bullet>[-+])(?=\s)|(?<rule>(?:={2,}|-{2,}|~{3,})\s*$))",
            RegexOptions.Compiled);

        private static readonly Regex HeadingName =
            new Regex(@"^heading\s*(\d)$", RegexOptions.Compiled | RegexOptions.IgnoreCase);

        // Built-in list styles such as "List Bullet 2" or "List Number 3" express nesting through their name
        private static readonly Regex ListStyleLevel =
            new Regex(@"^list\s+(?:bullet|number)\s+(\d)$", RegexOptions.Compiled | RegexOptions.IgnoreCase);

        private static readonly HashSet<string> MonospaceFonts = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "Consolas", "Courier New", "Courier", "Lucida Console", "Lucida Sans Typewriter",
            "Cascadia Code", "Cascadia Mono", "Source Code Pro", "Menlo", "Monaco",
            "DejaVu Sans Mono", "Fira Code", "JetBrains Mono"
        };

        private readonly MainDocumentPart mainPart;
        private readonly Dictionary<string, W.Style> styles = new Dictionary<string, W.Style>();
        private readonly Dictionary<int, W.NumberingInstance> numberingInstances = new Dictionary<int, W.NumberingInstance>();
        private readonly Dictionary<int, W.AbstractNum> abstractNumbers = new Dictionary<int, W.AbstractNum>();

        // Last number used per abstract list and level (null = not used yet)
        private readonly Dictionary<int, int?[]> counters = new Dictionary<int, int?[]>();
        private readonly HashSet<int> seenNumberingIds = new HashSet<int>();

        private readonly List<string> blocks = new List<string>();
        private readonly List<string> listLines = new List<string>();
        private readonly List<string> codeLines = new List<string>();
        private int lastListLevel = -1;

        private WordToMarkdownConverter(MainDocumentPart mainPart)
        {
            this.mainPart = mainPart;

            var styleList = mainPart.StyleDefinitionsPart?.Styles?.Elements<W.Style>();
            if (styleList != null)
            {
                foreach (var style in styleList)
                {
                    string id = style.StyleId?.Value;
                    if (!string.IsNullOrEmpty(id))
                    {
                        styles[id] = style;
                    }
                }
            }

            var numbering = mainPart.NumberingDefinitionsPart?.Numbering;
            if (numbering != null)
            {
                foreach (var instance in numbering.Elements<W.NumberingInstance>())
                {
                    int? id = instance.NumberID?.Value;
                    if (id.HasValue)
                    {
                        numberingInstances[id.Value] = instance;
                    }
                }

                foreach (var abstractNum in numbering.Elements<W.AbstractNum>())
                {
                    int? id = abstractNum.AbstractNumberId?.Value;
                    if (id.HasValue)
                    {
                        abstractNumbers[id.Value] = abstractNum;
                    }
                }
            }
        }

        public static string ConvertFile(string docxPath)
        {
            // FileShare.ReadWrite: allows reading even if the file is open in Word
            using (var stream = new FileStream(docxPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var document = WordprocessingDocument.Open(stream, false))
            {
                var mainPart = document.MainDocumentPart;
                var body = mainPart?.Document?.Body;
                if (body == null)
                {
                    return string.Empty;
                }

                return new WordToMarkdownConverter(mainPart).Build(body);
            }
        }

        private string Build(W.Body body)
        {
            ProcessBlocks(body);
            FlushList();
            FlushCode();

            return blocks.Count == 0 ? string.Empty : string.Join("\n\n", blocks) + "\n";
        }

        // ---------------------------------------------------------------- blocks

        private void ProcessBlocks(OpenXmlElement container)
        {
            foreach (var child in container.ChildElements)
            {
                if (child is W.Paragraph paragraph)
                {
                    ProcessParagraph(paragraph);
                }
                else if (child is W.Table table)
                {
                    FlushList();
                    FlushCode();

                    string markdown = RenderTable(table);
                    if (!string.IsNullOrWhiteSpace(markdown))
                    {
                        blocks.Add(markdown);
                    }
                }
                else if (child.LocalName == "sdt" || child.LocalName == "sdtContent" || child.LocalName == "customXml")
                {
                    ProcessBlocks(child); // content controls and custom XML wrappers
                }
            }
        }

        private void ProcessParagraph(W.Paragraph paragraph)
        {
            var segments = new List<Segment>();
            CollectSegments(paragraph, segments, null);

            bool hasText = segments.Any(s => !string.IsNullOrWhiteSpace(s.Text));
            if (!hasText)
            {
                // An empty paragraph only separates blocks, except inside a code block where it is a blank line
                if (codeLines.Count > 0)
                {
                    codeLines.Add(string.Empty);
                }

                return;
            }

            List<W.Style> chain = GetStyleChain(paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value);

            // 1. Headings
            int headingLevel = GetHeadingLevel(paragraph, chain);
            if (headingLevel > 0)
            {
                FlushList();
                FlushCode();

                string title = RenderPlain(segments);
                if (title.Length > 0)
                {
                    blocks.Add(new string('#', Math.Min(headingLevel, 6)) + " " + title);
                }

                return;
            }

            // 2. List items
            ListInfo list = GetListInfo(paragraph, chain);
            if (list != null)
            {
                FlushCode();
                AddListItem(list, segments);
                return;
            }

            // 3. Code blocks
            bool allMonospace = segments
                .Where(s => !string.IsNullOrWhiteSpace(s.Text))
                .All(s => s.Code);

            if (allMonospace || chain.Any(IsCodeStyle))
            {
                FlushList();

                string code = string.Concat(segments.Select(s => s.Text)).Replace("\t", "    ");
                codeLines.AddRange(code.TrimEnd('\r', '\n').Split('\n'));
                return;
            }

            // 4. Regular paragraphs and quotes
            FlushList();
            FlushCode();

            string text = RenderSegments(segments).Trim();
            if (text.Length == 0)
            {
                return;
            }

            text = EscapeLineStarts(text).Replace("\n", "  \n");

            if (chain.Any(IsQuoteStyle))
            {
                text = "> " + text.Replace("\n", "\n> ");
            }

            blocks.Add(text);
        }

        private void FlushList()
        {
            if (listLines.Count > 0)
            {
                blocks.Add(string.Join("\n", listLines));
                listLines.Clear();
            }

            lastListLevel = -1;
        }

        private void FlushCode()
        {
            // Blank lines at the end of a code block are only separators
            while (codeLines.Count > 0 && codeLines[codeLines.Count - 1].Trim().Length == 0)
            {
                codeLines.RemoveAt(codeLines.Count - 1);
            }

            if (codeLines.Count > 0)
            {
                string code = string.Join("\n", codeLines);
                string fence = new string('`', Math.Max(3, LongestBacktickRun(code) + 1));
                blocks.Add(fence + "\n" + code + "\n" + fence);
            }

            codeLines.Clear();
        }

        // ---------------------------------------------------------------- styles

        private List<W.Style> GetStyleChain(string styleId)
        {
            var chain = new List<W.Style>();
            int guard = 0;

            while (!string.IsNullOrEmpty(styleId) && guard++ < 20 && styles.TryGetValue(styleId, out var style))
            {
                chain.Add(style);
                styleId = style.BasedOn?.Val?.Value;
            }

            return chain;
        }

        private static int GetHeadingLevel(W.Paragraph paragraph, List<W.Style> chain)
        {
            foreach (var style in chain)
            {
                string name = style.StyleName?.Val?.Value;
                if (string.IsNullOrEmpty(name))
                {
                    continue;
                }

                var match = HeadingName.Match(name);
                if (match.Success)
                {
                    return int.Parse(match.Groups[1].Value, CultureInfo.InvariantCulture);
                }

                if (name.Equals("Title", StringComparison.OrdinalIgnoreCase))
                {
                    return 1;
                }
            }

            // Fallback: outline level set on the paragraph or on its styles (0 = level 1, 9 = body text)
            int? outline = paragraph.ParagraphProperties?.OutlineLevel?.Val?.Value;
            if (!outline.HasValue)
            {
                foreach (var style in chain)
                {
                    outline = style.StyleParagraphProperties?.OutlineLevel?.Val?.Value;
                    if (outline.HasValue)
                    {
                        break;
                    }
                }
            }

            return outline.HasValue && outline.Value >= 0 && outline.Value < 9 ? outline.Value + 1 : 0;
        }

        private static bool IsCodeStyle(W.Style style)
        {
            return IsCodeStyleName(style.StyleName?.Val?.Value);
        }

        private static bool IsCodeStyleName(string name)
        {
            if (string.IsNullOrEmpty(name))
            {
                return false;
            }

            string lower = name.ToLowerInvariant();
            return lower.Contains("code") || lower.Contains("preformatted") || lower == "plain text";
        }

        private static bool IsQuoteStyle(W.Style style)
        {
            string name = style.StyleName?.Val?.Value;
            return name != null &&
                   (name.Equals("Quote", StringComparison.OrdinalIgnoreCase) ||
                    name.Equals("Intense Quote", StringComparison.OrdinalIgnoreCase) ||
                    name.Equals("Block Text", StringComparison.OrdinalIgnoreCase));
        }

        // ---------------------------------------------------------------- lists

        private sealed class ListInfo
        {
            public int Level { get; set; }

            public bool Ordered { get; set; }

            public int Number { get; set; }
        }

        private ListInfo GetListInfo(W.Paragraph paragraph, List<W.Style> chain)
        {
            var numPr = paragraph.ParagraphProperties?.NumberingProperties;
            int? numId = numPr?.NumberingId?.Val?.Value;
            int? ilvl = numPr?.NumberingLevelReference?.Val?.Value;

            // Numbering can also be inherited from the paragraph style chain
            foreach (var style in chain)
            {
                var styleNumPr = style.StyleParagraphProperties?.NumberingProperties;
                if (styleNumPr == null)
                {
                    continue;
                }

                if (!numId.HasValue)
                {
                    numId = styleNumPr.NumberingId?.Val?.Value;
                }

                if (!ilvl.HasValue)
                {
                    ilvl = styleNumPr.NumberingLevelReference?.Val?.Value;
                }
            }

            // numId 0 means "no numbering"
            if (!numId.HasValue || numId.Value == 0)
            {
                return null;
            }

            if (!numberingInstances.TryGetValue(numId.Value, out var instance))
            {
                return null;
            }

            int? abstractId = instance.AbstractNumId?.Val?.Value;
            if (!abstractId.HasValue || !abstractNumbers.TryGetValue(abstractId.Value, out var abstractNum))
            {
                return null;
            }

            int level = Math.Max(0, Math.Min(ilvl ?? 0, 8));
            W.Level definition = FindLevel(instance, abstractNum, level);
            string format = definition?.NumberingFormat?.Val?.InnerText ?? "bullet";

            if (format == "none")
            {
                return null;
            }

            var info = new ListInfo { Level = Math.Max(level, GetStyleListLevel(chain)), Ordered = format != "bullet" };
            if (info.Ordered)
            {
                info.Number = NextNumber(numId.Value, abstractId.Value, instance, definition, level);
            }

            return info;
        }

        /// <summary>
        /// Nesting level (0-based) implied by a style name like "List Bullet 2", or 0 if there is none.
        /// </summary>
        private static int GetStyleListLevel(List<W.Style> chain)
        {
            foreach (var style in chain)
            {
                string name = style.StyleName?.Val?.Value;
                if (string.IsNullOrEmpty(name))
                {
                    continue;
                }

                var match = ListStyleLevel.Match(name);
                if (match.Success)
                {
                    return Math.Min(int.Parse(match.Groups[1].Value, CultureInfo.InvariantCulture) - 1, 8);
                }
            }

            return 0;
        }

        private static W.Level FindLevel(W.NumberingInstance instance, W.AbstractNum abstractNum, int level)
        {
            W.Level overrideLevel = instance.Elements<W.LevelOverride>()
                .FirstOrDefault(o => o.LevelIndex?.Value == level)?.Level;

            return overrideLevel ??
                   abstractNum.Elements<W.Level>().FirstOrDefault(l => l.LevelIndex?.Value == level);
        }

        private int NextNumber(int numId, int abstractId, W.NumberingInstance instance, W.Level definition, int level)
        {
            if (!counters.TryGetValue(abstractId, out var state))
            {
                state = new int?[9];
                counters[abstractId] = state;
            }

            // The first time a list instance is seen, a start override can restart the numbering
            if (seenNumberingIds.Add(numId))
            {
                foreach (var levelOverride in instance.Elements<W.LevelOverride>())
                {
                    int? index = levelOverride.LevelIndex?.Value;
                    int? start = levelOverride.StartOverrideNumberingValue?.Val?.Value;

                    if (index.HasValue && start.HasValue && index.Value >= 0 && index.Value < state.Length)
                    {
                        state[index.Value] = start.Value - 1;
                    }
                }
            }

            int first = definition?.StartNumberingValue?.Val?.Value ?? 1;
            int number = state[level].HasValue ? state[level].Value + 1 : first;

            state[level] = number;
            for (int i = level + 1; i < state.Length; i++)
            {
                state[i] = null; // deeper levels restart after an item of a higher level
            }

            return number;
        }

        private void AddListItem(ListInfo list, List<Segment> segments)
        {
            string text = RenderSegments(segments).Trim();
            if (text.Length == 0)
            {
                return;
            }

            if (listLines.Count == 0)
            {
                lastListLevel = -1;
            }

            // Avoid level jumps (e.g. 0 -> 2), which Markdown would interpret as a code block
            int level = Math.Max(0, Math.Min(list.Level, lastListLevel + 1));
            lastListLevel = level;

            string indent = new string(' ', level * 4);
            string marker = list.Ordered ? list.Number.ToString(CultureInfo.InvariantCulture) + ". " : "- ";
            string continuation = "  \n" + indent + new string(' ', marker.Length);

            listLines.Add(indent + marker + text.Replace("\n", continuation));
        }

        // ---------------------------------------------------------------- inline content

        private sealed class Segment
        {
            public string Text { get; set; }

            public bool Bold { get; set; }

            public bool Italic { get; set; }

            public bool Strike { get; set; }

            public bool Code { get; set; }

            public string Url { get; set; }
        }

        private void CollectSegments(OpenXmlElement container, List<Segment> segments, string url)
        {
            foreach (var child in container.ChildElements)
            {
                if (child is W.Run run)
                {
                    AddRun(run, segments, url);
                }
                else if (child is W.Hyperlink hyperlink)
                {
                    CollectSegments(hyperlink, segments, GetHyperlinkTarget(hyperlink) ?? url);
                }
                else
                {
                    switch (child.LocalName)
                    {
                        // Inserted text, content controls, simple fields: their runs are part of the paragraph.
                        // Deleted text (del, moveFrom) is skipped.
                        case "ins":
                        case "moveTo":
                        case "sdt":
                        case "sdtContent":
                        case "fldSimple":
                        case "smartTag":
                        case "customXml":
                            CollectSegments(child, segments, url);
                            break;
                    }
                }
            }
        }

        private string GetHyperlinkTarget(W.Hyperlink hyperlink)
        {
            string id = hyperlink.Id?.Value;
            if (string.IsNullOrEmpty(id))
            {
                return null; // internal link (anchor): keep only the text
            }

            var relationship = mainPart.HyperlinkRelationships.FirstOrDefault(r => r.Id == id);
            return relationship?.Uri?.OriginalString;
        }

        private void AddRun(W.Run run, List<Segment> segments, string url)
        {
            var properties = run.RunProperties;
            if (IsOn(properties?.Vanish))
            {
                return; // hidden text
            }

            var text = new StringBuilder();
            foreach (var child in run.ChildElements)
            {
                if (child is W.Text t)
                {
                    text.Append(t.Text);
                }
                else if (child is W.Break b)
                {
                    string type = b.Type?.InnerText;
                    if (type != "page" && type != "column")
                    {
                        text.Append('\n');
                    }
                }
                else
                {
                    switch (child.LocalName)
                    {
                        case "tab":
                            text.Append('\t');
                            break;
                        case "cr":
                            text.Append('\n');
                            break;
                        case "noBreakHyphen":
                            text.Append('-');
                            break;
                    }
                }
            }

            if (text.Length == 0)
            {
                return;
            }

            bool bold = IsOn(properties?.Bold);
            bool italic = IsOn(properties?.Italic);
            bool strike = IsOn(properties?.Strike);
            bool code = false;

            // Character styles such as Strong, Emphasis or a code style
            foreach (var style in GetStyleChain(properties?.RunStyle?.Val?.Value))
            {
                string name = style.StyleName?.Val?.Value ?? string.Empty;

                if (name.Equals("Strong", StringComparison.OrdinalIgnoreCase))
                {
                    bold = true;
                }
                else if (name.Equals("Emphasis", StringComparison.OrdinalIgnoreCase))
                {
                    italic = true;
                }
                else if (IsCodeStyleName(name))
                {
                    code = true;
                }
            }

            string font = properties?.RunFonts?.Ascii?.Value ?? properties?.RunFonts?.HighAnsi?.Value;
            if (font != null && MonospaceFonts.Contains(font))
            {
                code = true;
            }

            // Merge adjacent segments with the same formatting to avoid "**a****b**"
            if (segments.Count > 0)
            {
                var last = segments[segments.Count - 1];
                if (last.Bold == bold && last.Italic == italic && last.Strike == strike &&
                    last.Code == code && last.Url == url)
                {
                    last.Text += text.ToString();
                    return;
                }
            }

            segments.Add(new Segment
            {
                Text = text.ToString(),
                Bold = bold,
                Italic = italic,
                Strike = strike,
                Code = code,
                Url = url
            });
        }

        private static bool IsOn(W.OnOffType element)
        {
            // <w:b/> means on, <w:b w:val="0"/> means off
            return element != null && (element.Val?.Value ?? true);
        }

        private static string RenderSegments(List<Segment> segments)
        {
            var sb = new StringBuilder();
            foreach (var segment in segments)
            {
                sb.Append(RenderSegment(segment));
            }

            return sb.ToString();
        }

        private static string RenderSegment(Segment segment)
        {
            string text = segment.Text;
            string core = text.Trim();

            if (core.Length == 0)
            {
                return text.Replace("\t", " ");
            }

            // Whitespace must stay outside the markers, otherwise Markdown does not recognize them
            string leading = text.Substring(0, text.Length - text.TrimStart().Length).Replace("\t", " ");
            string trailing = text.Substring(text.TrimEnd().Length).Replace("\t", " ");
            core = core.Replace("\t", " ");

            string wrapped;
            if (segment.Code)
            {
                wrapped = WrapInlineCode(core);
            }
            else
            {
                wrapped = EscapeInline(core);

                if (segment.Bold && segment.Italic)
                {
                    wrapped = "***" + wrapped + "***";
                }
                else if (segment.Bold)
                {
                    wrapped = "**" + wrapped + "**";
                }
                else if (segment.Italic)
                {
                    wrapped = "*" + wrapped + "*";
                }

                if (segment.Strike)
                {
                    wrapped = "~~" + wrapped + "~~";
                }
            }

            if (!string.IsNullOrEmpty(segment.Url))
            {
                string url = segment.Url.Replace(" ", "%20").Replace("(", "%28").Replace(")", "%29");
                wrapped = "[" + wrapped + "](" + url + ")";
            }

            return leading + wrapped + trailing;
        }

        private static string RenderPlain(List<Segment> segments)
        {
            string text = string.Concat(segments.Select(s => s.Text));
            text = Regex.Replace(text, @"\s+", " ").Trim();
            return EscapeInline(text);
        }

        // ---------------------------------------------------------------- escaping

        private static string WrapInlineCode(string text)
        {
            string fence = new string('`', LongestBacktickRun(text) + 1);
            bool pad = text.StartsWith("`", StringComparison.Ordinal) || text.EndsWith("`", StringComparison.Ordinal);

            return pad ? fence + " " + text + " " + fence : fence + text + fence;
        }

        private static int LongestBacktickRun(string text)
        {
            int longest = 0;
            int current = 0;

            foreach (char c in text)
            {
                current = c == '`' ? current + 1 : 0;
                longest = Math.Max(longest, current);
            }

            return longest;
        }

        /// <summary>
        /// Escapes characters that Markdown would interpret inside a line of text.
        /// </summary>
        private static string EscapeInline(string text)
        {
            var sb = new StringBuilder(text.Length + 8);

            foreach (char c in text)
            {
                if (c == '\\' || c == '*' || c == '`' || c == '[' || c == ']' || c == '<')
                {
                    sb.Append('\\');
                }

                sb.Append(c);
            }

            return UnderscoreAtWordEdge.Replace(sb.ToString(), "\\_");
        }

        /// <summary>
        /// Escapes constructs that Markdown would interpret at the beginning of a line.
        /// </summary>
        private static string EscapeLineStarts(string text)
        {
            var lines = text.Split('\n');

            for (int i = 0; i < lines.Length; i++)
            {
                lines[i] = LineStart.Replace(lines[i], match =>
                {
                    if (match.Groups["num"].Success)
                    {
                        return match.Groups["num"].Value + "\\" + match.Groups["dot"].Value;
                    }

                    return "\\" + match.Value;
                });
            }

            return string.Join("\n", lines);
        }

        // ---------------------------------------------------------------- tables

        private string RenderTable(W.Table table)
        {
            var rows = new List<List<string>>();

            foreach (var row in table.Elements<W.TableRow>())
            {
                var cells = new List<string>();

                foreach (var cell in row.Elements<W.TableCell>())
                {
                    cells.Add(RenderCell(cell));

                    // A merged cell spans several grid columns: pad to keep the columns aligned
                    int span = cell.TableCellProperties?.GridSpan?.Val?.Value ?? 1;
                    for (int i = 1; i < span; i++)
                    {
                        cells.Add(string.Empty);
                    }
                }

                rows.Add(cells);
            }

            if (rows.Count == 0)
            {
                return string.Empty;
            }

            int columns = rows.Max(r => r.Count);
            if (columns == 0)
            {
                return string.Empty;
            }

            var sb = new StringBuilder();
            for (int i = 0; i < rows.Count; i++)
            {
                var cells = rows[i];
                while (cells.Count < columns)
                {
                    cells.Add(string.Empty);
                }

                sb.Append("| ").Append(string.Join(" | ", cells)).Append(" |");

                if (i == 0)
                {
                    sb.Append("\n|").Append(string.Concat(Enumerable.Repeat(" --- |", columns)));
                }

                if (i < rows.Count - 1)
                {
                    sb.Append("\n");
                }
            }

            return sb.ToString();
        }

        private string RenderCell(W.TableCell cell)
        {
            var lines = new List<string>();

            foreach (var paragraph in cell.Elements<W.Paragraph>())
            {
                var segments = new List<Segment>();
                CollectSegments(paragraph, segments, null);

                string text = RenderSegments(segments).Replace("\n", " ").Trim();
                if (text.Length > 0)
                {
                    lines.Add(text);
                }
            }

            return string.Join(" ", lines).Replace("|", "\\|");
        }
    }
}
