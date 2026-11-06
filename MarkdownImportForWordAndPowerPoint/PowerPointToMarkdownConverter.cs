using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace MarkdownImportForWordAndPowerPoint
{
    /// <summary>
    /// Converts a .pptx file to Markdown using the Open XML SDK.
    /// Rules:
    ///  - slide title = "# title"
    ///  - body text = bulleted list, indented according to the paragraph level
    ///  - tables = Markdown tables
    ///  - slides separated by "---"
    ///  - speaker notes optional (blockquote)
    /// Ignored: pictures, charts, SmartArt, footer text, date and slide number.
    /// Markdown special characters found in the slide text are escaped.
    /// </summary>
    internal static class PowerPointToMarkdownConverter
    {
        // An underscore at the start or end of a word could be read as emphasis; intra-word underscores are safe.
        private static readonly Regex UnderscoreAtWordEdge =
            new Regex(@"(?<![\p{L}\p{N}])_|_(?![\p{L}\p{N}])", RegexOptions.Compiled);

        // Constructs that Markdown would interpret at the beginning of a line (heading, quote, list, rule).
        private static readonly Regex LineStart = new Regex(
            @"^(?:(?<num>\d+)(?<dot>[.)])(?=\s)|(?<hash>#{1,6})(?=\s|$)|(?<quote>>)|(?<bullet>[-+])(?=\s)|(?<rule>(?:={2,}|-{2,}|~{3,})\s*$))",
            RegexOptions.Compiled);

        public static string ConvertFile(string pptxPath, bool includeNotes = false)
        {
            // FileShare.ReadWrite: allows reading even if the file is open in PowerPoint
            using (var stream = new FileStream(pptxPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var document = PresentationDocument.Open(stream, false))
            {
                var presentationPart = document.PresentationPart;
                var slideIds = presentationPart?.Presentation?.SlideIdList?.Elements<P.SlideId>();
                if (slideIds == null)
                {
                    return string.Empty;
                }

                var slides = new List<string>();
                int number = 0;

                foreach (var slideId in slideIds)
                {
                    number++;
                    var slidePart = presentationPart.GetPartById(slideId.RelationshipId.Value) as SlidePart;
                    if (slidePart == null)
                    {
                        continue;
                    }

                    slides.Add(ConvertSlide(slidePart, number, includeNotes));
                }

                return string.Join("\n\n---\n\n", slides) + "\n";
            }
        }

        // ---------------------------------------------------------------- slide

        private sealed class SlideContext
        {
            public SlideContext(SlidePart slidePart)
            {
                SlidePart = slidePart;
            }

            public SlidePart SlidePart { get; }

            public string Title { get; set; }

            public List<string> Blocks { get; } = new List<string>();
        }

        private static string ConvertSlide(SlidePart slidePart, int number, bool includeNotes)
        {
            var context = new SlideContext(slidePart);
            var shapeTree = slidePart.Slide?.CommonSlideData?.ShapeTree;
            if (shapeTree != null)
            {
                Walk(shapeTree, context);
            }

            var parts = new List<string>();

            if (!string.IsNullOrWhiteSpace(context.Title))
            {
                parts.Add("# " + context.Title);
            }
            else
            {
                parts.Add("<!-- Slide " + number + " -->");
            }

            parts.AddRange(context.Blocks);

            if (includeNotes)
            {
                string notes = ReadNotes(slidePart);
                if (!string.IsNullOrWhiteSpace(notes))
                {
                    parts.Add(notes);
                }
            }

            return string.Join("\n\n", parts);
        }

        private static void Walk(OpenXmlElement container, SlideContext context)
        {
            foreach (var child in container.ChildElements)
            {
                if (child is P.Shape shape)
                {
                    HandleShape(shape, context);
                }
                else if (child is P.GroupShape group)
                {
                    Walk(group, context);
                }
                else if (child is P.GraphicFrame frame)
                {
                    HandleGraphicFrame(frame, context);
                }
            }
        }

        private static string GetPlaceholderType(P.Shape shape)
        {
            var placeholder = shape.NonVisualShapeProperties?
                .ApplicationNonVisualDrawingProperties?
                .GetFirstChild<P.PlaceholderShape>();

            if (placeholder == null)
            {
                return null; // not a placeholder
            }

            string type = placeholder.Type?.InnerText;
            return string.IsNullOrEmpty(type) ? "obj" : type;
        }

        private static void HandleShape(P.Shape shape, SlideContext context)
        {
            string placeholderType = GetPlaceholderType(shape);
            bool bulletsByDefault = false;

            if (placeholderType != null)
            {
                switch (placeholderType)
                {
                    case "title":
                    case "ctrTitle":
                        if (context.Title == null)
                        {
                            context.Title = RenderTitle(shape.TextBody);
                            return;
                        }
                        break;

                    case "sldNum":
                    case "ftr":
                    case "dt":
                    case "hdr":
                        return; // decorative elements

                    case "body":
                    case "obj":
                        bulletsByDefault = true;
                        break;
                }
            }

            string text = RenderTextBody(shape.TextBody, context.SlidePart, bulletsByDefault);
            if (!string.IsNullOrWhiteSpace(text))
            {
                context.Blocks.Add(text);
            }
        }

        private static void HandleGraphicFrame(P.GraphicFrame frame, SlideContext context)
        {
            var table = frame.Descendants<A.Table>().FirstOrDefault();
            if (table == null)
            {
                return; // charts, SmartArt, OLE objects: ignored
            }

            string markdown = RenderTable(table, context.SlidePart);
            if (!string.IsNullOrWhiteSpace(markdown))
            {
                context.Blocks.Add(markdown);
            }
        }

        // ---------------------------------------------------------------- text

        private static string RenderTitle(P.TextBody body)
        {
            if (body == null)
            {
                return null;
            }

            var lines = body.Elements<A.Paragraph>()
                .Select(p => RenderParagraph(p, null, false).Replace("\n", " ").Trim())
                .Where(s => s.Length > 0);

            return string.Join(" ", lines);
        }

        private static string RenderTextBody(P.TextBody body, SlidePart slidePart, bool bulletsByDefault)
        {
            if (body == null)
            {
                return string.Empty;
            }

            var sb = new StringBuilder();
            bool lastWasList = false;
            int previousLevel = -1;

            foreach (var paragraph in body.Elements<A.Paragraph>())
            {
                string text = RenderParagraph(paragraph, slidePart, true).Trim();
                if (text.Length == 0)
                {
                    continue;
                }

                ClassifyBullet(paragraph, bulletsByDefault, out bool isList, out bool ordered);

                if (sb.Length > 0)
                {
                    sb.Append(lastWasList && isList ? "\n" : "\n\n");
                }

                if (isList)
                {
                    int level = paragraph.ParagraphProperties?.Level?.Value ?? 0;

                    if (!lastWasList)
                    {
                        previousLevel = -1;
                    }

                    // Avoid level jumps (e.g. 0 -> 2), which Markdown would interpret as a code block
                    level = Math.Max(0, Math.Min(level, previousLevel + 1));
                    previousLevel = level;

                    string indent = new string(' ', level * 4);
                    string marker = ordered ? "1. " : "- ";
                    string continuation = "  \n" + indent + new string(' ', marker.Length);

                    sb.Append(indent).Append(marker).Append(text.Replace("\n", continuation));
                }
                else
                {
                    previousLevel = -1;

                    // Plain paragraph: text starting with "#", "-", "1." etc. must not turn into a heading or list
                    sb.Append(EscapeLineStarts(text).Replace("\n", "  \n"));
                }

                lastWasList = isList;
            }

            return sb.ToString();
        }

        private static void ClassifyBullet(A.Paragraph paragraph, bool bulletsByDefault, out bool isList, out bool ordered)
        {
            var properties = paragraph.ParagraphProperties;
            isList = bulletsByDefault;
            ordered = false;

            if (properties == null)
            {
                return;
            }

            if (properties.GetFirstChild<A.NoBullet>() != null)
            {
                isList = false;
            }
            else if (properties.GetFirstChild<A.AutoNumberedBullet>() != null)
            {
                isList = true;
                ordered = true;
            }
            else if (properties.GetFirstChild<A.CharacterBullet>() != null ||
                     properties.GetFirstChild<A.PictureBullet>() != null)
            {
                isList = true;
            }
        }

        // ---------------------------------------------------------------- escaping

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

        // ---------------------------------------------------------------- inline formatting

        private sealed class Segment
        {
            public string Text { get; set; }

            public bool Bold { get; set; }

            public bool Italic { get; set; }

            public string Url { get; set; }
        }

        private static string RenderParagraph(A.Paragraph paragraph, SlidePart slidePart, bool applyFormatting)
        {
            var segments = new List<Segment>();

            foreach (var child in paragraph.ChildElements)
            {
                if (child is A.Run run)
                {
                    AddSegment(segments, run.Text?.Text, run.RunProperties, slidePart, applyFormatting);
                }
                else if (child is A.Field field)
                {
                    AddSegment(segments, field.Text?.Text, field.RunProperties, slidePart, applyFormatting);
                }
                else if (child is A.Break)
                {
                    if (segments.Count > 0)
                    {
                        segments[segments.Count - 1].Text += "\n";
                    }
                    else
                    {
                        segments.Add(new Segment { Text = "\n" });
                    }
                }
            }

            var sb = new StringBuilder();
            foreach (var segment in segments)
            {
                sb.Append(RenderSegment(segment));
            }

            return sb.ToString();
        }

        private static void AddSegment(
            List<Segment> segments,
            string text,
            A.RunProperties properties,
            SlidePart slidePart,
            bool applyFormatting)
        {
            if (string.IsNullOrEmpty(text))
            {
                return;
            }

            bool bold = false;
            bool italic = false;
            string url = null;

            if (applyFormatting && properties != null)
            {
                bold = properties.Bold?.Value == true;
                italic = properties.Italic?.Value == true;
                url = GetHyperlink(properties, slidePart);
            }

            // Merge adjacent segments with the same formatting to avoid "**a****b**"
            if (segments.Count > 0)
            {
                var last = segments[segments.Count - 1];
                if (last.Bold == bold && last.Italic == italic && last.Url == url)
                {
                    last.Text += text;
                    return;
                }
            }

            segments.Add(new Segment { Text = text, Bold = bold, Italic = italic, Url = url });
        }

        private static string GetHyperlink(A.RunProperties properties, SlidePart slidePart)
        {
            if (slidePart == null)
            {
                return null;
            }

            string id = properties.GetFirstChild<A.HyperlinkOnClick>()?.Id?.Value;
            if (string.IsNullOrEmpty(id))
            {
                return null;
            }

            var relationship = slidePart.HyperlinkRelationships.FirstOrDefault(r => r.Id == id);
            return relationship?.Uri?.OriginalString;
        }

        private static string RenderSegment(Segment segment)
        {
            string text = segment.Text;
            string core = text.Trim();

            if (core.Length == 0)
            {
                return text;
            }

            // Whitespace must stay outside the markers, otherwise Markdown does not recognize them
            string leading = text.Substring(0, text.Length - text.TrimStart().Length);
            string trailing = text.Substring(text.TrimEnd().Length);

            core = EscapeInline(core);

            string wrapped = core;
            if (segment.Bold && segment.Italic)
            {
                wrapped = "***" + core + "***";
            }
            else if (segment.Bold)
            {
                wrapped = "**" + core + "**";
            }
            else if (segment.Italic)
            {
                wrapped = "*" + core + "*";
            }

            if (!string.IsNullOrEmpty(segment.Url))
            {
                string url = segment.Url.Replace(" ", "%20").Replace("(", "%28").Replace(")", "%29");
                wrapped = "[" + wrapped + "](" + url + ")";
            }

            return leading + wrapped + trailing;
        }

        // ---------------------------------------------------------------- tables

        private static string RenderTable(A.Table table, SlidePart slidePart)
        {
            var rows = new List<List<string>>();

            foreach (var row in table.Elements<A.TableRow>())
            {
                var cells = new List<string>();
                foreach (var cell in row.Elements<A.TableCell>())
                {
                    cells.Add(RenderCell(cell, slidePart));
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

        private static string RenderCell(A.TableCell cell, SlidePart slidePart)
        {
            var paragraphs = cell.TextBody?.Elements<A.Paragraph>();
            if (paragraphs == null)
            {
                return string.Empty;
            }

            var lines = paragraphs
                .Select(p => RenderParagraph(p, slidePart, true).Replace("\n", " ").Trim())
                .Where(s => s.Length > 0);

            return string.Join(" ", lines).Replace("|", "\\|");
        }

        // ---------------------------------------------------------------- speaker notes

        private static string ReadNotes(SlidePart slidePart)
        {
            var shapeTree = slidePart.NotesSlidePart?.NotesSlide?.CommonSlideData?.ShapeTree;
            if (shapeTree == null)
            {
                return null;
            }

            var texts = new List<string>();

            foreach (var shape in shapeTree.Descendants<P.Shape>())
            {
                if (GetPlaceholderType(shape) != "body")
                {
                    continue;
                }

                string text = RenderTextBody(shape.TextBody, null, false);
                if (!string.IsNullOrWhiteSpace(text))
                {
                    texts.Add(text);
                }
            }

            if (texts.Count == 0)
            {
                return null;
            }

            var quoted = string.Join("\n\n", texts)
                .Split('\n')
                .Select(line => line.Length == 0 ? ">" : "> " + line);

            return "> **Notes:**\n>\n" + string.Join("\n", quoted);
        }
    }
}
