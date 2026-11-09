# Markdown Import for Word and PowerPoint

A Visual Studio extension that converts Word (`.docx`) and PowerPoint (`.pptx`) files to Markdown and adds the result to your project, straight from Solution Explorer.

<!-- TODO: add the Marketplace badge and link once the extension is published -->
<!-- TODO: add a screenshot or GIF of the context menu in action -->

## Features

- Right-click any folder in Solution Explorer and choose **Add Markdown file from Word** or **Add Markdown file from PowerPoint**.
- Pick a `.docx` or `.pptx` file. The extension creates a `.md` file with the same name in that folder and adds it to the project.
- Preserves the structure of the source: headings, lists, tables, links, text formatting and code (see [What gets converted](#what-gets-converted)).
- Escapes Markdown special characters found in the source text, so the output renders the way the author intended.
- Asks before overwriting an existing `.md` file.
- Built on the official [Open XML SDK](https://github.com/dotnet/Open-XML-SDK). Microsoft Office does not need to be installed.
- Writes UTF-8 files without a byte order mark.

## Requirements

- Visual Studio 2022 version 17.14 or later, or Visual Studio 2026 (64-bit).
- `.docx` and `.pptx` files. The legacy binary formats (`.doc` and `.ppt`) are not supported.

## Installation

- **Marketplace**: <!-- TODO: add the Visual Studio Marketplace link -->
- **Manual**: download the `.vsix` file from the [Releases](../../releases) page and double-click it.

## Usage

1. In Solution Explorer, right-click a folder.
2. Choose **Add Markdown file from Word** or **Add Markdown file from PowerPoint**.
3. Select the file in the dialog.
4. The Markdown file is created in the folder and added to the project.

## What gets converted

### Word (`.docx`)

| Word | Markdown |
| --- | --- |
| Title, Heading 1 to 9, or any paragraph with an outline level | `#` headings (up to six levels) |
| Bulleted and numbered lists, including nested levels | `-` and `1.` lists with indentation |
| Bold, italic, strikethrough | `**bold**`, `*italic*`, `~~strikethrough~~` |
| Hyperlinks | `[text](url)` |
| Manual line breaks | Line breaks inside the paragraph |
| Monospace fonts (Consolas, Courier New and similar) or code styles | Inline code, or fenced code blocks for whole paragraphs |
| Quote and Block Text styles | `>` quotes |
| Tables | Markdown tables (the first row becomes the header row) |

### PowerPoint (`.pptx`)

| PowerPoint | Markdown |
| --- | --- |
| Slide title | `#` heading |
| Slide without a title | `<!-- Slide N -->` comment |
| Body text | Bulleted list, indented according to the paragraph level |
| Numbered paragraphs | `1.` lists |
| Bold, italic, hyperlinks | `**bold**`, `*italic*`, `[text](url)` |
| Tables | Markdown tables |
| Slide boundaries | `---` |

### Example

A Word document with a heading, formatted text and inline code:

```markdown
# Inline formatting

Normal text, **bold**, *italic* and ***both***, ~~struck~~.

A [link to Microsoft](https://www.microsoft.com) in the middle of a sentence, then `Console.WriteLine()` as inline code.
```

## Known limitations

**Both formats**

- Images, and the content of text boxes, are not converted.
- Conversion runs on the Visual Studio UI thread, so very large files may freeze the IDE for a moment.

**Word**

- Footnotes, endnotes, headers, footers and comments are ignored.
- Deleted text in tracked changes is ignored.
- Tables nested inside table cells are ignored.
- Hyperlinks stored as complex fields (instead of regular hyperlinks) are kept as plain text.
- Title and Heading 1 both map to a level 1 heading.

**PowerPoint**

- Charts and SmartArt are ignored, as are footer, date and slide number placeholders.
- Speaker notes are not included.
- Text in body placeholders is treated as a bulleted list unless the paragraph explicitly turns bullets off. Bullet settings inherited from the slide layout or master are not resolved.
- Text follows the order of the shapes in the file, which can differ from the visual order on the slide.
- Hidden slides are included.

## Building from source

Prerequisites: Visual Studio with the **Visual Studio extension development** workload.

1. Clone the repository and open the solution.
2. Restore the NuGet packages (`DocumentFormat.OpenXml`, `Microsoft.VisualStudio.SDK`, `Microsoft.VSSDK.BuildTools`).
3. Press **F5** to start the Experimental Instance of Visual Studio with the extension loaded.

### Project layout

| File | Purpose |
| --- | --- |
| `WordToMarkdownConverter.cs` | `.docx` to Markdown conversion |
| `PowerPointToMarkdownConverter.cs` | `.pptx` to Markdown conversion |
| `MarkdownImportHelper.cs` | Shared flow: selected folder, file dialog, saving the file, adding it to the project |
| `AddMarkdownFromWordCommand.cs` | Command handler for the Word command |
| `AddMarkdownFromPowerPointCommand.cs` | Command handler for the PowerPoint command |
| `*.vsct` | Command table: menu placement, texts and icons |
| `*Package.cs` | Package class that registers the commands |

## Contributing

Bug reports and feature requests are welcome in the [issue tracker](../../issues). If a conversion does not look right, please attach a small sample file that reproduces the problem (without confidential content).

## Author

Alessandro Del Sole
