# Markdown Chunks for Retrieval

[![Docs](https://img.shields.io/badge/Docs-2865E0?style=for-the-badge&logo=Hugo&logoColor=white)](https://docs.groupdocs.com/parser/net/)
[![Blog](https://img.shields.io/badge/Blog-2865E0?style=for-the-badge&logo=WordPress&logoColor=white)](https://blog.groupdocs.com/categories/groupdocs.parser-product-family/)
[![Free Support](https://img.shields.io/badge/Free%20Support-2865E0?style=for-the-badge&logo=Discourse&logoColor=white)](https://forum.groupdocs.com/c/parser/17)
[![Temporary License](https://img.shields.io/badge/Temporary%20License-2865E0?style=for-the-badge&logo=rocket&logoColor=white)](https://purchase.groupdocs.com/temp-license/100308)

## Overview

`parse-docx-to-markdown-chunks-dotnet` is a runnable .NET 8 console demo that parses Word DOCX to Markdown with GroupDocs.Parser 26.10.0, falls back to plain text for PDF, and splits heading sections into `chunk-NN.md` files for a retrieval pipeline. Built with C# and the `GetFormattedText` Markdown mode, the sample shows how a service can inspect a file, extract structured text, and emit document chunks without a separate Office converter. Designed for backend developers who need Markdown extraction and heading split in the same process.

I wired this after a search index swallowed a three-page DOCX as one embedding and then ranked the wrong section for every query.

## Technology Stack

- **Platform**: .NET 8.0
- **Product**: GroupDocs.Parser 26.10.0
- **Language**: C#
- **Framework**: .NET SDK-style console project

## Problem Statement

Developers who feed Office and PDF files into embedding indexes often receive one giant string per file. That string mixes cover pages, policy clauses, and footnotes, so cosine search cannot tell which heading answered the question. Building a custom DOCX-to-Markdown path means reading Open XML styles, mapping list levels, and still failing when the next upload is a scanned-looking PDF with only a text layer.

A second pain point is format coverage: `GetFormattedText` in Markdown mode returns structured headings for Word, yet many PDF files return null from formatted extraction. Without a `GetText` fallback, the pipeline either crashes or silently drops those files. Teams also want cheap facts (file type, page count, size) before paying for a full extract.

GroupDocs.Parser for .NET answers these needs with `Parser.GetDocumentInfo`, `GetFormattedText(FormattedTextMode.Markdown)`, `GetText`, and a small heading splitter that writes one Markdown file per `#` / `##` section.

## Solution Overview

GroupDocs.Parser addresses these challenges through a single `Parser` instance per source path plus explicit Markdown and plain-text extraction options. Key technical features include:

- **Document facts first**: `GetDocumentInfo` reports file type, page count, and size so you can skip or batch before extraction.
- **Markdown extraction**: `GetFormattedText` with `FormattedTextMode.Markdown` keeps headings, lists, and emphasis as Markdown.
- **PDF text fallback**: `GetText` still yields a usable text layer when Markdown mode returns null.
- **Heading split**: `ChunkMarkdownByHeadings` writes `chunk-01.md`, `chunk-02.md`, and so on from ATX headings.
- **On-prem .NET 8**: NuGet package `GroupDocs.Parser` 26.10.0 with sample DOCX and PDF under `documents/`.

## When should I use this demo?

Use this demo when a retrieval pipeline must turn Word DOCX into Markdown sections and still accept PDF through a plain-text path. Prefer `ExtractMarkdown` for Office files that support formatted extraction, then call `ExtractPlainTextFallback` when the reader is null. Always run `GetDocumentInfo` first if page count or size can short-circuit work, and finish with `ChunkMarkdownByHeadings` so each embedding maps to one `#` or `##` block instead of the whole file.

## Prerequisites

Before running these examples, ensure you have:

- **Runtime** – .NET 8 SDK
- **Package** – GroupDocs.Parser 26.10.0 (restored from the `.csproj`)
- **License (optional)** – point `REPLACE_WITH_YOUR_LICENSE_PATH` at a valid `.lic`, or request a [temporary license](https://purchase.groupdocs.com/temp-license/100308)
- **Sample input** – `documents/sample.docx` and `documents/sample.pdf` ship with the repo

## Getting Started

### Installation

**Using Package Manager:**

```bash
dotnet add package GroupDocs.Parser --version 26.10.0
```

**Manual Installation:**

Clone this repository, open `ParseMarkdownChunks.csproj`, and restore packages with `dotnet restore`.

### Configuration

1. Optionally set the license path constant in `Program.cs` to your `.lic` file.
2. Confirm `documents/sample.docx` (and optionally `sample.pdf`) are present.
3. Run `dotnet run --project ParseMarkdownChunks.csproj -c Release`.

## Repository Structure

```
parse-docx-to-markdown-chunks-dotnet/
│
├── .gitignore
├── ParseMarkdownChunks.csproj
├── Program.cs
├── documents/
│   ├── sample.docx
│   └── sample.pdf
└── Result/
    ├── document-info.txt
    ├── sample.md
    ├── sample-pdf.txt
    └── chunks/
        ├── chunk-01.md
        ├── chunk-02.md
        ├── chunk-03.md
        └── chunk-04.md
```

### File Descriptions

- **ParseMarkdownChunks.csproj** – .NET 8 project referencing GroupDocs.Parser 26.10.0
- **Program.cs** – Console entry point and the four feature methods
- **documents/sample.docx** – Word sample used for Markdown extraction
- **documents/sample.pdf** – PDF sample used for plain-text fallback
- **Result/** – Generated info, Markdown, PDF text, and heading chunks

## Code Implementation

### Implementation: GetDocumentInfo

Writes basic document facts (file type, page count, size) so a retrieval pipeline can decide how to chunk or skip before full extraction.

```csharp
using var parser = new Parser(sourcePath);
var info = parser.GetDocumentInfo();

var sb = new StringBuilder();
sb.AppendLine($"Source: {Path.GetFileName(sourcePath)}");
sb.AppendLine($"File type: {info?.FileType}");
sb.AppendLine($"Pages: {info?.PageCount}");
sb.AppendLine($"Size: {info?.Size} bytes");
File.WriteAllText(outputPath, sb.ToString());
```

#### Technical Details

`Parser.GetDocumentInfo` returns an info object without reading the full text layer. On the shipped `sample.docx` the demo writes a 3-page Microsoft Word Open XML Document entry of 14820 bytes into `Result/document-info.txt`.

**Key Components:**
- `Parser`: Opens the source path
- `GetDocumentInfo`: Returns file type, page count, and size
- `StringBuilder`: Formats the facts for disk

**Parameters:**
- `sourcePath`: Input DOCX or PDF path
- `outputPath`: Text file for the facts dump

**Output:** `Result/document-info.txt` with source name, type, pages, and size.

---

### Implementation: ExtractMarkdown

Extracts the document as Markdown via `GetFormattedText(FormattedTextMode.Markdown)` and saves UTF-8 text. Returns null when the format does not support formatted extraction.

```csharp
using var parser = new Parser(sourcePath);
using TextReader? reader = parser.GetFormattedText(
    new FormattedTextOptions(FormattedTextMode.Markdown));

if (reader == null)
{
    return null;
}

string markdown = reader.ReadToEnd();
File.WriteAllText(outputPath, markdown, Encoding.UTF8);
return markdown;
```

#### Technical Details

Markdown mode maps Word heading styles to ATX markers so later splits follow author sections. Dispose the `TextReader` promptly; the sample writes `Result/sample.md` for the DOCX path.

**Key Components:**
- `FormattedTextOptions`: Selects Markdown mode
- `GetFormattedText`: Returns a reader or null
- `Encoding.UTF8`: Keeps Markdown portable for chunk files

**Parameters:**
- `sourcePath`: Usually a Word DOCX
- `outputPath`: Full Markdown file path

**Output:** Markdown string and `Result/sample.md`, or null.

---

### Implementation: ExtractPlainTextFallback

Calls `GetText` when Markdown or formatted extraction is unavailable. PDF samples typically take this path.

```csharp
using var parser = new Parser(sourcePath);
using TextReader? reader = parser.GetText();
if (reader == null)
{
    return null;
}

string text = reader.ReadToEnd();
File.WriteAllText(outputPath, text, Encoding.UTF8);
return text;
```

#### Technical Details

Use this after a null Markdown reader so PDF text still reaches the index. The demo writes `Result/sample-pdf.txt` for `documents/sample.pdf`.

**Key Components:**
- `GetText`: Plain text layer extraction
- `TextReader`: Streams the extracted text

**Parameters:**
- `sourcePath`: PDF or other format without Markdown mode
- `outputPath`: Plain-text output path

**Output:** Plain text string and file, or null.

---

### Implementation: ChunkMarkdownByHeadings

Splits on `#` / `##` headings into `chunk-NN.md` files for embedding-sized document chunks.

```csharp
Directory.CreateDirectory(outputDirectory);
foreach (string stale in Directory.EnumerateFiles(
    outputDirectory, "chunk-*.md"))
{
    File.Delete(stale);
}

var chunks = new List<string>();
var current = new StringBuilder();
using (var reader = new StringReader(markdown))
{
    string? line;
    while ((line = reader.ReadLine()) != null)
    {
        bool isHeading = line.StartsWith("# ", StringComparison.Ordinal)
            || line.StartsWith("## ", StringComparison.Ordinal);
        if (isHeading && current.Length > 0)
        {
            chunks.Add(current.ToString().TrimEnd());
            current.Clear();
        }
        current.AppendLine(line);
    }
}
```

```csharp
if (current.Length > 0)
{
    chunks.Add(current.ToString().TrimEnd());
}
if (chunks.Count == 0 && !string.IsNullOrWhiteSpace(markdown))
{
    chunks.Add(markdown.Trim());
}
for (int i = 0; i < chunks.Count; i++)
{
    string path = Path.Combine(
        outputDirectory, $"chunk-{i + 1:00}.md");
    File.WriteAllText(path, chunks[i], Encoding.UTF8);
}
return chunks.Count;
```

#### Technical Details

Only ATX `#` and `##` start a new chunk; deeper headings stay inside the current section. The sample DOCX produces four files under `Result/chunks/`.

**Key Components:**
- `StringReader`: Line walk over Markdown
- `chunk-{i:00}.md`: Zero-padded chunk file names

**Parameters:**
- `markdown`: Full Markdown string from extraction
- `outputDirectory`: Folder for chunk files

**Output:** Integer chunk count and `Result/chunks/chunk-*.md`.

## Best Practices

When implementing Markdown Chunks for Retrieval, consider these best practices:

- **Inspect before extract**: Call `GetDocumentInfo` and skip empty or oversized files by page count and size.
- **Null-check Markdown readers**: Treat a null `GetFormattedText` result as a signal to call `GetText`, not as a hard failure.
- **Prefer heading chunks over fixed tokens**: Split on `#` / `##` so embeddings match author sections.
- **Clear stale chunks**: Delete old `chunk-*.md` files before writing a new set for the same document.
- **Dispose parsers and readers**: Keep `using` scopes tight in long-running hosts.
- **License for production volume**: Evaluation mode can watermark or limit output; set a real `.lic` for batch jobs.

## Additional Resources

For more in-depth information about Markdown extraction and document chunks with GroupDocs.Parser for .NET, explore these technical resources:

* **Parse documents to Markdown chunks (docs use case)** – Comparison of inspect, Markdown extract, PDF fallback, and heading split: [Read the article →](https://docs.groupdocs.com/parser/net/use-cases/parse-documents-to-markdown-chunks/)

* **Parse documents to Markdown chunks (.NET blog)** – Tutorial comparing the four methods with runnable C# samples: [Read the article →](https://blog.groupdocs.com/parser/parse-documents-to-markdown-chunks-net/)

* **Extract formatted text from documents** – API guide for `GetFormattedText` modes including Markdown: [Read the article →](https://docs.groupdocs.com/parser/net/extract-formatted-text-from-documents/)

* **Markdown extraction with GroupDocs.Parser** – Product docs on Markdown output: [Read the article →](https://docs.groupdocs.com/parser/net/markdown/)

* **Prepare Documents for RAG in Java** – Related blog on claim files to heading-scoped Markdown chunks (Java): [Read the article →](https://blog.groupdocs.com/parser/prepare-documents-rag-markdown-chunks-java/)

* **GroupDocs.Parser for .NET examples** – Broader sample projects on GitHub: [Read the article →](https://github.com/groupdocs-parser/GroupDocs.Parser-for-.NET)

## Keywords

`GroupDocs.Parser`, `.NET`, `Markdown extraction`, `GetFormattedText`, `document chunks`, `heading split`, `PDF text fallback`, `Word DOCX parse`, `C# parser`, `retrieval pipeline`, `FormattedTextMode.Markdown`, `GetText`, `GetDocumentInfo`, `chunk-NN.md`, `ATX headings`, `embedding index`, `.NET 8`, `NuGet GroupDocs.Parser`, `plain text fallback`, `section-based chunking`

## Support

For technical support, visit:
- [Free Support Forum](https://forum.groupdocs.com/c/parser/17)
- [Product Documentation](https://docs.groupdocs.com/parser/net/)
- [Get Temporary License](https://purchase.groupdocs.com/temp-license/100308)
