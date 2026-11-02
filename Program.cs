// Topic: Parse Office/PDF documents to Markdown and split the text into heading-based chunks for RAG pipelines.
// Uses GroupDocs.Parser for .NET (GetFormattedText Markdown mode, with GetText fallback when formatted extraction is unavailable).

using System.Text;
using GroupDocs.Parser;
using GroupDocs.Parser.Options;

namespace Demo.ParseMarkdownChunks;

internal static class Program
{
    private const string DocsFolder = "documents";
    private const string ResultFolder = "Result";

    private static readonly string SourceDocx = Path.Combine(DocsFolder, "sample.docx");
    private static readonly string SourcePdf = Path.Combine(DocsFolder, "sample.pdf");

    private static int Main()
    {
        Directory.CreateDirectory(DocsFolder);
        Directory.CreateDirectory(ResultFolder);
        ApplyLicense();

        if (!File.Exists(SourceDocx))
        {
            Console.Error.WriteLine($"Missing source document: {Path.GetFullPath(SourceDocx)}");
            return 1;
        }

        string infoPath = Path.Combine(ResultFolder, "document-info.txt");
        GetDocumentInfo(SourceDocx, infoPath);
        Console.WriteLine($"Document info: {Path.GetFullPath(infoPath)}");

        string markdownPath = Path.Combine(ResultFolder, "sample.md");
        string? markdown = ExtractMarkdown(SourceDocx, markdownPath);
        if (markdown == null)
        {
            Console.Error.WriteLine("Markdown extraction returned null for the DOCX sample.");
            return 2;
        }

        Console.WriteLine($"Markdown written: {Path.GetFullPath(markdownPath)} ({markdown.Length} chars)");

        string chunksDir = Path.Combine(ResultFolder, "chunks");
        int chunkCount = ChunkMarkdownByHeadings(markdown, chunksDir);
        Console.WriteLine($"Chunks written: {chunkCount} -> {Path.GetFullPath(chunksDir)}");

        // PDF often has no formatted Markdown path; demonstrate the plain-text fallback.
        if (File.Exists(SourcePdf))
        {
            string pdfTextPath = Path.Combine(ResultFolder, "sample-pdf.txt");
            string? plain = ExtractPlainTextFallback(SourcePdf, pdfTextPath);
            Console.WriteLine(plain == null
                ? "PDF plain-text fallback: not supported for this file"
                : $"PDF plain-text fallback: {Path.GetFullPath(pdfTextPath)} ({plain.Length} chars)");
        }

        return chunkCount > 0 ? 0 : 3;
    }

    private static void ApplyLicense()
    {
        // Point this at your .lic file to remove evaluation limits.
        // Get a free temporary licence: https://purchase.groupdocs.com/temporary-license
        const string licensePath = "REPLACE_WITH_YOUR_LICENSE_PATH";
        if (File.Exists(licensePath))
        {
            new License().SetLicense(licensePath);
            Console.WriteLine("[license] applied");
        }
        else
        {
            Console.WriteLine("[license] no licence set - running in evaluation mode");
        }
    }

    /// <summary>
    /// Writes basic document facts (file type, page count, size) for a source file.
    /// </summary>
    /// <remarks>
    /// Calls <see cref="Parser.GetDocumentInfo"/> so a RAG pipeline can decide how to
    /// chunk or skip a file before paying for full text extraction.
    /// </remarks>
    private static void GetDocumentInfo(string sourcePath, string outputPath)
    {
        using var parser = new Parser(sourcePath);
        var info = parser.GetDocumentInfo();

        var sb = new StringBuilder();
        sb.AppendLine($"Source: {Path.GetFileName(sourcePath)}");
        sb.AppendLine($"File type: {info?.FileType}");
        sb.AppendLine($"Pages: {info?.PageCount}");
        sb.AppendLine($"Size: {info?.Size} bytes");
        File.WriteAllText(outputPath, sb.ToString());
    }

    /// <summary>
    /// Extracts the document as Markdown and saves it to disk.
    /// </summary>
    /// <remarks>
    /// Uses <c>GetFormattedText</c> with <see cref="FormattedTextMode.Markdown"/> so
    /// headings, lists, emphasis, hyperlinks and tables stay structured for embedding.
    /// Returns null when the format does not support formatted extraction.
    /// </remarks>
    private static string? ExtractMarkdown(string sourcePath, string outputPath)
    {
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
    }

    /// <summary>
    /// Extracts plain text when Markdown/formatted extraction is unavailable.
    /// </summary>
    /// <remarks>
    /// PDF and some other formats return null from formatted Markdown extraction.
    /// <c>GetText</c> still yields a usable text layer for indexing or chunking.
    /// </remarks>
    private static string? ExtractPlainTextFallback(string sourcePath, string outputPath)
    {
        using var parser = new Parser(sourcePath);
        using TextReader? reader = parser.GetText();
        if (reader == null)
        {
            return null;
        }

        string text = reader.ReadToEnd();
        File.WriteAllText(outputPath, text, Encoding.UTF8);
        return text;
    }

    /// <summary>
    /// Splits Markdown into heading-based chunks and writes one file per chunk.
    /// </summary>
    /// <remarks>
    /// Splits on ATX headings (<c>#</c> / <c>##</c>) so each chunk maps to a section
    /// suitable for embedding and retrieval-augmented generation (RAG).
    /// </remarks>
    private static int ChunkMarkdownByHeadings(string markdown, string outputDirectory)
    {
        Directory.CreateDirectory(outputDirectory);
        foreach (string stale in Directory.EnumerateFiles(outputDirectory, "chunk-*.md"))
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
            string path = Path.Combine(outputDirectory, $"chunk-{i + 1:00}.md");
            File.WriteAllText(path, chunks[i], Encoding.UTF8);
        }

        return chunks.Count;
    }
}
