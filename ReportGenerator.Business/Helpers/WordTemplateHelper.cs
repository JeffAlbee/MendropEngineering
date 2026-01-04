using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using PDFtoImage;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.PixelFormats;
using SixLabors.ImageSharp.Processing;
using System.Text;
using System.Text.RegularExpressions;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using WColor = DocumentFormat.OpenXml.Wordprocessing.Color;

namespace ReportGenerator.Business.Helpers
{
    public static class WordTemplateHelper
    {
        #region Private Fields

        private static readonly Regex PlaceholderRegex = new(@"\{\{([A-Za-z0-9_]+)\}\}");

        private const bool DEBUG_HIGHLIGHT_REPLACED_FIELDS = false;

        #endregion Private Fields

        #region Public Methods

        public static List<string> GetPlaceholders(byte[] templateBytes)
        {
            var placeholders = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            using var memStream = new MemoryStream(templateBytes);

            using var doc = WordprocessingDocument.Open(memStream, false);

            var mainPart = doc.MainDocumentPart
                    ?? throw new Exception("MainDocumentPart not found in template.");

            // Combine *all* text from the document body
            var fullText = GetFullText(mainPart.Document.Body!);

            // Also include header/footer text
            foreach (var header in mainPart.HeaderParts)
                fullText += "\n" + GetFullText(header.Header);

            foreach (var footer in mainPart.FooterParts)
                fullText += "\n" + GetFullText(footer.Footer);

            // Find all {{placeholders}}
            foreach (Match match in PlaceholderRegex.Matches(fullText))
            {
                placeholders.Add(match.Groups[1].Value);
            }

            return placeholders.ToList();
        }

        public static byte[] ReplacePlaceholders(byte[] templateBytes, Dictionary<string, object?> values)
        {
            using var memStream = new MemoryStream();

            memStream.Write(templateBytes, 0, templateBytes.Length);

            memStream.Position = 0;

            using (var doc = WordprocessingDocument.Open(memStream, true))
            {
                var mainPart = doc.MainDocumentPart
                    ?? throw new Exception("MainDocumentPart not found in template.");

                // Split values into text & images
                var textValues = new Dictionary<string, string?>();
                var imageValues = new Dictionary<string, byte[]>();
                var pdfImageValues = new Dictionary<string, List<byte[]>>();
                var bulletValues = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);

                foreach (var kvp in values)
                {
                    switch (kvp.Value)
                    {
                        case byte[] bytes:
                            imageValues[kvp.Key] = bytes;
                            break;

                        case string s when IsImagePath(s):
                            if (File.Exists(s))
                            {
                                imageValues[kvp.Key] = File.ReadAllBytes(s);
                            }
                            else
                            {
                                textValues[kvp.Key] = $"[[IMAGE_NOT_FOUND::{kvp.Key}]]";
                            }
                            break;

                        case string s when Path.GetExtension(s).Equals(".pdf", StringComparison.OrdinalIgnoreCase):
                            if (File.Exists(s))
                            {
                                var pdfBytes = File.ReadAllBytes(s);
                                var pages = ConvertPdfToImages(pdfBytes);

                                if (pages.Count > 0)
                                    pdfImageValues[kvp.Key] = pages;
                                else
                                    textValues[kvp.Key] = $"[[PDF_CONVERSION_FAILED::{kvp.Key}]]";
                            }
                            else
                            {
                                textValues[kvp.Key] = $"[[PDF_NOT_FOUND::{kvp.Key}]]";
                            }
                            break;

                        case IEnumerable<string> list when kvp.Value is not string:
                            bulletValues[kvp.Key] = list.Where(x => !string.IsNullOrWhiteSpace(x)).ToList();
                            break;

                        case string s:
                            textValues[kvp.Key] = s;
                            break;

                        default:
                            textValues[kvp.Key] = kvp.Value?.ToString();
                            break;
                    }
                }

                var textOnlyValues = textValues
                    .Where(kv => !imageValues.ContainsKey(kv.Key) && !bulletValues.ContainsKey(kv.Key))
                    .ToDictionary(kv => kv.Key, kv => kv.Value);

                var imageKeys = imageValues.Keys.ToHashSet(StringComparer.OrdinalIgnoreCase);

                // Replace text placeholders
                ReplaceInElement(doc, mainPart.Document.Body!, textOnlyValues, pdfImageValues, bulletValues, imageKeys);

                foreach (var section in mainPart.Document.Body!.Elements<SectionProperties>())
                {
                    foreach (var headerRef in section.Elements<HeaderReference>())
                    {
                        var headerPart = (HeaderPart?)mainPart.GetPartById(headerRef.Id!);
                        if(headerPart != null)
                        {
                            ReplaceInElement(doc,headerPart.Header, textOnlyValues, pdfImageValues, bulletValues, imageKeys);
                        }
                    }

                    foreach (var footerRef in section.Elements<FooterReference>())
                    {
                        var footerPart = (FooterPart?)mainPart.GetPartById(footerRef.Id!);
                        if (footerPart != null)
                        {
                            ReplaceInElement(doc, footerPart.Footer, textOnlyValues, pdfImageValues, bulletValues, imageKeys);
                        }
                    }
                }

                // Replace in ALL header parts
                foreach (var headerPart in mainPart.HeaderParts)
                {
                    ReplaceInElement(doc, headerPart.Header, textOnlyValues, pdfImageValues, bulletValues, imageKeys);
                }

                // Replace in ALL footer parts
                foreach (var footerPart in mainPart.FooterParts)
                {
                    ReplaceInElement(doc, footerPart.Footer, textOnlyValues, pdfImageValues, bulletValues, imageKeys);
                }

                // Replace image placeholders
                InsertImagesFromBytes(doc, imageValues);

                mainPart.Document.Save();
            }

            return PerformGlobalXmlReplacement(memStream.ToArray(), values);
        }

        #endregion Public Methods

        #region Private Methods

        private static void InsertImagesFromBytes(WordprocessingDocument doc, Dictionary<string, byte[]> imageValues)
        {
            var mainPart = doc.MainDocumentPart!;
            var paragraphs = mainPart.Document.Descendants<Paragraph>().ToList();

            foreach (var paragraph in paragraphs)
            {
                // Combine text of the whole paragraph (handles split runs)
                var allText = string.Concat(paragraph.Descendants<Text>().Select(t => t.Text));

                foreach (var kvp in imageValues)
                {
                    string placeholder = kvp.Key;
                    byte[] imageBytes = kvp.Value;

                    string token = $"{{{{{placeholder}}}}}";
                    int idx = allText.IndexOf(token, StringComparison.OrdinalIgnoreCase);
                    if (idx == -1)
                        continue;

                    // Remove all runs from the paragraph
                    paragraph.RemoveAllChildren<Run>();

                    // --- BEFORE TEXT ---
                    string before = allText[..idx];
                    if (!string.IsNullOrEmpty(before))
                    {
                        paragraph.AppendChild(
                            new Run(new Text(before)
                            {
                                Space = SpaceProcessingModeValues.Preserve
                            })
                        );
                    }

                    // --- CREATE NEW PARAGRAPH FOR IMAGE ---
                    var imgParagraph = new Paragraph(
                        new ParagraphProperties(
                            new ParagraphStyleId { Val = "Caption" } // ensures correct rendering in Word
                        )
                    );

                    // Insert image paragraph AFTER the current paragraph
                    paragraph.Parent!.InsertAfter(imgParagraph, paragraph);

                    // Add image to the new paragraph
                    AddImageToParagraph(doc, imgParagraph, placeholder, imageBytes);

                    // --- AFTER TEXT ---
                    string after = allText[(idx + token.Length)..];
                    if (!string.IsNullOrEmpty(after))
                    {
                        var afterParagraph = new Paragraph(
                            new Run(new Text(after)
                            {
                                Space = SpaceProcessingModeValues.Preserve
                            })
                        );

                        // Insert the "after" paragraph after the image paragraph
                        paragraph.Parent.InsertAfter(afterParagraph, imgParagraph);
                    }

                    break;
                }
            }
        }

        private static void AddImageToParagraph(WordprocessingDocument doc, Paragraph paragraph, string name, byte[] bytes)
        {
            var mainPart = doc.MainDocumentPart!;
            var imagePart = mainPart.AddImagePart(ImagePartType.Png);

            using (var stream = new MemoryStream(bytes))
                imagePart.FeedData(stream);

            const long MaxWidthEmu = (long)(6.5 * 914400); // 6.5 inches
            int pixelWidth, pixelHeight;

            using (var img = Image.Load<Rgba32>(bytes))
            {
                pixelWidth = img.Width;
                pixelHeight = img.Height;
            }

            long emuWidth = pixelWidth * 9525L;
            long emuHeight = pixelHeight * 9525L;

            if (emuWidth > MaxWidthEmu)
            {
                double ratio = (double)MaxWidthEmu / emuWidth;
                emuWidth = MaxWidthEmu;
                emuHeight = (long)(emuHeight * ratio);
            }

            string relId = mainPart.GetIdOfPart(imagePart);

            var element =
                new Drawing(
                    new DW.Inline(
                        new DW.Extent() { Cx = emuWidth, Cy = emuHeight },
                        new DW.EffectExtent() { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
                        new DW.DocProperties() { Id = (UInt32Value)1U, Name = name },
                        new DW.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks() { NoChangeAspect = true }),
                        new A.Graphic(
                            new A.GraphicData(
                                new PIC.Picture(
                                    new PIC.NonVisualPictureProperties(
                                        new PIC.NonVisualDrawingProperties() { Id = 0U, Name = name },
                                        new PIC.NonVisualPictureDrawingProperties()
                                    ),
                                    new PIC.BlipFill(
                                        new A.Blip() { Embed = relId },
                                        new A.Stretch(new A.FillRectangle())
                                    ),
                                    new PIC.ShapeProperties(
                                        new A.Transform2D(
                                            new A.Offset() { X = 0L, Y = 0L },
                                            new A.Extents() { Cx = emuWidth, Cy = emuHeight }
                                        ),
                                        new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }
                                    )
                                )
                            )
                            { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }
                        )
                    )
                    {
                        DistanceFromTop = 0U,
                        DistanceFromBottom = 0U,
                        DistanceFromLeft = 0U,
                        DistanceFromRight = 0U
                    });

            paragraph.AppendChild(new Run(element));
        }

        private static void ReplaceInElement(
            WordprocessingDocument doc,
            OpenXmlElement element, 
            Dictionary<string, string?> values,
            Dictionary<string, List<byte[]>> pdfImageValues,
            Dictionary<string, List<string>> bulletValues,
            HashSet<string>? imageKeys = null)
        {
            var paragraphs = element.Descendants<Paragraph>().ToList();

            foreach (var paragraph in paragraphs)
            {
                ReplaceTextPlaceholders(doc, paragraph, values, pdfImageValues, bulletValues, imageKeys);
            }
        }

        private static byte[] PerformGlobalXmlReplacement(byte[] docBytes, Dictionary<string, object?> values)
        {
            using var result = new MemoryStream();

            using (var zip = new System.IO.Compression.ZipArchive(
                new MemoryStream(docBytes),
                System.IO.Compression.ZipArchiveMode.Read))
            {
                using var newZip = new System.IO.Compression.ZipArchive(
                    result,
                    System.IO.Compression.ZipArchiveMode.Create,
                    true);

                foreach (var entry in zip.Entries)
                {
                    var newEntry = newZip.CreateEntry(entry.FullName);

                    using var entryStream = entry.Open();
                    using var newEntryStream = newEntry.Open();

                    if (!entry.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
                    {
                        entryStream.CopyTo(newEntryStream);
                        continue;
                    }

                    using var reader = new StreamReader(entryStream);
                    string xml = reader.ReadToEnd();

                    // Do final replacement on all tokens
                    foreach (var kvp in values)
                    {
                        if (kvp.Value is byte[] || kvp.Value is IEnumerable<string>)
                            continue;

                        string token = $"{{{{{kvp.Key}}}}}";
                        string replacement = kvp.Value?.ToString() ?? "";
                        xml = xml.Replace(token, replacement, StringComparison.OrdinalIgnoreCase);
                    }

                    using var writer = new StreamWriter(newEntryStream);
                    writer.Write(xml);
                }
            }

            return result.ToArray();
        }

        private static string GetFullText(OpenXmlElement element)
        {
            var sb = new StringBuilder();
            foreach (var text in element.Descendants<Text>())
            {
                sb.Append(text.Text);
            }
            return sb.ToString();
        }

        private static bool IsImagePath(string path)
        {
            var ext = Path.GetExtension(path)?.ToLowerInvariant();
            return ext is ".png" or ".jpg" or ".jpeg" or ".bmp";
        }

        private static void ReplaceTextPlaceholders(
            WordprocessingDocument doc,
            Paragraph paragraph,
            Dictionary<string, string?> values,
            Dictionary<string, List<byte[]>> pdfImageValues,
            Dictionary<string, List<string>> bulletValues,
            HashSet<string>? imageKeys)
        {
            var runs = paragraph.Elements<Run>().ToList();
            if (runs.Count == 0) return;

            // Build full paragraph text
            var buffer = new StringBuilder();
            var map = new List<(Run run, int start, int length)>();

            foreach (var run in runs)
            {
                var txt = run.GetFirstChild<Text>()?.Text ?? "";
                int start = buffer.Length;

                buffer.Append(txt);
                map.Add((run, start, txt.Length));
            }

            string fullText = buffer.ToString();

            // Match placeholders
            var matches = PlaceholderRegex.Matches(fullText);
            if (matches.Count == 0) return;

            // Replace from last → first
            for (int i = matches.Count - 1; i >= 0; i--)
            {
                var match = matches[i];
                string name = match.Groups[1].Value;

                // BULLET placeholder → replace paragraph with bullet paragraphs
                if (bulletValues != null && bulletValues.TryGetValue(name, out var bullets))
                {
                    InsertBulletParagraphs(doc, paragraph, bullets);
                    return;
                }

                // PDF placeholder → replace paragraph with PDF images
                if (pdfImageValues != null && pdfImageValues.ContainsKey(name))
                {
                    bool isLastPdf = IsLastPdfPlaceholder(i, matches, pdfImageValues);
                    InsertPdfImages(doc, paragraph, pdfImageValues[name], isLastPdf);
                    return;
                }

                // IMAGE placeholder → skip text replace
                if (imageKeys != null && imageKeys.Contains(name))
                    continue;

                bool hasValue = values.TryGetValue(name, out string? value) &&
                                !string.IsNullOrWhiteSpace(value);

                // Use either:
                string replacement = hasValue
                    ? value!
                    : $"<<{name}>>";  // missing placeholder marker

                var affectedRuns = ReplaceTextInRuns(map, match.Index, match.Length, replacement);

                // Apply formatting AFTER text replacement
                foreach (var run in affectedRuns)
                {
                    EnsureRunProperties(run);

                    // highlight replaced
                    if (hasValue)
                    {
                        if (DEBUG_HIGHLIGHT_REPLACED_FIELDS)
                        {
                            run.RunProperties.Highlight = new Highlight
                            {
                                Val = HighlightColorValues.LightGray
                            };
                        }
                    }
                    else
                    {
                        // missing value — red bold yellow highlight
                        run.RunProperties.Color = new WColor { Val = "FF0000" };
                        run.RunProperties.Bold = new Bold();
                        run.RunProperties.Highlight = new Highlight
                        {
                            Val = HighlightColorValues.Yellow
                        };
                    }
                }
            }
        }

        private static void InsertBulletParagraphs(WordprocessingDocument doc, Paragraph placeholderParagraph, List<string> bullets)
        {
            if (bullets == null || bullets.Count == 0)
            {
                placeholderParagraph.Remove();
                return;
            }

            var parent = placeholderParagraph.Parent;
            if (parent == null)
                return;

            int bulletNumId = EnsureBulletNumbering(doc);

            var basePPr = placeholderParagraph.ParagraphProperties != null
                ? (ParagraphProperties)placeholderParagraph.ParagraphProperties.CloneNode(true)
                : new ParagraphProperties();

            bool templateHasNumbering =
                basePPr.NumberingProperties?.NumberingId?.Val != null;

            // If the template placeholder paragraph is NOT actually a bullet/list item,
            // force bullets by creating numbering and applying NumberingProperties.
            int? forcedNumId = null;
            if (!templateHasNumbering)
                forcedNumId = EnsureBulletNumbering(doc);

            OpenXmlElement refNode = placeholderParagraph;

            foreach (var bullet in bullets)
            {
                Paragraph p;

                if (templateHasNumbering)
                {
                    // Clone template formatting (works if template paragraph truly has bullets)
                    p = new Paragraph();
                    p.AppendChild((ParagraphProperties)basePPr.CloneNode(true));
                }
                else
                {
                    // Force bullet formatting (works even if template isn't set correctly)
                    p = CreateBulletParagraph(bullet, forcedNumId!.Value);

                    // Optional: preserve style if your template uses a style like "ListParagraph"
                    if (basePPr.ParagraphStyleId != null)
                    {
                        var pPr = p.GetFirstChild<ParagraphProperties>() ?? p.PrependChild(new ParagraphProperties());
                        pPr.ParagraphStyleId = (ParagraphStyleId)basePPr.ParagraphStyleId.CloneNode(true);
                    }
                }

                // Add text if we created paragraph via template clone
                if (templateHasNumbering)
                {
                    p.RemoveAllChildren<Run>();
                    p.AppendChild(new Run(new Text(bullet) { Space = SpaceProcessingModeValues.Preserve }));
                }

                parent.InsertAfter(p, refNode);
                refNode = p;
            }

            placeholderParagraph.Remove();
        }

        private static Paragraph CreateBulletParagraph(string text, int numberingId)
        {
            return new Paragraph(
                new ParagraphProperties(
                    new NumberingProperties(
                        new NumberingLevelReference { Val = 0 },
                        new NumberingId { Val = numberingId }
                    )
                ),
                new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve })
            );
        }

        private static int EnsureBulletNumbering(WordprocessingDocument doc)
        {
            var mainPart = doc.MainDocumentPart!;

            // If the template already has numbering (it does, since headings are numbered),
            // we must NOT overwrite it.
            var numberingPart = mainPart.NumberingDefinitionsPart;

            if (numberingPart == null)
            {
                // No numbering part at all (rare for your template). Create it safely.
                numberingPart = mainPart.AddNewPart<NumberingDefinitionsPart>();
                numberingPart.Numbering = new Numbering();
            }
            else
            {
                // Numbering part exists → MUST be readable; otherwise do not overwrite.
                if (numberingPart.Numbering == null)
                    throw new InvalidOperationException(
                        "Template has a NumberingDefinitionsPart but it could not be loaded. " +
                        "Refusing to overwrite numbering.xml because it would remove heading/list numbering."
                    );
            }

            var numbering = numberingPart.Numbering;

            // 1) Reuse an existing bullet numbering instance if one exists
            var existingBulletNumId = FindExistingBulletNumberId(numbering);
            if (existingBulletNumId.HasValue)
                return existingBulletNumId.Value;

            // 2) Otherwise create a new bullet numbering definition (append only)
            int nextAbstractNumId =
                numbering.Elements<AbstractNum>()
                    .Select(a => (int)a.AbstractNumberId!.Value)
                    .DefaultIfEmpty(0)
                    .Max() + 1;

            int nextNumId =
                numbering.Elements<NumberingInstance>()
                    .Select(n => (int)n.NumberID!.Value)
                    .DefaultIfEmpty(0)
                    .Max() + 1;

            var abstractNum = new AbstractNum(
                new Level(
                    new StartNumberingValue { Val = 1 },
                    new NumberingFormat { Val = NumberFormatValues.Bullet },
                    new LevelText { Val = "•" },
                    new LevelJustification { Val = LevelJustificationValues.Left },
                    new ParagraphProperties(
                        new Indentation { Left = "360", Hanging = "180" }
                    )
                )
                { LevelIndex = 0 }
            )
            { AbstractNumberId = nextAbstractNumId };

            var numInstance = new NumberingInstance(new AbstractNumId { Val = nextAbstractNumId })
            { NumberID = nextNumId };

            numbering.Append(abstractNum);
            numbering.Append(numInstance);

            numbering.Save();
            return nextNumId;
        }

        private static int? FindExistingBulletNumberId(Numbering numbering)
        {
            foreach (var inst in numbering.Elements<NumberingInstance>())
            {
                var numId = (int)inst.NumberID!.Value;
                var absId = inst.GetFirstChild<AbstractNumId>()?.Val?.Value;
                if (absId == null) continue;

                var abs = numbering.Elements<AbstractNum>()
                    .FirstOrDefault(a => a.AbstractNumberId?.Value == absId.Value);

                var lvl0 = abs?.Elements<Level>().FirstOrDefault(l => l.LevelIndex?.Value == 0);
                var fmt = lvl0?.GetFirstChild<NumberingFormat>()?.Val?.Value;

                if (fmt == NumberFormatValues.Bullet)
                    return numId;
            }

            return null;
        }

        private static List<Run> ReplaceTextInRuns(
            List<(Run run, int start, int length)> map,
            int matchStart,
            int matchLength,
            string replacement)
        {
            int matchEnd = matchStart + matchLength;
            int replaceIndex = 0;

            var affectedRuns = new List<Run>();

            foreach (var (run, start, len) in map)
            {
                int runEnd = start + len;
                if (runEnd <= matchStart || start >= matchEnd)
                    continue;

                var textEl = run.GetFirstChild<Text>();
                if (textEl == null) continue;

                string old = textEl.Text;
                int localStart = Math.Max(0, matchStart - start);
                int localEnd = Math.Min(len, matchEnd - start);

                string newText =
                    old[..localStart]
                    + replacement
                    + old[localEnd..];

                textEl.Text = newText;

                affectedRuns.Add(run);

                // After first run, empty replacement for remaining runs
                replacement = "";
            }

            return affectedRuns;
        }

        private static void EnsureRunProperties(Run run)
        {
            if (run.RunProperties == null)
                run.RunProperties = new RunProperties();
            else
                run.RunProperties = (RunProperties)run.RunProperties.CloneNode(true);
        }

        private static List<byte[]> ConvertPdfToImages(byte[] pdfBytes)
        {
            using var pdfStream = new MemoryStream(pdfBytes);

            var bitmaps = Conversion.ToImages(pdfStream);

            var result = new List<byte[]>();

            foreach (var skBitmap in bitmaps)
            {
                using var bitmap = skBitmap;

                // Convert SKBitmap → ImageSharp Image<Rgba32>
                using var img = Image.LoadPixelData<Rgba32>(
                    bitmap.Bytes,
                    bitmap.Width,
                    bitmap.Height
                );

                // Resize BEFORE encoding
                img.Mutate(x => x.Resize(new ResizeOptions
                {
                    Size = new SixLabors.ImageSharp.Size(2200, 0),
                    Mode = ResizeMode.Max
                }));

                // Encode as JPEG for huge size reduction
                using var ms = new MemoryStream();

                img.Save(ms, new SixLabors.ImageSharp.Formats.Jpeg.JpegEncoder
                {
                    Quality = 80
                });

                result.Add(ms.ToArray());
            }

            return result;
        }

        private static void InsertPdfImages(
            WordprocessingDocument doc,
            Paragraph placeholderParagraph,
            List<byte[]> pdfPages,
            bool isLastPdf)
        {
            if (pdfPages == null || pdfPages.Count == 0)
                return;

            var parent = placeholderParagraph.Parent;
            if (parent == null)
                return;

            OpenXmlElement refNode = placeholderParagraph;

            // insert a page break BEFORE the first PDF page
            var preBreak = new Paragraph(
                new ParagraphProperties(
                    new ParagraphStyleId() { Val = "Caption" }
                ),
                new Run(new Break() { Type = BreakValues.Page })
            );
            parent.InsertAfter(preBreak, refNode);
            refNode = preBreak;

            // Insert all PDF pages
            for (int i = 0; i < pdfPages.Count; i++)
            {
                // Paragraph for image (always Caption style)
                var imgParagraph = new Paragraph(
                    new ParagraphProperties(
                        new ParagraphStyleId() { Val = "Caption" }
                    )
                );

                parent.InsertAfter(imgParagraph, refNode);
                AddImageToParagraph(doc, imgParagraph, $"pdf_page_{i + 1}", pdfPages[i]);
                refNode = imgParagraph;

                // Page break between pages (but not after last page)
                if (i < pdfPages.Count - 1)
                {
                    var midBreak = new Paragraph(
                        new ParagraphProperties(
                            new ParagraphStyleId() { Val = "Caption" }
                        ),
                        new Run(new Break() { Type = BreakValues.Page })
                    );
                    parent.InsertAfter(midBreak, refNode);
                    refNode = midBreak;
                }
            }

            if (!isLastPdf)
            {
                var postBreak = new Paragraph(
                    new ParagraphProperties(
                        new ParagraphStyleId() { Val = "Caption" }),
                    new Run(new Break() { Type = BreakValues.Page })
                );

                parent.InsertAfter(postBreak, refNode);
            }

            // Remove original placeholder paragraph so its style does not affect layout
            placeholderParagraph.Remove();
        }

        private static bool IsLastPdfPlaceholder(
            int currentIndex,
            MatchCollection matches,
            Dictionary<string, List<byte[]>> pdfImageValues)
        {
            // Look forward in matches to find ANY remaining PDF placeholders
            for (int j = currentIndex - 1; j >= 0; j--)
            {
                string nextName = matches[j].Groups[1].Value;
                if (pdfImageValues.ContainsKey(nextName))
                    return false;
            }

            return true;
        }

        #endregion Private Methods
    }
}
