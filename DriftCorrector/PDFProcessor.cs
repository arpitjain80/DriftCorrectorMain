using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text.RegularExpressions;
using Docnet.Core;
using Docnet.Core.Models;
using iText.Forms;
using iText.IO.Font;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf.Canvas.Parser.Data;
using iText.Kernel.Pdf.Canvas.Parser.Listener;

namespace DocumentProcessor
{
    public class PdfComparer
    {
        //IConfiguration _config;
        private readonly bool _drawIgnoredRegion;
        private readonly bool _ignoreWordsNextToTextinBracketForComparison;
        private readonly bool _considerRedTextInBracketAsOneWord;
        private readonly bool _ignoreTextinBracketForComparison;
        private readonly bool _generateOnePageDifference;
        public List<PdfComparer.IgnoreRegion> lstIgnoreAnnotationRegions;
        string _logPath;
        CustomLogger _logger;
        /// <summary>
        /// Set to true for the MODIFIED comparison (V14 vs corrected PDF): Gate 3 (x-drift threshold)
        /// marks unapplied diffs as clean (green) because the corrector couldn't fix them.
        /// Set to false (default) for the ORIGINAL comparison (V14 vs V23 original): Gate 3 is
        /// disabled so real correctable diffs remain red, showing what needed correction.
        /// </summary>
        public bool IsModifiedComparison { get; set; } = false;
        /// <summary>When true, X/Y/W diff category labels are drawn on report rectangle images.
        /// Default false. Pass showXYWInDiffReport=true to CompareFolders to activate.</summary>
        public bool ShowXYWInDiffReport { get; set; } = false;
        public PdfComparer(
            bool drawIgnoredRegion,
            bool ignoreWordsNextToTextinBracketForComparison,
            bool considerRedTextInBracketAsOneWord,
            bool ignoreTextinBracketForComparison,
            bool GenerateOnePageDifference = true, string logPath = "")
        {
            _drawIgnoredRegion = drawIgnoredRegion;
            _ignoreWordsNextToTextinBracketForComparison = ignoreWordsNextToTextinBracketForComparison;
            _considerRedTextInBracketAsOneWord = considerRedTextInBracketAsOneWord;
            _ignoreTextinBracketForComparison = ignoreTextinBracketForComparison;
            _generateOnePageDifference = GenerateOnePageDifference;
            _logPath = logPath;
            _logger = new CustomLogger(_logPath);
        }

        public class FieldInfo
        {
            public string Name;
            public string Type;
            public string Value;
            public int Page;
            public float X, Y, Width, Height;
        }

        public class TextInfo
        {
            public string Text;
            public string Font;
            public float FontSize;
            public int Page;
            public float X, Y;
            public bool IsRedTextInBracketWord;
        }

        public class FieldDiff
        {
            public string Name;
            public string Issue;
            public FieldInfo Info1, Info2;
            public DiffCategory Categories;
        }

        public class TextDiff
        {
            public string Text;
            public string Issue;
            public TextInfo Info1, Info2;
            public DiffCategory Categories;
        }

        public class IgnoreRegion
        {
            public int Page;
            public float XStart, XEnd, Y, Height;
        }

        /// <summary>Reasons why a diff rectangle was identified.</summary>
        [Flags]
        public enum DiffCategory
        {
            None = 0,
            /// <summary>Difference in X coordinate (horizontal position shift).</summary>
            XDiff = 1,
            /// <summary>Difference in Y coordinate (vertical position shift).</summary>
            YDiff = 2,
            /// <summary>Word/text content difference at the same or nearby position.</summary>
            WDiff = 4,
            /// <summary>X difference that was identified but ignored (XI = X Ignore) because
            /// the text is already aligned with the computed pivot left-margin X.
            /// Rendered as green with "XI" label. Does not count as an unresolved diff.</summary>
            XIgnore = 8
        }

        /// <summary>A single classified diff rectangle (pdf1/baseline coordinates).</summary>
        public class DiffRectangleInfo
        {
            public int Page { get; set; }
            public float X { get; set; }
            public float Y { get; set; }
            public float Width { get; set; }
            public float Height { get; set; }
            /// <summary>true = clean/acceptable (green); false = unresolved (red).</summary>
            public bool IsGreen { get; set; }
            /// <summary>Why this diff was identified (XDiff, YDiff, WDiff — can be combined).</summary>
            public DiffCategory Categories { get; set; }
        }

        /// <summary>Per-file correction statistics. Populated only when IsModifiedComparison=true
        /// AND prior diff rectangles from the unmodified comparison are supplied.</summary>
        public class FileDiffStats
        {
            public int TotalDiffs { get; set; }
            public int RectifiedDiffs { get; set; }
            public int AcceptableDiffs { get; set; }
            public double PercentRectified { get; set; }
            public double PercentRectifiedOrAcceptable { get; set; }
            /// <summary>True when word templates were available so green rects (acceptables) can exist.</summary>
            public bool HasAcceptableContext { get; set; }
        }

        /// <summary>Extended per-file comparison result returned by CompareFolders.</summary>
        public class FileComparisonResult
        {
            public string PdfPath { get; set; }
            public bool HasDifference { get; set; }
            public bool AllDiffsAreClean { get; set; }
            public bool IsTextPresent { get; set; }
            /// <summary>Full path to the individual HTML diff report for this file.</summary>
            public string IndividualReportPath { get; set; }
            /// <summary>Full path to the folder-level summary report.</summary>
            public string MainReportPath { get; set; }
            /// <summary>All classified diff rectangles (pdf1 baseline coords). Empty when no diffs.</summary>
            public List<DiffRectangleInfo> DiffRectangles { get; set; } = new List<DiffRectangleInfo>();
            /// <summary>Correction stats. Null unless IsModifiedComparison=true and prior rects supplied.</summary>
            public FileDiffStats Stats { get; set; }
        }

        private HashSet<string> HighlightedRegions = new HashSet<string>();

        string _ReportBasePath = "";
        string _ReportFileName = "";
        public string ReportBasePath { get { return _ReportBasePath; } set { _ReportBasePath = value; } }
        public string ReportFileName { get { return _ReportFileName; } set { _ReportFileName = value; } }

        public static List<IgnoreRegion> GetAnnotations(string pdfPath, string AnnnotationName)
        {
            var ignoreRegions = new List<IgnoreRegion>();

            using (PdfReader reader = new PdfReader(pdfPath))
            using (iText.Kernel.Pdf.PdfDocument pdfDoc = new iText.Kernel.Pdf.PdfDocument(reader))
            {
                int numberOfPages = pdfDoc.GetNumberOfPages();

                for (int i = 1; i <= numberOfPages; i++)
                {
                    var page = pdfDoc.GetPage(i);
                    var annotations = page.GetAnnotations();

                    foreach (var annotation in annotations)
                    {
                        //PdfDictionary annotDict = annotation.GetPdfObject();

                        // Assuming you have: 
                        string pdfObjString = annotation.GetPdfObject().ToString();

                        // Regex pattern to match: /Subj Ignore (case-insensitive, optional whitespace)
                        string pattern = $@"\/Subj\s+{AnnnotationName}";

                        // Perform case-insensitive match
                        bool hasIgnoreSubject = Regex.IsMatch(pdfObjString, pattern, RegexOptions.IgnoreCase);

                        if (hasIgnoreSubject)
                        {
                            iText.Kernel.Geom.Rectangle rect = annotation.GetRectangle().ToRectangle();

                            if (rect != null)
                            {
                                ignoreRegions.Add(new IgnoreRegion
                                {
                                    Page = i,
                                    XStart = rect.GetX(),
                                    Y = rect.GetY(),
                                    XEnd = (rect.GetX() + rect.GetWidth()),
                                    Height = rect.GetHeight()
                                });
                            }
                        }
                    }
                }
            }

            return ignoreRegions;
        }

        public static List<iText.Kernel.Geom.Rectangle> GetAnnotations(iText.Kernel.Pdf.PdfDocument pdfDoc, int PageNum, string annotationName)
        {
            var rectangles = new List<iText.Kernel.Geom.Rectangle>();
            int numberOfPages = pdfDoc.GetNumberOfPages();

            for (int i = 1; i <= numberOfPages; i++)
            {
                if (i == PageNum)
                {
                    var page = pdfDoc.GetPage(i);
                    var annotations = page.GetAnnotations();

                    foreach (var annotation in annotations)
                    {
                        string pdfObjString = annotation.GetPdfObject().ToString();
                        string pattern = $@"\/Subj\s+{annotationName}";

                        bool hasMatchingSubject = Regex.IsMatch(pdfObjString, pattern, RegexOptions.IgnoreCase);

                        if (hasMatchingSubject)
                        {
                            var rect = annotation.GetRectangle().ToRectangle();
                            if (rect != null)
                            {
                                rectangles.Add(rect);
                            }
                        }
                    }
                }
            }

            return rectangles;
        }

        private List<IgnoreRegion> ComputeIgnoreRegions(List<TextInfo> words)
        {
            var regions = new List<IgnoreRegion>();
            var ordered = words.OrderBy(w => w.Page).ThenBy(w => w.Y).ThenBy(w => w.X).ToList();

            int XExtraIgnore = 0;
            int YExtraIgnore = 0;

            // bool blnIgnoreWordsNextToTextinBracketForComparison = bool.Parse(_config["IgnoreWordsNextToTextinBracketForComparison"]);
            if (_ignoreWordsNextToTextinBracketForComparison)
            {
                YExtraIgnore = 10;
                XExtraIgnore = 10;
            }

            for (int i = 0; i < ordered.Count - 1; i++)
            {
                var current = ordered[i];
                var next = ordered[i + 1];

                if (current.IsRedTextInBracketWord && current.Page == next.Page)
                {
                    float xStart = current.X;
                    float xEnd = next.X - 1;
                    if (xEnd <= xStart) xEnd = current.X + current.FontSize * current.Text.Length;

                    regions.Add(new IgnoreRegion
                    {
                        Page = current.Page,
                        XStart = xStart,
                        XEnd = xEnd + XExtraIgnore,
                        Y = current.Y - 2,
                        Height = current.FontSize + YExtraIgnore  // Use font size as height for drawing
                    });
                }
            }

            return regions;
        }

        public List<FieldInfo> ExtractFields(string pdfPath)
        {
            var fields = new List<FieldInfo>();
            using (var pdfDoc = new iText.Kernel.Pdf.PdfDocument(new PdfReader(pdfPath)))
            {
                var form = PdfAcroForm.GetAcroForm(pdfDoc, false);
                if (form == null || form.GetAllFormFields().Count == 0) return fields;

                foreach (var kv in form.GetAllFormFields())
                {
                    var field = kv.Value;
                    foreach (var widget in field.GetWidgets())
                    {
                        var rect = widget.GetRectangle().ToRectangle();
                        var page = widget.GetPage();
                        fields.Add(new FieldInfo
                        {
                            Name = kv.Key,
                            Type = field.GetFormType()?.ToString() ?? "Unknown",
                            Value = field.GetValueAsString() ?? string.Empty,
                            Page = pdfDoc.GetPageNumber(page),
                            X = rect.GetX(),
                            Y = rect.GetY(),
                            Width = rect.GetWidth(),
                            Height = rect.GetHeight()
                        });
                    }
                }
            }

            return fields;
        }

        public List<TextInfo> ExtractText(string pdfPath)
        {
            var texts = new List<TextInfo>();
            using (var pdf = new iText.Kernel.Pdf.PdfDocument(new PdfReader(pdfPath)))
            {
                for (int i = 1; i <= pdf.GetNumberOfPages(); i++)
                {
                    var page = pdf.GetPage(i);
                    bool blbConsiderRedTextInBracketAsOneWord = _considerRedTextInBracketAsOneWord; //bool.Parse(_config["ConsiderRedTextInBracketAsOneWord"]);

                    WordLevelExtractionStrategy strategy;

                    List<iText.Kernel.Geom.Rectangle> lstWhiteMasked = PdfComparer.GetAnnotations(pdf, i, "WhiteMasked");

                    if (blbConsiderRedTextInBracketAsOneWord)
                    {
                        strategy = new WordLevelExtractionStrategy(true, lstWhiteMasked);
                    }
                    else
                    {
                        strategy = new WordLevelExtractionStrategy(false, lstWhiteMasked);
                    }

                    var processor = new PdfCanvasProcessor(strategy);
                    processor.ProcessPageContent(page);

                    foreach (var word in strategy.Words)
                    {
                        texts.Add(new TextInfo
                        {
                            Text = word.Text,
                            Font = word.FontName,
                            FontSize = word.FontSize,
                            Page = i,
                            X = word.BoundingBox.GetX(),
                            Y = word.BoundingBox.GetY(),
                            IsRedTextInBracketWord = word.IsRedTextInBracketWord
                        });
                    }
                }
            }

            return texts;
        }

        //Version 3
        public bool RemoveText(string pdfPath, List<TextInfo> textInfos, string textToRemove, bool removeMultiple, string outputDirectory)
        {
            var wordsToRemove = textToRemove.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            var occurrences = new List<(int Page, List<TextInfo> Sequence)>();
            var pagesSeen = new HashSet<int>();

            for (int i = 0; i < textInfos.Count; i++)
            {
                if (!string.Equals(textInfos[i].Text, wordsToRemove[0], StringComparison.OrdinalIgnoreCase))
                    continue;

                bool match = true;
                var sequence = new List<TextInfo> { textInfos[i] };

                for (int j = 1; j < wordsToRemove.Length && i + j < textInfos.Count; j++)
                {
                    if (!string.Equals(textInfos[i + j].Text, wordsToRemove[j], StringComparison.OrdinalIgnoreCase) ||
                        textInfos[i + j].Page != textInfos[i].Page)
                    {
                        match = false;
                        break;
                    }
                    sequence.Add(textInfos[i + j]);
                }

                if (match)
                {
                    if (!removeMultiple)
                    {
                        if (!pagesSeen.Contains(textInfos[i].Page))
                        {
                            occurrences.Add((textInfos[i].Page, sequence));
                            pagesSeen.Add(textInfos[i].Page);
                        }
                    }
                    else
                    {
                        occurrences.Add((textInfos[i].Page, sequence));
                    }
                }
            }

            if (occurrences.Count == 0)
                return false;

            // Construct the output file path using the output directory and the original file name
            string outputPath = System.IO.Path.Combine(outputDirectory, System.IO.Path.GetFileName(pdfPath));

            var reader = new PdfReader(pdfPath);
            var writer = new PdfWriter(outputPath);
            var pdfDoc = new iText.Kernel.Pdf.PdfDocument(reader, writer);

            foreach (var (pageNumber, sequence) in occurrences)
            {
                var page = pdfDoc.GetPage(pageNumber);
                var canvas = new PdfCanvas(page);

                foreach (var word in sequence)
                {
                    float estimatedWidth = word.Text.Length * word.FontSize * 0.5f;
                    float estimatedHeight = word.FontSize * 1.2f;

                    canvas.SaveState()
                          .SetFillColorRgb(1, 1, 1)
                          .Rectangle(word.X, word.Y, estimatedWidth, estimatedHeight)
                          .Fill()
                          .RestoreState();

                    var rect = new iText.Kernel.Geom.Rectangle(word.X, word.Y, estimatedWidth, estimatedHeight);

                    var annotation = new iText.Kernel.Pdf.Annot.PdfSquareAnnotation(rect)
                        .SetFlags(iText.Kernel.Pdf.Annot.PdfAnnotation.PRINT | 32) // PRINT + NoView
                        .SetColor(iText.Kernel.Colors.ColorConstants.WHITE)
                        .SetBorder(new iText.Kernel.Pdf.PdfArray(new float[] { 0, 0, 0 }));

                    annotation.GetPdfObject().Remove(iText.Kernel.Pdf.PdfName.AP);
                    annotation.GetPdfObject().Put(new iText.Kernel.Pdf.PdfName("CA"), new iText.Kernel.Pdf.PdfNumber(0));
                    annotation.GetPdfObject().Put(new iText.Kernel.Pdf.PdfName("ca"), new iText.Kernel.Pdf.PdfNumber(0));
                    annotation.GetPdfObject().Put(new iText.Kernel.Pdf.PdfName("H"), new iText.Kernel.Pdf.PdfName("N"));
                    annotation.Put(iText.Kernel.Pdf.PdfName.Subj, new iText.Kernel.Pdf.PdfString("WhiteMasked"));

                    page.AddAnnotation(annotation);
                }
            }

            pdfDoc.Close();
            return true;
        }

        public static bool CotainsText(List<TextInfo> textInfos, string textToSearch)
        {
            var wordsToSearch = textToSearch.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            var occurrences = new List<(int Page, List<TextInfo> Sequence)>();
            var pagesSeen = new HashSet<int>();

            for (int i = 0; i < textInfos.Count; i++)
            {
                if (!string.Equals(textInfos[i].Text, wordsToSearch[0], StringComparison.OrdinalIgnoreCase))
                    continue;

                bool match = true;
                var sequence = new List<TextInfo> { textInfos[i] };

                for (int j = 1; j < wordsToSearch.Length && i + j < textInfos.Count; j++)
                {
                    if (!string.Equals(textInfos[i + j].Text, wordsToSearch[j], StringComparison.OrdinalIgnoreCase) ||
                        textInfos[i + j].Page != textInfos[i].Page)
                    {
                        match = false;
                        break;
                    }
                    sequence.Add(textInfos[i + j]);
                }

                if (match)
                {
                    return true;
                }
            }

            return false;
        }

        public List<FieldDiff> CompareFields(List<FieldInfo> f1, List<FieldInfo> f2)
        {
            var diffs = new List<FieldDiff>();
            var map2 = f2.ToDictionary(f => f.Name);
            var matched = new HashSet<string>();

            foreach (var f in f1)
            {
                if (!map2.TryGetValue(f.Name, out var f2Match) && !string.IsNullOrEmpty(f.Name) && !string.IsNullOrWhiteSpace(f.Name))
                {
                    diffs.Add(new FieldDiff
                    {
                        Name = f.Name,
                        Issue = "Missing in PDF2",
                        Info1 = f,
                        Info2 = null,  // So we can still highlight region in PDF2 at same location
                        Categories = DiffCategory.WDiff
                    });
                }
                else
                {
                    matched.Add(f.Name);
                    //No X Tolerance
                    //bool xMismatch = f.X != f2Match.X;
                    //bool yMismatch = Math.Abs(f.Y - f2Match.Y) > 2;

                    // if ((xMismatch || yMismatch) && !string.IsNullOrEmpty(f.Name) && !string.IsNullOrWhiteSpace(f.Name))
                    float dx = Math.Abs(f.X - f2Match.X), dy = Math.Abs(f.Y - f2Match.Y);
                    if ((dx > 1 || dy > 2) && !string.IsNullOrEmpty(f.Name) && !string.IsNullOrWhiteSpace(f.Name))
                    {
                        var posCats = DiffCategory.None;
                        if (dx > 1) posCats |= DiffCategory.XDiff;
                        if (dy > 2) posCats |= DiffCategory.YDiff;
                        diffs.Add(new FieldDiff
                        {
                            Name = f.Name,
                            Issue = "Field position mismatch",
                            Info1 = f,
                            Info2 = f2Match,
                            Categories = posCats
                        });
                    }
                    else if (!string.IsNullOrEmpty(f.Type) && !string.IsNullOrEmpty(f2Match.Type) && f.Type != f2Match.Type && !string.IsNullOrEmpty(f.Name) && !string.IsNullOrWhiteSpace(f.Name))
                    {
                        diffs.Add(new FieldDiff
                        {
                            Name = f.Name,
                            Issue = "Field type mismatch",
                            Info1 = f,
                            Info2 = f2Match,
                            Categories = DiffCategory.WDiff
                        });
                    }
                }
            }

            // Fields present in f2 but missing in f1
            foreach (var f in f2)
            {
                if (!matched.Contains(f.Name) && !string.IsNullOrEmpty(f.Name) && !string.IsNullOrWhiteSpace(f.Name))
                {
                    diffs.Add(new FieldDiff
                    {
                        Name = f.Name,
                        Issue = "Extra in PDF2",
                        Info1 = null,
                        Info2 = f,
                        Categories = DiffCategory.WDiff
                    });
                }
            }

            return diffs;
        }

        public List<TextDiff> CompareText(List<TextInfo> t1, List<TextInfo> t2)
        {
            var diffs = new List<TextDiff>();
            var matched = new HashSet<TextInfo>();
            var ignoreRegions = ComputeIgnoreRegions(t1);
            ignoreRegions = ignoreRegions.Concat(this.lstIgnoreAnnotationRegions).ToList();

            bool IsInIgnoreRegion(TextInfo t)
            {
                return ignoreRegions.Any(r =>
                    t.Page == r.Page &&
                    t.X >= r.XStart && t.X <= r.XEnd &&
                    Math.Abs(t.Y - r.Y) < 2
                );
            }

            foreach (var a in t1)
            {
                if (string.IsNullOrWhiteSpace(a.Text) && string.IsNullOrEmpty(a.Text))
                    continue; //Skip empty or whitespace-only text

                if (IsInIgnoreRegion(a)) continue;

                var candidates = t2.Where(b =>
                    b.Page == a.Page &&
                    Math.Abs(a.X - b.X) <= 2 &&    //Tolerance of 2 px, can be adjusted
                                                   //a.X == b.X &&                              // Strict X position
                    Math.Abs(a.Y - b.Y) <= 3 &&
                    !matched.Contains(b)
                ).ToList();

                TextInfo bestMatch = null;
                float bestDist = float.MaxValue;

                foreach (var b in candidates)
                {
                    float dx = Math.Abs(a.X - b.X);
                    float dy = Math.Abs(a.Y - b.Y);
                    float dist = dx + dy;

                    if (dist < bestDist && string.Equals(a.Text, b.Text, StringComparison.Ordinal))
                    {
                        bestDist = dist;
                        bestMatch = b;
                    }
                }

                if (bestMatch != null)
                {
                    matched.Add(bestMatch);

                    bool isFontSizeMismatch = Math.Abs(a.FontSize - bestMatch.FontSize) < .25 ? false : true;   //tolerate fontsize mismatch upto .25
                    string aFont = a.Font.Contains("+") ? a.Font.Split('+')[1] : a.Font;
                    string bestMatchFont = bestMatch.Font.Contains("+") ? bestMatch.Font.Split('+')[1] : bestMatch.Font;

                    if ((isFontSizeMismatch) || (!string.IsNullOrEmpty(aFont) && !string.IsNullOrEmpty(bestMatchFont) && aFont != bestMatchFont) && !string.IsNullOrEmpty(a.Text) && !string.IsNullOrWhiteSpace(a.Text))
                    {
                        // Formatting mismatch: same word at same position but different font/size.
                        // WDiff only — XDiff/YDiff are determined at the merged-rect level, not per-word.
                        diffs.Add(new TextDiff
                        {
                            Text = a.Text,
                            Issue = "Formatting mismatch",
                            Info1 = a,
                            Info2 = bestMatch,
                            Categories = DiffCategory.WDiff
                        });
                    }
                }
                else if (!string.IsNullOrEmpty(a.Text) && !string.IsNullOrWhiteSpace(a.Text))
                {
                    // Word in t1 not found in t2 at same position.
                    // Mark WDiff at word level; XDiff/YDiff will be determined from the merged rect positions.
                    diffs.Add(new TextDiff
                    {
                        Text = a.Text,
                        Issue = "Missing or shifted in PDF2",
                        Info1 = a,
                        Info2 = null,  // So we can still highlight region on same page in PDF2
                        Categories = DiffCategory.WDiff
                    });
                }
            }

            foreach (var b in t2)
            {
                if (!matched.Contains(b) && !string.IsNullOrEmpty(b.Text) && !string.IsNullOrWhiteSpace(b.Text) && !IsInIgnoreRegion(b))
                {
                    // Word in t2 not matched to any t1 word at same position.
                    // Mark WDiff at word level; XDiff/YDiff will be determined from the merged rect positions.
                    diffs.Add(new TextDiff
                    {
                        Text = b.Text,
                        Issue = "Extra in PDF2",
                        Info1 = null,
                        Info2 = b,
                        Categories = DiffCategory.WDiff
                    });
                }
            }

            return diffs;
        }

        /// <summary>
        /// Computes per-file correction stats by comparing prior (unmodified) red rects
        /// against the current (modified) rects. No logic change to rect generation.
        /// </summary>
        /// <summary>
        /// When category data is available: X-diffs are red (structural position change);
        /// W/Y-only diffs are green (acceptable rendering/content difference).
        /// Falls back to <paramref name="fallbackIsGreen"/> when no category data.
        /// </summary>
        private static bool CategoryBasedIsGreen(DiffCategory cats, bool fallbackIsGreen)
            => cats != DiffCategory.None ? (cats & DiffCategory.XDiff) == 0 : fallbackIsGreen;

        /// <summary>
        /// Computes the "pivot left X" for XI (X-Ignore) reconciliation.
        /// Returns the canonical left-margin X that leftmost text should align to.
        ///
        /// Priority:
        ///   0. Clean word positions (words from IsCleanDiff=true phrase groups) — highest priority.
        ///   A. Leftmost free text (words not in any diff rect) → mode X of leftmost words.
        ///   B. Leftmost green rects (no XDiff) → mode X of leftmost rects.
        ///
        /// If neither clean positions, free text, nor green rects are available (i.e. all
        /// leftmost rects have X diffs), no pivot is returned and XI is NOT applied — those
        /// rects remain marked X and are corrected by the word drift corrector as normal.
        /// </summary>
        private static float? ComputePivotX(
            IEnumerable<float> freeTextXValues,
            IEnumerable<float> greenRectXValues,
            float leftmostWindowPt = 10f,
            float clusterTolPt = 3f,
            IEnumerable<float> cleanPosXValues = null)
        {
            // Priority: freeTextX first (stable text not in any X-diff rect), then greenRectX, then cleanPosX as last resort.
            // freeTextX must be first: if the XDiff rect bx equals the free-text minimum, the rect is AT the reference
            // margin (not to the left of it), meaning it drifted in V23 and needs correction — so the pivot must
            // reflect the true leftmost stable reference, not the rect's own position via cleanPosX.
            foreach (var source in new[] { freeTextXValues?.ToList(), greenRectXValues?.ToList(), cleanPosXValues?.ToList() })
            {
                if (source == null || source.Count == 0) continue;
                float minX = source.Min();
                var leftmost = source.Where(x => x <= minX + leftmostWindowPt).ToList();
                if (leftmost.Count == 0) continue;
                float? pivot = ClusterModeX(leftmost, clusterTolPt);
                if (pivot.HasValue) return pivot;
            }
            return null;
        }

        /// <summary>
        /// Returns the center of the largest X-value cluster (values within <paramref name="tol"/> of
        /// each other). Ties are broken by choosing the leftmost (smallest) cluster anchor.
        /// </summary>
        private static float? ClusterModeX(IList<float> values, float tol)
        {
            if (values == null || values.Count == 0) return null;
            if (values.Count == 1) return values[0];

            var anchors = new List<float>();
            var counts  = new List<int>();
            foreach (float v in values)
            {
                int best = -1;
                float bestDist = tol;
                for (int ci = 0; ci < anchors.Count; ci++)
                {
                    float d = Math.Abs(anchors[ci] - v);
                    if (d <= bestDist) { bestDist = d; best = ci; }
                }
                if (best >= 0) counts[best]++;
                else { anchors.Add(v); counts.Add(1); }
            }

            int maxCount = counts.Max();
            float pivot = float.MaxValue;
            for (int ci = 0; ci < anchors.Count; ci++)
                if (counts[ci] == maxCount && anchors[ci] < pivot)
                    pivot = anchors[ci];
            return pivot == float.MaxValue ? (float?)null : pivot;
        }

        /// <summary>
        /// Removes rects that are fully contained within a larger rect on the same page.
        /// Sorts by area descending; for each rect, keeps it only if no already-kept rect contains it.
        /// This prevents nested rects from inflating counts when the same diff area has both a
        /// parent and a child rectangle.
        /// </summary>
        private static List<DiffRectangleInfo> DeduplicateContained(List<DiffRectangleInfo> rects)
        {
            const float CONTAIN_TOL = 2f; // pt slop for containment check
            var sorted = rects.OrderByDescending(r => r.Width * r.Height).ToList();
            var kept = new List<DiffRectangleInfo>();
            foreach (var r in sorted)
            {
                bool containedByKept = kept.Any(k =>
                    k.Page == r.Page &&
                    k.X - CONTAIN_TOL <= r.X &&
                    k.Y - CONTAIN_TOL <= r.Y &&
                    k.X + k.Width  + CONTAIN_TOL >= r.X + r.Width &&
                    k.Y + k.Height + CONTAIN_TOL >= r.Y + r.Height);
                if (!containedByKept)
                    kept.Add(r);
            }
            return kept;
        }

        private FileDiffStats ComputeStats(
            List<DiffRectangleInfo> priorRects,
            List<DiffRectangleInfo> currentRects,
            bool hasWordContext)
        {
            const float MATCH_TOL = 10f; // pt — spatial tolerance for matching rects across comparisons

            // Deduplicate nested rects (parent + child for the same area) in both lists
            // so that a single logical diff is not counted multiple times.
            var priorDedup   = DeduplicateContained(priorRects);
            var currentDedup = DeduplicateContained(currentRects);

            // 1. TotalDiffs = ALL deduplicated rectangles from the original comparison (red + green).
            int totalDiffs = priorDedup.Count;
            if (totalDiffs == 0)
                return new FileDiffStats { TotalDiffs = 0, HasAcceptableContext = hasWordContext };

            // 2. Rectified = original rects that have NO spatially-matching rect in modified
            //    (the diff was completely resolved — the rectangle disappeared entirely).
            int rectified = 0;
            foreach (var prior in priorDedup)
            {
                bool hasMatch = currentDedup.Any(c =>
                    c.Page == prior.Page &&
                    c.X < prior.X + prior.Width  + MATCH_TOL &&
                    c.X + c.Width  > prior.X     - MATCH_TOL &&
                    c.Y < prior.Y + prior.Height + MATCH_TOL &&
                    c.Y + c.Height > prior.Y     - MATCH_TOL);
                if (!hasMatch) rectified++;
            }

            // 3. Acceptable = green rectangles in the deduplicated modified comparison.
            int acceptable = currentDedup.Count(r => CategoryBasedIsGreen(r.Categories, r.IsGreen));

            double pctRectified = (double)rectified / totalDiffs * 100.0;
            double pctRectifiedOrAcceptable = (double)(rectified + acceptable) / totalDiffs * 100.0;
            return new FileDiffStats
            {
                TotalDiffs = totalDiffs,
                RectifiedDiffs = rectified,
                AcceptableDiffs = acceptable,
                PercentRectified = Math.Round(pctRectified, 1),
                PercentRectifiedOrAcceptable = Math.Round(pctRectifiedOrAcceptable, 1),
                HasAcceptableContext = hasWordContext
            };
        }

        private bool AreRectanglesClose(RectangleF r1, RectangleF r2, float padding)
        {
            var expanded1 = ExpandRectangle(r1, padding);
            var expanded2 = ExpandRectangle(r2, padding);
            return expanded1.IntersectsWith(r2) || r2.IntersectsWith(expanded1);
        }

        private RectangleF ExpandRectangle(RectangleF rect, float padding)
        {
            return new RectangleF(
                rect.X - padding,
                rect.Y - padding,
                rect.Width + 2 * padding,
                rect.Height + 2 * padding
            );
        }

        public (List<FileComparisonResult>, List<string>, bool, bool) CompareFiles(
            string file1Path, string file2Path,
            List<string> IsTextPresent = null,
            string sourceTemplateDir = null)
        {
            List<string> logMsgs = new List<string>();
            var fileResults = new List<FileComparisonResult>();
            bool hasAtleastOneDiff = false;
            bool isTextPresentInAtleastOneFile = false;
            string logMsg = "";

            if (!File.Exists(file1Path))
            {
                logMsg = $"File1 file not found : '{file1Path}'.";
                logMsgs.Add(logMsg);
            }
            if (!File.Exists(file2Path))
            {
                logMsg = $"File2 file not found : '{file2Path}'.";
                logMsgs.Add(logMsg);
            }

            logMsg = $"\nComparing Files: '{file1Path}' and '{file2Path}'";
            logMsgs.Add(logMsg);

            var fields1 = ExtractFields(file1Path);
            var fields2 = ExtractFields(file2Path);
            var text1 = ExtractText(file1Path);
            var text2 = ExtractText(file2Path);
            lstIgnoreAnnotationRegions = PdfComparer.GetAnnotations(file1Path, "Ignore");

            var fieldDiffs = CompareFields(fields1, fields2);
            var textDiffs = CompareText(text1, text2);

            // ── JSON diff FIRST — classifies phrase groups as IsCleanDiff ─────────────
            string jsonSubDir = Path.Combine(ReportBasePath, "JSONDiff");
            var (phraseGroups, cleanWordPositions) = GenerateJsonDiffReport(
                file1Path, file2Path, text1, text2,
                jsonSubDir, sourceTemplateDir);

            // ── HTML report — uses phrase groups for green/red rect colouring ─────────
            bool hasDiff = GenerateHtmlReport(
                file1Path, file2Path, fieldDiffs, textDiffs,
                out string reportFileName, out bool allDiffsAreClean,
                out List<DiffRectangleInfo> fileDiffRects,
                text1, text2, phraseGroups, cleanWordPositions);

            bool blnTxtPresent = false;
            string srchTextPresent = "";
            if (IsTextPresent != null && IsTextPresent.Count > 0)
            {
                foreach (string srchText in IsTextPresent)
                {
                    blnTxtPresent = CotainsText(text2, srchText);
                    if (blnTxtPresent)
                    {
                        isTextPresentInAtleastOneFile = true;
                        srchTextPresent = srchText;
                        break;
                    }
                }
            }

            string mainReportPath = Path.Combine(ReportBasePath, _ReportFileName);
            fileResults.Add(new FileComparisonResult
            {
                PdfPath = file1Path,
                HasDifference = hasDiff,
                AllDiffsAreClean = allDiffsAreClean,
                IsTextPresent = blnTxtPresent,
                IndividualReportPath = reportFileName,
                MainReportPath = mainReportPath,
                DiffRectangles = fileDiffRects
            });

            string flname = System.IO.Path.GetFileName(file1Path);
            logMsg = hasDiff
                ? (allDiffsAreClean ? $"Clean differences (green) found for form : {flname}" : $"Difference found and Report saved for form : {flname}")
                : $"No differences found for form : {flname}";
            logMsgs.Add(logMsg);
            logMsgs.Add("Individual Form Difference Reports Generated.");

            if (blnTxtPresent)
                logMsgs.Add($"Text {srchTextPresent} found in form {flname}");

            if (hasDiff) hasAtleastOneDiff = true;

            logMsgs.Add(hasAtleastOneDiff ? "Atleast one form has difference" : "No differences found");

            if (isTextPresentInAtleastOneFile)
                logMsgs.Add($"Atleast one form has text present : {string.Join(",", IsTextPresent)}");
            else
                logMsgs.Add("No text search matches found");

            if (ReportBasePath != "")
            {
                logMsgs.Add("Generating Summary Report");
                GenerateSummaryReport(ReportBasePath, _ReportFileName, fileResults);
                logMsgs.Add("Completed Generating Summary Report");
            }
            else
            {
                logMsgs.Add("ReportBasePath is empty, skipped Generating Summary Report");
            }

            return (fileResults, logMsgs, hasAtleastOneDiff, isTextPresentInAtleastOneFile);
        }

        private PdfComparer CreateSubComparer(string subReportBasePath)
        {
            var sub = new PdfComparer(
                _drawIgnoredRegion,
                _ignoreWordsNextToTextinBracketForComparison,
                _considerRedTextInBracketAsOneWord,
                _ignoreTextinBracketForComparison,
                _generateOnePageDifference,
                _logPath);
            sub.ReportBasePath = subReportBasePath;
            sub.ReportFileName = _ReportFileName;
            sub.IsModifiedComparison = IsModifiedComparison;
            sub.ShowXYWInDiffReport = ShowXYWInDiffReport;
            return sub;
        }

        public (List<FileComparisonResult>, List<string>, bool, bool) CompareFolders(
            string SourceFolder, string DestinationFolder,
            List<string> IsTextPresent = null,
            string sourceTemplateDir = null,
            Dictionary<string, List<DiffRectangleInfo>> priorDiffRects = null,
            bool showXYWInDiffReport = false)
        {
            ShowXYWInDiffReport = showXYWInDiffReport;
            if (string.IsNullOrEmpty(_ReportFileName))
                _ReportFileName = "Report.html";

            List<string> logMsgs = new List<string>();
            var fileResults = new List<FileComparisonResult>();
            bool hasAtleastOneDiff = false;
            bool isTextPresentInAtleastOneFile = false;
            string logMsg = "";

            var files1 = Directory.GetFiles(SourceFolder, "*.pdf");
            var files2 = Directory.GetFiles(DestinationFolder, "*.pdf")
                                  .Select(System.IO.Path.GetFileName)
                                  .ToHashSet(StringComparer.OrdinalIgnoreCase);

            logMsg = $"Total Forms in Baseline Folder : {files1.Length}";
            logMsgs.Add(logMsg);
            logMsg = $"Total Forms in Publish Folder : {files2.Count}";
            logMsgs.Add(logMsg);

            if (files1.Length > 0 && files2.Count > 0)
            {
                foreach (var file1Path in files1)
                {
                    string fileName = System.IO.Path.GetFileName(file1Path);
                    string file2Path = System.IO.Path.Combine(DestinationFolder, fileName);

                    if (!files2.Contains(fileName) || !File.Exists(file2Path))
                    {
                        logMsg = $"Skipped: No matching file for '{fileName}' found in second directory.";
                        logMsgs.Add(logMsg);
                        continue;
                    }

                    logMsg = $"\nComparing form: {fileName}";
                    logMsgs.Add(logMsg);

                    var fields1 = ExtractFields(file1Path);
                    var fields2 = ExtractFields(file2Path);
                    var text1 = ExtractText(file1Path);
                    var text2 = ExtractText(file2Path);
                    lstIgnoreAnnotationRegions = PdfComparer.GetAnnotations(file1Path, "Ignore");

                    var fieldDiffs = CompareFields(fields1, fields2);
                    var textDiffs = CompareText(text1, text2);

                    // ── JSON diff FIRST — classifies phrase groups as IsCleanDiff ─────
                    string jsonSubDir = Path.Combine(ReportBasePath, "JSONDiff");
                    var (phraseGroups, cleanWordPositions) = GenerateJsonDiffReport(
                        file1Path, file2Path, text1, text2,
                        jsonSubDir, sourceTemplateDir);

                    // ── HTML report — uses phrase groups for green/red rect colouring ──
                    bool hasDiff = GenerateHtmlReport(
                        file1Path, file2Path, fieldDiffs, textDiffs,
                        out string reportFileName, out bool allDiffsAreClean,
                        out List<DiffRectangleInfo> fileDiffRects,
                        text1, text2, phraseGroups, cleanWordPositions);

                    bool blnTxtPresent = false;
                    string srchTextPresent = "";
                    if (IsTextPresent != null && IsTextPresent.Count > 0)
                    {
                        foreach (string srchText in IsTextPresent)
                        {
                            blnTxtPresent = CotainsText(text2, srchText);
                            if (blnTxtPresent)
                            {
                                isTextPresentInAtleastOneFile = true;
                                srchTextPresent = srchText;
                                break;
                            }
                        }
                    }

                    // Compute correction stats when this is a modified comparison and prior rects are supplied
                    FileDiffStats fileStats = null;
                    if (IsModifiedComparison && priorDiffRects != null &&
                        priorDiffRects.TryGetValue(file1Path, out var priorFileRects))
                    {
                        fileStats = ComputeStats(priorFileRects, fileDiffRects, sourceTemplateDir != null);
                    }

                    string mainReportPath = Path.Combine(ReportBasePath, _ReportFileName);
                    fileResults.Add(new FileComparisonResult
                    {
                        PdfPath = file1Path,
                        HasDifference = hasDiff,
                        AllDiffsAreClean = allDiffsAreClean,
                        IsTextPresent = blnTxtPresent,
                        IndividualReportPath = reportFileName,
                        MainReportPath = mainReportPath,
                        DiffRectangles = fileDiffRects,
                        Stats = fileStats
                    });

                    string flname = System.IO.Path.GetFileName(file1Path);
                    logMsg = hasDiff
                        ? (allDiffsAreClean ? $"Clean differences (green) found for form : {flname}" : $"Difference found and Report saved for form : {flname}")
                        : $"No differences found for form : {flname}";
                    logMsgs.Add(logMsg);
                    logMsgs.Add("Individual Form Difference Reports Generated.");

                    if (blnTxtPresent)
                        logMsgs.Add($"Text {srchTextPresent} found in form {flname}");

                    if (hasDiff) hasAtleastOneDiff = true;
                }

                logMsgs.Add(hasAtleastOneDiff ? "Atleast one form has difference" : "No differences found");

                if (isTextPresentInAtleastOneFile)
                    logMsgs.Add($"Atleast one form has text present : {string.Join(",", IsTextPresent)}");
                else
                    logMsgs.Add("No text search matches found");
            }
            else
            {
                logMsgs.Add("No forms available");
            }

            // ── RECURSIVE SUBFOLDER PROCESSING ───────────────────────────────────────
            var subFolderSummaries = new List<(string SubFolderName, bool HasAnyDiff, bool AllClean)>();
            // Collect all subfolder file results so they are included in the returned list
            // (needed by callers building priorDiffRects for the modified comparison).
            // They are appended AFTER GenerateSummaryReport so the root-level report
            // only shows root-level files (unchanged visual behaviour).
            var allSubFileResults = new List<FileComparisonResult>();

            foreach (var sourceSubDir in Directory.GetDirectories(SourceFolder))
            {
                string subDirName = Path.GetFileName(sourceSubDir);
                string destSubDir = Path.Combine(DestinationFolder, subDirName);

                if (!Directory.Exists(destSubDir))
                {
                    string skipMsg = $"[INFO] Skipping subfolder '{subDirName}': not found in destination folder.";
                    logMsgs.Add(skipMsg);
                    Console.WriteLine(skipMsg);
                    continue;
                }

                string subReportBase = Path.Combine(ReportBasePath, subDirName);
                Directory.CreateDirectory(subReportBase);

                string subWordDir = null;
                if (sourceTemplateDir != null)
                {
                    string candidate = Path.Combine(sourceTemplateDir, subDirName);
                    if (Directory.Exists(candidate))
                        subWordDir = candidate;
                    else
                    {
                        string warnMsg = $"[INFO] No word template subfolder '{subDirName}' found; comparison proceeds without word context.";
                        logMsgs.Add(warnMsg);
                        Console.WriteLine(warnMsg);
                    }
                }

                var sub = CreateSubComparer(subReportBase);
                var (subRows, subLogs, subHasDiff, _) = sub.CompareFolders(
                    sourceSubDir, destSubDir, IsTextPresent, subWordDir, priorDiffRects, ShowXYWInDiffReport);
                logMsgs.AddRange(subLogs);

                bool subAllClean = !subHasDiff || subRows.All(r => !r.HasDifference || r.AllDiffsAreClean);
                subFolderSummaries.Add((subDirName, subHasDiff, subAllClean));
                if (subHasDiff) hasAtleastOneDiff = true;

                // subRows is already the flat list for that subtree (recursive calls do the same)
                allSubFileResults.AddRange(subRows);
            }

            // ── GENERATE SUMMARY REPORT (files + subfolder links) ────────────────────
            if (ReportBasePath != "")
            {
                logMsgs.Add("Generating Summary Report");
                GenerateSummaryReport(ReportBasePath, _ReportFileName, fileResults,
                                      subFolderSummaries, _ReportFileName);
                logMsgs.Add("Completed Generating Summary Report");
            }
            else
            {
                logMsgs.Add("ReportBasePath is empty, skipped Generating Summary Report");
            }

            // Append subfolder results AFTER report generation so the caller has a flat
            // list of every processed file (root + all subfolders, any depth).
            fileResults.AddRange(allSubFileResults);

            return (fileResults, logMsgs, hasAtleastOneDiff, isTextPresentInAtleastOneFile);
        }

        public static void GenerateSummaryReport(
            string summaryFileFolder,
            string summaryFileName,
            List<FileComparisonResult> rows)
        {
            int total = rows.Count;
            int withDiff = rows.Count(r => r.HasDifference && !r.AllDiffsAreClean);
            int cleanDiff = rows.Count(r => r.HasDifference && r.AllDiffsAreClean);
            int clean = total - rows.Count(r => r.HasDifference);
            string generated = DateTime.Now.ToString("MMMM dd, yyyy  h:mm tt");

            var sb = new StringBuilder();
            sb.AppendLine(@"<!DOCTYPE html>
<html lang=""en"">
<head>
<meta charset=""UTF-8"" />
<meta name=""viewport"" content=""width=device-width, initial-scale=1.0"" />
<title>PDF Comparison – Summary</title>
<link rel=""preconnect"" href=""https://fonts.googleapis.com"" />
<link rel=""preconnect"" href=""https://fonts.gstatic.com"" crossorigin />
<link href=""https://fonts.googleapis.com/css2?family=Lora:wght@400;600&family=Source+Sans+3:wght@300;400;500;600&display=swap"" rel=""stylesheet"" />
<style>
  *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
  :root {
    --bg:          #f5f4f0;
    --surface:     #ffffff;
    --border:      #e2e0db;
    --text-main:   #1c1b18;
    --text-muted: #6b6860;
    --accent:      #2a5bd7;
    --accent-lt:  #eef2fd;
    --green:       #1a7a4a;
    --green-lt:    #eaf5ef;
    --red:         #c0392b;
    --red-lt:      #fdf0ee;
    --shadow-sm:  0 1px 3px rgba(0,0,0,.07), 0 1px 2px rgba(0,0,0,.05);
    --shadow-md:  0 4px 12px rgba(0,0,0,.08), 0 2px 4px rgba(0,0,0,.04);
    --radius:      10px;
  }
  body {
    font-family: 'Source Sans 3', sans-serif;
    font-weight: 400;
    background: var(--bg);
    color: var(--text-main);
    min-height: 100vh;
  }
  .topbar {
    background: var(--surface);
    border-bottom: 1px solid var(--border);
    padding: 0 40px;
    height: 60px;
    display: flex;
    align-items: center;
    justify-content: space-between;
    position: sticky;
    top: 0;
    z-index: 100;
    box-shadow: var(--shadow-sm);
  }
  .topbar-brand {
    font-family: 'Lora', serif;
    font-size: 17px;
    font-weight: 600;
    color: var(--text-main);
    letter-spacing: -.01em;
  }
  .topbar-brand span { color: var(--accent); }
  .topbar-meta { font-size: 12.5px; color: var(--text-muted); font-weight: 300; }
  .page { max-width: 1060px; margin: 0 auto; padding: 48px 32px 80px; }
  .page-heading { margin-bottom: 36px; }
  .page-heading h1 {
    font-family: 'Lora', serif;
    font-size: 28px; font-weight: 600;
    letter-spacing: -.02em; color: var(--text-main); margin-bottom: 6px;
  }
  .page-heading p { font-size: 14px; color: var(--text-muted); font-weight: 300; }
  .stats {
    display: grid;
    grid-template-columns: repeat(4, 1fr);
    gap: 16px; margin-bottom: 36px;
  }
  .stat-card {
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: var(--radius);
    padding: 22px 24px;
    box-shadow: var(--shadow-sm);
  }
  .stat-card .label {
    font-size: 11.5px; font-weight: 500;
    letter-spacing: .06em; text-transform: uppercase;
    color: var(--text-muted); margin-bottom: 8px;
  }
  .stat-card .value {
    font-family: 'Lora', serif;
    font-size: 36px; font-weight: 600;
    line-height: 1; color: var(--text-main);
  }
  .stat-card.accent .value { color: var(--accent); }
  .stat-card.green  .value { color: var(--green);  }
  .stat-card.red    .value { color: var(--red);    }
  .table-card {
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: var(--radius);
    box-shadow: var(--shadow-sm);
    overflow: hidden;
  }
  .table-card-header {
    padding: 18px 24px;
    border-bottom: 1px solid var(--border);
    display: flex; align-items: center; gap: 10px;
  }
  .table-card-header h2 { font-size: 14px; font-weight: 600; color: var(--text-main); }
  .badge-count {
    background: var(--accent-lt); color: var(--accent);
    font-size: 11px; font-weight: 600;
    padding: 2px 8px; border-radius: 20px;
  }
  table { width: 100%; border-collapse: collapse; }
  thead th {
    font-size: 11px; font-weight: 600;
    letter-spacing: .07em; text-transform: uppercase;
    color: var(--text-muted); padding: 11px 24px;
    text-align: left; background: #fafaf8;
    border-bottom: 1px solid var(--border);
  }
  tbody tr { border-bottom: 1px solid var(--border); transition: background .15s; }
  tbody tr:last-child { border-bottom: none; }
  tbody tr:hover { background: #fafaf8; }
  tbody td {
    padding: 14px 24px; font-size: 13.5px;
    vertical-align: middle; color: var(--text-main);
  }
  .col-num    { width: 52px; color: var(--text-muted); font-size: 12.5px; }
  .col-file   { font-family: 'Source Code Pro', monospace; font-size: 12.5px; }
  .col-status { width: 140px; }
  .col-report { width: 140px; }
  .badge {
    display: inline-flex; align-items: center; gap: 5px;
    font-size: 12px; font-weight: 600;
    padding: 3px 10px; border-radius: 20px; white-space: nowrap;
  }
  .badge-ok  { background: var(--green-lt); color: var(--green); }
  .badge-err { background: var(--red-lt);   color: var(--red);   }
  .badge-clean-diff {
    background: #e6f8ed; color: #1aaf5d;
  }
  .badge-ok::before            { content: '✓'; font-size: 11px; }
  .badge-err::before           { content: '●'; font-size: 7px;  }
  .badge-clean-diff::before    { content: '◈'; font-size: 11px; }
  .report-link {
    font-size: 12.5px; color: var(--accent); text-decoration: none;
    font-weight: 500; display: inline-flex; align-items: center; gap: 4px;
    padding: 4px 10px; border-radius: 6px;
    border: 1px solid #c9d8f5; background: var(--accent-lt);
    transition: background .15s, border-color .15s;
  }
  .report-link:hover { background: #dde7fb; border-color: #a8bfee; }
  .report-link::after { content: '→'; }
  .jump-nav {
    background: var(--surface); border: 1px solid var(--border);
    border-radius: var(--radius); box-shadow: var(--shadow-sm);
    padding: 18px 24px; margin-bottom: 28px;
  }
  .jump-nav h3 {
    font-size: 11px; font-weight: 600;
    letter-spacing: .07em; text-transform: uppercase;
    color: var(--text-muted); margin-bottom: 12px;
  }
  .jump-links { display: flex; flex-wrap: wrap; gap: 8px; }
  .jump-link {
    font-size: 12px; color: var(--accent); text-decoration: none;
    padding: 4px 10px; border-radius: 6px;
    border: 1px solid #c9d8f5; background: var(--accent-lt);
    font-weight: 500; transition: background .15s;
  }
  .jump-link:hover { background: #dde7fb; }
  .footer {
    margin-top: 56px; text-align: center;
    font-size: 12px; color: var(--text-muted); font-weight: 300;
  }
</style>
</head>
<body>");

            sb.AppendLine($@"
<div class=""topbar"">
  <div class=""topbar-brand"">PDF<span>Compare</span> &mdash; Summary Report</div>
  <div class=""topbar-meta"">Generated {generated}</div>
</div>

<div class=""page"">
  <div class=""page-heading"">
    <h1>Comparison Summary</h1>
    <p>All baseline PDF files compared against the publish folder. Click a report link to inspect differences.</p>
  </div>

  <div class=""stats"">
    <div class=""stat-card accent"">
      <div class=""label"">Total Files</div>
      <div class=""value"">{total}</div>
    </div>
    <div class=""stat-card red"">
      <div class=""label"">Differences Found</div>
      <div class=""value"">{withDiff}</div>
    </div>
    <div class=""stat-card"" style=""--val-color:#1aaf5d"">
      <div class=""label"">Clean Differences</div>
      <div class=""value"" style=""color:#1aaf5d"">{cleanDiff}</div>
    </div>
    <div class=""stat-card green"">
      <div class=""label"">No Differences</div>
      <div class=""value"">{clean}</div>
    </div>
  </div>");

            // ── Jump nav (files with diffs only) ──────────────────────────────────────
            var diffRows = rows.Where(r => r.HasDifference).ToList();
            if (diffRows.Count > 0)
            {
                sb.AppendLine(@"  <div class=""jump-nav"">
    <h3>Jump to Files with Differences</h3>
    <div class=""jump-links"">");
                foreach (var row in diffRows)
                {
                    string fname = System.IO.Path.GetFileName(row.PdfPath);
                    string anchor = "row-" + System.Text.RegularExpressions.Regex.Replace(
                        System.IO.Path.GetFileNameWithoutExtension(row.PdfPath), @"[^a-zA-Z0-9]", "-");
                    sb.AppendLine($@"      <a class=""jump-link"" href=""#{anchor}"">{System.Net.WebUtility.HtmlEncode(fname)}</a>");
                }
                sb.AppendLine(@"    </div>
  </div>");
            }

            // ── Main results table ────────────────────────────────────────────────────
            sb.AppendLine($@"
  <div class=""table-card"">
    <div class=""table-card-header"">
      <h2>File Results</h2>
      <span class=""badge-count"">{total} files</span>
    </div>
    <table>
      <thead>
        <tr>
          <th class=""col-num"">#</th>
          <th>File Name</th>
          <th class=""col-status"">Status</th>
          <th class=""col-report"">Report</th>
        </tr>
      </thead>
      <tbody>");

            int idx = 1;
            foreach (var row in rows)
            {
                string fname = System.IO.Path.GetFileName(row.PdfPath);
                string reportFileName = System.IO.Path.GetFileName(row.IndividualReportPath);
                string anchor = "row-" + System.Text.RegularExpressions.Regex.Replace(
                    System.IO.Path.GetFileNameWithoutExtension(row.PdfPath), @"[^a-zA-Z0-9]", "-");
                string encodedFname = System.Net.WebUtility.HtmlEncode(fname);

                string statusBadge = row.HasDifference && row.AllDiffsAreClean
                    ? @"<span class=""badge badge-clean-diff"">Clean Difference</span>"
                    : row.HasDifference
                        ? @"<span class=""badge badge-err"">Differences</span>"
                        : @"<span class=""badge badge-ok"">Clean</span>";

                // ── CHANGED: href points into Reports/ subfolder ──────────────────────
                string reportCell = !string.IsNullOrEmpty(reportFileName)
                    ? $@"<a class=""report-link"" href=""Reports/{System.Net.WebUtility.HtmlEncode(reportFileName)}"" target=""_blank"">View Report</a>"
                    : "<span style=\"color:var(--text-muted);font-size:12px;\">—</span>";

                sb.AppendLine($@"
        <tr id=""{anchor}"">
          <td class=""col-num"">{idx++}</td>
          <td class=""col-file"">{encodedFname}</td>
          <td class=""col-status"">{statusBadge}</td>
          <td class=""col-report"">{reportCell}</td>
        </tr>");
            }

            sb.AppendLine(@"
      </tbody>
    </table>
  </div>");

            // ── Correction Statistics section (only when at least one file has stats) ──
            var statsRows = rows.Where(r => r.Stats != null).ToList();
            if (statsRows.Count > 0)
            {
                bool showAcceptable = true; // always show all columns

                int totalDiffsAll = statsRows.Sum(r => r.Stats.TotalDiffs);
                int totalRectified = statsRows.Sum(r => r.Stats.RectifiedDiffs);
                int totalAcceptable = statsRows.Sum(r => r.Stats.AcceptableDiffs);
                double overallPctRectified = totalDiffsAll > 0 ? Math.Round((double)totalRectified / totalDiffsAll * 100, 1) : 100.0;
                double overallPctRectifiedOrAcceptable = totalDiffsAll > 0 ? Math.Round((double)(totalRectified + totalAcceptable) / totalDiffsAll * 100, 1) : 100.0;

                sb.AppendLine($@"
  <style>
    .corr-stats-card {{ margin-top: 28px; background: var(--surface); border: 1px solid var(--border); border-radius: var(--radius); box-shadow: var(--shadow-sm); overflow: hidden; }}
    .corr-stats-card .table-card-header {{ display: flex; align-items: center; justify-content: space-between; padding: 16px 20px; border-bottom: 1px solid var(--border); background: #f9f8f5; }}
    .corr-stats-card .table-card-header h2 {{ font-size: 14px; font-weight: 600; color: var(--text-main); }}
    .corr-stats-card table {{ width: 100%; border-collapse: collapse; font-size: 13px; }}
    .corr-stats-card th, .corr-stats-card td {{ padding: 10px 14px; text-align: left; border-bottom: 1px solid var(--border); }}
    .corr-stats-card th {{ font-size: 11px; font-weight: 600; letter-spacing: .05em; text-transform: uppercase; color: var(--text-muted); background: #fafaf8; }}
    .corr-stats-card td.num, .corr-stats-card th.num {{ text-align: center; font-variant-numeric: tabular-nums; }}
    .corr-stats-card tr.summary-row {{ background: #f0f4ff; font-weight: 600; }}
    .pct-green {{ color: #1a7a4a; font-weight: 600; }}
    .pct-na    {{ color: var(--text-muted); font-style: italic; }}
  </style>
  <div class=""corr-stats-card"">
    <div class=""table-card-header"">
      <h2>Correction Statistics</h2>
      <span class=""badge-count"">vs prior comparison</span>
    </div>
    <table>
      <thead>
        <tr>
          <th style=""width:40px"">#</th>
          <th>File</th>
          <th class=""num"">Total Diffs</th>
          <th class=""num"">Rectified</th>
          <th class=""num"">Acceptable</th>
          <th class=""num"">% Rectified</th>
          <th class=""num"">% Rect. or Acceptable</th>
        </tr>
      </thead>
      <tbody>");

                int si = 1;
                foreach (var row in statsRows)
                {
                    string sfname = System.Net.WebUtility.HtmlEncode(System.IO.Path.GetFileName(row.PdfPath));
                    var s = row.Stats;
                    // % columns: N/A when no prior diffs; always green when there were diffs
                    string naSpan = @"<span class=""pct-na"">N/A</span>";
                    string pctRectCell = s.TotalDiffs == 0
                        ? naSpan
                        : $@"<span class=""pct-green"">{s.PercentRectified:0.#}%</span>";
                    string pctCombCell = s.TotalDiffs == 0 ? naSpan : $@"<span class=""pct-green"">{s.PercentRectifiedOrAcceptable:0.#}%</span>";
                    string rectifiedCell = s.TotalDiffs == 0 ? "—" : s.RectifiedDiffs.ToString();
                    string acceptableCell = s.TotalDiffs == 0 ? "—" : s.AcceptableDiffs.ToString();
                    sb.AppendLine($@"
        <tr>
          <td>{si++}</td>
          <td>{sfname}</td>
          <td class=""num"">{s.TotalDiffs}</td>
          <td class=""num"">{rectifiedCell}</td>
          <td class=""num"">{acceptableCell}</td>
          <td class=""num"">{pctRectCell}</td>
          <td class=""num"">{pctCombCell}</td>
        </tr>");
                }

                // Summary totals row — N/A when no prior diffs at all, always green otherwise
                string naSpanSum = @"<span class=""pct-na"">N/A</span>";
                string sumPctRect = totalDiffsAll == 0
                    ? naSpanSum
                    : $@"<span class=""pct-green"">{overallPctRectified:0.#}%</span>";
                string sumPctComb = totalDiffsAll == 0 ? naSpanSum : $@"<span class=""pct-green"">{overallPctRectifiedOrAcceptable:0.#}%</span>";
                sb.AppendLine($@"
        <tr class=""summary-row"">
          <td></td>
          <td>TOTAL</td>
          <td class=""num"">{totalDiffsAll}</td>
          <td class=""num"">{(totalDiffsAll == 0 ? "—" : totalRectified.ToString())}</td>
          <td class=""num"">{(totalDiffsAll == 0 ? "—" : totalAcceptable.ToString())}</td>
          <td class=""num"">{sumPctRect}</td>
          <td class=""num"">{sumPctComb}</td>
        </tr>
      </tbody>
    </table>
  </div>");
            }

            sb.AppendLine($@"
  <div class=""footer"">PDFCompare &nbsp;·&nbsp; Report generated {generated}</div>
</div>
</body>
</html>");

            string summaryReportFileName = System.IO.Path.Combine(summaryFileFolder, summaryFileName);
            File.WriteAllText(summaryReportFileName, sb.ToString());
        }

        // ── Overload: includes Subfolders section in the summary report ───────────────
        public static void GenerateSummaryReport(
            string summaryFileFolder,
            string summaryFileName,
            List<FileComparisonResult> rows,
            List<(string SubFolderName, bool HasAnyDiff, bool AllClean)> subFolders,
            string subFolderReportFileName)
        {
            // Generate the base report (existing logic, unchanged)
            GenerateSummaryReport(summaryFileFolder, summaryFileName, rows);

            // No subfolders — nothing more to do
            if (subFolders == null || subFolders.Count == 0)
                return;

            // Build the subfolder section HTML
            var sbSub = new StringBuilder();
            sbSub.AppendLine($@"
  <div class=""table-card"" style=""margin-top:28px;"">
    <div class=""table-card-header"">
      <h2>Subfolders</h2>
      <span class=""badge-count"">{subFolders.Count} folder{(subFolders.Count == 1 ? "" : "s")}</span>
    </div>
    <table>
      <thead>
        <tr>
          <th class=""col-num"">#</th>
          <th>Folder Name</th>
          <th class=""col-status"">Status</th>
          <th class=""col-report"">Open</th>
        </tr>
      </thead>
      <tbody>");

            int folderIdx = 1;
            foreach (var sf in subFolders)
            {
                string encodedName = System.Net.WebUtility.HtmlEncode(sf.SubFolderName);
                string hrefLink = System.Net.WebUtility.HtmlEncode(
                    sf.SubFolderName + "/" + subFolderReportFileName);
                string badge = sf.HasAnyDiff && !sf.AllClean
                    ? @"<span class=""badge badge-err"">Differences</span>"
                    : sf.HasAnyDiff
                        ? @"<span class=""badge badge-clean-diff"">Clean Difference</span>"
                        : @"<span class=""badge badge-ok"">Clean</span>";
                sbSub.AppendLine($@"
        <tr>
          <td class=""col-num"">{folderIdx++}</td>
          <td class=""col-file"">&#128193; {encodedName}</td>
          <td class=""col-status"">{badge}</td>
          <td class=""col-report""><a class=""report-link"" href=""{hrefLink}"">Open &rarr;</a></td>
        </tr>");
            }
            sbSub.AppendLine(@"      </tbody>
    </table>
  </div>");

            // Inject the subfolder section before the footer in the already-written file
            string reportFilePath = System.IO.Path.Combine(summaryFileFolder, summaryFileName);
            string content = File.ReadAllText(reportFilePath);
            const string footerMarker = "\n  <div class=\"footer\">";
            int insertIdx = content.LastIndexOf(footerMarker, StringComparison.Ordinal);
            if (insertIdx >= 0)
            {
                content = content.Insert(insertIdx, "\n" + sbSub.ToString());
                File.WriteAllText(reportFilePath, content);
            }
        }

        public bool GenerateHtmlReport(
            string pdf1, string pdf2,
            List<FieldDiff> fieldDiffs, List<TextDiff> textDiffs,
            out string reportFileName,
            out bool allDiffsAreClean,
            out List<DiffRectangleInfo> diffRectangles,
            List<TextInfo> text1, List<TextInfo> text2,
            List<PhraseDiffGroup> phraseGroups = null,
            HashSet<(int page, float x, float y)> cleanWordPositions = null)
        {
            allDiffsAreClean = false;
            diffRectangles = new List<DiffRectangleInfo>(); // default empty; overwritten below
            if ((fieldDiffs == null || fieldDiffs.Count == 0) && (textDiffs == null || textDiffs.Count == 0))
            {
                Console.WriteLine("No differences found. Generating clean report.");
                var pdf2FileNameClean = System.IO.Path.GetFileNameWithoutExtension(pdf2);
                var safePdf2FileNameClean = string.Concat(pdf2FileNameClean.Split(System.IO.Path.GetInvalidFileNameChars()));
                string reportsSubDirClean = Path.Combine(ReportBasePath, "Reports");
                Directory.CreateDirectory(reportsSubDirClean);
                reportFileName = Path.Combine(reportsSubDirClean, $"{safePdf2FileNameClean}.html");
                string summaryBackLinkClean = string.IsNullOrEmpty(_ReportFileName) ? "../" : $"../{_ReportFileName}";
                string generatedClean = DateTime.Now.ToString("MMMM dd, yyyy  h:mm tt");
                string cleanHtml = $@"<!DOCTYPE html>
<html lang=""en"">
<head>
<meta charset=""UTF-8"" />
<title>Clean – {System.Net.WebUtility.HtmlEncode(pdf2FileNameClean)}</title>
<style>
  body {{ font-family: sans-serif; background: #f5f4f0; margin: 0; padding: 40px; }}
  .topbar {{ background: #fff; border-bottom: 1px solid #e8e6e0; padding: 12px 24px; display: flex; justify-content: space-between; align-items: center; margin: -40px -40px 40px; }}
  .back-link {{ font-size: 13px; color: #4a6fa5; text-decoration: none; }}
  .topbar-title {{ font-size: 14px; font-weight: 600; color: #2d2d2d; }}
  .topbar-meta {{ font-size: 11.5px; color: #999; }}
  .clean-card {{ background: #fff; border: 1px solid #e8e6e0; border-radius: 8px; padding: 48px; text-align: center; max-width: 600px; margin: 80px auto 0; }}
  .clean-icon {{ font-size: 48px; margin-bottom: 16px; color: #1aaf5d; }}
  h1 {{ font-size: 22px; font-weight: 600; color: #1aaf5d; margin-bottom: 8px; }}
  p {{ color: #666; font-size: 14px; }}
</style>
</head>
<body>
<div class=""topbar"">
  <a class=""back-link"" href=""{System.Net.WebUtility.HtmlEncode(summaryBackLinkClean)}"">← Summary</a>
  <div class=""topbar-title"">Diff Report — {System.Net.WebUtility.HtmlEncode(pdf2FileNameClean)}</div>
  <div class=""topbar-meta"">Generated {generatedClean}</div>
</div>
<div class=""clean-card"">
  <div class=""clean-icon"">✓</div>
  <h1>No Differences Found</h1>
  <p>{System.Net.WebUtility.HtmlEncode(pdf2FileNameClean)}</p>
  <p style=""margin-top:8px;font-size:12px;color:#999;"">V14 and V23 PDFs are identical.</p>
</div>
</body>
</html>";
                File.WriteAllText(reportFileName, cleanHtml);
                return false;
            }

            bool blnIsThereDifference = false;
            var pdf2FileName = System.IO.Path.GetFileNameWithoutExtension(pdf2);
            var safePdf2FileName = string.Concat(pdf2FileName.Split(System.IO.Path.GetInvalidFileNameChars()));

            // ── CHANGED: write HTML into Reports\ subfolder ──────────────────────────
            string reportsSubDir = Path.Combine(ReportBasePath, "Reports");
            Directory.CreateDirectory(reportsSubDir);
            var reportPath = Path.Combine(reportsSubDir, $"{safePdf2FileName}.html");
            reportFileName = reportPath;
            // ─────────────────────────────────────────────────────────────────────────

            string pdf1Name = System.IO.Path.GetFileName(pdf1);
            string pdf1Folder = System.IO.Path.GetDirectoryName(pdf1) ?? "";
            string pdf2Name = System.IO.Path.GetFileName(pdf2);
            string pdf2Folder = System.IO.Path.GetDirectoryName(pdf2) ?? "";

            // Back-link points one level up to the summary in ReportBasePath
            string summaryBackLink = string.IsNullOrEmpty(_ReportFileName)
                ? "../"
                : $"../{_ReportFileName}";

            var ignoreRegionsPdf1 = ComputeIgnoreRegions(text1);
            ignoreRegionsPdf1 = ignoreRegionsPdf1.Concat(this.lstIgnoreAnnotationRegions).ToList();
            var ignoreRegionsPdf2 = ComputeIgnoreRegions(text2);
            ignoreRegionsPdf2 = ignoreRegionsPdf2.Concat(this.lstIgnoreAnnotationRegions).ToList();

            string generated = DateTime.Now.ToString("MMMM dd, yyyy  h:mm tt");

            var html = new StringBuilder();
            html.AppendLine($@"<!DOCTYPE html>
<html lang=""en"">
<head>
<meta charset=""UTF-8"" />
<meta name=""viewport"" content=""width=device-width, initial-scale=1.0"" />
<title>Diff – {System.Net.WebUtility.HtmlEncode(pdf2FileName)}</title>
<link rel=""preconnect"" href=""https://fonts.googleapis.com"" />
<link rel=""preconnect"" href=""https://fonts.gstatic.com"" crossorigin />
<link href=""https://fonts.googleapis.com/css2?family=Lora:wght@400;600&family=Source+Sans+3:wght@300;400;500;600&display=swap"" rel=""stylesheet"" />
<style>
  *, *::before, *::after {{ box-sizing: border-box; margin: 0; padding: 0; }}
  :root {{
    --bg:           #f5f4f0;
    --surface:      #ffffff;
    --border:       #e2e0db;
    --text-main:    #1c1b18;
    --text-muted:   #6b6860;
    --accent:       #2a5bd7;
    --accent-lt:    #eef2fd;
    --red:          #c0392b;
    --red-lt:       #fdf0ee;
    --shadow-sm:    0 1px 3px rgba(0,0,0,.07), 0 1px 2px rgba(0,0,0,.05);
    --shadow-md:    0 4px 16px rgba(0,0,0,.09), 0 2px 4px rgba(0,0,0,.04);
    --radius:       10px;
  }}
  *, *::before, *::after {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{
    font-family: 'Source Sans 3', sans-serif;
    font-weight: 400;
    background: var(--bg);
    color: var(--text-main);
    min-height: 100vh;
  }}
  .topbar {{
    background: var(--surface);
    border-bottom: 1px solid var(--border);
    padding: 0 40px;
    height: 60px;
    display: flex;
    align-items: center;
    justify-content: space-between;
    position: sticky;
    top: 0;
    z-index: 100;
    box-shadow: var(--shadow-sm);
  }}
  .topbar-left {{ display: flex; align-items: center; gap: 18px; }}
  .back-link {{
    font-size: 13px;
    color: var(--accent);
    text-decoration: none;
    font-weight: 500;
    display: inline-flex;
    align-items: center;
    gap: 5px;
    padding: 5px 12px;
    border-radius: 7px;
    border: 1px solid #c9d8f5;
    background: var(--accent-lt);
    transition: background .15s;
  }}
  .back-link:hover {{ background: #dde7fb; }}
  .back-link::before {{ content: '←'; }}
  .topbar-title {{
    font-family: 'Lora', serif;
    font-size: 15px;
    font-weight: 600;
    color: var(--text-main);
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
    max-width: 480px;
  }}
  .topbar-meta {{
    font-size: 12px;
    color: var(--text-muted);
    font-weight: 300;
    white-space: nowrap;
  }}
  .page {{ max-width: 1300px; margin: 0 auto; padding: 40px 32px 80px; }}
  .page-heading {{ margin-bottom: 32px; display: flex; align-items: flex-start; gap: 14px; }}
  .heading-icon {{
    width: 42px; height: 42px; border-radius: 10px;
    background: var(--red-lt); display: flex; align-items: center; justify-content: center;
    font-size: 20px; flex-shrink: 0; margin-top: 2px;
  }}
  .page-heading h1 {{
    font-family: 'Lora', serif;
    font-size: 24px; font-weight: 600;
    letter-spacing: -.02em; margin-bottom: 4px;
  }}
  .page-heading p {{ font-size: 13.5px; color: var(--text-muted); font-weight: 300; }}
  .diff-count-bar {{
    display: flex; align-items: center; gap: 12px;
    margin-bottom: 28px;
    padding: 12px 18px;
    background: var(--red-lt);
    border: 1px solid #f0c9c5;
    border-radius: var(--radius);
    font-size: 13.5px; color: var(--red); font-weight: 500;
  }}
  .diff-count-bar .num {{
    font-family: 'Lora', serif; font-size: 22px; font-weight: 600;
  }}
  .diff-card {{
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: var(--radius);
    box-shadow: var(--shadow-md);
    margin-bottom: 28px;
    overflow: hidden;
  }}
  .panels {{
    display: grid;
    grid-template-columns: 1fr 1fr;
  }}
  .panel {{ padding: 0; }}
  .panel + .panel {{ border-left: 2px solid var(--border); }}
  .panel-header {{
    padding: 12px 18px;
    border-bottom: 1px solid var(--border);
    background: #fafaf8;
    display: flex;
    flex-direction: column;
    gap: 3px;
  }}
  .panel-label {{
    font-size: 10.5px;
    font-weight: 600;
    letter-spacing: .08em;
    text-transform: uppercase;
    color: var(--text-muted);
  }}
  .panel-label.baseline {{ color: #1a6e9e; }}
  .panel-label.publish  {{ color: #7a3ec0; }}
  .panel-filename {{
    font-size: 13px; font-weight: 600; color: var(--text-main);
    white-space: nowrap; overflow: hidden; text-overflow: ellipsis;
  }}
  .panel-folder {{
    font-size: 11px; color: var(--text-muted); font-weight: 300;
    white-space: nowrap; overflow: hidden; text-overflow: ellipsis;
  }}
  .panel-img {{
    padding: 16px;
    display: flex;
    justify-content: center;
    align-items: flex-start;
    background: #fbfaf7;
  }}
  .panel-img img {{
    max-width: 100%;
    border-radius: 4px;
    border: 1px solid var(--border);
    box-shadow: 0 2px 8px rgba(0,0,0,.06);
    display: block;
  }}
  .diff-nav {{
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: var(--radius);
    padding: 14px 20px;
    margin-bottom: 28px;
    display: flex;
    align-items: center;
    gap: 12px;
    flex-wrap: wrap;
  }}
  .diff-nav span {{
    font-size: 11.5px; font-weight: 600;
    letter-spacing: .06em; text-transform: uppercase;
    color: var(--text-muted); margin-right: 4px;
  }}
  .diff-anchor {{
    font-size: 12px; color: var(--accent); text-decoration: none;
    font-weight: 500; padding: 3px 9px; border-radius: 5px;
    border: 1px solid #c9d8f5; background: var(--accent-lt);
    transition: background .15s;
  }}
  .diff-anchor:hover {{ background: #dde7fb; }}
  .footer {{
    margin-top: 56px; text-align: center;
    font-size: 12px; color: var(--text-muted); font-weight: 300;
  }}
  .back-to-top {{
    display: inline-block; margin-top: 12px;
    font-size: 12.5px; color: var(--accent);
    text-decoration: none; font-weight: 500;
  }}
</style>
</head>
<body>");

            html.AppendLine($@"
<div class=""topbar"">
  <div class=""topbar-left"">
    <a class=""back-link"" href=""{System.Net.WebUtility.HtmlEncode(summaryBackLink)}"">Summary</a>
    <div class=""topbar-title"">Diff Report &mdash; {System.Net.WebUtility.HtmlEncode(pdf2FileName)}</div>
  </div>
  <div class=""topbar-meta"">Generated {generated}</div>
</div>

<div class=""page"">
  <div class=""page-heading"">
    <div class=""heading-icon"">⚡</div>
    <div>
      <h1>Difference Report</h1>
      <p>{System.Net.WebUtility.HtmlEncode(pdf2FileName)}</p>
    </div>
  </div>");

            // ── Build merged field diffs ──────────────────────────────────────────────
            var mergedFieldDiffs = new List<(RectangleF rect1, RectangleF rect2, int page1, int page2, string mergedText, DiffCategory categories)>();
            if (fieldDiffs != null)
            {
                foreach (var diff in fieldDiffs)
                {
                    int page1 = diff.Info1?.Page ?? diff.Info2.Page;
                    int page2 = diff.Info2?.Page ?? diff.Info1.Page;
                    var rect1 = new RectangleF(diff.Info1?.X ?? diff.Info2.X, diff.Info1?.Y ?? diff.Info2.Y,
                                               diff.Info1?.Width ?? diff.Info2.Width, diff.Info1?.Height ?? diff.Info2.Height);
                    var rect2 = new RectangleF(diff.Info2?.X ?? diff.Info1.X, diff.Info2?.Y ?? diff.Info1.Y,
                                               diff.Info2?.Width ?? diff.Info1.Width, diff.Info2?.Height ?? diff.Info1.Height);
                    bool merged = false;
                    for (int i = 0; i < mergedFieldDiffs.Count; i++)
                    {
                        var existing = mergedFieldDiffs[i];
                        if (existing.page1 == page1 && existing.page2 == page2 &&
                            AreRectanglesClose(existing.rect1, rect1, 10) && AreRectanglesClose(existing.rect2, rect2, 10))
                        {
                            mergedFieldDiffs[i] = (RectangleF.Union(existing.rect1, rect1),
                                                   RectangleF.Union(existing.rect2, rect2),
                                                   page1, page2,
                                                   existing.mergedText + "; " + $"[FieldDiff] {diff.Name}: {diff.Issue}",
                                                   existing.categories | diff.Categories);
                            merged = true; break;
                        }
                    }
                    if (!merged)
                        mergedFieldDiffs.Add((rect1, rect2, page1, page2, $"[FieldDiff] {diff.Name}: {diff.Issue}", diff.Categories));
                }
            }

            // ── Build merged text diffs ───────────────────────────────────────────────
            // actual1/actual2: bounding boxes built ONLY from words that truly exist in pdf1/pdf2.
            // These are used for X/Y classification (rect1/rect2 are polluted by null-fallback positions
            // that make "Missing or shifted" and "Extra in PDF2" diffs look like they're at the same spot).
            var mergedTextDiffs = new List<(RectangleF rect1, RectangleF rect2, int page1, int page2, List<string> textDescriptions, DiffCategory categories, RectangleF? actual1, RectangleF? actual2, List<TextDiff> wordDiffs)>();

            static RectangleF? UnionNullable(RectangleF? a, RectangleF? b) =>
                a.HasValue && b.HasValue ? RectangleF.Union(a.Value, b.Value) : a ?? b;

            if (textDiffs != null)
            {
                foreach (var diff in textDiffs)
                {
                    int page1 = diff.Info1?.Page ?? diff.Info2.Page;
                    int page2 = diff.Info2?.Page ?? diff.Info1.Page;
                    float x1 = diff.Info1?.X ?? diff.Info2.X, y1 = diff.Info1?.Y ?? diff.Info2.Y;
                    float x2 = diff.Info2?.X ?? diff.Info1.X, y2 = diff.Info2?.Y ?? diff.Info1.Y;
                    float w1 = EstimateWidth(diff.Info1?.Text ?? diff.Info2?.Text);
                    float h1 = EstimateHeight(diff.Info1?.FontSize ?? diff.Info2?.FontSize ?? 12);
                    float w2 = EstimateWidth(diff.Info2?.Text ?? diff.Info1?.Text);
                    float h2 = EstimateHeight(diff.Info2?.FontSize ?? diff.Info1?.FontSize ?? 12);
                    var rect1 = new RectangleF(x1, y1, w1, h1);
                    var rect2 = new RectangleF(x2, y2, w2, h2);
                    // actual1/actual2: only set when the side truly exists (no null-fallback)
                    RectangleF? actual1 = diff.Info1 != null ? new RectangleF(diff.Info1.X, diff.Info1.Y, w1, h1) : (RectangleF?)null;
                    RectangleF? actual2 = diff.Info2 != null ? new RectangleF(diff.Info2.X, diff.Info2.Y, w2, h2) : (RectangleF?)null;
                    string cleanedText = diff.Text?.Trim();
                    string desc = (!string.IsNullOrEmpty(cleanedText) && !string.IsNullOrWhiteSpace(cleanedText))
                                  ? $"[TextDiff] \"{cleanedText}\": {diff.Issue}" : null;
                    bool merged = false;
                    for (int i = 0; i < mergedTextDiffs.Count; i++)
                    {
                        var existing = mergedTextDiffs[i];
                        if (existing.page1 == page1 && existing.page2 == page2 &&
                            AreRectanglesClose(existing.rect1, rect1, 10) && AreRectanglesClose(existing.rect2, rect2, 10))
                        {
                            var dl = existing.textDescriptions;
                            if (!string.IsNullOrEmpty(desc)) dl.Add(desc);
                            existing.wordDiffs.Add(diff);
                            mergedTextDiffs[i] = (RectangleF.Union(existing.rect1, rect1),
                                                  RectangleF.Union(existing.rect2, rect2),
                                                  page1, page2, dl, existing.categories | diff.Categories,
                                                  UnionNullable(existing.actual1, actual1),
                                                  UnionNullable(existing.actual2, actual2),
                                                  existing.wordDiffs);
                            merged = true; break;
                        }
                    }
                    if (!merged)
                    {
                        var dl = new List<string>();
                        if (!string.IsNullOrEmpty(desc)) dl.Add(desc);
                        mergedTextDiffs.Add((rect1, rect2, page1, page2, dl, diff.Categories, actual1, actual2, new List<TextDiff> { diff }));
                    }
                }
            }

            // ── Secondary merge pass: pair MISSING-only groups with EXTRA-only groups ────
            // When MISSING (V14 only) and EXTRA (V23 only) words for the same text appear
            // at different Y positions (Y diff > 10pt primary tolerance), they end up in
            // separate mergedTextDiff groups and XDiff detection cannot pair them.
            // This pass merges such groups on the same page when they share matching text
            // and their actual X extents are close — enabling the existing XDiff logic to
            // detect the horizontal shift.
            {
                // Identify pure-MISSING and pure-EXTRA groups by index
                var missingOnlyIdx = new List<int>();
                var extraOnlyIdx   = new List<int>();
                for (int i = 0; i < mergedTextDiffs.Count; i++)
                {
                    var mt = mergedTextDiffs[i];
                    bool allMissing = mt.wordDiffs.Count > 0 && mt.wordDiffs.All(d => d.Info1 != null && d.Info2 == null);
                    bool allExtra   = mt.wordDiffs.Count > 0 && mt.wordDiffs.All(d => d.Info1 == null && d.Info2 != null);
                    if (allMissing) missingOnlyIdx.Add(i);
                    else if (allExtra) extraOnlyIdx.Add(i);
                }

                var mergedAwayIdx = new HashSet<int>();

                foreach (int mi in missingOnlyIdx)
                {
                    if (mergedAwayIdx.Contains(mi)) continue;
                    var missing = mergedTextDiffs[mi];

                    // Build text set from MISSING side (Info1)
                    var missingTexts = new HashSet<string>(
                        missing.wordDiffs.Select(d => d.Info1?.Text).Where(t => t != null),
                        StringComparer.Ordinal);

                    foreach (int ei in extraOnlyIdx)
                    {
                        if (mergedAwayIdx.Contains(ei)) continue;
                        var extra = mergedTextDiffs[ei];

                        // Must be on the same page
                        if (extra.page1 != missing.page1) continue;

                        // At least one word text must match between the two groups
                        bool hasTextMatch = extra.wordDiffs.Any(
                            d => d.Info2?.Text != null && missingTexts.Contains(d.Info2.Text));
                        if (!hasTextMatch) continue;

                        // Actual X extents must overlap or be within 30pt (guards against
                        // unrelated words that happen to share a common short text token)
                        const float X_CROSS_TOL = 30f;
                        bool xClose = missing.actual1.HasValue && extra.actual2.HasValue &&
                            missing.actual1.Value.Right  + X_CROSS_TOL > extra.actual2.Value.Left &&
                            extra.actual2.Value.Right    + X_CROSS_TOL > missing.actual1.Value.Left;
                        if (!xClose) continue;

                        // Merge the EXTRA group into the MISSING group
                        var combinedDesc = missing.textDescriptions
                            .Concat(extra.textDescriptions)
                            .Distinct()
                            .ToList();
                        var combinedWords = missing.wordDiffs.Concat(extra.wordDiffs).ToList();

                        mergedTextDiffs[mi] = (
                            RectangleF.Union(missing.rect1, extra.rect1),
                            RectangleF.Union(missing.rect2, extra.rect2),
                            missing.page1, missing.page2,
                            combinedDesc,
                            missing.categories | extra.categories,
                            UnionNullable(missing.actual1, extra.actual1),
                            UnionNullable(missing.actual2, extra.actual2),
                            combinedWords);

                        // Update local reference for subsequent EXTRA candidates
                        missing = mergedTextDiffs[mi];
                        missingTexts.UnionWith(
                            extra.wordDiffs.Select(d => d.Info2?.Text).Where(t => t != null));

                        mergedAwayIdx.Add(ei);
                        _logger.Log($"[CROSS-GROUP-MERGE] pg={missing.page1} merged EXTRA group " +
                            $"(rect=({extra.rect2.X:0.##},{extra.rect2.Y:0.##}), {extra.wordDiffs.Count} word(s)) " +
                            $"into MISSING group (rect=({mergedTextDiffs[mi].rect1.X:0.##},{mergedTextDiffs[mi].rect1.Y:0.##})) " +
                            $"— Y-displaced MISSING+EXTRA pair, enables XDiff detection");
                    }
                }

                // Remove groups that were merged away (descending order to preserve lower indices)
                foreach (int idx in mergedAwayIdx.OrderByDescending(x => x))
                    mergedTextDiffs.RemoveAt(idx);
            }

            // ── Post-process mergedTextDiffs: compute XDiff/YDiff ─────────────────────────
            //
            // XDiff — text-content pairing (reliable at any shift magnitude, no false positives):
            //   Pair each "Missing or shifted" word (Info1 set, Info2 null) with an "Extra" word
            //   (Info1 null, Info2 set) in the same group that has IDENTICAL text content.
            //   These pairs are the same word appearing at different X positions → their X delta
            //   is the true block-level X shift for that word. Mean across all pairs = block shift.
            //   · Genuine block X shift (table, paragraph, cell shifted): all words match up;
            //     consistent non-zero X delta → mean ≠ 0 → XDiff. Works at any magnitude (2pt+).
            //   · Content change (different words replaced): Missing and Extra texts differ →
            //     few/no pairs → no XDiff. Immune to false positives.
            //   · Word wrapping (same words, same hanging-indent X, different Y):
            //     pairs found but X delta ≈ 0 → mean ≈ 0 → no XDiff.
            //   Formatting-mismatch diffs (both sides set, matched within 2pt X) are excluded —
            //   they contribute ~0 X delta and would dilute the mean.
            //   Threshold: 2pt on the mean (stable average; not a per-word noisy comparison).
            //
            // YDiff — union bounding-box approach (reliable for Y):
            //   actual1.Y vs actual2.Y. Threshold: 3pt.
            for (int i = 0; i < mergedTextDiffs.Count; i++)
            {
                var mt = mergedTextDiffs[i];
                var cats = mt.categories; // carries WDiff from word level

                // XDiff: pair Missing↔Extra words by text content, then validate with four
                // criteria that together distinguish genuine block-level X shifts from artefacts:
                //
                // 1. pairRatio ≥ 0.5: at least half of Missing words find a same-text match in
                //    Extra. Genuine block shifts have the same words on both sides → high ratio.
                //    Content changes produce different replacement words → few pairs → low ratio.
                //
                // 2. |mean| in (4pt, 50pt): the average X delta must be a plausible block shift.
                //    Word wrapping moves words 100pt+ across a line → caught by the 50pt upper bound.
                //    Noise/rendering variation is < 4pt → caught by the lower bound.
                //
                // 3. stddev < 2pt: a true uniform block shift has all pairs at nearly the same
                //    delta. Rendering accumulation produces a rising series of deltas (stddev 3–5pt)
                //    and is rejected here even if mean happens to be high.
                //
                // 4. firstPairRatio > 0.6: for a genuine block shift ALL words (including the
                //    leftmost) shift by D → leftmost pair delta ≈ mean → ratio ≈ 1.0.
                //    For font-rendering accumulation the leftmost words are matched within 2pt
                //    (only later words exceed the threshold) → leftmost pair delta ≈ 2–3pt << mean
                //    → ratio ≈ 0.3–0.4. This is the key discriminator that allows mean to be as
                //    low as 4pt while still rejecting accumulation false positives.
                var missing = mt.wordDiffs.Where(d => d.Info1 != null && d.Info2 == null).ToList();
                var extra   = mt.wordDiffs.Where(d => d.Info1 == null && d.Info2 != null).ToList();
                var usedExtra = new HashSet<int>();
                var pairDeltas = new List<(float missingX, float delta)>();
                float sumX = 0f, sumXSq = 0f;
                int pairCount = 0;
                foreach (var m in missing)
                {
                    for (int j = 0; j < extra.Count; j++)
                    {
                        if (!usedExtra.Contains(j) &&
                            string.Equals(m.Info1.Text, extra[j].Info2.Text, StringComparison.Ordinal))
                        {
                            float delta = extra[j].Info2.X - m.Info1.X;
                            pairDeltas.Add((m.Info1.X, delta));
                            sumX += delta;
                            sumXSq += delta * delta;
                            pairCount++;
                            usedExtra.Add(j);
                            break;
                        }
                    }
                }
                if (pairCount >= 1 && missing.Count > 0)
                {
                    float pairRatio = (float)pairCount / missing.Count;
                    float mean = sumX / pairCount;
                    float variance = (sumXSq / pairCount) - (mean * mean);
                    float stddev = (float)Math.Sqrt(Math.Max(0f, variance));
                    // firstPairRatio: leftmost paired word's delta relative to mean.
                    // Genuine block shift → all words shift uniformly → leftmost delta ≈ mean → ratio ≈ 1.0.
                    // Rendering accumulation → leftmost words matched (within 2pt), only later words
                    // paired → leftmost delta << mean → ratio ≈ 0.3–0.4.
                    var leftmost = pairDeltas.MinBy(p => p.missingX);
                    bool sameDir = (mean > 0 && leftmost.delta > 0) || (mean < 0 && leftmost.delta < 0);
                    float firstPairRatio = (sameDir && Math.Abs(mean) > 0f)
                        ? Math.Abs(leftmost.delta) / Math.Abs(mean) : 0f;
                    if (pairRatio >= 0.5f && Math.Abs(mean) > 4f && Math.Abs(mean) < 50f
                        && stddev < 2f && firstPairRatio > 0.6f)
                        cats |= DiffCategory.XDiff;
                }

                // YDiff: union bounding-box approach.
                if (mt.actual1.HasValue && mt.actual2.HasValue)
                {
                    if (Math.Abs(mt.actual1.Value.Y - mt.actual2.Value.Y) > 3f) cats |= DiffCategory.YDiff;
                }
                mergedTextDiffs[i] = (mt.rect1, mt.rect2, mt.page1, mt.page2, mt.textDescriptions, cats, mt.actual1, mt.actual2, mt.wordDiffs);
            }

            // ── XI: X-Ignore reconciliation — leftmost XDiff rects already at pivot X ──
            //
            // After all X/W/Y categories are computed, check whether leftmost-aligned XDiff
            // rects (both text diffs AND field diffs) have their baseline position already at
            // the canonical "pivot left X".  If so, applying an X drift correction would shift
            // text AWAY from proper alignment.  Such rects are marked XI (XDiff cleared) and
            // treated as green/acceptable.
            //
            // Pivot X is determined per-page in this priority order:
            //   A. Leftmost baseline text NOT overlapping any diff rect → mode X of those words.
            //   B. Leftmost green rects (no XDiff) from either text or field diffs → mode X.
            //   C. Majority X among all XDiff rects; ties → leftmost center.
            //
            // Note: field diffs use rect1.X directly; text diffs prefer actual1.X (true word
            // position) over rect1.X.  Both are baseline (V14/PDF1) coordinates.
            {
                const float XI_LEFTMOST_TOL = 5f;   // rects within 5pt of page min-X = "leftmost"
                const float XI_ALIGN_TOL    = 3f;   // baseline X within 3pt of pivot = "already aligned"
                const float XI_FREE_WINDOW  = 10f;  // leftmost free-text window

                // ── Build per-page indices for both diff lists ────────────────────────
                var pageToTextIdx = new Dictionary<int, List<int>>();
                for (int i = 0; i < mergedTextDiffs.Count; i++)
                {
                    int pg = mergedTextDiffs[i].page1;
                    if (!pageToTextIdx.TryGetValue(pg, out var lst)) { lst = new List<int>(); pageToTextIdx[pg] = lst; }
                    lst.Add(i);
                }
                var pageToFieldIdx = new Dictionary<int, List<int>>();
                for (int i = 0; i < mergedFieldDiffs.Count; i++)
                {
                    int pg = mergedFieldDiffs[i].page1;
                    if (!pageToFieldIdx.TryGetValue(pg, out var lst)) { lst = new List<int>(); pageToFieldIdx[pg] = lst; }
                    lst.Add(i);
                }

                var allPages = pageToTextIdx.Keys.Union(pageToFieldIdx.Keys).ToList();
                _logger.Log($"[XI-ENTRY] XI block start: totalTextDiffs={mergedTextDiffs.Count} totalFieldDiffs={mergedFieldDiffs.Count} pages={string.Join(",", allPages)}");

                foreach (int page in allPages)
                {
                    pageToTextIdx.TryGetValue(page, out var textIndices);
                    pageToFieldIdx.TryGetValue(page, out var fieldIndices);
                    textIndices  = textIndices  ?? new List<int>();
                    fieldIndices = fieldIndices ?? new List<int>();

                    _logger.Log($"[XI-PAGE] pg={page} textDiffs={textIndices.Count} fieldDiffs={fieldIndices.Count}");

                    // Log all text diffs on this page
                    foreach (int i in textIndices)
                    {
                        var mt = mergedTextDiffs[i];
                        float bx = mt.actual1.HasValue ? mt.actual1.Value.X : mt.rect1.X;
                        _logger.Log($"  [XI-TEXT] idx={i} rect=({mt.rect1.X:0.##},{mt.rect1.Y:0.##}) bx={bx:0.##} cats={mt.categories} actual1={( mt.actual1.HasValue ? $"({mt.actual1.Value.X:0.##},{mt.actual1.Value.Y:0.##})" : "null")}");
                    }

                    // Log all field diffs on this page
                    foreach (int i in fieldIndices)
                    {
                        var mf = mergedFieldDiffs[i];
                        _logger.Log($"  [XI-FIELD] idx={i} rect=({mf.rect1.X:0.##},{mf.rect1.Y:0.##}) bx={mf.rect1.X:0.##} cats={mf.categories} text={mf.mergedText}");
                    }

                    // Baseline X getter: text diffs prefer actual1.X; field diffs use rect1.X
                    float TextDiffX(int i)  => mergedTextDiffs[i].actual1.HasValue
                                                ? mergedTextDiffs[i].actual1.Value.X
                                                : mergedTextDiffs[i].rect1.X;
                    float FieldDiffX(int i) => mergedFieldDiffs[i].rect1.X;

                    // Collect XDiff indices from both lists on this page
                    var xDiffTextIdx  = textIndices .Where(i => (mergedTextDiffs[i] .categories & DiffCategory.XDiff) != 0).ToList();
                    var xDiffFieldIdx = fieldIndices.Where(i => (mergedFieldDiffs[i].categories & DiffCategory.XDiff) != 0).ToList();
                    if (xDiffTextIdx.Count == 0 && xDiffFieldIdx.Count == 0)
                    {
                        _logger.Log($"[XI-SKIP] pg={page} no XDiff items on this page -> skip");
                        continue;
                    }

                    _logger.Log($"[XI-XDIFF] pg={page} xDiffText={xDiffTextIdx.Count} xDiffField={xDiffFieldIdx.Count}");
                    foreach (int i in xDiffTextIdx)
                        _logger.Log($"  [XI-XDIFF-TEXT] idx={i} bx={TextDiffX(i):0.##} cats={mergedTextDiffs[i].categories}");
                    foreach (int i in xDiffFieldIdx)
                        _logger.Log($"  [XI-XDIFF-FIELD] idx={i} bx={FieldDiffX(i):0.##} cats={mergedFieldDiffs[i].categories} text={mergedFieldDiffs[i].mergedText}");

                    // Combined minimum X across all XDiff items on this page
                    var allXDiffX = xDiffTextIdx.Select(TextDiffX).Concat(xDiffFieldIdx.Select(FieldDiffX)).ToList();
                    float minXDiffX = allXDiffX.Min();

                    // Leftmost XDiff candidates from both lists
                    var leftmostTextIdx  = xDiffTextIdx .Where(i => TextDiffX(i)  <= minXDiffX + XI_LEFTMOST_TOL).ToList();
                    var leftmostFieldIdx = xDiffFieldIdx.Where(i => FieldDiffX(i) <= minXDiffX + XI_LEFTMOST_TOL).ToList();

                    _logger.Log($"[XI-LEFTMOST] pg={page} minXDiffX={minXDiffX:0.##} tol={XI_LEFTMOST_TOL}pt leftmostText={leftmostTextIdx.Count} leftmostField={leftmostFieldIdx.Count}");

                    if (leftmostTextIdx.Count == 0 && leftmostFieldIdx.Count == 0)
                    {
                        _logger.Log($"[XI-SKIP] pg={page} no leftmost XDiff candidates -> skip");
                        continue;
                    }

                    // Option A: baseline text words not inside any X-diff rect on this page.
                    // Words inside Y-only or W-only diff rects still have a valid, unchanged X
                    // position and are legitimate pivot references — exclude only X-diff rects.
                    var xDiffRects1OnPage = textIndices
                        .Where(i => (mergedTextDiffs[i].categories & DiffCategory.XDiff) != 0)
                        .Select(i => mergedTextDiffs[i].rect1)
                        .Concat(fieldIndices
                            .Where(i => (mergedFieldDiffs[i].categories & DiffCategory.XDiff) != 0)
                            .Select(i => mergedFieldDiffs[i].rect1))
                        .ToList();
                    var freeTextX = text1
                        .Where(t => t.Page == page && !xDiffRects1OnPage.Any(r =>
                            t.X >= r.X - 2f && t.X <= r.Right + 2f &&
                            t.Y >= r.Y - 2f && t.Y <= r.Bottom + 2f))
                        .Select(t => t.X)
                        .ToList();

                    // Option B: non-XDiff rects (green) from both lists on this page
                    var greenRectX = textIndices .Where(i => (mergedTextDiffs[i] .categories & DiffCategory.XDiff) == 0).Select(TextDiffX)
                              .Concat(fieldIndices.Where(i => (mergedFieldDiffs[i].categories & DiffCategory.XDiff) == 0).Select(FieldDiffX))
                              .ToList();

                    // Diagnostic: log xDiff exclusion rects and raw text1 entries on this page
                    _logger.Log($"[XI-FREETXT-DBG] pg={page} xDiffExcludeRects={xDiffRects1OnPage.Count} text1OnPage={text1.Count(t => t.Page == page)}");
                    foreach (var r in xDiffRects1OnPage)
                        _logger.Log($"  [XI-EXCL-RECT] rect=({r.X:0.##},{r.Y:0.##},{r.Right:0.##},{r.Bottom:0.##})");
                    foreach (var t in text1.Where(t => t.Page == page).OrderBy(t => t.X).Take(30))
                    {
                        bool excluded = xDiffRects1OnPage.Any(r =>
                            t.X >= r.X - 2f && t.X <= r.Right + 2f &&
                            t.Y >= r.Y - 2f && t.Y <= r.Bottom + 2f);
                        _logger.Log($"  [XI-TEXT1] x={t.X:0.##} y={t.Y:0.##} text={t.Text} excluded={excluded}");
                    }

                    var cleanPosX = cleanWordPositions == null ? new List<float>() :
                        cleanWordPositions.Where(p => p.page == page).Select(p => p.x).ToList();

                    _logger.Log($"[XI-OPTIONS] pg={page} optClean(cleanPos)={cleanPosX.Count} optA(freeText)={freeTextX.Count} optB(green)={greenRectX.Count} (optC/XDiff removed per requirement)");
                    if (cleanPosX.Count > 0) _logger.Log($"  [XI-OPT-CLEAN] cleanPosX min={cleanPosX.Min():0.##} values=[{string.Join(",", cleanPosX.Take(20).Select(x => x.ToString("0.##")))}]");
                    if (freeTextX.Count > 0) _logger.Log($"  [XI-OPT-A] freeTextX min={freeTextX.Min():0.##} max={freeTextX.Max():0.##} allValues=[{string.Join(",", freeTextX.OrderBy(x => x).Select(x => x.ToString("0.##")))}]");
                    if (greenRectX.Count > 0) _logger.Log($"  [XI-OPT-B] greenRectX values=[{string.Join(",", greenRectX.Select(x => x.ToString("0.##")))}]");

                    float? pivotX = ComputePivotX(freeTextX, greenRectX, XI_FREE_WINDOW, 3f, cleanPosX);
                    if (!pivotX.HasValue)
                    {
                        _logger.Log($"[XI-SKIP] pg={page} ComputePivotX returned null -> skip");
                        continue;
                    }

                    _logger.Log($"[XI] pg={page} pivotX={pivotX.Value:0.##} leftmostText={leftmostTextIdx.Count} leftmostField={leftmostFieldIdx.Count} freeText={freeTextX.Count} green={greenRectX.Count} alignTol={XI_ALIGN_TOL}pt");

                    // Clear XDiff from leftmost text diffs that are strictly LEFT of the pivot.
                    // XI applies only when bx < pivot - tolerance (rect IS the leftmost element).
                    // When bx == pivot the rect is AT the reference margin and V23 drift is real -> keep X.
                    foreach (int i in leftmostTextIdx)
                    {
                        float tx = TextDiffX(i);
                        if (tx < pivotX.Value - XI_ALIGN_TOL)
                        {
                            var mt = mergedTextDiffs[i];
                            mergedTextDiffs[i] = (mt.rect1, mt.rect2, mt.page1, mt.page2,
                                mt.textDescriptions, (mt.categories & ~DiffCategory.XDiff) | DiffCategory.XIgnore, mt.actual1, mt.actual2, mt.wordDiffs);
                            _logger.Log($"[XI-TEXT-CLEAR] pg={page} idx={i} rect=({mt.rect1.X:0.##},{mt.rect1.Y:0.##}) " +
                                $"bx={tx:0.##} pivotX={pivotX.Value:0.##} leftOf={(pivotX.Value - tx):0.##}pt -> XDiff cleared, XIgnore set (XI)");
                        }
                        else
                        {
                            var mt = mergedTextDiffs[i];
                            _logger.Log($"[XI-TEXT-KEEP] pg={page} idx={i} rect=({mt.rect1.X:0.##},{mt.rect1.Y:0.##}) " +
                                $"bx={tx:0.##} pivotX={pivotX.Value:0.##} notLeftOf (need bx < pivot-{XI_ALIGN_TOL}pt={pivotX.Value - XI_ALIGN_TOL:0.##}) -> keep X");
                        }
                    }

                    // Clear XDiff from leftmost field diffs that are strictly LEFT of the pivot.
                    foreach (int i in leftmostFieldIdx)
                    {
                        float fx = FieldDiffX(i);
                        if (fx < pivotX.Value - XI_ALIGN_TOL)
                        {
                            var mf = mergedFieldDiffs[i];
                            mergedFieldDiffs[i] = (mf.rect1, mf.rect2, mf.page1, mf.page2,
                                mf.mergedText, (mf.categories & ~DiffCategory.XDiff) | DiffCategory.XIgnore);
                            _logger.Log($"[XI-FIELD-CLEAR] pg={page} idx={i} field=({mf.rect1.X:0.##},{mf.rect1.Y:0.##}) " +
                                $"bx={fx:0.##} pivotX={pivotX.Value:0.##} leftOf={(pivotX.Value - fx):0.##}pt -> XDiff cleared, XIgnore set (XI)");
                        }
                        else
                        {
                            var mf = mergedFieldDiffs[i];
                            _logger.Log($"[XI-FIELD-KEEP] pg={page} idx={i} field=({mf.rect1.X:0.##},{mf.rect1.Y:0.##}) " +
                                $"bx={fx:0.##} pivotX={pivotX.Value:0.##} notLeftOf (need bx < pivot-{XI_ALIGN_TOL}pt={pivotX.Value - XI_ALIGN_TOL:0.##}) -> keep X");
                        }
                    }
                }
            }

            // ── Pass 2: Propagate XI to sibling X-diff rects in the same Word table ────
            // If a table contains an XI rect (already-correct column), the word corrector
            // will skip the whole table. Any other X-diff rects in that same table will
            // therefore also never be corrected → mark them XI too (renders green).
            //
            // Hybrid identification strategy:
            //   PRIMARY  — DocxTableIndex (Word document table index, precise):
            //     Phrase groups whose words have docxContext.TableIndex set carry a
            //     DocxTableIndex.  When available, use this to identify which Word table
            //     an XI group belongs to, and propagate only to phrase groups with the
            //     SAME DocxTableIndex.  This prevents cross-table contamination when
            //     DetectTableGroups spatially clusters groups from different Word tables
            //     into the same TableGroupId (e.g. Company/Insured tableIdx=1 and Firm
            //     Name tableIdx=2 sharing TableGroupId=1 on the same PDF page).
            //
            //   FALLBACK — TableGroupId (PDF spatial cluster):
            //     Merge-field tokens (e.g. «PolicyNumber», «CompanyName») often have no
            //     docxContext match, so DocxTableIndex is null.  For these groups fall back
            //     to TableGroupId: include any phrase group whose DocxTableIndex is null
            //     AND whose TableGroupId appears in the XI groups' TableGroupId set.
            //     This preserves the original behaviour for documents where docx matching
            //     is unavailable — Policy Number and Endorsement Date still inherit XI.
            if (phraseGroups != null && phraseGroups.Count > 0)
            {
                // XI Word tables by DocxTableIndex (primary, precise)
                var xiDocxTableIndices = new HashSet<int>(
                    phraseGroups
                        .Where(pg => pg.IsXIgnore && pg.DocxTableIndex.HasValue)
                        .Select(pg => pg.DocxTableIndex.Value));

                // XI table clusters by TableGroupId (fallback for groups without DocxTableIndex)
                var xiTableGroupIds = new HashSet<int>(
                    phraseGroups
                        .Where(pg => pg.IsXIgnore && pg.TableGroupId.HasValue)
                        .Select(pg => pg.TableGroupId.Value));

                // Tier 3: Y-row proximity to Pass 1 XI rects.
                // When phrase groups are reclassified as INLINE_RUN (table group removed by
                // AugmentTableGroups because docxContext couldn't confirm Word table membership),
                // they lose their TableGroupId and IsXIgnore flags.  Pass 1 of GenerateHtmlReport
                // still correctly marks the leftmost column's rects as XI (DiffCategory.XIgnore)
                // by operating directly on rects.  To propagate that XI to the other-column rects
                // in the same visual row (e.g. Policy Number at same Y as Company «XI»), collect
                // the Y extents of all Pass 1 XI rects.  Any phrase group on the same page within
                // XI_SIB_TOL pt in Y of a Pass 1 XI rect is treated as a sibling anchor — its
                // baseline region is added to xiTableRegions so RectInXiTable can cover it.
                var xiRectYRanges = mergedTextDiffs
                    .Where(mt => (mt.categories & DiffCategory.XIgnore) != 0)
                    .Select(mt => (page: mt.page1, top: mt.rect1.Top, bottom: mt.rect1.Bottom))
                    .Concat(mergedFieldDiffs
                        .Where(mf => (mf.categories & DiffCategory.XIgnore) != 0)
                        .Select(mf => (page: mf.page1, top: mf.rect1.Top, bottom: mf.rect1.Bottom)))
                    .ToList();

                const float XI_SIB_TOL = 8f; // pt overlap tolerance

                if (xiDocxTableIndices.Count > 0 || xiTableGroupIds.Count > 0 || xiRectYRanges.Count > 0)
                {
                    // All baseline regions belonging to XI tables — three-tier membership:
                    //   Tier 1: DocxTableIndex IS set → match by docx table index (cross-table-safe)
                    //   Tier 2: DocxTableIndex is null → fallback to TableGroupId spatial cluster
                    //   Tier 3: phrase group Y is within XI_SIB_TOL of any Pass 1 XI rect on same page
                    //           (catches INLINE_RUN groups that lost their table group identifier)
                    var xiTableRegions = phraseGroups
                        .Where(pg => pg.BaselineRegion != null &&
                                     (
                                         // Tier 1: same Word table by docx index
                                         (pg.DocxTableIndex.HasValue &&
                                          xiDocxTableIndices.Contains(pg.DocxTableIndex.Value))
                                         ||
                                         // Tier 2: no docx context — use spatial cluster
                                         (!pg.DocxTableIndex.HasValue &&
                                          pg.TableGroupId.HasValue &&
                                          xiTableGroupIds.Contains(pg.TableGroupId.Value))
                                         ||
                                         // Tier 3: same Y-row as a Pass 1 XI rect
                                         xiRectYRanges.Any(r =>
                                             r.page == pg.Page &&
                                             r.bottom + XI_SIB_TOL > pg.BaselineRegion.Y &&
                                             r.top    - XI_SIB_TOL < pg.BaselineRegion.Y)
                                     ))
                        .ToList();

                    bool RectInXiTable(int page, RectangleF rect) =>
                        xiTableRegions.Any(pg =>
                            pg.Page == page &&
                            pg.BaselineRegion.XStart < rect.Right  + XI_SIB_TOL &&
                            pg.BaselineRegion.XEnd   > rect.Left   - XI_SIB_TOL &&
                            pg.BaselineRegion.Y      < rect.Bottom + XI_SIB_TOL &&
                            pg.BaselineRegion.Y      > rect.Top    - XI_SIB_TOL);

                    // Propagate to text diffs
                    for (int i = 0; i < mergedTextDiffs.Count; i++)
                    {
                        var mt = mergedTextDiffs[i];
                        if ((mt.categories & DiffCategory.XDiff)   == 0) continue; // no X diff
                        if ((mt.categories & DiffCategory.XIgnore) != 0) continue; // already XI
                        if (!RectInXiTable(mt.page1, mt.rect1)) continue;
                        mergedTextDiffs[i] = (mt.rect1, mt.rect2, mt.page1, mt.page2,
                            mt.textDescriptions,
                            (mt.categories & ~DiffCategory.XDiff) | DiffCategory.XIgnore,
                            mt.actual1, mt.actual2, mt.wordDiffs);
                        _logger.Log($"[XI-SIBLING-TEXT] pg={mt.page1} rect=({mt.rect1.X:0.##},{mt.rect1.Y:0.##}) " +
                            $"-> same Word table as XI phrase group -> XDiff cleared, XIgnore set (XI)");
                    }

                    // Propagate to field diffs
                    for (int i = 0; i < mergedFieldDiffs.Count; i++)
                    {
                        var mf = mergedFieldDiffs[i];
                        if ((mf.categories & DiffCategory.XDiff)   == 0) continue;
                        if ((mf.categories & DiffCategory.XIgnore) != 0) continue;
                        if (!RectInXiTable(mf.page1, mf.rect1)) continue;
                        mergedFieldDiffs[i] = (mf.rect1, mf.rect2, mf.page1, mf.page2,
                            mf.mergedText,
                            (mf.categories & ~DiffCategory.XDiff) | DiffCategory.XIgnore);
                        _logger.Log($"[XI-SIBLING-FIELD] pg={mf.page1} rect=({mf.rect1.X:0.##},{mf.rect1.Y:0.##}) " +
                            $"-> same Word table as XI phrase group -> XDiff cleared, XIgnore set (XI)");
                    }
                }
            }

            // ── Build a spatial lookup of phrase groups for rect classification ─────────
            // A rect is "clean" only if it overlaps EXCLUSIVELY IsCleanDiff phrase groups.
            // A single non-clean overlapping group forces the rect to remain red.
            bool IsPhraseGroupClean(int page, RectangleF rect)
            {
                // ORIGINAL comparison: every visual diff between V14 and V23 original is a real finding.
                // No gate classification is applied — all rects are red by definition.
                // Only the MODIFIED comparison (IsModifiedComparison=true) uses gate logic to
                // classify what remains red after correction vs what is a rendering artifact.
                if (!IsModifiedComparison)
                    return false;

                const float OVERLAP_TOL = 8f;   // pt tolerance for spatial overlap
                const float POS_TOL = 12f;  // pt tolerance for direct position match

                // -- Primary path: phrase-group spatial overlap
                if (phraseGroups != null && phraseGroups.Count > 0)
                {
                    var overlapping = phraseGroups.Where(pg =>
                        pg.Page == page &&
                        pg.BaselineRegion != null &&
                        pg.BaselineRegion.XStart < rect.Right + OVERLAP_TOL &&
                        pg.BaselineRegion.XEnd > rect.Left - OVERLAP_TOL &&
                        pg.BaselineRegion.Y < rect.Top + rect.Height + OVERLAP_TOL &&
                        pg.BaselineRegion.Y > rect.Top - OVERLAP_TOL).ToList();

                    _logger.Log($"[ISPGCLEAN-PRI] pg={page} rect=({rect.X:0.##},{rect.Y:0.##} {rect.Width:0.##}x{rect.Height:0.##}) " +
                        $"phraseGroups.Count={phraseGroups.Count} overlapping={overlapping.Count}");

                    if (overlapping.Count > 0)
                    {
                        bool result = overlapping.All(pg => pg.IsCleanDiff);
                        _logger.Log($"  [PRI-RESULT] allClean={result} " +
                            string.Join(" ", overlapping.Select(g => $"pg{g.GroupId}={g.IsCleanDiff}")));
                        return result;
                    }
                    else
                    {
                        // Zero overlapping phrase groups means GenerateJsonDiffReport found NO word
                        // position diffs in this rect's area — nothing was corrected or attempted.
                        // The CompareText pixel-diff here is a rendering artifact (spacing, anti-aliasing)
                        // that the Word corrector cannot touch. Return clean (green).
                        var samePage = phraseGroups.Where(pg => pg.Page == page && pg.BaselineRegion != null)
                            .OrderBy(pg => Math.Abs(pg.BaselineRegion.Y - rect.Y) + Math.Abs(pg.BaselineRegion.XStart - rect.X))
                            .Take(3).ToList();
                        foreach (var sp in samePage)
                            _logger.Log($"  [PRI-NEAREST] pgId={sp.GroupId} bl.XStart={sp.BaselineRegion.XStart:0.##} bl.Y={sp.BaselineRegion.Y:0.##} IsClean={sp.IsCleanDiff}");

                        bool hasPageGroups = phraseGroups.Any(pg => pg.Page == page);
                        if (hasPageGroups)
                        {
                            _logger.Log($"  [PRI-NO-OVERLAP] pg={page} rect=({rect.X:0.##},{rect.Y:0.##}) -- 0 overlapping groups but page has diffs; area is clean (no word diffs here)");
                            return true;   // rect area had no word position diffs => clean
                        }
                        else
                        {
                            // phraseGroups exist on other pages but not on this page.
                            // This diff is either content that Y-drifted from a previous page (page overflow)
                            // or a pre-existing V14/V23 rendering difference on a page the corrector never touched.
                            // Either way the corrector had no jurisdiction here → classify as clean.
                            _logger.Log($"  [PRI-OTHER-PAGE] pg={page} rect=({rect.X:0.##},{rect.Y:0.##}) -- phraseGroups exist on other pages but none on pg={page}; pre-existing V14/V23 diff => clean");
                            return true;
                        }
                    }
                }

                // -- Fallback path: used only when this page has no phrase groups at all
                // (i.e. file produced zero word diffs). Direct position match for Gate-1 reflow words.
                if (cleanWordPositions != null && cleanWordPositions.Count > 0)
                {
                    bool found = cleanWordPositions.Any(p =>
                        p.page == page &&
                        Math.Abs(p.x - rect.X) < POS_TOL &&
                        Math.Abs(p.y - rect.Y) < POS_TOL);

                    _logger.Log($"[ISPGCLEAN-FALL] pg={page} rect.X={rect.X:0.##} rect.Y={rect.Y:0.##} " +
                        $"cleanWordPositions.Count={cleanWordPositions.Count} matchFound={found}");

                    if (!found)
                    {
                        var nearest = cleanWordPositions.Where(p => p.page == page)
                            .OrderBy(p => Math.Abs(p.x - rect.X) + Math.Abs(p.y - rect.Y))
                            .Take(3);
                        foreach (var n in nearest)
                            _logger.Log($"  [FALL-NEAREST] X={n.x:0.##} Y={n.y:0.##} dX={Math.Abs(n.x - rect.X):0.##} dY={Math.Abs(n.y - rect.Y):0.##}");
                    }
                    return found;
                }

                _logger.Log($"[ISPGCLEAN-MISS] pg={page} rect=({rect.X:0.##},{rect.Y:0.##}) -- no phraseGroups and no cleanWordPositions; GJDR found 0 word diffs => green");
                return true;   // GJDR found 0 word diffs for this document — Word corrector applied nothing
            }

            // ── Collect renderable items ──────────────────────────────────────────────
            var renderItems = new List<(string img1, string img2, int diffIndex)>();

            if (_generateOnePageDifference)
            {
                // Page buckets now have separate red and green rect lists
                var pageBuckets = new Dictionary<(int, int), (List<(RectangleF r, DiffCategory c)> red1, List<(RectangleF r, DiffCategory c)> green1,
                                                               List<(RectangleF r, DiffCategory c)> red2, List<(RectangleF r, DiffCategory c)> green2,
                                                               List<string> descriptions)>();

                foreach (var item in mergedFieldDiffs)
                {
                    var ir1 = ignoreRegionsPdf1.Where(r => r.Page == item.page1).Select(r => new RectangleF(r.XStart, r.Y, r.XEnd - r.XStart, r.Height)).ToList();
                    var ir2 = ignoreRegionsPdf2.Where(r => r.Page == item.page2).Select(r => new RectangleF(r.XStart, r.Y, r.XEnd - r.XStart, r.Height)).ToList();
                    if (IsSignificantlyOverlapping(item.rect1, ir1) || IsSignificantlyOverlapping(item.rect2, ir2)) continue;
                    var key = (item.page1, item.page2);
                    if (!pageBuckets.TryGetValue(key, out var val))
                        val = (new List<(RectangleF, DiffCategory)>(), new List<(RectangleF, DiffCategory)>(), new List<(RectangleF, DiffCategory)>(), new List<(RectangleF, DiffCategory)>(), new List<string>());
                    // In modified comparison, field diffs covered by clean phrase groups are also clean
                    // (same uncorrected positional drift that produced the word-level clean diffs).
                    // CategoryBasedIsGreen is also checked so XI-cleared XDiff (no longer XDiff in
                    // categories) is respected the same way text diffs are.
                    bool isFieldClean = IsModifiedComparison && IsPhraseGroupClean(item.page1, item.rect1);
                    bool effectiveFieldClean = CategoryBasedIsGreen(item.categories, isFieldClean);
                    _logger.Log($"[BUCKET-FIELD] pg={item.page1} rect=({item.rect1.X:0.##},{item.rect1.Y:0.##}) cats={item.categories} isFieldClean={isFieldClean} effectiveFieldClean={effectiveFieldClean} IsModified={IsModifiedComparison} text={item.mergedText}");
                    if (effectiveFieldClean) { val.green1.Add((item.rect1, item.categories)); val.green2.Add((item.rect2, item.categories)); }
                    else { val.red1.Add((item.rect1, item.categories)); val.red2.Add((item.rect2, item.categories)); }
                    val.descriptions.Add(item.mergedText);
                    pageBuckets[key] = val;
                }

                // Pass 1: classify every mergedTextDiff rect independently and stage results
                var stagedTextDiffs = new List<(int page1, int page2, RectangleF rect1, RectangleF rect2, bool isClean, List<string> descriptions, DiffCategory categories)>();
                foreach (var item in mergedTextDiffs)
                {
                    if (item.textDescriptions == null || item.textDescriptions.Count == 0) continue;
                    var ir1 = ignoreRegionsPdf1.Where(r => r.Page == item.page1).Select(r => new RectangleF(r.XStart, r.Y, r.XEnd - r.XStart, r.Height)).ToList();
                    var ir2 = ignoreRegionsPdf2.Where(r => r.Page == item.page2).Select(r => new RectangleF(r.XStart, r.Y, r.XEnd - r.XStart, r.Height)).ToList();
                    if (IsSignificantlyOverlapping(item.rect1, ir1) || IsSignificantlyOverlapping(item.rect2, ir2)) continue;
                    bool isClean = IsPhraseGroupClean(item.page1, item.rect1);
                    stagedTextDiffs.Add((item.page1, item.page2, item.rect1, item.rect2, isClean, item.textDescriptions, item.categories));
                }

                // Pass 2: green-inheritance — a red rect fully contained inside a green rect inherits green.
                // Rationale: if the containing rect is clean (no correction applied to the whole area),
                // any sub-rect within that area must also be clean by definition.
                const float CONTAIN_TOL = 2f;   // pt slop for containment check
                var greenRectsByPage = stagedTextDiffs
                    .Where(s => s.isClean)
                    .GroupBy(s => s.page1)
                    .ToDictionary(g => g.Key, g => g.Select(s => s.rect1).ToList());

                for (int si = 0; si < stagedTextDiffs.Count; si++)
                {
                    var s = stagedTextDiffs[si];
                    if (s.isClean) continue;   // already green, skip
                    if (!greenRectsByPage.TryGetValue(s.page1, out var greenRects)) continue;
                    bool contained = greenRects.Any(g =>
                        g.X - CONTAIN_TOL <= s.rect1.X &&
                        g.Y - CONTAIN_TOL <= s.rect1.Y &&
                        g.Right + CONTAIN_TOL >= s.rect1.Right &&
                        g.Bottom + CONTAIN_TOL >= s.rect1.Bottom);
                    if (contained)
                    {
                        _logger.Log($"[GREEN-INHERIT] pg={s.page1} rect ({s.rect1.X:0.##},{s.rect1.Y:0.##} {s.rect1.Width:0.##}x{s.rect1.Height:0.##}) upgraded to green — contained inside green parent");
                        stagedTextDiffs[si] = (s.page1, s.page2, s.rect1, s.rect2, true, s.descriptions, s.categories);
                    }
                }

                foreach (var item in stagedTextDiffs)
                {
                    var key = (item.page1, item.page2);
                    if (!pageBuckets.TryGetValue(key, out var val))
                        val = (new List<(RectangleF, DiffCategory)>(), new List<(RectangleF, DiffCategory)>(), new List<(RectangleF, DiffCategory)>(), new List<(RectangleF, DiffCategory)>(), new List<string>());
                    bool effectivelyClean = CategoryBasedIsGreen(item.categories, item.isClean);
                    if (effectivelyClean) { val.green1.Add((item.rect1, item.categories)); val.green2.Add((item.rect2, item.categories)); }
                    else { val.red1.Add((item.rect1, item.categories)); val.red2.Add((item.rect2, item.categories)); }
                    foreach (var d in item.descriptions) val.descriptions.Add(d);
                    pageBuckets[key] = val;
                }

                // Capture classified rectangles for caller (pdf1/baseline coords; no logic change)
                diffRectangles = new List<DiffRectangleInfo>();
                foreach (var kv in pageBuckets)
                {
                    int captPg = kv.Key.Item1;
                    if (kv.Value.red1.Count + kv.Value.green1.Count == 0) continue;
                    foreach (var (r, c) in kv.Value.red1)
                        diffRectangles.Add(new DiffRectangleInfo { Page = captPg, X = r.X, Y = r.Y, Width = r.Width, Height = r.Height, IsGreen = false, Categories = c });
                    foreach (var (g, c) in kv.Value.green1)
                        diffRectangles.Add(new DiffRectangleInfo { Page = captPg, X = g.X, Y = g.Y, Width = g.Width, Height = g.Height, IsGreen = true, Categories = c });
                }

                int di = 1;
                int totalRed = 0, totalGreen = 0;
                foreach (var kv in pageBuckets)
                {
                    var pg1 = kv.Key.Item1; var pg2 = kv.Key.Item2;
                    var red1 = kv.Value.red1; var red2 = kv.Value.red2;
                    var grn1 = kv.Value.green1; var grn2 = kv.Value.green2;
                    if ((red1.Count + grn1.Count == 0) && (red2.Count + grn2.Count == 0)) continue;
                    string i1 = CapturePageWithRectsToBase64(pdf1, pg1, red1, grn1, ignoreRegionsPdf1.Where(r => r.Page == pg1).ToList());
                    string i2 = CapturePageWithRectsToBase64(pdf2, pg2, red2, grn2, ignoreRegionsPdf2.Where(r => r.Page == pg2).ToList());
                    renderItems.Add((i1, i2, di++));
                    blnIsThereDifference = true;
                    totalRed += red1.Count + red2.Count;
                    totalGreen += grn1.Count + grn2.Count;
                }
                allDiffsAreClean = blnIsThereDifference && totalRed == 0 && totalGreen > 0;
            }
            else
            {
                int di = 1;
                int totalRed = 0, totalGreen = 0;
                if (mergedFieldDiffs != null)
                {
                    foreach (var item in mergedFieldDiffs)
                    {
                        var ir1 = ignoreRegionsPdf1.Where(r => r.Page == item.page1).Select(r => new RectangleF(r.XStart, r.Y, r.XEnd - r.XStart, r.Height)).ToList();
                        var ir2 = ignoreRegionsPdf2.Where(r => r.Page == item.page2).Select(r => new RectangleF(r.XStart, r.Y, r.XEnd - r.XStart, r.Height)).ToList();
                        if (IsSignificantlyOverlapping(item.rect1, ir1) || IsSignificantlyOverlapping(item.rect2, ir2)) continue;
                        // In modified comparison, field diffs covered by clean phrase groups are also clean.
                        bool isFieldClean = IsModifiedComparison && IsPhraseGroupClean(item.page1, item.rect1);
                        bool effectiveFieldClean = CategoryBasedIsGreen(item.categories, isFieldClean);
                        diffRectangles.Add(new DiffRectangleInfo { Page = item.page1, X = item.rect1.X, Y = item.rect1.Y, Width = item.rect1.Width, Height = item.rect1.Height, IsGreen = effectiveFieldClean, Categories = item.categories });
                        if (effectiveFieldClean)
                        {
                            string i1 = CaptureToBase64(pdf1, item.page1, item.rect1.X, item.rect1.Y, item.rect1.Width, item.rect1.Height,
                                ignoreRegionsPdf1.Where(r => r.Page == item.page1).ToList(),
                                new List<RectangleF> { item.rect1 }, item.categories);
                            string i2 = CaptureToBase64(pdf2, item.page2, 0, 0, 0, 0,
                                ignoreRegionsPdf2.Where(r => r.Page == item.page2).ToList(),
                                new List<RectangleF> { item.rect2 }, item.categories);
                            renderItems.Add((i1, i2, di++));
                            blnIsThereDifference = true;
                            totalGreen++;
                        }
                        else
                        {
                            string i1 = CaptureToBase64(pdf1, item.page1, item.rect1.X, item.rect1.Y, item.rect1.Width, item.rect1.Height, ignoreRegionsPdf1.Where(r => r.Page == item.page1).ToList(), diffCats: item.categories);
                            string i2 = CaptureToBase64(pdf2, item.page2, item.rect2.X, item.rect2.Y, item.rect2.Width, item.rect2.Height, ignoreRegionsPdf2.Where(r => r.Page == item.page2).ToList(), diffCats: item.categories);
                            renderItems.Add((i1, i2, di++));
                            blnIsThereDifference = true;
                            totalRed++;
                        }
                    }
                }
                if (mergedTextDiffs != null)
                {
                    foreach (var item in mergedTextDiffs)
                    {
                        if (item.textDescriptions.Count == 0) continue;
                        var ir1 = ignoreRegionsPdf1.Where(r => r.Page == item.page1).Select(r => new RectangleF(r.XStart, r.Y, r.XEnd - r.XStart, r.Height)).ToList();
                        var ir2 = ignoreRegionsPdf2.Where(r => r.Page == item.page2).Select(r => new RectangleF(r.XStart, r.Y, r.XEnd - r.XStart, r.Height)).ToList();
                        if (IsSignificantlyOverlapping(item.rect1, ir1) || IsSignificantlyOverlapping(item.rect2, ir2)) continue;
                        bool isClean = IsPhraseGroupClean(item.page1, item.rect1);
                        bool effectiveClean = CategoryBasedIsGreen(item.categories, isClean);
                        diffRectangles.Add(new DiffRectangleInfo { Page = item.page1, X = item.rect1.X, Y = item.rect1.Y, Width = item.rect1.Width, Height = item.rect1.Height, IsGreen = effectiveClean, Categories = item.categories });
                        if (effectiveClean)
                        {
                            // Render with only green rect (no red)
                            string i1 = CaptureToBase64(pdf1, item.page1, item.rect1.X, item.rect1.Y, item.rect1.Width, item.rect1.Height,
                                ignoreRegionsPdf1.Where(r => r.Page == item.page1).ToList(),
                                new List<RectangleF> { item.rect1 }, item.categories);
                            // For clean diffs we pass the rect as cleanRect on pdf2 but NOT as the main red rect
                            string i2 = CaptureToBase64(pdf2, item.page2, 0, 0, 0, 0,
                                ignoreRegionsPdf2.Where(r => r.Page == item.page2).ToList(),
                                new List<RectangleF> { item.rect2 }, item.categories);
                            renderItems.Add((i1, i2, di++));
                            blnIsThereDifference = true;
                            totalGreen++;
                        }
                        else
                        {
                            string i1 = CaptureToBase64(pdf1, item.page1, item.rect1.X, item.rect1.Y, item.rect1.Width, item.rect1.Height, ignoreRegionsPdf1.Where(r => r.Page == item.page1).ToList(), diffCats: item.categories);
                            string i2 = CaptureToBase64(pdf2, item.page2, item.rect2.X, item.rect2.Y, item.rect2.Width, item.rect2.Height, ignoreRegionsPdf2.Where(r => r.Page == item.page2).ToList(), diffCats: item.categories);
                            renderItems.Add((i1, i2, di++));
                            blnIsThereDifference = true;
                            totalRed++;
                        }
                    }
                }
                allDiffsAreClean = blnIsThereDifference && totalRed == 0 && totalGreen > 0;
            }

            if (!blnIsThereDifference)
            {
                Console.WriteLine($"No actual visual/textual differences found (all diffs in ignore regions). Generating clean report for: {reportPath}");
                string cleanHtml2 = $@"<!DOCTYPE html>
<html lang=""en"">
<head>
<meta charset=""UTF-8"" />
<title>Clean – {System.Net.WebUtility.HtmlEncode(pdf2FileName)}</title>
<style>
  body {{ font-family: sans-serif; background: #f5f4f0; margin: 0; padding: 40px; }}
  .topbar {{ background: #fff; border-bottom: 1px solid #e8e6e0; padding: 12px 24px; display: flex; justify-content: space-between; align-items: center; margin: -40px -40px 40px; }}
  .back-link {{ font-size: 13px; color: #4a6fa5; text-decoration: none; }}
  .topbar-title {{ font-size: 14px; font-weight: 600; color: #2d2d2d; }}
  .topbar-meta {{ font-size: 11.5px; color: #999; }}
  .clean-card {{ background: #fff; border: 1px solid #e8e6e0; border-radius: 8px; padding: 48px; text-align: center; max-width: 600px; margin: 80px auto 0; }}
  .clean-icon {{ font-size: 48px; margin-bottom: 16px; color: #1aaf5d; }}
  h1 {{ font-size: 22px; font-weight: 600; color: #1aaf5d; margin-bottom: 8px; }}
  p {{ color: #666; font-size: 14px; }}
</style>
</head>
<body>
<div class=""topbar"">
  <a class=""back-link"" href=""{System.Net.WebUtility.HtmlEncode(summaryBackLink)}"">← Summary</a>
  <div class=""topbar-title"">Diff Report — {System.Net.WebUtility.HtmlEncode(pdf2FileName)}</div>
  <div class=""topbar-meta"">Generated {generated}</div>
</div>
<div class=""clean-card"">
  <div class=""clean-icon"">✓</div>
  <h1>No Differences Found</h1>
  <p>{System.Net.WebUtility.HtmlEncode(pdf2FileName)}</p>
  <p style=""margin-top:8px;font-size:12px;color:#999;"">V14 and V23 PDFs are identical (or all detected differences are in ignored regions).</p>
</div>
</body>
</html>";
                File.WriteAllText(reportPath, cleanHtml2);
                return false;
            }


            if (renderItems.Count > 1)
            {
                html.AppendLine(@"  <div class=""diff-nav""><span>Jump to</span>");
                foreach (var (_, _, dIdx) in renderItems)
                    html.AppendLine($@"    <a class=""diff-anchor"" href=""#diff-{dIdx}"">Diff {dIdx}</a>");
                html.AppendLine(@"  </div>");
            }

            // ── Diff cards ────────────────────────────────────────────────────────────
            foreach (var (img1, img2, dIdx) in renderItems)
            {
                html.AppendLine($@"
  <div class=""diff-card"" id=""diff-{dIdx}"">
    <div class=""panels"">
      <div class=""panel"">
        <div class=""panel-header"">
          <div class=""panel-label baseline"">BASELINE · PDF 1</div>
          <div class=""panel-filename"">{System.Net.WebUtility.HtmlEncode(pdf1Name)}</div>
          <div class=""panel-folder"">{System.Net.WebUtility.HtmlEncode(pdf1Folder)}</div>
        </div>
        <div class=""panel-img""><img src=""{img1}"" alt=""PDF 1 diff region"" /></div>
      </div>
      <div class=""panel"">
        <div class=""panel-header"">
          <div class=""panel-label publish"">PUBLISH · PDF 2</div>
          <div class=""panel-filename"">{System.Net.WebUtility.HtmlEncode(pdf2Name)}</div>
          <div class=""panel-folder"">{System.Net.WebUtility.HtmlEncode(pdf2Folder)}</div>
        </div>
        <div class=""panel-img""><img src=""{img2}"" alt=""PDF 2 diff region"" /></div>
      </div>
    </div>
  </div>");
            }

            // ── Footer ────────────────────────────────────────────────────────────────
            html.AppendLine($@"
  <div class=""footer"">
    PDFCompare &nbsp;·&nbsp; {System.Net.WebUtility.HtmlEncode(pdf2FileName)} &nbsp;·&nbsp; Generated {generated}
    <br/><a class=""back-to-top"" href=""#"">↑ Back to top</a>
  </div>
</div>
</body>
</html>");

            File.WriteAllText(reportPath, html.ToString());
            Console.WriteLine($"Report generated: {reportPath}");
            return true;
        }

        // ─── Helper: render a PDF page to Bitmap via Docnet  ───
        private Bitmap RenderPdfPageToBitmap(string pdfPath, int pageIndex0Based, int renderWidth, int renderHeight)
        {
            // DocLib.Instance is a singleton — use GetDocReader with using, but do NOT dispose the singleton itself.
            using var docReader = DocLib.Instance.GetDocReader(
                pdfPath,
                new PageDimensions(renderWidth, renderHeight));

            using var pageReader = docReader.GetPageReader(pageIndex0Based);

            byte[] rawBgra = pageReader.GetImage(); // BGRA byte array
            int w = pageReader.GetPageWidth();
            int h = pageReader.GetPageHeight();

            var bmp = new Bitmap(w, h, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
            var bmpData = bmp.LockBits(
                new Rectangle(0, 0, w, h),
                System.Drawing.Imaging.ImageLockMode.WriteOnly,
                System.Drawing.Imaging.PixelFormat.Format32bppArgb);

            System.Runtime.InteropServices.Marshal.Copy(rawBgra, 0, bmpData.Scan0, rawBgra.Length);
            bmp.UnlockBits(bmpData);

            return bmp;
        }

        // ─── CHANGED: Safely get page dimensions, returning a boolean if the page exists ───
        private (SizeF Size, bool Exists) GetPdfPageSizePtsSafe(string pdfPath, int pageNumber1Based)
        {
            using var pdfDoc = new iText.Kernel.Pdf.PdfDocument(new PdfReader(pdfPath));
            int totalPages = pdfDoc.GetNumberOfPages();

            if (pageNumber1Based < 1 || pageNumber1Based > totalPages)
            {
                // Fallback size: use page 1 size if the document has any pages, else generic A4 size
                var fallbackSize = totalPages > 0
                    ? pdfDoc.GetPage(1).GetPageSize()
                    : new iText.Kernel.Geom.Rectangle(0, 0, 612, 792);
                return (new SizeF(fallbackSize.GetWidth(), fallbackSize.GetHeight()), false);
            }

            var ps = pdfDoc.GetPage(pageNumber1Based).GetPageSize();
            return (new SizeF(ps.GetWidth(), ps.GetHeight()), true);
        }

        // ─── ADDED: Generates a placeholder image for missing pages to prevent crashes ───
        private string GenerateMissingPageImageToBase64(int width, int height, int pageNumber)
        {
            int safeWidth = width > 0 ? width : 1275;
            int safeHeight = height > 0 ? height : 1650;

            using var image = new Bitmap(safeWidth, safeHeight, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
            using var g = Graphics.FromImage(image);
            g.Clear(System.Drawing.Color.WhiteSmoke);

            using var pen = new Pen(System.Drawing.Color.LightGray, 3);
            pen.DashStyle = System.Drawing.Drawing2D.DashStyle.Dash;
            g.DrawRectangle(pen, 20, 20, safeWidth - 40, safeHeight - 40);

            string text = $"Page {pageNumber} is not present in this PDF.";
            using var font = new Font("Arial", 20, FontStyle.Bold);
            using var brush = new SolidBrush(System.Drawing.Color.DarkGray);

            var textSize = g.MeasureString(text, font);
            float x = (safeWidth - textSize.Width) / 2f;
            float y = (safeHeight - textSize.Height) / 2f;

            g.DrawString(text, font, brush, x, y);

            using var ms = new MemoryStream();
            image.Save(ms, ImageFormat.Png);
            return $"data:image/png;base64,{Convert.ToBase64String(ms.ToArray())}";
        }

        // ─── CHANGED: Drop-in replacement for CaptureToBase64 ───
        private string CaptureToBase64(string pdfPath, int page, float x, float y, float width, float height,
            List<IgnoreRegion> ignoreRegions = null,
            List<RectangleF> cleanRects = null,
            DiffCategory diffCats = DiffCategory.None)
        {
            int dpi = 150;
            // Utilize the safe extraction method
            var (pageSize, pageExists) = GetPdfPageSizePtsSafe(pdfPath, page);
            int renderWidth = (int)(pageSize.Width * dpi / 72);
            int renderHeight = (int)(pageSize.Height * dpi / 72);

            // Gracefully handle if the page doesn't exist by returning a placeholder
            if (!pageExists)
            {
                return GenerateMissingPageImageToBase64(renderWidth, renderHeight, page);
            }

            using var image = RenderPdfPageToBitmap(pdfPath, page - 1, renderWidth, renderHeight);
            using var g = Graphics.FromImage(image);

            // Get actual rendered dimensions (Docnet may clip to fit aspect ratio)
            float actualW = image.Width;
            float actualH = image.Height;

            // Coordinate transform: PDF origin is bottom-left; image origin is top-left
            float scaleX = actualW / pageSize.Width;
            float scaleY = actualH / pageSize.Height;

            // Draw clean (green) rects first — red wins on overlap
            if (cleanRects != null)
            {
                using var cleanPen = new Pen(System.Drawing.Color.LimeGreen, 3);
                foreach (var cr in cleanRects)
                {
                    float crX = cr.X * scaleX;
                    float crY = (pageSize.Height - cr.Y - cr.Height) * scaleY;
                    float crW = cr.Width * scaleX;
                    float crH = cr.Height * scaleY;
                    g.DrawRectangle(cleanPen, crX, crY, crW, crH);
                    DrawDiffCategoryLabel(g, crX, crY, crW, crH, diffCats);
                }
            }

            float rectX = x * scaleX;
            float rectY = (pageSize.Height - y - height) * scaleY;
            float rectWidth = width * scaleX;
            float rectHeight = height * scaleY;

            using var pen = new Pen(System.Drawing.Color.Red, 3);
            g.DrawRectangle(pen, rectX, rectY, rectWidth, rectHeight);
            if (width > 0 && height > 0)
                DrawDiffCategoryLabel(g, rectX, rectY, rectWidth, rectHeight, diffCats);

            if (_drawIgnoredRegion && ignoreRegions != null)
            {
                using var greenPen = new Pen(System.Drawing.Color.LimeGreen, 2);
                foreach (var region in ignoreRegions)
                {
                    float igX = region.XStart * scaleX;
                    float igY = (pageSize.Height - region.Y - region.Height) * scaleY;
                    float igWidth = (region.XEnd - region.XStart) * scaleX;
                    float igHeight = region.Height * scaleY;
                    g.DrawRectangle(greenPen, igX, igY, igWidth, igHeight);
                }
            }

            using var ms = new MemoryStream();
            image.Save(ms, ImageFormat.Png);
            return $"data:image/png;base64,{Convert.ToBase64String(ms.ToArray())}";
        }

        // ─── CHANGED: Drop-in replacement for CapturePageWithRectsToBase64 ───
        private string CapturePageWithRectsToBase64(string pdfPath, int page,
            List<(RectangleF r, DiffCategory c)> rects,
            List<(RectangleF r, DiffCategory c)> cleanRects = null,
            List<IgnoreRegion> ignoreRegions = null)
        {
            int dpi = 150;
            // Utilize the safe extraction method
            var (pageSize, pageExists) = GetPdfPageSizePtsSafe(pdfPath, page);
            int renderWidth = (int)(pageSize.Width * dpi / 72);
            int renderHeight = (int)(pageSize.Height * dpi / 72);

            // Gracefully handle if the page doesn't exist by returning a placeholder
            if (!pageExists)
            {
                return GenerateMissingPageImageToBase64(renderWidth, renderHeight, page);
            }

            using var image = RenderPdfPageToBitmap(pdfPath, page - 1, renderWidth, renderHeight);
            using var g = Graphics.FromImage(image);

            float actualW = image.Width;
            float actualH = image.Height;
            float scaleX = actualW / pageSize.Width;
            float scaleY = actualH / pageSize.Height;

            // Draw clean (intentionally-skipped) diffs in green FIRST.
            // Red rects are drawn on top so they always win on any overlap.
            if (cleanRects != null && cleanRects.Count > 0)
            {
                using var cleanPen = new Pen(System.Drawing.Color.LimeGreen, 3);
                foreach (var (cr, cats) in cleanRects)
                {
                    float crX = cr.X * scaleX;
                    float crY = (pageSize.Height - cr.Y - cr.Height) * scaleY;
                    float crW = cr.Width * scaleX;
                    float crH = cr.Height * scaleY;
                    g.DrawRectangle(cleanPen, crX, crY, crW, crH);
                    DrawDiffCategoryLabel(g, crX, crY, crW, crH, cats);
                }
            }

            if (rects != null && rects.Count > 0)
            {
                using var pen = new Pen(System.Drawing.Color.Red, 3);
                foreach (var (r, cats) in rects)
                {
                    float rectX = r.X * scaleX;
                    float rectY = (pageSize.Height - r.Y - r.Height) * scaleY;
                    float rectWidth = r.Width * scaleX;
                    float rectHeight = r.Height * scaleY;
                    g.DrawRectangle(pen, rectX, rectY, rectWidth, rectHeight);
                    DrawDiffCategoryLabel(g, rectX, rectY, rectWidth, rectHeight, cats);
                }
            }

            if (_drawIgnoredRegion && ignoreRegions != null)
            {
                using var ignorePen = new Pen(System.Drawing.Color.LimeGreen, 2);
                foreach (var region in ignoreRegions)
                {
                    float igX = region.XStart * scaleX;
                    float igY = (pageSize.Height - region.Y - region.Height) * scaleY;
                    float igWidth = (region.XEnd - region.XStart) * scaleX;
                    float igHeight = region.Height * scaleY;
                    g.DrawRectangle(ignorePen, igX, igY, igWidth, igHeight);
                }
            }

            using var ms = new MemoryStream();
            image.Save(ms, ImageFormat.Png);
            return $"data:image/png;base64,{Convert.ToBase64String(ms.ToArray())}";
        }

        /// <summary>
        /// Draws a small category label (e.g. "X,Y" or "W") in the top-left corner of a
        /// rectangle on the given Graphics context. No-op when cats == None.
        /// </summary>
        private void DrawDiffCategoryLabel(Graphics g, float rectX, float rectY, float rectW, float rectH, DiffCategory cats)
        {
            if (cats == DiffCategory.None || !ShowXYWInDiffReport) return;
            var parts = new System.Collections.Generic.List<string>();
            if ((cats & DiffCategory.XIgnore) != 0) parts.Add("XI");
            else if ((cats & DiffCategory.XDiff) != 0) parts.Add("X");
            if ((cats & DiffCategory.YDiff) != 0) parts.Add("Y");
            if ((cats & DiffCategory.WDiff) != 0) parts.Add("W");
            string label = string.Join(",", parts);
            using var font = new Font("Arial", 8f, FontStyle.Bold, GraphicsUnit.Point);
            var size = g.MeasureString(label, font);
            float lx = rectX;
            float ly = rectY - size.Height - 2;
            // Semi-transparent black background for readability
            using var bgBrush = new SolidBrush(System.Drawing.Color.FromArgb(180, 0, 0, 0));
            g.FillRectangle(bgBrush, lx - 1, ly - 1, size.Width + 2, size.Height + 2);
            // White label text
            using var textBrush = new SolidBrush(System.Drawing.Color.White);
            g.DrawString(label, font, textBrush, lx, ly);
        }

        private float EstimateWidth(string text, float fontSize = 12)
        {
            if (string.IsNullOrEmpty(text)) return 50;
            // Approximate width: assume average character width of 0.5 * fontSize
            return text.Length * fontSize * 0.5f;
        }

        private float EstimateHeight(float fontSize = 12)
        {
            // Typical height is ~1.2x the font size
            return fontSize * 1.2f;
        }

        private bool IsSignificantlyOverlapping(RectangleF rect, List<RectangleF> ignoreRegions)
        {
            if (ignoreRegions == null || ignoreRegions.Count == 0)
                return false;

            foreach (var ignore in ignoreRegions)
            {
                if (ignore.IntersectsWith(rect) || rect.IntersectsWith(ignore) ||
                    ignore.Contains(rect) || rect.Contains(ignore))
                {
                    return true;
                }
            }

            return false;
        }

        public static void PrintAllWordsFromPdf(string pdfPath)
        {
            string outputFilePath = System.IO.Path.ChangeExtension(pdfPath, ".txt");

            var pdfDoc = new iText.Kernel.Pdf.PdfDocument(new PdfReader(pdfPath));
            int totalPages = pdfDoc.GetNumberOfPages();

            using (StreamWriter writer = new StreamWriter(outputFilePath))
            {
                for (int pageNum = 1; pageNum <= totalPages; pageNum++)
                {
                    writer.WriteLine($"\n--- Page {pageNum} ---");
                    List<iText.Kernel.Geom.Rectangle> lstWhiteMasked = PdfComparer.GetAnnotations(pdfDoc, pageNum, "WhiteMasked");
                    var strategy = new WordLevelExtractionStrategy(true, lstWhiteMasked);
                    var processor = new PdfCanvasProcessor(strategy);

                    processor.ProcessPageContent(pdfDoc.GetPage(pageNum));

                    foreach (var word in strategy.Words)
                    {
                        var rect = word.BoundingBox;
                        writer.WriteLine($"Text: \"{word.Text}\"");
                        writer.WriteLine($"  Position: X={rect.GetX():0.##}, Y={rect.GetY():0.##}");
                        writer.WriteLine($"  Width: {rect.GetWidth():0.##}, Height: {rect.GetHeight():0.##}");
                        writer.WriteLine($"  Font: {word.FontName}, Size: {word.FontSize:0.##}");
                        writer.WriteLine();
                    }
                }
            }
        }

        #region "Methods specific to calculating JSONDiffReport"
        // ─── Helpers ─────────────────────────────────────────────────────────────────────
        /// <summary>
        /// Computes correction targets for a word diff entry.
        ///
        /// FIXED vs previous version:
        ///    - Checks CellPadding strategy first (cellPaddingPt ≈ driftX) before TableIndent
        ///    - Uses actual colWidthPt from docxContext as currentValuePt (was always 0)
        ///    - For uniform drift: always suggests CellPadding or TableIndent, never column width
        ///    - Column width suggestion only appears when drifts are genuinely non-uniform
        ///    - Correction description is now accurate and non-contradictory
        /// </summary>
        private List<DocxCorrectionTarget> BuildCorrectionTargets(
            WordDiffEntry entry,
            DocxElementInfo elem)
        {
            var targets = new List<DocxCorrectionTarget>();
            if (entry.Delta == null) return targets;

            float dx = entry.Delta.X;   // positive = publish word is to the RIGHT
            float dy = entry.Delta.Y;   // positive = publish word is ABOVE baseline (PDF Y-up)

            const float STRATEGY_TOL = 0.5f;

            // ── X correction ─────────────────────────────────────────────────────────
            if (Math.Abs(dx) > 0.5f)
            {
                DocxCorrectionTarget xTarget = null;

                if (elem.ElementType == "TABLE_CELL")
                {
                    float cellPad = elem.CellLeftPaddingPt ?? 0f;
                    float tableInd = elem.TableLeftIndentPt ?? 0f;

                    // PRIORITY 1: Cell padding equals drift → padding was increased, reduce it back
                    if (Math.Abs(cellPad - Math.Abs(dx)) <= STRATEGY_TOL && cellPad > 0)
                    {
                        float newPad = Math.Max(0f, cellPad - Math.Abs(dx));
                        xTarget = new DocxCorrectionTarget
                        {
                            Axis = "X",
                            Property = "CellFormat.LeftPadding",
                            CurrentValuePt = cellPad,
                            NewValuePt = newPad,
                            CurrentValueTwips = (int)Math.Round(cellPad * 20),
                            NewValueTwips = (int)Math.Round(newPad * 20),
                            Description = $"[CELL_PADDING] Reduce LeftPadding on ALL cells in table from " +
                                               $"{cellPad:0.##}pt → {newPad:0.##}pt to cancel +{dx:0.##}pt drift. " +
                                               $"Apply once to all rows/cells — NOT per word."
                        };
                    }
                    // PRIORITY 2: Table left indent equals drift
                    else if (Math.Abs(tableInd - Math.Abs(dx)) <= STRATEGY_TOL && tableInd > 0)
                    {
                        float newInd = tableInd - dx;
                        xTarget = new DocxCorrectionTarget
                        {
                            Axis = "X",
                            Property = "Table.LeftIndent",
                            CurrentValuePt = tableInd,
                            NewValuePt = newInd,
                            CurrentValueTwips = (int)Math.Round(tableInd * 20),
                            NewValueTwips = (int)Math.Round(newInd * 20),
                            Description = $"[TABLE_INDENT] Reduce Table.LeftIndent from " +
                                               $"{tableInd:0.##}pt → {newInd:0.##}pt to cancel +{dx:0.##}pt drift. " +
                                               $"Apply once per table — NOT per word."
                        };
                    }
                    // PRIORITY 3: Column-width differential (non-uniform drifts; col-1+ only)
                    else if (elem.ColIndex > 0)
                    {
                        // The previous column's width needs shrinking to pull this column left.
                        // currentValuePt uses actual colWidth of the PREVIOUS column.
                        float prevColWidth = 0f; // Will be filled if we have it — for now use what we know
                                                 // We don't store prev-col width; use a descriptive fallback
                        xTarget = new DocxCorrectionTarget
                        {
                            Axis = "X",
                            Property = $"Table.Rows[*].Cells[{elem.ColIndex - 1}].CellFormat.Width",
                            CurrentValuePt = prevColWidth,   // 0 = unknown; corrector will read from doc
                            NewValuePt = prevColWidth - dx,
                            CurrentValueTwips = 0,
                            NewValueTwips = (int)Math.Round(-dx * 20),
                            Description = $"[COLUMN_WIDTH] Reduce col[{elem.ColIndex - 1}] width by {dx:0.##}pt " +
                                               $"to shift col[{elem.ColIndex}] left. " +
                                               $"Current width must be read from the Word document at runtime."
                        };
                    }
                    // PRIORITY 4: Whole table drifted for unknown reason — safest fallback
                    else
                    {
                        float newInd = tableInd - dx;
                        xTarget = new DocxCorrectionTarget
                        {
                            Axis = "X",
                            Property = "Table.LeftIndent",
                            CurrentValuePt = tableInd,
                            NewValuePt = newInd,
                            CurrentValueTwips = (int)Math.Round(tableInd * 20),
                            NewValueTwips = (int)Math.Round(newInd * 20),
                            Description = $"[TABLE_INDENT_FALLBACK] Table.LeftIndent: " +
                                               $"{tableInd:0.##}pt → {newInd:0.##}pt. " +
                                               $"Root cause unknown; padding and indent did not match drift."
                        };
                    }

                    if (xTarget != null) targets.Add(xTarget);
                }
                else if (elem.ElementType == "LIST_ITEM")
                {
                    float cur = elem.ListTextPositionPt ?? elem.LeftIndentPt;
                    float nw = cur - dx;
                    targets.Add(new DocxCorrectionTarget
                    {
                        Axis = "X",
                        Property = "ListLevel.TextPosition (w:lvl/w:ind/@w:left)",
                        CurrentValuePt = cur,
                        NewValuePt = nw,
                        CurrentValueTwips = (int)Math.Round(cur * 20),
                        NewValueTwips = (int)Math.Round(nw * 20),
                        Description = $"[LIST] Adjust list TextPosition {cur:0.##}pt → {nw:0.##}pt. " +
                                           $"Also adjust NumberPosition by same amount to keep bullet aligned."
                    });
                }
                else if (elem.ElementType == "IMAGE" && elem.IsInlineImage == false)
                {
                    float cur = elem.ImageAnchorXPt ?? 0f;
                    float nw = cur - dx;
                    long emu = (long)Math.Round(-dx * 12700);
                    targets.Add(new DocxCorrectionTarget
                    {
                        Axis = "X",
                        Property = "Shape.Left (wp:posOffset horizontal)",
                        CurrentValuePt = cur,
                        NewValuePt = nw,
                        CurrentValueTwips = (int)Math.Round(cur * 20),
                        NewValueTwips = (int)Math.Round(nw * 20),
                        CorrectionEmu = emu,
                        Description = $"[FLOATING_IMAGE] Adjust X anchor: {cur:0.##}pt → {nw:0.##}pt ({emu} EMU)."
                    });
                }
                else
                {
                    // Regular paragraph, header, footer, merge field outside table
                    float cur = elem.LeftIndentPt;
                    float nw = cur - dx;
                    targets.Add(new DocxCorrectionTarget
                    {
                        Axis = "X",
                        Property = "ParagraphFormat.LeftIndent (w:ind/@w:left)",
                        CurrentValuePt = cur,
                        NewValuePt = nw,
                        CurrentValueTwips = (int)Math.Round(cur * 20),
                        NewValueTwips = (int)Math.Round(nw * 20),
                        Description = $"[PARAGRAPH] LeftIndent: {cur:0.##}pt → {nw:0.##}pt " +
                                           $"({(int)Math.Round(-dx * 20)} twips)."
                    });
                }
            }

            // ── Y correction ─────────────────────────────────────────────────────────
            if (Math.Abs(dy) > 0.5f)
            {
                DocxCorrectionTarget yTarget;

                if (elem.ElementType == "TABLE_CELL")
                {
                    float curH = elem.RowHeightPt ?? 0f;
                    if (curH > 0)
                    {
                        float newH = Math.Max(0f, curH - dy);
                        yTarget = new DocxCorrectionTarget
                        {
                            Axis = "Y",
                            Property = "Row.RowFormat.Height (w:trHeight/@w:val)",
                            CurrentValuePt = curH,
                            NewValuePt = newH,
                            CurrentValueTwips = (int)Math.Round(curH * 20),
                            NewValueTwips = (int)Math.Round(newH * 20),
                            Description = $"[ROW_HEIGHT] Row height: {curH:0.##}pt → {newH:0.##}pt."
                        };
                    }
                    else
                    {
                        float cur = elem.SpaceBeforePt;
                        float nw = Math.Max(0f, cur - dy);
                        yTarget = new DocxCorrectionTarget
                        {
                            Axis = "Y",
                            Property = "ParagraphFormat.SpaceBefore inside cell (w:spacing/@w:before)",
                            CurrentValuePt = cur,
                            NewValuePt = nw,
                            CurrentValueTwips = (int)Math.Round(cur * 20),
                            NewValueTwips = (int)Math.Round(nw * 20),
                            Description = $"[ROW_SPACINGBEFORE] Auto-height row; adjust paragraph SpaceBefore: " +
                                               $"{cur:0.##}pt → {nw:0.##}pt."
                        };
                    }
                }
                else if (elem.ElementType == "IMAGE" && elem.IsInlineImage == false)
                {
                    float cur = elem.ImageAnchorYPt ?? 0f;
                    float nw = cur - dy;
                    long emu = (long)Math.Round(-dy * 12700);
                    yTarget = new DocxCorrectionTarget
                    {
                        Axis = "Y",
                        Property = "Shape.Top (wp:posOffset vertical)",
                        CurrentValuePt = cur,
                        NewValuePt = nw,
                        CurrentValueTwips = (int)Math.Round(cur * 20),
                        NewValueTwips = (int)Math.Round(nw * 20),
                        CorrectionEmu = emu,
                        Description = $"[FLOATING_IMAGE] Adjust Y anchor: {cur:0.##}pt → {nw:0.##}pt ({emu} EMU)."
                    };
                }
                else
                {
                    float cur = elem.SpaceBeforePt;
                    float nw = Math.Max(0f, cur - dy);
                    yTarget = new DocxCorrectionTarget
                    {
                        Axis = "Y",
                        Property = "ParagraphFormat.SpaceBefore (w:spacing/@w:before)",
                        CurrentValuePt = cur,
                        NewValuePt = nw,
                        CurrentValueTwips = (int)Math.Round(cur * 20),
                        NewValueTwips = (int)Math.Round(nw * 20),
                        Description = $"[PARAGRAPH_Y] SpaceBefore: {cur:0.##}pt → {nw:0.##}pt. " +
                                           $"If already 0, adjust SpaceAfter of the preceding paragraph instead."
                    };
                }

                targets.Add(yTarget);
            }

            return targets.Count > 0 ? targets : null;
        }

        /// <summary>
        /// After phrase groups and table groups are built using PDF heuristics,
        /// uses docxContext data to:
        ///
        ///    1. Assign correct tableGroupId / tableColumn / tableRow to phrase groups
        ///       the PDF heuristic missed.
        ///
        ///    2. Ensure tableGroups.phraseGroupIds includes ALL phrase groups that belong
        ///       to the same Word table.
        ///
        ///    3. Set a single unambiguous correctionStrategy on each tableGroup.
        ///
        ///    4. FIX: When the PDF heuristic missed a table entirely (existingGroup == null),
        ///    create a new TableDiffGroup from the accurate docxContext word data instead
        ///    of doing nothing. This handles single-row tables and tables with full-width
        ///    merged header rows that collapse to a single phrase group.
        ///
        /// Call this AFTER BuildPhraseGroups, TagInlineRuns, DetectTableGroups, and
        /// BackFillTableInfoToWords — as the very last step before serialising the report.
        /// </summary>
        private void AugmentTableGroupsFromDocxContext(
             List<WordDiffEntry> wordDiffs,
             List<PhraseDiffGroup> phraseGroups,
             List<TableDiffGroup> tableGroups)
        {
            if (wordDiffs == null || phraseGroups == null) return;

            // ── Step 1: Build a map of (docx tableIndex) → list of phrase group IDs ──
            var phraseToDocxTable = new Dictionary<int, int>(); // phraseGroupId → docxTableIndex

            foreach (var phrase in phraseGroups)
            {
                if (phrase.WordIds == null) continue;

                var votes = wordDiffs
                    .Where(w => phrase.WordIds.Contains(w.DifferenceId) &&
                                w.DocxContext?.TableIndex != null)
                    .GroupBy(w => w.DocxContext.TableIndex.Value)
                    .OrderByDescending(g => g.Count())
                    .FirstOrDefault();

                if (votes != null)
                    phraseToDocxTable[phrase.GroupId] = votes.Key;
            }

            // ── Step 2: Group phrase groups by docxContext table index ─────────────
            var docxTableToPhraseIds = phraseToDocxTable
                .GroupBy(kv => kv.Value)
                .ToDictionary(g => g.Key, g => g.Select(kv => kv.Key).ToList());

            // ── Step 3: For each docx table, find or create a matching JSON tableGroup
            foreach (var kvp in docxTableToPhraseIds)
            {
                int docxTableIdx = kvp.Key;
                var allPhraseIds = kvp.Value;

                // Find existing tableGroup that overlaps with these phrase IDs
                TableDiffGroup existingGroup = tableGroups?.FirstOrDefault(tg =>
                    tg.PhraseGroupIds != null &&
                    tg.PhraseGroupIds.Any(id => allPhraseIds.Contains(id)));

                if (existingGroup != null)
                {
                    // ── Existing group: add missing phrase IDs and stamp strategy ────
                    foreach (int pgId in allPhraseIds)
                    {
                        if (!existingGroup.PhraseGroupIds.Contains(pgId))
                            existingGroup.PhraseGroupIds.Add(pgId);
                    }

                    // Ensure tableGroupId is backfilled on phrase groups that were missed
                    foreach (int pgId in allPhraseIds)
                    {
                        var pg = phraseGroups.FirstOrDefault(p => p.GroupId == pgId);
                        if (pg != null && !pg.TableGroupId.HasValue)
                            pg.TableGroupId = existingGroup.TableId;
                    }

                    StampCorrectionStrategy(existingGroup, wordDiffs, allPhraseIds);
                }
                else if (tableGroups != null)
                {
                    // ── FIX: Heuristic missed this table entirely ──────────────────
                    foreach (int pgId in allPhraseIds)
                    {
                        var pg = phraseGroups.FirstOrDefault(p => p.GroupId == pgId);
                        if (pg == null) continue;

                        var repWord = wordDiffs
                            .Where(w => w.PhraseGroupId == pgId &&
                                        w.DocxContext?.TableIndex == docxTableIdx)
                            .OrderBy(w => w.DocxContext.ColIndex ?? 99)
                            .FirstOrDefault();

                        if (repWord?.DocxContext != null)
                        {
                            pg.TableColumn = repWord.DocxContext.ColIndex;
                            pg.TableRow = repWord.DocxContext.RowIndex;
                        }
                    }

                    int newId = (tableGroups.Count > 0 ? tableGroups.Max(t => t.TableId) : 0) + 1;

                    var newGroup = BuildTableGroupFromDocxContext(
                        docxTableIdx, allPhraseIds, phraseGroups, wordDiffs, newId);

                    if (newGroup != null)
                    {
                        StampCorrectionStrategy(newGroup, wordDiffs, allPhraseIds);
                        tableGroups.Add(newGroup);

                        foreach (int pgId in allPhraseIds)
                        {
                            var pg = phraseGroups.FirstOrDefault(p => p.GroupId == pgId);
                            if (pg != null)
                            {
                                pg.TableGroupId = newGroup.TableId;
                                pg.LayoutContext = "TABLE_CELL";

                                if (pg.NegativeCorrection != null)
                                    pg.NegativeCorrection.WordCorrectionHint =
                                        $"Table {newGroup.TableId} (docx-derived), " +
                                        $"Col {pg.TableColumn}, Row {pg.TableRow}: " +
                                        $"see tableGroups for correction strategy.";
                            }
                        }

                        foreach (var w in wordDiffs)
                        {
                            if (w.PhraseGroupId.HasValue && allPhraseIds.Contains(w.PhraseGroupId.Value))
                            {
                                w.TableGroupId = newGroup.TableId;
                                w.TableColumn = w.DocxContext?.ColIndex ?? 0;
                                w.TableRow = w.DocxContext?.RowIndex ?? 0;
                                w.LayoutContext = "TABLE_CELL";
                            }
                        }
                    }
                }
            }

            // ── Step 4: Stamp tableColumn/tableRow on phrase groups using docxContext ─
            foreach (var phrase in phraseGroups)
            {
                if (!phrase.TableGroupId.HasValue || phrase.WordIds == null) continue;

                var repWord = wordDiffs
                    .Where(w => phrase.WordIds.Contains(w.DifferenceId) &&
                                w.DocxContext?.TableIndex != null)
                    .OrderBy(w => w.DocxContext.ColIndex ?? 99)
                    .FirstOrDefault();

                if (repWord?.DocxContext == null) continue;

                phrase.TableColumn = repWord.DocxContext.ColIndex;
                phrase.TableRow = repWord.DocxContext.RowIndex;

                // Backfill on word-level entries too
                foreach (int wordId in phrase.WordIds)
                {
                    var word = wordDiffs.FirstOrDefault(w => w.DifferenceId == wordId);
                    if (word == null) continue;
                    word.TableGroupId = phrase.TableGroupId;
                    word.TableColumn = word.DocxContext?.ColIndex ?? phrase.TableColumn;
                    word.TableRow = word.DocxContext?.RowIndex ?? phrase.TableRow;
                }
            }

            // ═════════════════════════════════════════════════════════════════════════
            // ── Step 5 (NEW): Validate phrase groups in PDF-heuristic table groups
            //                  against docxContext — remove false-positive table cells
            // ═════════════════════════════════════════════════════════════════════════
            //
            // The PDF heuristic can assign body-paragraph text as "table cells" when
            // wrapping justified paragraphs produce word positions that look like table
            // columns. Cross-check each phrase group with docxContext:
            //   - Has docxContext.TableIndex → genuinely in a Word table  → keep ✓
            //   - No docxContext.TableIndex  → body paragraph fragment      → purge ✗
            //
            // Purged phrase groups get layoutContext = "INLINE_RUN" so the corrector's
            // PATH 3 skips them (IsParaType() returns false for INLINE_RUN). This
            // prevents incorrect paragraph-indent corrections being applied to them.
            // ═════════════════════════════════════════════════════════════════════════
            if (tableGroups != null)
            {
                var tableGroupsToRemove = new List<TableDiffGroup>();

                foreach (var tg in tableGroups)
                {
                    if (tg.PhraseGroupIds == null || tg.PhraseGroupIds.Count == 0)
                        continue;

                    var phraseIdsToRemove = new List<int>();

                    foreach (int pgId in tg.PhraseGroupIds)
                    {
                        var pg = phraseGroups.FirstOrDefault(p => p.GroupId == pgId);
                        if (pg?.WordIds == null) continue;

                        // Does ANY word in this phrase group have confirmed docxContext.TableIndex?
                        bool confirmedInWordTable = wordDiffs
                            .Any(w => pg.WordIds.Contains(w.DifferenceId) &&
                                      w.DocxContext?.TableIndex != null);

                        string phraseTextStr = string.Join(" ", wordDiffs.Where(w => pg.WordIds.Contains(w.DifferenceId)).Select(w => w.Text));
                        _logger.Log($"[PDFComparer] Evaluating TableGroup {tg.TableId}, PhraseGroup {pg.GroupId} '{phraseTextStr}'. Confirmed in Word Table? {confirmedInWordTable}");

                        if (confirmedInWordTable && phraseTextStr.Contains("Subpoena"))
                        {
                            var offendingWords = wordDiffs.Where(w => pg.WordIds.Contains(w.DifferenceId) && w.DocxContext?.TableIndex != null).ToList();
                            foreach (var ow in offendingWords)
                            {
                                _logger.Log($"   -> Offending Word: '{ow.Text}' mapped to DocxTableIndex: {ow.DocxContext?.TableIndex}");
                            }
                        }

                        if (!confirmedInWordTable)
                        {
                            _logger.Log($"[PDFComparer] -> Purging PhraseGroup {pg.GroupId} from TableGroup {tg.TableId}. Converting to INLINE_RUN.");
                            // ── Body-paragraph fragment misclassified as TABLE_CELL ─
                            // Remove from this table group and reclassify to INLINE_RUN.
                            // Rationale for INLINE_RUN (not PARAGRAPH):
                            //   • PATH 3 only processes PARAGRAPH / LIST_ITEM / MERGE_FIELD.
                            //   • INLINE_RUN is explicitly excluded from PATH 3.
                            //   • Intra-paragraph word X-position variation (e.g., a word
                            //     in the middle of a wrapped line) is not correctable by a
                            //     paragraph LeftIndent adjustment — attempting it would
                            //     corrupt the indent of the whole paragraph.
                            phraseIdsToRemove.Add(pgId);

                            pg.TableGroupId = null;
                            pg.TableColumn = null;
                            pg.TableRow = null;

                            // Only override if the PDF heuristic forced it to TABLE_CELL.
                            // If TagInlineRuns had already set INLINE_RUN, keep that label.
                            if (pg.LayoutContext == "TABLE_CELL")
                                pg.LayoutContext = "INLINE_RUN";

                            // Update individual word entries
                            foreach (int wordId in pg.WordIds)
                            {
                                var word = wordDiffs.FirstOrDefault(w => w.DifferenceId == wordId);
                                if (word == null) continue;
                                if (word.LayoutContext == "TABLE_CELL")
                                    word.LayoutContext = "INLINE_RUN";
                                // Clear table-group metadata from word entries
                                word.TableGroupId = null;
                                word.TableColumn = null;
                                word.TableRow = null;
                            }

                            // Update the phrase group's negativeCorrection hint to reflect
                            // that this is now an INLINE_RUN (no Word correction possible)
                            if (pg.NegativeCorrection != null)
                                pg.NegativeCorrection.WordCorrectionHint =
                                    "[RECLASSIFIED] This phrase group was initially placed in a " +
                                    $"PDF-heuristic table group ({tg.TableId}) but no docxContext " +
                                    "confirmed it is inside a Word table. It has been reclassified " +
                                    "as INLINE_RUN. This is likely intra-paragraph word-position " +
                                    "variation caused by V13→V23 text-layout changes (character " +
                                    "spacing, kerning, justification). No Word correction is possible.";
                        }
                    }

                    // Remove purged phrase IDs from the table group
                    foreach (int pgId in phraseIdsToRemove)
                        tg.PhraseGroupIds.Remove(pgId);

                    // If the table group now has NO docxContext-confirmed members it was
                    // entirely a PDF-heuristic false positive → schedule for removal.
                    if (tg.PhraseGroupIds.Count == 0)
                    {
                        tableGroupsToRemove.Add(tg);
                        Console.WriteLine(
                            $"[AugmentTableGroups] TABLE GROUP {tg.TableId} removed: " +
                            $"all {phraseIdsToRemove.Count} phrase group(s) were PDF-heuristic " +
                            $"false-positives (no docxContext.TableIndex confirmed). " +
                            $"Body-paragraph text was mistaken for table columns.");
                    }
                    else if (phraseIdsToRemove.Count > 0)
                    {
                        Console.WriteLine(
                            $"[AugmentTableGroups] TABLE GROUP {tg.TableId}: removed " +
                            $"{phraseIdsToRemove.Count} false-positive phrase group(s) " +
                            $"(reclassified as INLINE_RUN). " +
                            $"{tg.PhraseGroupIds.Count} genuine table phrase group(s) remain.");

                        // Re-stamp correction strategy now that only genuine table members remain
                        StampCorrectionStrategy(tg, wordDiffs, tg.PhraseGroupIds);

                        // Recompute ColumnDrifts from genuine phrase groups only.
                        // The PDF heuristic may include spurious body-paragraph words in the
                        // per-column drift averages (e.g. an INLINE_RUN word near a column X
                        // position gets assigned to that column, skewing avgDeltaX). After
                        // purging those false positives, recalculate so ColumnDrifts accurately
                        // reflect the genuine table drift — preventing PATH 2 from over-correcting
                        // if it were to use this data (e.g. avgX = 9.74pt instead of true 5.4pt).
                        var genuineColGroups = phraseGroups
                            .Where(p => tg.PhraseGroupIds.Contains(p.GroupId) &&
                                        p.TableColumn.HasValue &&
                                        p.SharedDelta != null)
                            .GroupBy(p => p.TableColumn.Value)
                            .OrderBy(g => g.Key)
                            .ToList();

                        if (genuineColGroups.Count > 0)
                        {
                            var updatedColDrifts = new List<TableColumnDrift>();

                            // For COLUMN_WIDTHS strategy where no col[0] drift is observed
                            // (only higher-index columns have shifted), prepend a synthetic
                            // col[0] entry with avgDeltaX=0. This ensures that ApplyColumnWidths
                            // (via BuildTablePlanFromJsonGroup) does not shift the table indent —
                            // only the preceding column's width is adjusted.
                            if (tg.CorrectionStrategy == "COLUMN_WIDTHS" && genuineColGroups[0].Key > 0)
                            {
                                updatedColDrifts.Add(new TableColumnDrift
                                {
                                    ColumnIndex = 0,
                                    BaselineXStart = 0f,
                                    AvgDeltaX = 0f,
                                    WidthCorrectionTwips = 0,
                                    WidthCorrectionPt = "0pt"
                                });
                            }

                            float prevColDx = 0f;
                            foreach (var cg in genuineColGroups)
                            {
                                float avgDx = (float)cg.Average(p => p.SharedDelta.X);
                                float widthCorr = -(avgDx - prevColDx);
                                float baseX = tg.ColumnDrifts?.FirstOrDefault(c => c.ColumnIndex == cg.Key)?.BaselineXStart ?? 0f;
                                updatedColDrifts.Add(new TableColumnDrift
                                {
                                    ColumnIndex = cg.Key,
                                    BaselineXStart = baseX,
                                    AvgDeltaX = avgDx,
                                    WidthCorrectionTwips = (int)Math.Round(widthCorr * 20),
                                    WidthCorrectionPt = $"{widthCorr:+0.##;-0.##;0}pt"
                                });
                                prevColDx = avgDx;
                            }
                            tg.ColumnDrifts = updatedColDrifts;
                            Console.WriteLine(
                                $"[AugmentTableGroups] TABLE GROUP {tg.TableId}: recomputed ColumnDrifts " +
                                $"from {genuineColGroups.Count} genuine column(s) — " +
                                string.Join(", ", updatedColDrifts.Select(c => $"col[{c.ColumnIndex}]={c.AvgDeltaX:+0.##;-0.##;0}pt")));
                        }
                    }
                }

                // Remove wholly spurious table groups
                foreach (var tg in tableGroupsToRemove)
                    tableGroups.Remove(tg);
            }
        }

        /// <summary>
        /// Creates a TableDiffGroup from docxContext word metadata when the PDF
        /// heuristic (DetectTableGroups) missed the table entirely.
        ///
        /// This is called from AugmentTableGroupsFromDocxContext when no existing
        /// tableGroup overlaps with the phrase groups belonging to a docx table.
        ///
        /// Uses w.DocxContext.ColIndex and w.DocxContext.RowIndex as the ground-truth
        /// column/row assignments instead of the PDF XStart-clustering heuristic.
        /// This is more accurate because it reflects the actual Word document structure.
        ///
        /// Handles the following cases that defeat the PDF heuristic:
        ///   • Single-row tables with only content row (merged header collapses to 1 phrase)
        ///   • Tables where all rows happen to be in a single Y-cluster
        ///   • Very narrow tables with only 2 columns and 1 data row
        /// </summary>
        private TableDiffGroup BuildTableGroupFromDocxContext(
            int docxTableIdx,
            List<int> phraseGroupIds,
            List<PhraseDiffGroup> phraseGroups,
            List<WordDiffEntry> wordDiffs,
            int newTableId)
        {
            // Gather all words in this docx table that have drift data
            var tableWords = wordDiffs
                .Where(w => w.DocxContext?.TableIndex == docxTableIdx &&
                            w.Delta != null &&
                            w.PhraseGroupId.HasValue &&
                            phraseGroupIds.Contains(w.PhraseGroupId.Value))
                .ToList();

            if (tableWords.Count == 0) return null;

            // ── Build per-column X drifts from docx column index ─────────────────────
            var xByDocxCol = tableWords
                .Where(w => w.DocxContext.ColIndex != null)
                .GroupBy(w => w.DocxContext.ColIndex.Value)
                .OrderBy(g => g.Key)
                .ToDictionary(g => g.Key, g => g.Average(w => w.Delta.X));

            if (xByDocxCol.Count == 0) return null;

            // ── Build per-row Y drifts from docx row index ───────────────────────────
            var yByDocxRow = tableWords
                .Where(w => w.DocxContext.RowIndex != null && Math.Abs(w.Delta.Y) > 0.3f)
                .GroupBy(w => w.DocxContext.RowIndex.Value)
                .ToDictionary(g => g.Key, g => g.Average(w => w.Delta.Y));

            // ── Build TableColumnDrift list ───────────────────────────────────────────
            var columnDrifts = new List<TableColumnDrift>();
            float prevAvgDx = 0f;

            foreach (var kv in xByDocxCol.OrderBy(k => k.Key))
            {
                // Width correction = differential between this col's drift and the
                // previous col's drift. Uniform drift → widthCorrection = 0 (whole-table shift).
                float widthCorr = -(kv.Value - prevAvgDx);

                // ── Primary: phrase-group baseline X lookup ───────────────────────────
                float baseX = phraseGroups
                    .Where(pg => phraseGroupIds.Contains(pg.GroupId) &&
                                 pg.TableColumn == kv.Key)
                    .Select(pg => pg.BaselineRegion.XStart)
                    .DefaultIfEmpty(0f)
                    .Average();

                // ── Fallback: word-level baseline X ──────────────────────────────────
                // Used when no phrase group has pg.TableColumn matching this column.
                // This happens when a phrase group spans multiple docx columns but was
                // assigned the representative column index only (e.g. phrase group 2
                // "Shared Retro Date Retention" → tableColumn=2, leaving cols 3 and 4
                // with baselineXStart=0 from the phrase lookup above).
                if (baseX < 1f)
                {
                    var colWords = tableWords
                        .Where(w => w.DocxContext.ColIndex == kv.Key &&
                                    w.Baseline != null &&
                                    w.Baseline.X > 0f)
                        .ToList();

                    if (colWords.Count > 0)
                        baseX = colWords.Average(w => w.Baseline.X);
                }

                columnDrifts.Add(new TableColumnDrift
                {
                    ColumnIndex = kv.Key,
                    BaselineXStart = baseX,
                    AvgDeltaX = kv.Value,
                    WidthCorrectionTwips = (int)Math.Round(widthCorr * 20),
                    WidthCorrectionPt = $"{widthCorr:+0.##;-0.##;0}pt"
                });

                prevAvgDx = kv.Value;
            }

            // ── Build TableRowDrift list ──────────────────────────────────────────────
            var rowDrifts = new List<TableRowDrift>();
            float prevAvgDy = 0f;

            foreach (var kv in yByDocxRow.OrderBy(k => k.Key))
            {
                float heightCorr = -(kv.Value - prevAvgDy);

                float baseY = phraseGroups
                    .Where(pg => phraseGroupIds.Contains(pg.GroupId) &&
                                 pg.TableRow == kv.Key)
                    .Select(pg => pg.BaselineRegion.Y)
                    .DefaultIfEmpty(0f)
                    .Average();

                // Fallback: word-level baseline Y for rows with no matching phrase group
                if (baseY < 1f)
                {
                    var rowWords = tableWords
                        .Where(w => w.DocxContext.RowIndex == kv.Key &&
                                    w.Baseline != null &&
                                    w.Baseline.Y > 0f)
                        .ToList();

                    if (rowWords.Count > 0)
                        baseY = rowWords.Average(w => w.Baseline.Y);
                }

                rowDrifts.Add(new TableRowDrift
                {
                    RowIndex = kv.Key,
                    BaselineY = baseY,
                    AvgDeltaY = kv.Value,
                    HeightCorrectionTwips = (int)Math.Round(heightCorr * 20),
                    HeightCorrectionPt = $"{heightCorr:+0.##;-0.##;0}pt"
                });

                prevAvgDy = kv.Value;
            }

            // ── Compute bounds from phrase groups ─────────────────────────────────────
            var relevantPhrases = phraseGroups
                .Where(pg => phraseGroupIds.Contains(pg.GroupId))
                .ToList();

            if (relevantPhrases.Count == 0) return null;

            float bxStart = relevantPhrases.Min(p => p.BaselineRegion.XStart);
            float bxEnd = relevantPhrases.Max(p => p.BaselineRegion.XEnd);
            float byStart = relevantPhrases.Min(p => p.BaselineRegion.Y);
            float pxStart = relevantPhrases.Min(p => p.PublishRegion.XStart);
            float pxEnd = relevantPhrases.Max(p => p.PublishRegion.XEnd);
            float pyStart = relevantPhrases.Min(p => p.PublishRegion.Y);

            int distinctRows = tableWords
                .Where(w => w.DocxContext.RowIndex != null)
                .Select(w => w.DocxContext.RowIndex.Value)
                .Distinct()
                .Count();

            int page = relevantPhrases[0].Page;

            return new TableDiffGroup
            {
                TableId = newTableId,
                Page = page,
                RowCount = Math.Max(1, distinctRows),
                ColumnCount = columnDrifts.Count,
                ColumnDrifts = columnDrifts,
                RowDrifts = rowDrifts.Count > 0 ? rowDrifts : new List<TableRowDrift>(),
                BaselineBounds = new PhraseRegion { XStart = bxStart, XEnd = bxEnd, Y = byStart },
                PublishBounds = new PhraseRegion { XStart = pxStart, XEnd = pxEnd, Y = pyStart },
                WordCorrectionHint =
                    $"[DOCX-CONTEXT-DERIVED] Table {docxTableIdx} ({columnDrifts.Count} cols, " +
                    $"{Math.Max(1, distinctRows)} rows): detected via docxContext after PDF " +
                    $"heuristic missed it (e.g. merged/colspan header row or single data row).",
                PhraseGroupIds = phraseGroupIds.ToList()
                // CorrectionStrategy stamped by StampCorrectionStrategy after this returns
            };
        }

        /// <summary>
        /// Examines the words in a table group and stamps a single
        /// correctionStrategy string onto the tableGroup.
        ///
        /// CELL_PADDING         — cellLeftPaddingPt ≈ uniformDriftX (most common)
        /// TABLE_INDENT        — tableLeftIndentPt magnitude ≈ uniformDriftX
        /// COLUMN_WIDTHS       — drifts are non-uniform across columns
        /// TABLE_INDENT_FALLBACK — uniform drift, but root cause unclear
        ///
        /// FIX: The original condition `tableInd > 0` prevented TABLE_INDENT from
        /// ever being selected when the table has a negative left indent (a common
        /// pattern in professional letter templates). Fixed to use Math.Abs comparison.
        /// Also added tableInd >= 0 guard to CELL_PADDING selection to prevent
        /// over-correction into the left margin for negative-indent tables.
        /// </summary>
        private void StampCorrectionStrategy(
            TableDiffGroup tg,
            List<WordDiffEntry> wordDiffs,
            List<int> phraseGroupIds)
        {
            // TOL = 1.0f: raised from 0.5f to match WordDriftCorrector.UNIFORM_TOL.
            // Prevents float-precision false-negatives where PDF coordinate rounding
            // pushes a genuinely uniform column spread just above 0.5f.
            const float TOL = 1.0f;

            // CELL_PADDING is only valid if the corrected padding remains above this value.
            // Setting cell margin to ≤ 1pt is almost always a wrong correction.
            const float MIN_REMAINING_PADDING_PT = 1.0f;

            // Collect words from this table's phrase groups that have drift and docxContext data
            var words = wordDiffs
                .Where(w => w.PhraseGroupId.HasValue &&
                            phraseGroupIds.Contains(w.PhraseGroupId.Value) &&
                            w.Delta != null &&
                            w.DocxContext != null)
                .ToList();

            if (words.Count == 0) return;

            // Build per-column average X drift
            var xByCol = words
                .Where(w => w.DocxContext.ColIndex != null)
                .GroupBy(w => w.DocxContext.ColIndex.Value)
                .ToDictionary(g => g.Key, g => g.Average(w => w.Delta.X));

            if (xByCol.Count == 0) return;

            float minX = xByCol.Values.Min();
            float maxX = xByCol.Values.Max();
            bool isUniform = (maxX - minX) <= TOL;
            float avgX = xByCol.Values.Average();

            // Sample a col-0 word for cell padding and table indent context
            var sample = words.FirstOrDefault(w => w.DocxContext.ColIndex == 0)?.DocxContext
                         ?? words.First().DocxContext;

            float cellPad = sample.CellLeftPaddingPt ?? 0f;
            float tableInd = sample.TableLeftIndentPt ?? 0f;

            string strategy;
            string hint;

            if (!isUniform)
            {
                // ── COLUMN_WIDTHS ─────────────────────────────────────────────────────
                // Genuine per-column drift differences → column widths need adjustment.
                strategy = "COLUMN_WIDTHS";
                var colLines = xByCol
                    .OrderBy(kv => kv.Key)
                    .Select(kv => $"  Col[{kv.Key}]: drift={kv.Value:+0.##;-0.##}pt");
                hint = "WORD TABLE CORRECTIONS (COLUMN_WIDTHS strategy):\r\n" +
                       string.Join("\r\n", colLines) + "\r\n" +
                       "  Apply differential column-width corrections. See columnDrifts[] for per-column values.";
            }
            else if (cellPad > 0 &&
                     Math.Abs(avgX) < cellPad &&                         // *** THIS ROUND'S FIX (part 1) ***
                     (cellPad - Math.Abs(avgX)) > MIN_REMAINING_PADDING_PT && // *** THIS ROUND'S FIX (part 2) ***
                     tableInd >= 0f)                                     // round 1 fix: skip for negative-indent tables
            {
                // ── CELL_PADDING ──────────────────────────────────────────────────────
                // Uniform drift is LESS than the current cell left padding AND the
                // corrected padding stays meaningfully positive (> 1pt) AND table has
                // non-negative indent.
                //
                // WHAT CHANGED from round 1 (was: |cellPad - avgX| <= TOL):
                //   The old condition detected drift ≈ full cellPad, which could not
                //   coexist with a reasonable remaining padding. The new conditions are:
                //
                //   (a) Math.Abs(avgX) < cellPad
                //       Physical check: drift can't exceed the full cell padding.
                //       If avgX >= cellPad, removing all cell padding still isn't enough
                //       to explain the drift → root cause is not cell padding.
                //       For ManagementLiability: avgX=2.783 > cellPad=2.7 → FALSE → skip.
                //
                //   (b) (cellPad - Math.Abs(avgX)) > MIN_REMAINING_PADDING_PT (1.0pt)
                //       Sanity check: the corrected padding must be meaningful.
                //       If the correction would reduce padding to ≤ 1pt, the drift is
                //       almost certainly not caused by a padding change — use TABLE_INDENT_FALLBACK.
                //
                // NOTE: The old condition |cellPad-avgX| <= TOL and the guard (cellPad-avgX) > 1pt
                //       were mathematically mutually exclusive (CELL_PADDING could never be selected).
                //       The restructured conditions above are both logically sound and non-contradictory.
                strategy = "CELL_PADDING";
                float newPad = cellPad - Math.Abs(avgX); // guaranteed > MIN_REMAINING_PADDING_PT here
                hint = $"WORD TABLE CORRECTIONS (CELL_PADDING strategy):\r\n" +
                       $"  Root cause: cell LeftPadding increased from {newPad:0.##}pt to {cellPad:0.##}pt between versions.\r\n" +
                       $"  Fix: set LeftPadding on ALL cells to {newPad:0.##}pt ({(int)Math.Round(newPad * 20)} twips).\r\n" +
                       $"  Apply ONCE to all cells — do NOT apply per word or per phrase.";
            }
            else if (Math.Abs(tableInd) > 0f &&               // round 1 fix: was `tableInd > 0`, handles negative indents
                     Math.Abs(Math.Abs(tableInd) - Math.Abs(avgX)) <= TOL)
            {
                // ── TABLE_INDENT ──────────────────────────────────────────────────────
                // Uniform drift ≈ table left indent magnitude (positive or negative).
                strategy = "TABLE_INDENT";
                float newInd = tableInd - avgX;
                hint = $"WORD TABLE CORRECTIONS (TABLE_INDENT strategy):\r\n" +
                       $"  Root cause: Table.LeftIndent changed from {newInd:0.##}pt to {tableInd:0.##}pt between versions.\r\n" +
                       $"  Fix: set Table.LeftIndent to {newInd:0.##}pt ({(int)Math.Round(newInd * 20)} twips).\r\n" +
                       $"  Apply ONCE per table — do NOT apply per word or per phrase." +
                       (tableInd < 0f
                           ? "\r\n  Note: table has negative indent (intentional margin positioning)."
                           : "");
            }
            else if (xByCol.Count == 1 && xByCol.Keys.Single() > 0)
            {
                // ── COLUMN_WIDTHS (single non-first column drift) ─────────────────────
                // Only one column's drift is observed AND it is not col[0].
                // Root cause: the preceding column's width changed between versions,
                // which pushed this column's content left or right. Shifting the entire
                // table indent (TABLE_INDENT_FALLBACK) would wrongly move ALL columns,
                // including col[0] content that was never displaced.
                // Correct fix: adjust col[K-1] width by -drift so that col[K] returns
                // to its original X position while col[0..K-1] content is untouched.
                int singleColKey = xByCol.Keys.Single();
                strategy = "COLUMN_WIDTHS";
                hint = $"WORD TABLE CORRECTIONS (COLUMN_WIDTHS strategy — single non-first column drift):\r\n" +
                       $"  Only Col[{singleColKey}] has observed drift: {avgX:+0.##;-0.##}pt.\r\n" +
                       $"  Root cause: Col[{singleColKey - 1}] width changed between versions, shifting Col[{singleColKey}] content.\r\n" +
                       $"  Fix: widen Col[{singleColKey - 1}] by {-avgX:+0.##;-0.##}pt ({(int)Math.Round(-avgX * 20)} twips).\r\n" +
                       $"  Do NOT shift Table.LeftIndent — Col[0] content is correctly positioned.";
            }
            else
            {
                // ── TABLE_INDENT_FALLBACK ─────────────────────────────────────────────
                // Uniform drift, but root cause (cell padding / table indent) is unclear.
                // Safest fix: shift the entire table by the average drift.
                // This is the correct strategy for whole-table shifts caused by V13→V23
                // Aspose rendering engine changes.
                strategy = "TABLE_INDENT_FALLBACK";
                hint = $"WORD TABLE CORRECTIONS (TABLE_INDENT_FALLBACK strategy):\r\n" +
                       $"  Uniform drift of {avgX:0.##}pt across all {xByCol.Count} columns.\r\n" +
                       $"  Root cause unclear (cell padding and table indent do not match drift).\r\n" +
                       $"  Fix: reduce Table.LeftIndent by {avgX:0.##}pt ({(int)Math.Round(avgX * 20)} twips).\r\n" +
                       $"  This shifts the entire table left as a unit — do NOT apply per word or per phrase.\r\n" +
                       $"  Verify: Table.LeftIndent in source .docx = {tableInd:0.##}pt. " +
                       $"New value should be {tableInd - avgX:0.##}pt.";
            }

            tg.CorrectionStrategy = strategy;
            tg.WordCorrectionHint = hint;
        }

        /// <summary>
        /// Opens a .doc or .docx file with Aspose.Words and extracts every content element
        /// (paragraphs, table cells, headers, footers, list items, images, merge fields)
        /// into a flat list of DocxElementInfo for subsequent PDF word matching.
        ///
        /// NOTE: Set your Aspose.Words license before calling this method:
        ///    new Aspose.Words.License().SetLicense("Aspose.Words.lic");
        /// </summary>
        private List<DocxElementInfo> ParseDocxElements(string docxPath)
        {
            var elements = new List<DocxElementInfo>();

            try
            {
                var doc = new Aspose.Words.Document(docxPath);

                // Page left margin in points — used for approximate X calculation
                float pageLeftMarginPt = (float)doc.FirstSection.PageSetup.LeftMargin;

                // Build table index map (document order, 1-based)
                var tableIndexMap = new Dictionary<Aspose.Words.Tables.Table, int>();
                int tIdx = 0;
                foreach (Aspose.Words.Tables.Table tbl in
                         doc.GetChildNodes(Aspose.Words.NodeType.Table, true))
                    tableIndexMap[tbl] = ++tIdx;

                int elementId = 1;

                // ── Walk every paragraph in document order ────────────────────────────
                // GetChildNodes(Paragraph, true) returns paragraphs from main body,
                // headers, footers, table cells, text boxes — in document order.
                foreach (Aspose.Words.Paragraph para in
                         doc.GetChildNodes(Aspose.Words.NodeType.Paragraph, true))
                {
                    var elem = new DocxElementInfo { ElementId = elementId++ };

                    // ── Collect plain-text runs ───────────────────────────────────────
                    var sbText = new StringBuilder();
                    foreach (var node in para.ChildNodes)
                    {
                        if (node is Aspose.Words.Run run)
                            sbText.Append(run.Text);
                    }

                    // ── Collect merge fields ──────────────────────────────────────────
                    // In a template the field result is empty; we use the field name.
                    var mergeFieldNames = new List<string>();
                    foreach (Aspose.Words.Fields.Field field in para.Range.Fields)
                    {
                        if (field.Type == Aspose.Words.Fields.FieldType.FieldMergeField)
                        {
                            var mf = (Aspose.Words.Fields.FieldMergeField)field;
                            if (!string.IsNullOrWhiteSpace(mf.FieldName))
                                mergeFieldNames.Add(mf.FieldName);
                        }
                    }

                    elem.PlainText = sbText.ToString().Trim();
                    elem.MergeFieldNames = mergeFieldNames;

                    // Build word list for per-word matching
                    var wordList = elem.PlainText
                        .Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
                        .ToList();
                    // Add merge field representations (as they appear in PDF: «Name»)
                    foreach (var mfn in mergeFieldNames)
                        wordList.Add($"«{mfn}»");
                    elem.Words = wordList;

                    // Skip completely empty elements
                    if (string.IsNullOrWhiteSpace(elem.PlainText) && mergeFieldNames.Count == 0)
                        continue;

                    // ── Style ─────────────────────────────────────────────────────────
                    elem.StyleName = para.ParagraphFormat.Style?.Name ?? "Normal";

                    // ── Header / footer detection ─────────────────────────────────────
                    var hf = (Aspose.Words.HeaderFooter)para.GetAncestor(
                                  Aspose.Words.NodeType.HeaderFooter);
                    if (hf != null)
                    {
                        elem.IsInHeader = hf.HeaderFooterType ==
                                              Aspose.Words.HeaderFooterType.HeaderFirst ||
                                          hf.HeaderFooterType ==
                                              Aspose.Words.HeaderFooterType.HeaderEven ||
                                          hf.HeaderFooterType ==
                                              Aspose.Words.HeaderFooterType.HeaderPrimary;
                        elem.IsInFooter = !elem.IsInHeader;
                    }

                    // ── Table cell detection ──────────────────────────────────────────
                    var cell = (Aspose.Words.Tables.Cell)para.GetAncestor(
                                    Aspose.Words.NodeType.Cell);

                    if (cell != null)
                    {
                        var row = cell.ParentRow;
                        var table = row.ParentTable;

                        elem.TableIndex = tableIndexMap.TryGetValue(table, out int ti) ? ti : (int?)null;
                        elem.RowIndex = table.IndexOf(row);
                        elem.ColIndex = row.IndexOf(cell);
                        elem.ColWidthPt = (float)cell.CellFormat.Width;
                        elem.RowHeightPt = row.RowFormat.Height > 0 ? (float)row.RowFormat.Height : (float?)null;
                        elem.CellLeftPaddingPt = (float)cell.CellFormat.LeftPadding;
                        elem.TableLeftIndentPt = (float)table.LeftIndent;

                        // Sum widths of all columns to the left of this cell
                        float colOffset = 0f;
                        for (int c = 0; c < elem.ColIndex; c++)
                            if (row.Cells.Count > c)
                                colOffset += (float)row.Cells[c].CellFormat.Width;
                        elem.ColumnOffsetPt = colOffset;

                        elem.ElementType = "TABLE_CELL";
                        elem.ApproximateXPt = pageLeftMarginPt
                                              + (float)table.LeftIndent
                                              + colOffset
                                              + (float)cell.CellFormat.LeftPadding
                                              + (float)para.ParagraphFormat.LeftIndent;
                    }
                    else
                    {
                        // Not in a table
                        if (elem.IsInHeader)
                            elem.ElementType = "HEADER";
                        else if (elem.IsInFooter)
                            elem.ElementType = "FOOTER";
                        else
                            elem.ElementType = "PARAGRAPH";

                        elem.ApproximateXPt = pageLeftMarginPt
                                              + (float)para.ParagraphFormat.LeftIndent;
                    }

                    // ── List / bullet detection ───────────────────────────────────────
                    if (para.ListFormat.IsListItem)
                    {
                        elem.IsList = true;
                        elem.ListId = para.ListFormat.List?.ListId;
                        elem.ListLevel = para.ListFormat.ListLevelNumber;

                        // Keep TABLE_CELL if in a table; otherwise mark as LIST_ITEM
                        if (elem.ElementType == "PARAGRAPH")
                            elem.ElementType = "LIST_ITEM";

                        try
                        {
                            var lvl = para.ListFormat.ListLevel;
                            elem.ListNumberPositionPt = (float)lvl.NumberPosition;
                            elem.ListTextPositionPt = (float)lvl.TextPosition;
                        }
                        catch { /* list level details unavailable */ }
                    }

                    // ── Paragraph formatting ──────────────────────────────────────────
                    elem.LeftIndentPt = (float)para.ParagraphFormat.LeftIndent;
                    elem.SpaceBeforePt = (float)para.ParagraphFormat.SpaceBefore;
                    elem.SpaceAfterPt = (float)para.ParagraphFormat.SpaceAfter;
                    try { elem.LineSpacingPt = (float)para.ParagraphFormat.LineSpacing; } catch { }

                    elements.Add(elem);
                }

                // ── Walk all shapes (inline and floating images) ──────────────────────
                int imgCounter = 1;
                foreach (Aspose.Words.Drawing.Shape shape in
                         doc.GetChildNodes(Aspose.Words.NodeType.Shape, true))
                {
                    if (!shape.HasImage) continue;

                    var imgElem = new DocxElementInfo
                    {
                        ElementId = elementId++,
                        ElementType = "IMAGE",
                        PlainText = $"[IMAGE_{imgCounter}]",
                        Words = new List<string> { $"[IMAGE_{imgCounter}]" },
                        MergeFieldNames = new List<string>(),
                        IsInlineImage = shape.IsInline,
                        ImageWidthPt = (float)shape.Width,
                        ImageHeightPt = (float)shape.Height
                    };

                    if (!shape.IsInline)
                    {
                        // Floating: explicit anchor position
                        imgElem.ImageAnchorXPt = (float)shape.Left;
                        imgElem.ImageAnchorYPt = (float)shape.Top;
                        imgElem.ApproximateXPt = (float)shape.Left;
                    }
                    else
                    {
                        // Inline: inherits paragraph X
                        var parentPara = (Aspose.Words.Paragraph)shape.GetAncestor(
                                             Aspose.Words.NodeType.Paragraph);
                        if (parentPara != null)
                            imgElem.ApproximateXPt = pageLeftMarginPt
                                                     + (float)parentPara.ParagraphFormat.LeftIndent;
                    }

                    // Header / footer check for image
                    var imgHf = (Aspose.Words.HeaderFooter)shape.GetAncestor(
                                     Aspose.Words.NodeType.HeaderFooter);
                    if (imgHf != null)
                    {
                        imgElem.IsInHeader = imgHf.HeaderFooterType ==
                                                 Aspose.Words.HeaderFooterType.HeaderFirst ||
                                             imgHf.HeaderFooterType ==
                                                 Aspose.Words.HeaderFooterType.HeaderEven ||
                                             imgHf.HeaderFooterType ==
                                                 Aspose.Words.HeaderFooterType.HeaderPrimary;
                        imgElem.IsInFooter = !imgElem.IsInHeader;
                    }

                    elements.Add(imgElem);
                    imgCounter++;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[ParseDocxElements] Error parsing {docxPath}: {ex.Message}");
            }

            return elements;
        }

        /// <summary>
        /// Finds the best-matching DocxElementInfo for a single PDF word.
        ///
        /// Matching strategy:
        ///  1. For merge fields (PDF text has «» guillemets): match by field name.
        ///  2. For regular text: match by word content.
        ///  3. Among text matches, pick the one with closest ApproximateXPt to the PDF word X.
        ///  4. X must be within 80pt (generous to handle drift).
        ///  5. Once matched, the element is added to 'consumed' so it cannot be matched again
        ///     by a different word (prevents double-matching repeated words like "Date:").
        /// </summary>
        private DocxElementInfo MatchWordToDocxElement(
            string pdfText,
            float pdfX,
            List<DocxElementInfo> elements,
            HashSet<int> consumed)
        {
            if (elements == null || string.IsNullOrWhiteSpace(pdfText)) return null;

            string normalized = pdfText.Trim();
            bool isMergeField = normalized.StartsWith("«") && normalized.EndsWith("»");
            string mergeFieldName = isMergeField ? normalized.Trim('«', '»') : null;

            DocxElementInfo best = null;
            float bestScore = float.MaxValue;

            foreach (var elem in elements)
            {
                if (consumed.Contains(elem.ElementId)) continue;

                bool textMatches = false;

                if (isMergeField && mergeFieldName != null)
                {
                    textMatches = elem.MergeFieldNames.Any(mf =>
                        string.Equals(mf, mergeFieldName, StringComparison.OrdinalIgnoreCase));
                }
                else
                {
                    textMatches = elem.Words.Any(w =>
                        string.Equals(w, normalized, StringComparison.Ordinal));
                }

                if (!textMatches) continue;

                float xDist = Math.Abs(elem.ApproximateXPt - pdfX);
                if (xDist > 80f) continue; // outside plausible range

                // Score = X distance; tie-break by element ID (document order)
                float score = xDist * 10 + elem.ElementId * 0.001f;
                if (score < bestScore)
                {
                    bestScore = score;
                    best = elem;
                }
            }

            return best;
        }

        /// <summary>
        /// Builds the DocxContext object for a matched word diff entry,
        /// and computes precise correction targets for each drift axis present.
        /// </summary>
        private DocxContext BuildDocxContext(WordDiffEntry entry, DocxElementInfo elem)
        {
            if (elem == null) return null;

            // Determine element type label (merge field is a sub-type, not a separate type)
            bool isMergeField = entry.Text != null &&
                                entry.Text.StartsWith("«") && entry.Text.EndsWith("»");
            string elementTypeLabel = isMergeField ? "MERGE_FIELD" : elem.ElementType;

            var ctx = new DocxContext
            {
                ElementType = elementTypeLabel,
                StyleName = elem.StyleName,
                IsInHeader = elem.IsInHeader,
                IsInFooter = elem.IsInFooter,
                IsList = elem.IsList,
                ListId = elem.ListId,
                ListLevel = elem.ListLevel,
                ListNumberPositionPt = elem.ListNumberPositionPt,
                ListTextPositionPt = elem.ListTextPositionPt,
                MergeFieldName = isMergeField
                                            ? entry.Text?.Trim('«', '»')
                                            : null,
                ParagraphLeftIndentPt = elem.LeftIndentPt,
                SpaceBeforePt = elem.SpaceBeforePt,
                SpaceAfterPt = elem.SpaceAfterPt,
                LineSpacingPt = elem.LineSpacingPt,
                TableIndex = elem.TableIndex,
                RowIndex = elem.RowIndex,
                ColIndex = elem.ColIndex,
                ColWidthPt = elem.ColWidthPt,
                RowHeightPt = elem.RowHeightPt,
                CellLeftPaddingPt = elem.CellLeftPaddingPt,
                TableLeftIndentPt = elem.TableLeftIndentPt,
                ColumnOffsetPt = elem.ColumnOffsetPt,
                IsInlineImage = elem.IsInlineImage,
                ImageWidthPt = elem.ImageWidthPt,
                ImageHeightPt = elem.ImageHeightPt,
                ImageAnchorXPt = elem.ImageAnchorXPt,
                ImageAnchorYPt = elem.ImageAnchorYPt
            };

            // ── Compute correction targets ────────────────────────────────────────────
            if (entry.Delta != null)
                ctx.CorrectionTargets = BuildCorrectionTargets(entry, elem);

            return ctx;
        }

        private static string NormFont(string font)
        {
            if (string.IsNullOrEmpty(font)) return font;
            return font.Contains("+") ? font.Split('+')[1] : font;
        }

        private static string PrimaryDiffType(List<string> issues)
        {
            if (issues.Contains("POSITION_SHIFT")) return "POSITION_SHIFT";
            if (issues.Contains("FONT_SIZE_DIFF")) return "FONT_SIZE_DIFF";
            if (issues.Contains("FONT_DIFF")) return "FONT_DIFF";
            if (issues.Contains("MISSING")) return "MISSING";
            if (issues.Contains("EXTRA")) return "EXTRA";
            return issues.FirstOrDefault() ?? "UNKNOWN";
        }

        private static NegativeCorrection BuildCorrection(float dx, float dy, string hint)
        {
            float negX = -dx;
            float negY = -dy;
            return new NegativeCorrection
            {
                X = negX,
                Y = negY,
                XTwips = (int)Math.Round(negX * 20),   // 1pt = 20 twips
                YTwips = (int)Math.Round(negY * 20),
                XPt = $"{negX:+0.##;-0.##;0}pt",
                YPt = $"{negY:+0.##;-0.##;0}pt",
                WordCorrectionHint = hint
            };
        }

        private static string BuildParagraphHint(float dx, float dy)
        {
            var parts = new List<string>();
            if (Math.Abs(dx) > 0.5f)
                parts.Add($"Adjust paragraph LeftIndent by {-dx:+0.##;-0.##;0}pt ({(int)Math.Round(-dx * 20)} twips)");
            if (Math.Abs(dy) > 0.5f)
                parts.Add($"Adjust paragraph SpaceBefore by {-dy:+0.##;-0.##;0}pt ({(int)Math.Round(-dy * 20)} twips)");
            return parts.Count > 0 ? string.Join("; ", parts) : "No Word correction needed";
        }

        // ─── Phrase grouping ─────────────────────────────────────────────────────────────

        private List<PhraseDiffGroup> BuildPhraseGroups(List<WordDiffEntry> wordDiffs)
        {
            // Only group words that have a position shift — those drive Word corrections
            var shiftable = wordDiffs
                .Where(w => w.Issues != null && w.Issues.Contains("POSITION_SHIFT")
                            && w.Baseline != null && w.Publish != null)
                .OrderBy(w => w.Page)
                .ThenByDescending(w => w.Baseline.Y)
                .ThenBy(w => w.Baseline.X)
                .ToList();

            var groups = new List<PhraseDiffGroup>();
            var assigned = new HashSet<int>();
            int groupId = 1;

            for (int i = 0; i < shiftable.Count; i++)
            {
                var seed = shiftable[i];
                if (assigned.Contains(seed.DifferenceId)) continue;

                var groupWords = new List<WordDiffEntry> { seed };
                assigned.Add(seed.DifferenceId);

                for (int j = i + 1; j < shiftable.Count; j++)
                {
                    var cand = shiftable[j];
                    if (assigned.Contains(cand.DifferenceId)) continue;

                    var last = groupWords[groupWords.Count - 1];

                    if (cand.Page != seed.Page) break;

                    // Same line: Y within 3pt
                    if (Math.Abs(cand.Baseline.Y - seed.Baseline.Y) > 3f) break;

                    // Adjacent: candidate X starts within 30pt of last word's estimated end
                    float estLastEnd = last.Baseline.X + EstimateWidth(last.Text, last.Baseline.FontSize);
                    if (cand.Baseline.X > estLastEnd + 30f) continue;
                    if (cand.Baseline.X < last.Baseline.X) continue;

                    // Same drift vector within 1pt tolerance
                    if (Math.Abs(cand.Delta.X - seed.Delta.X) > 1f) continue;
                    if (Math.Abs(cand.Delta.Y - seed.Delta.Y) > 1f) continue;

                    groupWords.Add(cand);
                    assigned.Add(cand.DifferenceId);
                }

                string phraseText = string.Join(" ", groupWords.Select(w => w.Text));
                var allIssues = groupWords.SelectMany(w => w.Issues).Distinct().ToList();

                float bxStart = groupWords.Min(w => w.Baseline.X);
                float bxEnd = groupWords.Max(w => w.Baseline.X + EstimateWidth(w.Text, w.Baseline.FontSize));
                float by = groupWords[0].Baseline.Y;
                float pxStart = groupWords.Min(w => w.Publish.X);
                float pxEnd = groupWords.Max(w => w.Publish.X + EstimateWidth(w.Text, w.Publish.FontSize));

                float dx = seed.Delta.X;
                float dy = seed.Delta.Y;

                // Determine if all words on the same line share the same drift (whole-line shift)
                // vs partial-line (inline run shift — different X drift from other groups on same Y)
                string layoutCtx = seed.LayoutContext ?? "PARAGRAPH";
                string hint = BuildParagraphHint(dx, dy);
                if (layoutCtx == "TABLE_CELL")
                    hint = $"Table cell: see tableGroups for column/row width correction";

                var pg = new PhraseDiffGroup
                {
                    GroupId = groupId,
                    Page = seed.Page,
                    PhraseText = phraseText,
                    WordCount = groupWords.Count,
                    Issues = allIssues,
                    LayoutContext = layoutCtx,
                    BaselineRegion = new PhraseRegion
                    {
                        XStart = bxStart,
                        XEnd = bxEnd,
                        Y = by,
                        Font = NormFont(groupWords[0].Baseline?.Font),
                        FontSize = groupWords[0].Baseline?.FontSize ?? 0
                    },
                    PublishRegion = new PhraseRegion
                    {
                        XStart = pxStart,
                        XEnd = pxEnd,
                        Y = groupWords[0].Publish?.Y ?? by,
                        Font = NormFont(groupWords[0].Publish?.Font),
                        FontSize = groupWords[0].Publish?.FontSize ?? 0
                    },
                    SharedDelta = new DeltaValues { X = dx, Y = dy },
                    NegativeCorrection = BuildCorrection(dx, dy, hint),
                    WordIds = groupWords.Select(w => w.DifferenceId).ToList()
                };

                foreach (var w in groupWords) w.PhraseGroupId = groupId;

                // Populate DocxTableIndex: consensus Word table index for the phrase group.
                // If all words agree on the same tableIndex, record it; mixed-table groups get null.
                var tableIndices = groupWords
                    .Select(w => w.DocxContext?.TableIndex)
                    .Where(t => t.HasValue)
                    .Distinct()
                    .ToList();
                pg.DocxTableIndex = tableIndices.Count == 1 ? tableIndices[0] : (int?)null;

                groups.Add(pg);
                groupId++;
            }

            return groups;
        }

        /// <summary>
        /// Detects which phrase groups form a table grid.
        /// A table is identified when phrase groups share repeating X-column positions
        /// across multiple distinct Y-row positions (at least 2 rows × 2 columns).
        ///
        /// FIX: Also detects single-row tables (e.g. a row whose header row above it
        /// is a full-width merged/colspan cell that appears as only 1 phrase group).
        /// A single row with 3+ phrase groups at distinct X positions is treated as
        /// a valid single-row table. This prevents the PDF heuristic from silently
        /// missing tables that have merged header rows.
        /// </summary>
        private List<TableDiffGroup> DetectTableGroups(
            List<PhraseDiffGroup> phraseGroups,
            List<WordDiffEntry> wordDiffs)
        {
            var tables = new List<TableDiffGroup>();
            var claimed = new HashSet<int>(); // phraseGroup GroupIds already in a table
            int tableId = 1;

            // Group phrase groups by page first
            var byPage = phraseGroups
                .Where(pg => pg.Issues.Contains("POSITION_SHIFT"))
                .GroupBy(pg => pg.Page);

            foreach (var pageGroups in byPage)
            {
                int page = pageGroups.Key;
                var pgList = pageGroups
                    .OrderByDescending(pg => pg.BaselineRegion.Y)
                    .ThenBy(pg => pg.BaselineRegion.XStart)
                    .ToList();

                // Cluster rows: phrase groups within 6pt Y of each other = same row
                const float rowYTolerance = 6f;
                var rowClusters = new List<List<PhraseDiffGroup>>();

                foreach (var pg in pgList)
                {
                    bool placed = false;
                    foreach (var row in rowClusters)
                    {
                        if (Math.Abs(row[0].BaselineRegion.Y - pg.BaselineRegion.Y) <= rowYTolerance)
                        {
                            row.Add(pg);
                            placed = true;
                            break;
                        }
                    }
                    if (!placed) rowClusters.Add(new List<PhraseDiffGroup> { pg });
                }

                // Only rows with 2+ phrase groups are candidate table rows
                var candidateRows = rowClusters.Where(r => r.Count >= 2).ToList();

                // ── FIX: Allow single-row tables ──────────────────────────────────────
                // A table whose header row is a full-width merged cell renders as a
                // single phrase group, leaving only 1 "candidate" content row.
                // If that single row has 3+ phrase groups at distinct X positions,
                // treat it as a valid single-row table rather than skipping it.
                bool isSingleRowTable = candidateRows.Count == 1 && candidateRows[0].Count >= 3;

                if (candidateRows.Count < 2 && !isSingleRowTable)
                    continue; // need at least 2 multi-phrase rows, OR 1 row with 3+ phrases

                // ── Column alignment ──────────────────────────────────────────────────
                // Find XStart values that appear in 2+ rows (within 8pt tolerance).
                // For single-row tables, every column counts (minColOccurrences = 1).
                const float colXTolerance = 8f;

                var allXStarts = candidateRows
                    .SelectMany(r => r.Select(pg => pg.BaselineRegion.XStart))
                    .OrderBy(x => x)
                    .ToList();

                // Cluster XStarts into columns
                var colClusters = new List<List<float>>();
                foreach (var x in allXStarts)
                {
                    bool placed = false;
                    foreach (var col in colClusters)
                    {
                        if (Math.Abs(col.Average() - x) <= colXTolerance)
                        {
                            col.Add(x);
                            placed = true;
                            break;
                        }
                    }
                    if (!placed) colClusters.Add(new List<float> { x });
                }

                // For multi-row tables: columns must appear in 2+ rows.
                // For single-row tables: every distinct column position is valid.
                int minColOccurrences = isSingleRowTable ? 1 : 2;

                var validCols = colClusters
                    .Where(c => c.Count >= minColOccurrences)
                    .OrderBy(c => c.Average())
                    .ToList();

                if (validCols.Count < 2) continue;

                // ── Grid assignment ───────────────────────────────────────────────────
                var gridAssignment = new List<(PhraseDiffGroup pg, int row, int col)>();
                for (int ri = 0; ri < candidateRows.Count; ri++)
                {
                    var row = candidateRows[ri];
                    foreach (var pg in row)
                    {
                        int bestCol = -1;
                        float bestDist = float.MaxValue;
                        for (int ci = 0; ci < validCols.Count; ci++)
                        {
                            float colX = validCols[ci].Average();
                            float dist = Math.Abs(pg.BaselineRegion.XStart - colX);
                            if (dist <= colXTolerance && dist < bestDist)
                            {
                                bestDist = dist;
                                bestCol = ci;
                            }
                        }
                        if (bestCol >= 0) gridAssignment.Add((pg, ri, bestCol));
                    }
                }

                // For multi-row tables: need 4+ cells. For single-row: need 3+ cells.
                int minCells = isSingleRowTable ? 3 : 4;
                if (gridAssignment.Count < minCells) continue;

                // ── Column drifts ─────────────────────────────────────────────────────
                var columnDrifts = new List<TableColumnDrift>();
                for (int ci = 0; ci < validCols.Count; ci++)
                {
                    var colPgs = gridAssignment.Where(g => g.col == ci).ToList();
                    if (colPgs.Count == 0) continue;

                    float avgDx = colPgs.Average(g => g.pg.SharedDelta?.X ?? 0);
                    float baselineXStart = colPgs.Average(g => g.pg.BaselineRegion.XStart);

                    float prevColAvgDx = columnDrifts.Count > 0 ? columnDrifts[columnDrifts.Count - 1].AvgDeltaX : 0f;
                    float widthCorrection = -(avgDx - prevColAvgDx);

                    columnDrifts.Add(new TableColumnDrift
                    {
                        ColumnIndex = ci,
                        BaselineXStart = baselineXStart,
                        AvgDeltaX = avgDx,
                        WidthCorrectionTwips = (int)Math.Round(widthCorrection * 20),
                        WidthCorrectionPt = $"{widthCorrection:+0.##;-0.##;0}pt"
                    });
                }

                // ── Row drifts ────────────────────────────────────────────────────────
                var rowDrifts = new List<TableRowDrift>();
                for (int ri = 0; ri < candidateRows.Count; ri++)
                {
                    var rowPgs = gridAssignment.Where(g => g.row == ri).ToList();
                    if (rowPgs.Count == 0) continue;

                    float avgDy = rowPgs.Average(g => g.pg.SharedDelta?.Y ?? 0);
                    float baselineY = rowPgs.Average(g => g.pg.BaselineRegion.Y);
                    float prevAvgDy = rowDrifts.Count > 0 ? rowDrifts[rowDrifts.Count - 1].AvgDeltaY : 0f;
                    float heightCorr = -(avgDy - prevAvgDy);

                    rowDrifts.Add(new TableRowDrift
                    {
                        RowIndex = ri,
                        BaselineY = baselineY,
                        AvgDeltaY = avgDy,
                        HeightCorrectionTwips = (int)Math.Round(heightCorr * 20),
                        HeightCorrectionPt = $"{heightCorr:+0.##;-0.##;0}pt"
                    });
                }

                // ── Build bounds ──────────────────────────────────────────────────────
                var allPgs = gridAssignment.Select(g => g.pg).ToList();
                var claimedIds = allPgs.Select(p => p.GroupId).ToList();

                float bxStart = allPgs.Min(p => p.BaselineRegion.XStart);
                float bxEnd = allPgs.Max(p => p.BaselineRegion.XEnd);
                float byStart = allPgs.Min(p => p.BaselineRegion.Y);
                float pxStart = allPgs.Min(p => p.PublishRegion.XStart);
                float pxEnd = allPgs.Max(p => p.PublishRegion.XEnd);

                string tableHint = BuildTableHint(columnDrifts, rowDrifts);

                var tableGroup = new TableDiffGroup
                {
                    TableId = tableId++,
                    Page = page,
                    RowCount = candidateRows.Count,
                    ColumnCount = validCols.Count,
                    ColumnDrifts = columnDrifts,
                    RowDrifts = rowDrifts,
                    BaselineBounds = new PhraseRegion { XStart = bxStart, XEnd = bxEnd, Y = byStart },
                    PublishBounds = new PhraseRegion { XStart = pxStart, XEnd = pxEnd, Y = allPgs.Min(p => p.PublishRegion.Y) },
                    WordCorrectionHint = tableHint,
                    PhraseGroupIds = claimedIds
                };

                // ── Tag phrase groups and words ───────────────────────────────────────
                foreach (var (pg, ri, ci) in gridAssignment)
                {
                    if (claimed.Contains(pg.GroupId)) continue;
                    claimed.Add(pg.GroupId);

                    pg.TableGroupId = tableGroup.TableId;
                    pg.TableColumn = ci;
                    pg.TableRow = ri;
                    pg.LayoutContext = "TABLE_CELL";

                    if (pg.NegativeCorrection != null)
                        pg.NegativeCorrection.WordCorrectionHint =
                            $"Table {tableGroup.TableId}, Col {ci}, Row {ri}: see tableGroups for column/row corrections";
                }

                tables.Add(tableGroup);
            }

            return tables;
        }

        private static string BuildTableHint(List<TableColumnDrift> colDrifts, List<TableRowDrift> rowDrifts)
        {
            var sb = new System.Text.StringBuilder();
            sb.AppendLine("WORD TABLE CORRECTIONS:");

            foreach (var cd in colDrifts)
            {
                if (Math.Abs(cd.WidthCorrectionTwips) > 0)
                    sb.AppendLine($"  Column {cd.ColumnIndex}: adjust column width by {cd.WidthCorrectionPt} ({cd.WidthCorrectionTwips} twips)");
            }
            foreach (var rd in rowDrifts)
            {
                if (Math.Abs(rd.HeightCorrectionTwips) > 0)
                    sb.AppendLine($"  Row {rd.RowIndex}: adjust row height by {rd.HeightCorrectionPt} ({rd.HeightCorrectionTwips} twips)");
            }

            float firstColDx = colDrifts.FirstOrDefault()?.AvgDeltaX ?? 0f;
            if (Math.Abs(firstColDx) > 0.5f)
                sb.AppendLine($"  Table left margin: adjust by {-firstColDx:+0.##;-0.##;0}pt ({(int)Math.Round(-firstColDx * 20)} twips)");

            return sb.ToString().TrimEnd();
        }

        // ─── Back-fill table info into word entries ──────────────────────────────────────
        private static void BackFillTableInfoToWords(
            List<WordDiffEntry> wordDiffs,
            List<PhraseDiffGroup> phraseGroups,
            List<TableDiffGroup> tableGroups)
        {
            // Build lookup: phraseGroupId -> phraseGroup
            var pgLookup = phraseGroups.ToDictionary(pg => pg.GroupId);

            foreach (var w in wordDiffs)
            {
                if (w.PhraseGroupId == null) continue;
                if (!pgLookup.TryGetValue(w.PhraseGroupId.Value, out var pg)) continue;

                w.LayoutContext = pg.LayoutContext;
                w.TableGroupId = pg.TableGroupId;
                w.TableColumn = pg.TableColumn;
                w.TableRow = pg.TableRow;
            }
        }

        // ─── Detect inline runs: same Y, different X drift on same line ──────────────────
        /// <summary>
        /// On lines where multiple phrase groups exist at the same Y but different X deltas,
        /// those are inline runs (separate tab segments or text runs). Flag them INLINE_RUN.
        /// </summary>
        private static void TagInlineRuns(List<PhraseDiffGroup> phraseGroups)
        {
            const float yTol = 3f;

            var lineGroups = new List<List<PhraseDiffGroup>>();
            foreach (var pg in phraseGroups.OrderBy(p => p.Page).ThenByDescending(p => p.BaselineRegion.Y))
            {
                bool placed = false;
                foreach (var line in lineGroups)
                {
                    if (line[0].Page == pg.Page &&
                        Math.Abs(line[0].BaselineRegion.Y - pg.BaselineRegion.Y) <= yTol)
                    {
                        line.Add(pg);
                        placed = true;
                        break;
                    }
                }
                if (!placed) lineGroups.Add(new List<PhraseDiffGroup> { pg });
            }

            foreach (var line in lineGroups)
            {
                if (line.Count <= 1) continue;

                // If multiple phrase groups on same line have DIFFERENT X deltas → INLINE_RUN
                var distinctXDeltas = line
                    .Select(pg => (float)Math.Round(pg.SharedDelta?.X ?? 0, 1))
                    .Distinct()
                    .ToList();

                if (distinctXDeltas.Count > 1)
                {
                    foreach (var pg in line)
                        if (pg.LayoutContext != "TABLE_CELL")
                            pg.LayoutContext = "INLINE_RUN";
                }
            }
        }

        /// <summary>
        /// Generates a word-level JSON diff report capturing position shifts and font changes.
        /// When sourceTemplateDir is provided, each diff entry is enriched with precise
        /// DOCX element context (element type, formatting values, and specific correction targets)
        /// by matching PDF words against the corresponding Aspose.Words-parsed template.
        /// </summary>
        /// <summary>
        /// Returns the classified phrase groups so GenerateHtmlReport can use IsCleanDiff
        /// for green/red rectangle colouring. Returns null on error.
        /// </summary>
        public (List<PhraseDiffGroup> phraseGroups, HashSet<(int page, float x, float y)> cleanWordPositions) GenerateJsonDiffReport(
            string pdf1,
            string pdf2,
            List<TextInfo> text1,
            List<TextInfo> text2,
            string jsonOutputFolder,
            string sourceTemplateDir = null)
        {
            try
            {
                Directory.CreateDirectory(jsonOutputFolder);

                // ── Load DOCX element map if source template directory is provided ────
                List<DocxElementInfo> docxElements = null;
                string docxPath = null;

                if (!string.IsNullOrWhiteSpace(sourceTemplateDir) &&
                    Directory.Exists(sourceTemplateDir))
                {
                    string baseName = Path.GetFileNameWithoutExtension(pdf2);
                    foreach (var ext in new[] { ".docx", ".doc" })
                    {
                        string candidate = Path.Combine(sourceTemplateDir, baseName + ext);
                        if (File.Exists(candidate)) { docxPath = candidate; break; }
                    }

                    if (docxPath != null)
                    {
                        Console.WriteLine($"[GenerateJsonDiffReport] Parsing template: {docxPath}");
                        docxElements = ParseDocxElements(docxPath);
                        Console.WriteLine($"[GenerateJsonDiffReport] Parsed {docxElements.Count} elements.");
                    }
                    else
                    {
                        Console.WriteLine($"[GenerateJsonDiffReport] No .docx/.doc found for '{Path.GetFileNameWithoutExtension(pdf2)}' in {sourceTemplateDir}");
                    }
                }

                // ── Ignore regions ────────────────────────────────────────────────────
                var ignoreRegions = ComputeIgnoreRegions(text1);
                if (this.lstIgnoreAnnotationRegions != null)
                    ignoreRegions = ignoreRegions.Concat(this.lstIgnoreAnnotationRegions).ToList();

                bool IsInIgnoreRegion(TextInfo t) =>
                    ignoreRegions.Any(r =>
                        t.Page == r.Page &&
                        t.X >= r.XStart && t.X <= r.XEnd &&
                        Math.Abs(t.Y - r.Y) < 2);

                var validText1 = text1.Where(t => !string.IsNullOrWhiteSpace(t.Text) && !IsInIgnoreRegion(t)).ToList();
                var validText2 = text2.Where(t => !string.IsNullOrWhiteSpace(t.Text) && !IsInIgnoreRegion(t)).ToList();

                // ── NEW: Pre-compute per-page left-margin estimate for reflow detection ─
                const float REFLOW_MIN_LEFTWARD_DRIFT_PT = 12f;   // minimum |dx| to suspect reflow
                const float REFLOW_MAX_PUBLISH_FROM_MARGIN = 36f;   // publish must be within 36pt of margin
                const float REFLOW_MIN_BASELINE_OFFSET = 15f;   // baseline must be > 15pt right of margin

                var pageMinXPublish = validText2
                    .GroupBy(t => t.Page)
                    .ToDictionary(
                        g => g.Key,
                        g => g.Select(t => t.X).Where(x => x > 0f).DefaultIfEmpty(45f).Min());

                // ── Build line-bucket lookup for t2 ──────────────────────────────────
                const float LINE_Y_TOLERANCE = 5f;
                const float MATCH_Y_TOLERANCE = 5f;
                const float MATCH_X_TOLERANCE = 60f;

                var t2LineGroups = new Dictionary<(int page, int lineKey), List<TextInfo>>();
                foreach (var b in validText2)
                {
                    int lineKey = (int)(Math.Round(b.Y / LINE_Y_TOLERANCE) * LINE_Y_TOLERANCE);
                    var key = (b.Page, lineKey);
                    if (!t2LineGroups.TryGetValue(key, out var list))
                    { list = new List<TextInfo>(); t2LineGroups[key] = list; }
                    list.Add(b);
                }

                TextInfo FindBestT2Match(TextInfo a, HashSet<TextInfo> alreadyMatched)
                {
                    var bucketsToCheck = new HashSet<int>();
                    int baseBucket = (int)(Math.Round(a.Y / LINE_Y_TOLERANCE) * LINE_Y_TOLERANCE);
                    bucketsToCheck.Add(baseBucket);
                    bucketsToCheck.Add(baseBucket + (int)LINE_Y_TOLERANCE);
                    bucketsToCheck.Add(baseBucket - (int)LINE_Y_TOLERANCE);

                    TextInfo best = null;
                    float bestDist = float.MaxValue;

                    foreach (int bucket in bucketsToCheck)
                    {
                        if (!t2LineGroups.TryGetValue((a.Page, bucket), out var candidates)) continue;
                        foreach (var b in candidates)
                        {
                            if (alreadyMatched.Contains(b)) continue;
                            if (!string.Equals(a.Text, b.Text, StringComparison.Ordinal)) continue;
                            if (Math.Abs(a.Y - b.Y) > MATCH_Y_TOLERANCE) continue;
                            if (Math.Abs(a.X - b.X) > MATCH_X_TOLERANCE) continue;
                            float dist = Math.Abs(a.Y - b.Y) * 10 + Math.Abs(a.X - b.X);
                            if (dist < bestDist) { bestDist = dist; best = b; }
                        }
                    }
                    return best;
                }

                // ── Match PDF words and compute diffs ─────────────────────────────────
                var wordDiffs = new List<WordDiffEntry>();
                var matched = new HashSet<TextInfo>();
                var consumedDocxIds = new HashSet<int>();
                int diffId = 1;
                // Records baseline (page, X, Y) for every word that is intentionally
                // suppressed/clean so GenerateHtmlReport can classify rects even when
                // no phrase group exists (e.g., reflow words excluded from wordDiffs).
                var cleanWordPositions = new HashSet<(int page, float x, float y)>();

                var orderedText1 = validText1
                    .OrderBy(t => t.Page).ThenByDescending(t => t.Y).ThenBy(t => t.X)
                    .ToList();

                foreach (var a in orderedText1)
                {
                    var bestMatch = FindBestT2Match(a, matched);

                    if (bestMatch != null)
                    {
                        matched.Add(bestMatch);

                        float dx = bestMatch.X - a.X;
                        float dy = bestMatch.Y - a.Y;

                        bool hasPosDiff = Math.Abs(dx) > 0.5f || Math.Abs(dy) > 0.5f;

                        if (hasPosDiff && dx < -REFLOW_MIN_LEFTWARD_DRIFT_PT)
                        {
                            float pageMinX = pageMinXPublish.TryGetValue(a.Page, out float pmx)
                                             ? pmx : 45f;
                            bool publishAtLineStart = bestMatch.X <= pageMinX + REFLOW_MAX_PUBLISH_FROM_MARGIN;
                            bool baselineWasFurtherRight = a.X > pageMinX + REFLOW_MIN_BASELINE_OFFSET;

                            _logger.Log($"[REFLOW-CHECK] '{a.Text}' pg={a.Page} " +
                                $"bl.X={a.X:0.##} bl.Y={a.Y:0.##} pub.X={bestMatch.X:0.##} pub.Y={bestMatch.Y:0.##} " +
                                $"dx={dx:0.##} pageMinX={pageMinX:0.##} " +
                                $"pubAtStart={publishAtLineStart}(thresh<={pageMinX + REFLOW_MAX_PUBLISH_FROM_MARGIN:0.##}) " +
                                $"blFurtherRight={baselineWasFurtherRight}(bl.X>{pageMinX + REFLOW_MIN_BASELINE_OFFSET:0.##})");

                            if (publishAtLineStart && baselineWasFurtherRight)
                            {
                                Console.WriteLine(
                                    $"[GenerateJsonDiffReport] Reflow artefact suppressed: " +
                                    $"'{a.Text}' pg={a.Page} baseline.X={a.X:0.#} -> publish.X={bestMatch.X:0.#} " +
                                    $"(dx={dx:0.#}pt, pageMinX~{pageMinX:0.#}pt). " +
                                    $"Word wrapped to new line in V23 - not a correctable drift.");
                                hasPosDiff = false;
                                cleanWordPositions.Add((a.Page, a.X, a.Y));
                                _logger.Log($"[REFLOW-SUPPRESSED] cleanWordPositions << pg={a.Page} X={a.X:0.##} Y={a.Y:0.##}");
                            }
                            else
                            {
                                _logger.Log($"[REFLOW-FAILED] conditions not met, stays as POSITION_SHIFT");
                            }
                        }
                        else if (hasPosDiff)
                        {
                            _logger.Log($"[POS-NOREFLOW] '{a.Text}' pg={a.Page} bl.X={a.X:0.##} bl.Y={a.Y:0.##} dx={dx:0.##} dy={dy:0.##} -- dx not <-{REFLOW_MIN_LEFTWARD_DRIFT_PT}");
                        }

                        bool hasFontSzDiff = Math.Abs(a.FontSize - bestMatch.FontSize) >= 0.25f;
                        string fontA = NormFont(a.Font);
                        string fontB = NormFont(bestMatch.Font);
                        bool hasFontDiff = !string.IsNullOrEmpty(fontA) &&
                                           !string.IsNullOrEmpty(fontB) &&
                                           fontA != fontB;

                        if (!hasPosDiff && !hasFontSzDiff && !hasFontDiff) continue;

                        var issues = new List<string>();
                        if (hasPosDiff) issues.Add("POSITION_SHIFT");
                        if (hasFontSzDiff) issues.Add("FONT_SIZE_DIFF");
                        if (hasFontDiff) issues.Add("FONT_DIFF");

                        var entry = new WordDiffEntry
                        {
                            DifferenceId = diffId++,
                            PrimaryType = PrimaryDiffType(issues),
                            Issues = issues,
                            Text = a.Text,
                            Page = a.Page,
                            LayoutContext = "PARAGRAPH",
                            Baseline = new WordPosition
                            { X = a.X, Y = a.Y, Font = fontA, FontSize = a.FontSize, Page = a.Page },
                            Publish = new WordPosition
                            {
                                X = bestMatch.X,
                                Y = bestMatch.Y,
                                Font = fontB,
                                FontSize = bestMatch.FontSize,
                                Page = bestMatch.Page
                            },
                            Delta = new DeltaValues
                            {
                                X = dx,
                                Y = dy,
                                FontSizeDiff = bestMatch.FontSize - a.FontSize,
                                FontChanged = hasFontDiff
                            },
                            NegativeCorrection = BuildCorrection(dx, dy, BuildParagraphHint(dx, dy))
                        };

                        // Enrich with DOCX context
                        if (docxElements != null)
                        {
                            var docxElem = MatchWordToDocxElement(a.Text, a.X, docxElements, consumedDocxIds);
                            if (docxElem != null)
                            {
                                consumedDocxIds.Add(docxElem.ElementId);
                                entry.DocxContext = BuildDocxContext(entry, docxElem);
                                entry.LayoutContext = docxElem.ElementType;
                            }
                        }

                        wordDiffs.Add(entry);
                    }
                    else
                    {
                        // Genuinely missing from publish
                        var entry = new WordDiffEntry
                        {
                            DifferenceId = diffId++,
                            PrimaryType = "MISSING",
                            Issues = new List<string> { "MISSING" },
                            Text = a.Text,
                            Page = a.Page,
                            LayoutContext = "PARAGRAPH",
                            Baseline = new WordPosition
                            {
                                X = a.X,
                                Y = a.Y,
                                Font = NormFont(a.Font),
                                FontSize = a.FontSize,
                                Page = a.Page
                            }
                        };

                        if (docxElements != null)
                        {
                            var docxElem = MatchWordToDocxElement(a.Text, a.X, docxElements, consumedDocxIds);
                            if (docxElem != null)
                            {
                                consumedDocxIds.Add(docxElem.ElementId);
                                entry.DocxContext = BuildDocxContext(entry, docxElem);
                                entry.LayoutContext = docxElem.ElementType;
                            }
                        }

                        wordDiffs.Add(entry);
                    }
                }

                // Extra words in publish
                foreach (var b in validText2)
                {
                    if (matched.Contains(b)) continue;
                    wordDiffs.Add(new WordDiffEntry
                    {
                        DifferenceId = diffId++,
                        PrimaryType = "EXTRA",
                        Issues = new List<string> { "EXTRA" },
                        Text = b.Text,
                        Page = b.Page,
                        LayoutContext = "PARAGRAPH",
                        Publish = new WordPosition
                        {
                            X = b.X,
                            Y = b.Y,
                            Font = NormFont(b.Font),
                            FontSize = b.FontSize,
                            Page = b.Page
                        }
                    });
                }

                // ── Build phrase groups, detect inline runs and tables ────────────────
                var phraseGroups = BuildPhraseGroups(wordDiffs);
                TagInlineRuns(phraseGroups);
                var tableGroups = DetectTableGroups(phraseGroups, wordDiffs);
                BackFillTableInfoToWords(wordDiffs, phraseGroups, tableGroups);

                // ── Augment table groups and phrase groups using docxContext ──────────
                if (docxElements != null)
                    AugmentTableGroupsFromDocxContext(wordDiffs, phraseGroups, tableGroups);

                // ── Stamp IsCleanDiff on each phrase group ────────────────────────────
                // Mirrors the three safety gates in WordDriftCorrector that intentionally
                // skip certain diffs. These are drawn GREEN; everything else stays RED.
                //
                // Gate 1 – Reflow: word wrapped to a new line in V23.  Detected by:
                //   large leftward drift AND publish position is near the page left margin.
                //   Uses the same page-margin lookup built above for word matching.
                // Gate 2 – In-cell / layout-table paragraph:
                //   ApplyParagraphCorrection skips all paragraphs inside table cells.
                //   These appear as phrase groups with TableGroupId set AND PARAGRAPH layout.
                // Gate 3 – X-drift > 10pt, non-table paragraph:
                //   MAX_PARAGRAPH_X_DRIFT_PT safety cap in ApplyParagraphCorrection.
                const double MAX_PARA_X_PT = 10.0;   // matches WordDriftCorrector constant
                const float REFLOW_MIN_LEFT_JSON = 12f;   // same as REFLOW_MIN_LEFTWARD_DRIFT_PT
                const float REFLOW_MAX_MARGIN_JSON = 36f;  // same as REFLOW_MAX_PUBLISH_FROM_MARGIN

                foreach (var pg in phraseGroups)
                {
                    float dx = pg.SharedDelta?.X ?? 0f;
                    float dy = pg.SharedDelta?.Y ?? 0f;
                    float pubX = pg.PublishRegion?.XStart ?? float.MaxValue;
                    float pageMinX = pageMinXPublish.TryGetValue(pg.Page, out float pmx) ? pmx : 45f;

                    // Gate 1: Reflow detection
                    bool isReflow = dx < -REFLOW_MIN_LEFT_JSON &&
                                    pubX <= pageMinX + REFLOW_MAX_MARGIN_JSON;

                    // Gate 2 (partial — table-cell arm deferred to Phase 3 below):
                    // • INLINE_RUN: word was purged from its table group by the Word DOM check.
                    //   ApplyParagraphCorrection still skips it (it lives in a table cell).
                    //   Applied here unconditionally — INLINE_RUN groups never have significant
                    //   X drift from an uncorrected table, so they cannot be XI candidates.
                    //
                    // • tableGroupId set + TABLE_CELL/PARAGRAPH/LIST_ITEM/MERGE_FIELD:
                    //   Deferred to Phase 3 (after XI detection) so that XI detection can see
                    //   uncorrected-table phrase groups in BOTH original AND modified comparisons.
                    //   If applied here, Gate 2 would set IsCleanDiff=true on all table groups
                    //   before the XI check runs, causing !pg.IsCleanDiff in xDriftPGs to exclude
                    //   them → IsXIgnore never set → pass-2 in GenerateHtmlReport finds nothing.
                    bool isInCellPara = pg.LayoutContext == "INLINE_RUN";

                    // Gate 3: non-table phrase groups are ALWAYS clean in the MODIFIED comparison.
                    // Paragraphs, footers, and inline-runs that have no tableGroupId are never
                    // directly corrected by any PATH in WordDriftCorrector — only table-scoped content
                    // (tableGroupId != null) is touched.  So any dx/dy drift on a non-table phrase
                    // group in the modified comparison is a pre-existing V14/V23 rendering difference
                    // (including Y-drift overflow from a corrected table pushing content down the page).
                    // In the original comparison Gate 3 is disabled so these remain red (real findings).
                    bool isLargeDrift = IsModifiedComparison && !pg.TableGroupId.HasValue;

                    // Gate 4: Y-drift (vertical overflow reflow).
                    // ApplyParagraphCorrection NEVER adjusts Y positions — it only corrects X.
                    // If the shared Y-delta exceeds one half-line (~8pt), the paragraph has shifted
                    // vertically because upstream text overflowed onto a new line, pushing all
                    // subsequent paragraphs down. This is definitionally unapplied — always clean.
                    const float Y_REFLOW_THRESHOLD_PT = 8f;
                    bool isYReflow = Math.Abs(dy) > Y_REFLOW_THRESHOLD_PT;

                    pg.IsCleanDiff = isReflow || isInCellPara || isLargeDrift || isYReflow;

                    _logger.Log($"[ISCLEAN-STAMP] PhraseGroup={pg.GroupId} pg={pg.Page} " +
                        $"layoutCtx='{pg.LayoutContext}' tableGroupId={pg.TableGroupId?.ToString() ?? "null"} " +
                        $"dx={dx:0.##} dy={dy:0.##} pubX={pubX:0.##} pageMinX={(pageMinXPublish.TryGetValue(pg.Page, out float _pmxLog) ? _pmxLog : 45f):0.##} " +
                        $"isReflow={isReflow} isInCellPara={isInCellPara} isLargeDrift={isLargeDrift} isYReflow={isYReflow} " +
                        $"=> IsCleanDiff={pg.IsCleanDiff}");

                    // Also record baseline word positions for Gate 2 and Gate 3 phrase
                    // groups so IsPhraseGroupClean can match by position as a fallback.
                    if (pg.IsCleanDiff && pg.WordIds != null)
                    {
                        foreach (int wid in pg.WordIds)
                        {
                            var wd = wordDiffs.FirstOrDefault(w => w.DifferenceId == wid);
                            if (wd?.Baseline != null)
                                cleanWordPositions.Add((wd.Baseline.Page, wd.Baseline.X, wd.Baseline.Y));
                        }
                    }
                }

                // ── XI: X-Ignore reconciliation for phrase groups ─────────────────────
                //
                // After all IsCleanDiff gates are stamped, check whether leftmost-aligned
                // phrase groups with X drift have their baseline text already at the pivot
                // left X.  Such groups must not receive an X correction from WordDriftCorrector
                // (correcting them would shift text away from the correct position).
                // Mark them IsCleanDiff=true (XI) so the corrector skips them.
                //
                // Pivot X priority per page:
                //   A. Leftmost baseline words NOT covered by any phrase group → mode X.
                //   B. Leftmost phrase groups with negligible X drift → mode XStart.
                //   C. Majority XStart among X-drift phrase groups; ties → leftmost.
                {
                    const float XI_XDRIFT_THRESH = 2f;    // |dx| >= this = "has X drift"
                    const float XI_LEFTMOST_TOL  = 5f;    // within 5pt of page min-X = "leftmost"
                    const float XI_ALIGN_TOL     = 3f;    // within 3pt of pivot = "already aligned"
                    const float XI_FREE_WINDOW   = 10f;   // leftmost free-text window

                    // Build a lookup: baseline (page, X, Y) positions of X-DRIFT phrase group words only.
                    // Words inside non-X-drift (clean) phrase groups still have valid X positions and
                    // are legitimate pivot references — we must not exclude them from free text.
                    var xDriftPGWordPositions = new HashSet<(int page, float x, float y)>();
                    foreach (var pg in phraseGroups.Where(p =>
                        !p.IsCleanDiff &&
                        p.BaselineRegion != null &&
                        p.SharedDelta != null &&
                        Math.Abs(p.SharedDelta.X) >= XI_XDRIFT_THRESH))
                    {
                        if (pg.WordIds == null) continue;
                        foreach (int wid in pg.WordIds)
                        {
                            var wd = wordDiffs.FirstOrDefault(w => w.DifferenceId == wid);
                            if (wd?.Baseline != null)
                                xDriftPGWordPositions.Add((wd.Baseline.Page, wd.Baseline.X, wd.Baseline.Y));
                        }
                    }

                    foreach (int page in phraseGroups.Select(pg => pg.Page).Distinct().ToList())
                    {
                        var pagePGs = phraseGroups.Where(pg => pg.Page == page).ToList();

                        // X-drift phrase groups: not already clean AND significant |dx|
                        var xDriftPGs = pagePGs
                            .Where(pg => !pg.IsCleanDiff &&
                                         pg.BaselineRegion != null &&
                                         pg.SharedDelta != null &&
                                         Math.Abs(pg.SharedDelta.X) >= XI_XDRIFT_THRESH)
                            .ToList();
                        if (xDriftPGs.Count == 0) continue;

                        // Leftmost X-drift phrase groups
                        float minXDriftBx = xDriftPGs.Min(pg => pg.BaselineRegion.XStart);
                        var leftmostXDrift = xDriftPGs
                            .Where(pg => pg.BaselineRegion.XStart <= minXDriftBx + XI_LEFTMOST_TOL)
                            .ToList();
                        if (leftmostXDrift.Count == 0) continue;

                        // Option A: baseline words not covered by any X-drift phrase group on this page.
                        // Words in non-X-drift phrase groups have valid X positions and count as reference.
                        var freeTextX = validText1
                            .Where(t => t.Page == page && !xDriftPGWordPositions.Contains((t.Page, t.X, t.Y)))
                            .Select(t => t.X)
                            .ToList();

                        // Option B: phrase groups with negligible X drift (green baseline)
                        var greenPGX = pagePGs
                            .Where(pg => pg.BaselineRegion != null &&
                                         (pg.SharedDelta == null || Math.Abs(pg.SharedDelta.X) < XI_XDRIFT_THRESH))
                            .Select(pg => pg.BaselineRegion.XStart)
                            .ToList();

                        // Option C removed: if only XDiff rects exist as pivot source, do not apply XI.
                        // Those rects stay marked X and are corrected by the word drift corrector as normal.
                        float? pivotX = ComputePivotX(freeTextX, greenPGX, XI_FREE_WINDOW, 3f);
                        if (!pivotX.HasValue) continue;

                        _logger.Log($"[XI-GJDR] pg={page} pivotX={pivotX.Value:0.##} leftmostXDrift={leftmostXDrift.Count} freeText={freeTextX.Count} green={greenPGX.Count}");

                        // XI applies only when bx < pivot - tolerance (rect IS the leftmost element).
                        // When bx == pivot the rect is AT the reference margin and V23 drift is real -> keep X.
                        foreach (var pg in leftmostXDrift)
                        {
                            float bx = pg.BaselineRegion.XStart;
                            if (bx < pivotX.Value - XI_ALIGN_TOL)
                            {
                                pg.IsCleanDiff = true;
                                pg.IsXIgnore = true;
                                _logger.Log($"[XI-GJDR-CLEAN] pg={page} PhraseGroup={pg.GroupId} bx={bx:0.##} " +
                                    $"pivotX={pivotX.Value:0.##} leftOf={(pivotX.Value - bx):0.##}pt -> IsCleanDiff=true IsXIgnore=true (XI)");

                                // Record baseline positions so IsPhraseGroupClean position-fallback works
                                if (pg.WordIds != null)
                                {
                                    foreach (int wid in pg.WordIds)
                                    {
                                        var wd = wordDiffs.FirstOrDefault(w => w.DifferenceId == wid);
                                        if (wd?.Baseline != null)
                                            cleanWordPositions.Add((wd.Baseline.Page, wd.Baseline.X, wd.Baseline.Y));
                                    }
                                }
                            }
                            else
                            {
                                _logger.Log($"[XI-GJDR-KEEP] pg={page} PhraseGroup={pg.GroupId} bx={bx:0.##} " +
                                    $"pivotX={pivotX.Value:0.##} notLeftOf (need bx < pivot-{XI_ALIGN_TOL}pt={pivotX.Value - XI_ALIGN_TOL:0.##}) -> keep X");
                            }
                        }
                    }
                }

                // ── Phase 3: Apply Gate 2 table-cell to non-XI phrase groups ──────────
                // Now that XI detection has run, stamp the remaining table phrase groups
                // as IsCleanDiff=true. Skips any group already marked IsXIgnore so that
                // XI groups keep their IsXIgnore flag intact. Applies to BOTH original and
                // modified comparisons (PATH 3 already guards with !p.TableGroupId.HasValue
                // so these groups are never double-corrected via paragraph correction).
                foreach (var pg in phraseGroups)
                {
                    if (pg.IsXIgnore) continue;   // XI detection already handled this group
                    if (pg.IsCleanDiff) continue; // already clean from Gate 1 / 3 / 4 / INLINE_RUN
                    if (!pg.TableGroupId.HasValue) continue; // only table-affiliated groups
                    bool isTableCell = pg.LayoutContext == "TABLE_CELL" ||
                                       pg.LayoutContext == "PARAGRAPH"  ||
                                       pg.LayoutContext == "LIST_ITEM"  ||
                                       pg.LayoutContext == "MERGE_FIELD";
                    if (!isTableCell) continue;

                    pg.IsCleanDiff = true;
                    _logger.Log($"[GATE2-PHASE3] PhraseGroup={pg.GroupId} pg={pg.Page} layoutCtx='{pg.LayoutContext}' " +
                        $"tableGroupId={pg.TableGroupId} -> IsCleanDiff=true (Gate 2, Phase 3)");

                    // Record baseline positions for IsPhraseGroupClean position-fallback
                    if (pg.WordIds != null)
                    {
                        foreach (int wid in pg.WordIds)
                        {
                            var wd = wordDiffs.FirstOrDefault(w => w.DifferenceId == wid);
                            if (wd?.Baseline != null)
                                cleanWordPositions.Add((wd.Baseline.Page, wd.Baseline.X, wd.Baseline.Y));
                        }
                    }
                }

                // ── Assemble and serialise report ─────────────────────────────────────
                var report = new WordDiffReport
                {
                    ReportMetadata = new DiffReportMeta
                    {
                        GeneratedAt = DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ss"),
                        BaselinePdf = pdf1,
                        PublishPdf = pdf2,
                        BaselineFolder = Path.GetDirectoryName(pdf1) ?? "",
                        PublishFolder = Path.GetDirectoryName(pdf2) ?? ""
                    },
                    Summary = new DiffReportSummary
                    {
                        TotalWordDifferences = wordDiffs.Count,
                        PositionShifts = wordDiffs.Count(w => w.Issues.Contains("POSITION_SHIFT")),
                        FontDifferences = wordDiffs.Count(w => w.Issues.Contains("FONT_DIFF")),
                        FontSizeDifferences = wordDiffs.Count(w => w.Issues.Contains("FONT_SIZE_DIFF")),
                        MissingFromPublish = wordDiffs.Count(w => w.PrimaryType == "MISSING"),
                        ExtraInPublish = wordDiffs.Count(w => w.PrimaryType == "EXTRA"),
                        TotalPhraseGroups = phraseGroups.Count,
                        TotalTableGroups = tableGroups.Count,
                        IsClean = wordDiffs.Count == 0
                    },
                    TableGroups = tableGroups,
                    PhraseGroups = phraseGroups,
                    WordLevelDifferences = wordDiffs
                };

                var opts = new JsonSerializerOptions
                {
                    WriteIndented = true,
                    DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
                };

                string json = JsonSerializer.Serialize(report, opts);
                string jsonName = Path.GetFileNameWithoutExtension(pdf2) + "_diff.json";
                string jsonPath = Path.Combine(jsonOutputFolder, jsonName);
                File.WriteAllText(jsonPath, json);
                Console.WriteLine($"JSON diff report: {jsonPath}");
                _logger.Log($"[GJDR-DONE] cleanWordPositions.Count={cleanWordPositions.Count} phraseGroups.Count={phraseGroups.Count}");
                foreach (var cp in cleanWordPositions)
                    _logger.Log($"  [CLEAN-POS] pg={cp.page} X={cp.x:0.##} Y={cp.y:0.##}");
                return (phraseGroups, cleanWordPositions);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[GenerateJsonDiffReport] Error: {ex.Message}");
                return (null, null);
            }
        }


        #endregion

    }

    #region "WordLevelExtractionStrategy:v4"
    internal class WordLevelExtractionStrategy : IEventListener
    {
        public class WordInfo
        {
            public string Text { get; set; }
            public iText.Kernel.Geom.Rectangle BoundingBox { get; set; }
            public float FontSize { get; set; }
            public string FontName { get; set; }
            public bool IsRedTextInBracketWord { get; set; } = false;
        }

        public List<WordInfo> Words { get; } = new List<WordInfo>();

        private StringBuilder currentWord = new StringBuilder();
        private iText.Kernel.Geom.Rectangle currentRect = null;
        private string currentFont = null;
        private float currentFontSize = 0;
        private iText.Kernel.Geom.Vector lastEnd = null;

        private readonly bool ConsiderRedTextInBracketAsOneWord;
        private bool isInRedBracketedWord = false;
        private StringBuilder redBracketBuffer = new StringBuilder();
        private iText.Kernel.Geom.Rectangle redBracketRect = null;

        private static readonly char[] BracketStartChars = new[] { '(', '[', '<' };
        private static readonly char[] BracketEndChars = new[] { ')', ']', '>' };

        // New field to hold white mask rectangles covering white painted regions
        private readonly List<iText.Kernel.Geom.Rectangle> whiteMaskAreas;

        // Modified constructor: accept whiteMaskAreas list
        public WordLevelExtractionStrategy(bool considerRedTextInBracketAsOneWord = false, List<iText.Kernel.Geom.Rectangle> whiteMaskAreas = null)
        {
            ConsiderRedTextInBracketAsOneWord = considerRedTextInBracketAsOneWord;
            this.whiteMaskAreas = whiteMaskAreas ?? new List<iText.Kernel.Geom.Rectangle>();
        }

        public static iText.Kernel.Geom.Rectangle Union(iText.Kernel.Geom.Rectangle r1, iText.Kernel.Geom.Rectangle r2)
        {
            float x = Math.Min(r1.GetX(), r2.GetX());
            float y = Math.Min(r1.GetY(), r2.GetY());
            float right = Math.Max(r1.GetX() + r1.GetWidth(), r2.GetX() + r2.GetWidth());
            float top = Math.Max(r1.GetY() + r1.GetHeight(), r2.GetY() + r2.GetHeight());
            return new iText.Kernel.Geom.Rectangle(x, y, right - x, top - y);
        }

        // New helper method: check if word rect intersects any white mask rect
        private bool IsCoveredByWhiteMask(iText.Kernel.Geom.Rectangle wordRect)
        {
            foreach (var whiteRect in whiteMaskAreas)
            {
                if (Intersects(whiteRect, wordRect))
                    return true;
            }
            return false;
        }

        private bool Intersects(iText.Kernel.Geom.Rectangle r1, iText.Kernel.Geom.Rectangle r2)
        {
            return r1.GetX() < r2.GetX() + r2.GetWidth() &&
                   r1.GetX() + r1.GetWidth() > r2.GetX() &&
                   r1.GetY() < r2.GetY() + r2.GetHeight() &&
                   r1.GetY() + r1.GetHeight() > r2.GetY();
        }


        public void EventOccurred(IEventData data, EventType type)
        {
            if (type != EventType.RENDER_TEXT)
                return;

            var renderInfo = (TextRenderInfo)data;

            foreach (TextRenderInfo info in renderInfo.GetCharacterRenderInfos())
            {
                string c = info.GetText();
                if (string.IsNullOrEmpty(c))
                    continue;

                var bbox = Union(info.GetAscentLine().GetBoundingRectangle(), info.GetDescentLine().GetBoundingRectangle());


                // CHANGED: Skip characters inside white mask areas
                if (IsCoveredByWhiteMask(bbox))
                {
                    // If mid-word and encounter masked character, flush what we have
                    SaveCurrentWord();
                    // Skip adding this char to any new word
                    lastEnd = null;
                    continue;
                }


                if (char.IsWhiteSpace(c[0]) && !isInRedBracketedWord)
                {
                    SaveCurrentWord();
                    continue;
                }

                float fontSize = info.GetFontSize();
                float xSpacing = lastEnd != null ? info.GetBaseline().GetStartPoint().Get(0) - lastEnd.Get(0) : 0;
                float yDistance = lastEnd != null ? Math.Abs(info.GetBaseline().GetStartPoint().Get(1) - lastEnd.Get(1)) : 0;

                bool isNewLine = yDistance > fontSize * 0.5f;
                bool isNewWord = xSpacing > fontSize * 0.3f;

                if ((isNewLine || isNewWord) && !isInRedBracketedWord)
                {
                    SaveCurrentWord();
                }

                bool isRed = ConsiderRedTextInBracketAsOneWord && IsRedColor(info.GetFillColor());

                if (isRed && BracketStartChars.Contains(c[0]) && !isInRedBracketedWord)
                {
                    isInRedBracketedWord = true;
                    redBracketBuffer.Clear();
                    redBracketRect = bbox;
                }

                if (isInRedBracketedWord)
                {
                    redBracketBuffer.Append(c);
                    redBracketRect = redBracketRect == null ? bbox : Union(redBracketRect, bbox);

                    if (BracketEndChars.Contains(c[0]))
                    {
                        if (!IsCoveredByWhiteMask(redBracketRect)) // ⬅️ skip if masked
                        {
                            Words.Add(new WordInfo
                            {
                                Text = redBracketBuffer.ToString(),
                                BoundingBox = redBracketRect,
                                FontName = info.GetFont().GetFontProgram().ToString(),
                                FontSize = info.GetFontSize(),
                                IsRedTextInBracketWord = true
                            });
                        }

                        isInRedBracketedWord = false;
                        redBracketBuffer.Clear();
                        redBracketRect = null;
                        lastEnd = info.GetBaseline().GetEndPoint();
                        continue;
                    }

                    lastEnd = info.GetBaseline().GetEndPoint();
                    continue;
                }

                // No need to recheck white mask here anymore, handled above

                var tempWordRect = currentRect == null ? bbox : Union(currentRect, bbox);
                currentWord.Append(c);
                currentRect = tempWordRect;
                currentFont = info.GetFont().GetFontProgram().ToString();
                currentFontSize = fontSize;
                lastEnd = info.GetBaseline().GetEndPoint();
            }

        }

        private void SaveCurrentWord()
        {
            if (currentWord.Length > 0 && currentRect != null)
            {
                // Check if final word box intersects any white mask area
                if (!IsCoveredByWhiteMask(currentRect))
                {
                    Words.Add(new WordInfo
                    {
                        Text = currentWord.ToString(),
                        BoundingBox = currentRect,
                        FontName = currentFont,
                        FontSize = currentFontSize
                    });
                }

                // Reset buffers
                currentWord.Clear();
                currentRect = null;
                currentFont = null;
                currentFontSize = 0;
                lastEnd = null;
            }
        }

        private bool IsRedColor(iText.Kernel.Colors.Color color)
        {
            if (color is iText.Kernel.Colors.DeviceRgb rgb)
            {
                float[] rgbValues = rgb.GetColorValue();
                float red = rgbValues[0] * 255;
                float green = rgbValues[1] * 255;
                float blue = rgbValues[2] * 255;

                return red > 200 && green < 80 && blue < 80;
            }
            return false;
        }

        public ICollection<EventType> GetSupportedEvents()
        {
            return null; // Listen to all events
        }
    }

    #endregion

    #region "JSON Diff Report Related - Models & Code"

    /// <summary>
    /// Internal representation of a parsed DOCX content element.
    /// Built from Aspose.Words document analysis and used for matching PDF words
    /// to their source document context.
    /// </summary>
    internal class DocxElementInfo
    {
        public int ElementId { get; set; }

        /// <summary>PARAGRAPH | TABLE_CELL | LIST_ITEM | HEADER | FOOTER | IMAGE</summary>
        public string ElementType { get; set; }

        /// <summary>Combined plain text of this element (runs only, no field codes)</summary>
        public string PlainText { get; set; }

        /// <summary>Individual words extracted from PlainText for per-word matching</summary>
        public List<string> Words { get; set; } = new List<string>();

        /// <summary>
        /// Merge field names found in this element, e.g. ["WritingCompany", "Admitted"].
        /// PDF renders these as «WritingCompany», «Admitted».
        /// </summary>
        public List<string> MergeFieldNames { get; set; } = new List<string>();

        public string StyleName { get; set; }

        // ── Approximate X position in points ─────────────────────────────────────
        public float ApproximateXPt { get; set; }

        // ── Paragraph formatting ──────────────────────────────────────────────────
        public float LeftIndentPt { get; set; }
        public float SpaceBeforePt { get; set; }
        public float SpaceAfterPt { get; set; }
        public float LineSpacingPt { get; set; }

        // ── List/bullet ───────────────────────────────────────────────────────────
        public bool IsList { get; set; }
        public int? ListId { get; set; }
        public int? ListLevel { get; set; }
        public float? ListNumberPositionPt { get; set; }
        public float? ListTextPositionPt { get; set; }

        // ── Table location ────────────────────────────────────────────────────────
        public int? TableIndex { get; set; }
        public int? RowIndex { get; set; }
        public int? ColIndex { get; set; }
        public float? ColWidthPt { get; set; }
        public float? RowHeightPt { get; set; }
        public float? CellLeftPaddingPt { get; set; }
        public float? TableLeftIndentPt { get; set; }
        public float? ColumnOffsetPt { get; set; }

        // ── Image ─────────────────────────────────────────────────────────────────
        public bool? IsInlineImage { get; set; }
        public float? ImageWidthPt { get; set; }
        public float? ImageHeightPt { get; set; }
        public float? ImageAnchorXPt { get; set; }
        public float? ImageAnchorYPt { get; set; }

        // ── Location flags ────────────────────────────────────────────────────────
        public bool IsInHeader { get; set; }
        public bool IsInFooter { get; set; }
    }


    public class DocxContext
    {
        [JsonPropertyName("elementType")]
        public string ElementType { get; set; }

        [JsonPropertyName("styleName")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public string StyleName { get; set; }

        [JsonPropertyName("isInHeader")]
        public bool IsInHeader { get; set; }

        [JsonPropertyName("isInFooter")]
        public bool IsInFooter { get; set; }

        // ── List / bullet ─────────────────────────────────────────────────────────
        [JsonPropertyName("isList")]
        public bool IsList { get; set; }

        [JsonPropertyName("listId")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public int? ListId { get; set; }

        [JsonPropertyName("listLevel")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public int? ListLevel { get; set; }

        [JsonPropertyName("listNumberPositionPt")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public float? ListNumberPositionPt { get; set; }

        [JsonPropertyName("listTextPositionPt")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public float? ListTextPositionPt { get; set; }

        // ── Merge field ───────────────────────────────────────────────────────────
        [JsonPropertyName("mergeFieldName")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public string MergeFieldName { get; set; }

        // ── Paragraph formatting ──────────────────────────────────────────────────
        [JsonPropertyName("paragraphLeftIndentPt")]
        public float ParagraphLeftIndentPt { get; set; }

        [JsonPropertyName("spaceBeforePt")]
        public float SpaceBeforePt { get; set; }

        [JsonPropertyName("spaceAfterPt")]
        public float SpaceAfterPt { get; set; }

        [JsonPropertyName("lineSpacingPt")]
        public float LineSpacingPt { get; set; }

        // ── Table location ────────────────────────────────────────────────────────
        [JsonPropertyName("tableIndex")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public int? TableIndex { get; set; }

        [JsonPropertyName("rowIndex")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public int? RowIndex { get; set; }

        [JsonPropertyName("colIndex")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public int? ColIndex { get; set; }

        [JsonPropertyName("colWidthPt")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public float? ColWidthPt { get; set; }

        [JsonPropertyName("rowHeightPt")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public float? RowHeightPt { get; set; }

        [JsonPropertyName("cellLeftPaddingPt")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public float? CellLeftPaddingPt { get; set; }

        [JsonPropertyName("tableLeftIndentPt")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public float? TableLeftIndentPt { get; set; }

        [JsonPropertyName("columnOffsetPt")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public float? ColumnOffsetPt { get; set; }

        // ── Image ─────────────────────────────────────────────────────────────────
        [JsonPropertyName("isInlineImage")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public bool? IsInlineImage { get; set; }

        [JsonPropertyName("imageWidthPt")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public float? ImageWidthPt { get; set; }

        [JsonPropertyName("imageHeightPt")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public float? ImageHeightPt { get; set; }

        [JsonPropertyName("imageAnchorXPt")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public float? ImageAnchorXPt { get; set; }

        [JsonPropertyName("imageAnchorYPt")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public float? ImageAnchorYPt { get; set; }

        [JsonPropertyName("correctionTargets")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public List<DocxCorrectionTarget> CorrectionTargets { get; set; }
    }

    public class DocxCorrectionTarget
    {
        [JsonPropertyName("property")]
        public string Property { get; set; }

        [JsonPropertyName("axis")]
        public string Axis { get; set; }  // "X" | "Y"

        [JsonPropertyName("currentValuePt")]
        public float CurrentValuePt { get; set; }

        [JsonPropertyName("newValuePt")]
        public float NewValuePt { get; set; }

        [JsonPropertyName("currentValueTwips")]
        public int CurrentValueTwips { get; set; }

        [JsonPropertyName("newValueTwips")]
        public int NewValueTwips { get; set; }

        [JsonPropertyName("correctionEmu")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public long? CorrectionEmu { get; set; }

        [JsonPropertyName("description")]
        public string Description { get; set; }
    }
    public class WordDiffReport
    {
        [JsonPropertyName("reportMetadata")]
        public DiffReportMeta ReportMetadata { get; set; }

        [JsonPropertyName("summary")]
        public DiffReportSummary Summary { get; set; }

        [JsonPropertyName("tableGroups")]
        public List<TableDiffGroup> TableGroups { get; set; }

        [JsonPropertyName("phraseGroups")]
        public List<PhraseDiffGroup> PhraseGroups { get; set; }

        [JsonPropertyName("wordLevelDifferences")]
        public List<WordDiffEntry> WordLevelDifferences { get; set; }
    }

    public class DiffReportMeta
    {
        [JsonPropertyName("generatedAt")]
        public string GeneratedAt { get; set; }

        [JsonPropertyName("baselinePdf")]
        public string BaselinePdf { get; set; }

        [JsonPropertyName("publishPdf")]
        public string PublishPdf { get; set; }

        [JsonPropertyName("baselineFolder")]
        public string BaselineFolder { get; set; }

        [JsonPropertyName("publishFolder")]
        public string PublishFolder { get; set; }
    }

    public class DiffReportSummary
    {
        [JsonPropertyName("totalWordDifferences")]
        public int TotalWordDifferences { get; set; }

        [JsonPropertyName("positionShifts")]
        public int PositionShifts { get; set; }

        [JsonPropertyName("fontDifferences")]
        public int FontDifferences { get; set; }

        [JsonPropertyName("fontSizeDifferences")]
        public int FontSizeDifferences { get; set; }

        [JsonPropertyName("missingFromPublish")]
        public int MissingFromPublish { get; set; }

        [JsonPropertyName("extraInPublish")]
        public int ExtraInPublish { get; set; }

        [JsonPropertyName("totalPhraseGroups")]
        public int TotalPhraseGroups { get; set; }

        [JsonPropertyName("totalTableGroups")]
        public int TotalTableGroups { get; set; }

        [JsonPropertyName("isClean")]
        public bool IsClean { get; set; }
    }

    public class WordPosition
    {
        [JsonPropertyName("x")]
        public float X { get; set; }

        [JsonPropertyName("y")]
        public float Y { get; set; }

        [JsonPropertyName("font")]
        public string Font { get; set; }

        [JsonPropertyName("fontSize")]
        public float FontSize { get; set; }

        [JsonPropertyName("page")]
        public int Page { get; set; }
    }

    public class DeltaValues
    {
        [JsonPropertyName("x")]
        public float X { get; set; }

        [JsonPropertyName("y")]
        public float Y { get; set; }

        [JsonPropertyName("fontSizeDiff")]
        public float FontSizeDiff { get; set; }

        [JsonPropertyName("fontChanged")]
        public bool FontChanged { get; set; }
    }

    public class NegativeCorrection
    {
        [JsonPropertyName("x")]
        public float X { get; set; }

        [JsonPropertyName("y")]
        public float Y { get; set; }

        [JsonPropertyName("xTwips")]
        public int XTwips { get; set; }

        [JsonPropertyName("yTwips")]
        public int YTwips { get; set; }

        [JsonPropertyName("xPt")]
        public string XPt { get; set; }

        [JsonPropertyName("yPt")]
        public string YPt { get; set; }

        [JsonPropertyName("wordCorrectionHint")]
        public string WordCorrectionHint { get; set; }
    }

    public class WordDiffEntry
    {
        [JsonPropertyName("differenceId")]
        public int DifferenceId { get; set; }

        [JsonPropertyName("primaryType")]
        public string PrimaryType { get; set; }

        [JsonPropertyName("issues")]
        public List<string> Issues { get; set; }

        [JsonPropertyName("text")]
        public string Text { get; set; }

        [JsonPropertyName("page")]
        public int Page { get; set; }

        [JsonPropertyName("layoutContext")]
        public string LayoutContext { get; set; }

        [JsonPropertyName("baseline")]
        public WordPosition Baseline { get; set; }

        [JsonPropertyName("publish")]
        public WordPosition Publish { get; set; }

        [JsonPropertyName("delta")]
        public DeltaValues Delta { get; set; }

        [JsonPropertyName("negativeCorrection")]
        public NegativeCorrection NegativeCorrection { get; set; }

        [JsonPropertyName("phraseGroupId")]
        public int? PhraseGroupId { get; set; }

        [JsonPropertyName("tableGroupId")]
        public int? TableGroupId { get; set; }

        [JsonPropertyName("tableColumn")]
        public int? TableColumn { get; set; }

        [JsonPropertyName("tableRow")]
        public int? TableRow { get; set; }

        [JsonPropertyName("docxContext")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public DocxContext DocxContext { get; set; }
    }

    public class PhraseDiffGroup
    {
        [JsonPropertyName("groupId")]
        public int GroupId { get; set; }

        [JsonPropertyName("page")]
        public int Page { get; set; }

        [JsonPropertyName("phraseText")]
        public string PhraseText { get; set; }

        [JsonPropertyName("wordCount")]
        public int WordCount { get; set; }

        [JsonPropertyName("issues")]
        public List<string> Issues { get; set; }

        [JsonPropertyName("layoutContext")]
        public string LayoutContext { get; set; }

        [JsonPropertyName("baselineRegion")]
        public PhraseRegion BaselineRegion { get; set; }

        [JsonPropertyName("publishRegion")]
        public PhraseRegion PublishRegion { get; set; }

        [JsonPropertyName("sharedDelta")]
        public DeltaValues SharedDelta { get; set; }

        [JsonPropertyName("negativeCorrection")]
        public NegativeCorrection NegativeCorrection { get; set; }

        [JsonPropertyName("wordIds")]
        public List<int> WordIds { get; set; }

        [JsonPropertyName("tableGroupId")]
        public int? TableGroupId { get; set; }

        [JsonPropertyName("tableColumn")]
        public int? TableColumn { get; set; }

        [JsonPropertyName("tableRow")]
        public int? TableRow { get; set; }

        /// <summary>
        /// The Word document table index (1-based, from DocxContext.TableIndex) that all words
        /// in this phrase group belong to. Null when words come from different Word tables or
        /// when no DocxContext data is available.
        ///
        /// This is distinct from TableGroupId (which is a PDF-layer spatial cluster). Using this
        /// field in Pass 2 sibling propagation ensures XI propagation stays within the same Word
        /// table rather than bleeding into adjacent Word tables that share a PDF TableGroupId.
        /// </summary>
        [JsonPropertyName("docxTableIndex")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public int? DocxTableIndex { get; set; }

        /// <summary>
        /// True when this phrase group is intentionally not corrected by WordDriftCorrector
        /// (reflow, in-cell/layout-table paragraph, or X-drift > 10pt safety cap).
        /// False (default) means the corrector attempted or would attempt a fix → drawn in red.
        /// </summary>
        [JsonPropertyName("isCleanDiff")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
        public bool IsCleanDiff { get; set; }

        /// <summary>
        /// True ONLY when this phrase group was classified as XI (X-Ignore) — i.e. the leftmost
        /// X-drift group is already aligned at the pivot X and must NOT receive an X correction.
        /// Unlike IsCleanDiff (which covers Y-drift, reflow, in-cell etc.), IsXIgnore is exclusively
        /// set by the XI reconciliation pass and is what WordDriftCorrector uses to skip X correction.
        /// </summary>
        [JsonPropertyName("isXIgnore")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
        public bool IsXIgnore { get; set; }
    }

    public class PhraseRegion
    {
        [JsonPropertyName("xStart")]
        public float XStart { get; set; }

        [JsonPropertyName("xEnd")]
        public float XEnd { get; set; }

        [JsonPropertyName("y")]
        public float Y { get; set; }

        [JsonPropertyName("font")]
        public string Font { get; set; }

        [JsonPropertyName("fontSize")]
        public float FontSize { get; set; }
    }

    public class TableDiffGroup
    {
        [JsonPropertyName("correctionStrategy")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public string CorrectionStrategy { get; set; }

        [JsonPropertyName("tableId")]
        public int TableId { get; set; }

        [JsonPropertyName("page")]
        public int Page { get; set; }

        [JsonPropertyName("rowCount")]
        public int RowCount { get; set; }

        [JsonPropertyName("columnCount")]
        public int ColumnCount { get; set; }

        [JsonPropertyName("columnDrifts")]
        public List<TableColumnDrift> ColumnDrifts { get; set; }

        [JsonPropertyName("rowDrifts")]
        public List<TableRowDrift> RowDrifts { get; set; }

        [JsonPropertyName("baselineBounds")]
        public PhraseRegion BaselineBounds { get; set; }

        [JsonPropertyName("publishBounds")]
        public PhraseRegion PublishBounds { get; set; }

        [JsonPropertyName("wordCorrectionHint")]
        public string WordCorrectionHint { get; set; }

        [JsonPropertyName("phraseGroupIds")]
        public List<int> PhraseGroupIds { get; set; }
    }

    public class TableColumnDrift
    {
        [JsonPropertyName("columnIndex")]
        public int ColumnIndex { get; set; }

        [JsonPropertyName("baselineXStart")]
        public float BaselineXStart { get; set; }

        [JsonPropertyName("avgDeltaX")]
        public float AvgDeltaX { get; set; }

        [JsonPropertyName("widthCorrectionTwips")]
        public int WidthCorrectionTwips { get; set; }

        [JsonPropertyName("widthCorrectionPt")]
        public string WidthCorrectionPt { get; set; }
    }

    public class TableRowDrift
    {
        [JsonPropertyName("rowIndex")]
        public int RowIndex { get; set; }

        [JsonPropertyName("baselineY")]
        public float BaselineY { get; set; }

        [JsonPropertyName("avgDeltaY")]
        public float AvgDeltaY { get; set; }

        [JsonPropertyName("heightCorrectionTwips")]
        public int HeightCorrectionTwips { get; set; }

        [JsonPropertyName("heightCorrectionPt")]
        public string HeightCorrectionPt { get; set; }
    }

    #endregion

}