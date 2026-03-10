//extern alias AsposeWord23;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;
using Aspose.Words;
using System.Text;
using System.Xml.Linq;
using System.Diagnostics;
using System.Text.Json;
using static DocumentProcessor.PdfComparer;
using System.Text.RegularExpressions;


// Add reference to Aspose.Words DLL for your version
// For Version 14: Aspose.Words for .NET 14.x
// For Version 23: Aspose.Words for .NET 23.x

namespace DocumentProcessor
{

    public class PdfComparisonConfig
    {
        public string OriginalReportPath { get; set; }
        public string LogPath { get; set; }
        public string OriginalV14PDFFolderPath { get; set; }
        public string OriginalV23PDFFolderPath { get; set; }
        public string OriginalWordTemplateFolderDir { get; set; }
        public string OriginalJsonDiffFolder { get; set; }
        public string ModifiedWordTemplateFolder { get; set; }
        public string ModifiedV23PDFFolder { get; set; }
        public string ModifiedReportPath { get; set; }
    }
    public class FileState
    {
        public long Size { get; set; }
        public DateTime LastModified { get; set; }
    }



    public partial class DocumentService
    {
        public static void ApplyShifts(string templatePath, XDocument shiftsXml, string outputPath)
        {
            //Aspose.Words.License license = new Aspose.Words.License();
            //license.SetLicense(@"C:\Work organization\Aspose\Licence\Aspose.WordsProductFamily.lic");   // Looks in application folder

            Aspose.Words.Document doc = new Aspose.Words.Document(templatePath);

            foreach (var shift in shiftsXml.Descendants("textShift"))
            {
                string text = (string)shift.Attribute("text");
                double deltaX = (double)shift.Attribute("deltaX");
                double deltaY = (double)shift.Attribute("deltaY");

                // 1️⃣ Adjust Shapes (TextBoxes)
                foreach (Shape shape in doc.GetChildNodes(NodeType.Shape, true))
                {
                    if (shape.HasChildNodes && shape.GetText().Contains(text))
                    {
                        shape.Left += deltaX;          // OK
                        shape.Width -= deltaX;         // Shrink width instead of shifting text if needed
                        shape.WrapType = WrapType.None;
                    }
                }

                // 2️⃣ Adjust Paragraphs
                foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true))
                {
                    string paraText = para.GetText();
                    if (!string.IsNullOrEmpty(paraText) && paraText.Contains(text))
                    {
                        para.ParagraphFormat.LeftIndent = Math.Max(0, para.ParagraphFormat.LeftIndent + deltaX);
                        para.ParagraphFormat.SpaceBefore += deltaY;

                    }
                }

                // 3️⃣ Adjust Table Cells
                foreach (Table table in doc.GetChildNodes(NodeType.Table, true))
                {
                    foreach (Row row in table.Rows)
                    {
                        foreach (Cell cell in row.Cells)
                        {
                            if (cell.GetText().Contains(text))
                            {
                                cell.CellFormat.LeftPadding = Math.Max(0, cell.CellFormat.LeftPadding + deltaX);
                                cell.CellFormat.RightPadding = Math.Max(0, cell.CellFormat.RightPadding - deltaX);
                            }
                        }
                    }
                }
            }

            doc.Save(outputPath);
            Console.WriteLine($"DOCX template fixed and saved to: {outputPath}");
        }


     
        private static bool _encodingProviderRegistered = false;
        private static readonly object _lock = new object();

        private static bool _licenseSet = false;
        private static readonly object _licenseLock = new object();

        public static void DiagnoseAsposeVersion(string filePath)
        {
            try
            {
                // Use 'using' to ensure the file is closed and disposed of properly
                using (StreamWriter writer = new StreamWriter(filePath, append: false))
                {
                    writer.WriteLine("=== Aspose.Words Diagnostics ===");
                    writer.WriteLine($"Report Generated: {DateTime.Now}");

                    // Create a document to force assembly load
                    var doc = new Aspose.Words.Document();

                    // Get the actual loaded assembly
                    var asposeAssembly = typeof(Aspose.Words.Document).Assembly;

                    writer.WriteLine($"Assembly Name: {asposeAssembly.GetName().Name}");
                    writer.WriteLine($"Assembly Version: {asposeAssembly.GetName().Version}");
                    writer.WriteLine($"Assembly Location: {asposeAssembly.Location}");
                    writer.WriteLine($"Target Framework: {asposeAssembly.ImageRuntimeVersion}");

                    // Check custom attributes for target framework
                    var targetFrameworkAttr = asposeAssembly.GetCustomAttributes(
                        typeof(System.Runtime.Versioning.TargetFrameworkAttribute), false)
                        .FirstOrDefault() as System.Runtime.Versioning.TargetFrameworkAttribute;

                    if (targetFrameworkAttr != null)
                    {
                        writer.WriteLine($"Target Framework Attribute: {targetFrameworkAttr.FrameworkName}");
                    }

                    // Check if it references mscorlib (indicates .NET Framework)
                    var referencedAssemblies = asposeAssembly.GetReferencedAssemblies();
                    var hasMscorlib = referencedAssemblies.Any(a => a.Name == "mscorlib");
                    writer.WriteLine($"References mscorlib: {hasMscorlib} (true = .NET Framework, false = .NET Core)");

                    // Check for Remoting
                    var hasRemoting = referencedAssemblies.Any(a => a.Name.Contains("Remoting"));
                    writer.WriteLine($"References Remoting: {hasRemoting}");

                    writer.WriteLine("\nReferenced Assemblies:");
                    foreach (var refAssembly in referencedAssemblies.Take(10))
                    {
                        writer.WriteLine($"  - {refAssembly.Name} v{refAssembly.Version}");
                    }

                    writer.WriteLine("\n=== Your Runtime ===");
                    writer.WriteLine($"Runtime: {System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription}");
                    writer.WriteLine($"OS: {System.Runtime.InteropServices.RuntimeInformation.OSDescription}");
                }
            }
            catch (Exception ex)
            {
                // We log to Console here because if the file write fails, 
                // we need a fallback to see the error.
                Console.WriteLine($"Critical Diagnostic Error: {ex.Message}");
                File.WriteAllText(filePath + ".error.txt", ex.ToString());
            }
        }




        /// <summary>
        /// Sets the Aspose.Words license from a file path
        /// </summary>
        /// <param name="licensePath">Full path to the .lic file</param>
        public static void SetLicense(string licensePath)
        {
            EnsureEncodingProviderRegistered();
            if (!_licenseSet)
            {
                lock (_licenseLock)
                {
                    if (!_licenseSet)
                    {
                        try
                        {
                            Aspose.Words.License license = new Aspose.Words.License();
                            license.SetLicense(licensePath);
                            _licenseSet = true;
                            Console.WriteLine("Aspose.Words license applied successfully.");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Failed to set license: {ex.Message}");
                            throw;
                        }
                    }
                }
            }
        }


        /// <summary>
        /// Registers the code pages encoding provider required by Aspose.Words
        /// This must be called before any Aspose.Words operations in .NET Core/.NET 5+
        /// </summary>
        private static void EnsureEncodingProviderRegistered()
        {
            if (!_encodingProviderRegistered)
            {
                lock (_lock)
                {
                    if (!_encodingProviderRegistered)
                    {
                        // Register the code pages encoding provider
                        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
                        _encodingProviderRegistered = true;
                        Console.WriteLine("Encoding provider registered successfully.");
                    }
                }
            }
        }



        /// <summary>
        /// Converts a single Word document to PDF using Aspose.Words version 23
        /// </summary>
        /// <param name="sourceFilePath">Full path to the source .doc or .docx file</param>
        /// <param name="destinationFolderPath">Destination folder path where PDF will be saved</param>
        /// <returns>Full path to the created PDF file</returns>
        //public static string ConvertWordToPdf_V23(string sourceFilePath, string destinationFolderPath)
        //{
        //    try
        //    {
        //        EnsureEncodingProviderRegistered();
        //        // Validate inputs
        //        if (string.IsNullOrWhiteSpace(sourceFilePath))
        //            throw new ArgumentException("Source file path cannot be null or empty.", nameof(sourceFilePath));

        //        if (string.IsNullOrWhiteSpace(destinationFolderPath))
        //            throw new ArgumentException("Destination folder path cannot be null or empty.", nameof(destinationFolderPath));

        //        if (!File.Exists(sourceFilePath))
        //            throw new FileNotFoundException("Source file not found.", sourceFilePath);

        //        // Validate file extension
        //        string extension = Path.GetExtension(sourceFilePath).ToLower();
        //        if (extension != ".doc" && extension != ".docx")
        //            throw new ArgumentException("Source file must be a .doc or .docx file.", nameof(sourceFilePath));

        //        // Create destination folder if it doesn't exist
        //        if (!Directory.Exists(destinationFolderPath))
        //            Directory.CreateDirectory(destinationFolderPath);

        //        // Load the Word document
        //        var doc = new Aspose.Words.Document(sourceFilePath);

        //        // Generate output PDF path
        //        string fileName = Path.GetFileNameWithoutExtension(sourceFilePath);
        //        string outputPath = Path.Combine(destinationFolderPath, fileName + ".pdf");

        //        // Save as PDF
        //        // Note: In version 23, the API remains the same
        //        doc.Save(outputPath, Aspose.Words.SaveFormat.Pdf);

        //        Console.WriteLine($"Successfully converted: {Path.GetFileName(sourceFilePath)} -> {Path.GetFileName(outputPath)}");
        //        return outputPath;
        //    }
        //    catch (Exception ex)
        //    {
        //        Console.WriteLine($"Error converting {sourceFilePath}: {ex.Message}");
        //        throw;
        //    }
        //}

        //Converts source doc to destinatin PDF
        public static void ConvertDocToPdf(string sourceFilePath, string destinationPdfPath)
        {
            try
            {
                // 1. Load the Word document
                var doc = new Aspose.Words.Document(sourceFilePath);

                // 2. Ensure the destination directory exists
                string destDir = Path.GetDirectoryName(destinationPdfPath);
                if (!string.IsNullOrEmpty(destDir) && !Directory.Exists(destDir))
                {
                    Directory.CreateDirectory(destDir);
                }

                // 3. Save as PDF
                doc.Save(destinationPdfPath, SaveFormat.Pdf);

                Console.WriteLine($"✓ Successfully converted: {Path.GetFileName(sourceFilePath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ Failed to convert {Path.GetFileName(sourceFilePath)}: {ex.Message}");
                throw; // Re-throw if you want the calling loop to handle the error count
            }
        }

        public static void ConvertWordFolderToPDF_X(string WordFolderPath, string PDFFolderPath)
        {
            // Normalize paths (IMPORTANT)
            string sourceRoot = Path.GetFullPath(WordFolderPath).TrimEnd(Path.DirectorySeparatorChar);
            string destRoot = Path.GetFullPath(PDFFolderPath).TrimEnd(Path.DirectorySeparatorChar);
            var files = Directory.EnumerateFiles(sourceRoot, "*.*", SearchOption.AllDirectories)
                         .Where(f =>
                         {
                             string ext = Path.GetExtension(f);
                             return ext.Equals(".doc", StringComparison.OrdinalIgnoreCase) ||
                                    ext.Equals(".docx", StringComparison.OrdinalIgnoreCase);
                         });
            foreach (string sourceFile in files)
            {
                // Manually compute relative path (more reliable)
                string relativePath = sourceFile.Substring(sourceRoot.Length + 1);

                // Change extension to PDF
                relativePath = Path.ChangeExtension(relativePath, ".pdf");

                // Final destination path
                string destinationFile = Path.Combine(destRoot, relativePath);

                // Create directory structure if needed
                string destinationDir = Path.GetDirectoryName(destinationFile);
                Directory.CreateDirectory(destinationDir);

                // Convert
                ConvertDocToPdf(sourceFile, destinationFile);

                Console.WriteLine("Converted: " + relativePath);
            }

        }

        public static void ConvertWordFolderToPDF(string wordFolderPath, string pdfFolderPath, string jsonStateFilePath = null)
        {
            string sourceRoot = Path.GetFullPath(wordFolderPath);
            string destRoot = Path.GetFullPath(pdfFolderPath);

            // Create a safe root path with a trailing slash to safely scope our JSON cleanup later
            string sourceRootWithSlash = sourceRoot.TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar;

            bool useState = !string.IsNullOrWhiteSpace(jsonStateFilePath);

            Dictionary<string, FileState> fileStates = new Dictionary<string, FileState>(StringComparer.OrdinalIgnoreCase);
            bool stateChanged = false;

            if (useState && File.Exists(jsonStateFilePath))
            {
                try
                {
                    string json = File.ReadAllText(jsonStateFilePath);
                    fileStates = JsonSerializer.Deserialize<Dictionary<string, FileState>>(json) ?? fileStates;
                }
                catch (JsonException)
                {
                    fileStates = new Dictionary<string, FileState>(StringComparer.OrdinalIgnoreCase);
                }
            }

            // Track absolute paths for JSON cleanup, and relative paths for PDF mirroring
            HashSet<string> activeAbsoluteSourceFiles = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            HashSet<string> activeRelativePathsNoExt = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            var files = Directory.EnumerateFiles(sourceRoot, "*.*", SearchOption.AllDirectories)
                                 .Where(f =>
                                 {
                                     string ext = Path.GetExtension(f);
                                     return ext.Equals(".doc", StringComparison.OrdinalIgnoreCase) ||
                                            ext.Equals(".docx", StringComparison.OrdinalIgnoreCase);
                                 });

            foreach (string sourceFile in files)
            {
                // sourceFile is the absolute path, used as the unique JSON key
                activeAbsoluteSourceFiles.Add(sourceFile);

                string relativePath = Path.GetRelativePath(sourceRoot, sourceFile);
                activeRelativePathsNoExt.Add(Path.ChangeExtension(relativePath, null));

                FileInfo fileInfo = new FileInfo(sourceFile);
                long currentSize = fileInfo.Length;
                DateTime currentModified = fileInfo.LastWriteTimeUtc;

                bool requiresConversion = true;

                // Check state using the ABSOLUTE path
                if (useState && fileStates.TryGetValue(sourceFile, out FileState savedState))
                {
                    if (savedState.Size == currentSize && savedState.LastModified == currentModified)
                    {
                        requiresConversion = false;
                    }
                }

                if (requiresConversion)
                {
                    string destinationFile = Path.Combine(destRoot, Path.ChangeExtension(relativePath, ".pdf"));
                    string destinationDir = Path.GetDirectoryName(destinationFile);

                    if (!string.IsNullOrEmpty(destinationDir) && !Directory.Exists(destinationDir))
                    {
                        Directory.CreateDirectory(destinationDir);
                    }

                    ConvertDocToPdf(sourceFile, destinationFile);

                    if (useState)
                    {
                        // Save state using the ABSOLUTE path
                        fileStates[sourceFile] = new FileState { Size = currentSize, LastModified = currentModified };
                        stateChanged = true;
                    }
                }
            }

            // --- PHASE 2: ORPHANED PDF CLEANUP ---
            if (Directory.Exists(destRoot))
            {
                var existingPdfs = Directory.EnumerateFiles(destRoot, "*.pdf", SearchOption.AllDirectories);
                foreach (string pdf in existingPdfs)
                {
                    string relPdfPath = Path.GetRelativePath(destRoot, pdf);
                    string relPdfNoExt = Path.ChangeExtension(relPdfPath, null);

                    if (!activeRelativePathsNoExt.Contains(relPdfNoExt))
                    {
                        File.Delete(pdf);
                    }
                }
            }

            // --- PHASE 3: JSON STATE CLEANUP & SAVE ---
            if (useState)
            {
                // ONLY clean up keys that belong to the current source directory being processed
                var deletedKeys = fileStates.Keys.Where(k =>
                    k.StartsWith(sourceRootWithSlash, StringComparison.OrdinalIgnoreCase) &&
                    !activeAbsoluteSourceFiles.Contains(k)
                ).ToList();

                foreach (var key in deletedKeys)
                {
                    fileStates.Remove(key);
                    stateChanged = true;
                }

                if (stateChanged)
                {
                    string stateDir = Path.GetDirectoryName(jsonStateFilePath);
                    if (!string.IsNullOrEmpty(stateDir) && !Directory.Exists(stateDir))
                    {
                        Directory.CreateDirectory(stateDir);
                    }

                    string updatedJson = JsonSerializer.Serialize(fileStates, new JsonSerializerOptions { WriteIndented = true });
                    File.WriteAllText(jsonStateFilePath, updatedJson);
                }
            }
        }

        public static (List<FileComparisonResult>, List<string>, bool, bool) ExecuteDocumentComparisonWorkflow(PdfComparisonConfig config, bool applyDrifts = true, bool showXYZChar = true)
        {
            List<string> restrictedTextCheck = new List<string>();

            // 1. Initial Comparison (V14 vs V23)
            var comparer = new PdfComparer(false, false, false, false, true, config.LogPath);
            comparer.ReportBasePath = config.OriginalReportPath;
            comparer.ReportFileName = "Report.html";

            (List<PdfComparer.FileComparisonResult> ComparisonDetails, List<string> ComparisonLogs, bool hasDifference, bool isTextPresent) =
                comparer.CompareFolders(
                    config.OriginalV14PDFFolderPath,
                    config.OriginalV23PDFFolderPath,
                    restrictedTextCheck,
                    config.OriginalWordTemplateFolderDir,null, showXYZChar
                );

            if (applyDrifts)
            {
                // 2. Apply Word Drift Corrections
                var corrector = new WordDriftCorrector(config.LogPath);
                var results = corrector.ApplyDriftCorrections(
                    config.OriginalWordTemplateFolderDir,
                    config.OriginalJsonDiffFolder,
                    config.ModifiedWordTemplateFolder
                );

                // 3. Convert Modified Word Templates to PDF
                DocumentService.ConvertWordFolderToPDF(config.ModifiedWordTemplateFolder, config.ModifiedV23PDFFolder);
            }

            // 4. Generate Comparison Report with Modified PDF
            var comparer2 = new PdfComparer(false, false, false, false, true, config.LogPath);
            comparer2.ReportBasePath = config.ModifiedReportPath;
            comparer2.ReportFileName = "Report.html";
            comparer2.IsModifiedComparison = true;
            // Build prior-diff-rect 
            //If this is passed then the difference between prior and current can be sorted more accurately
            var priorDiffRects = ComparisonDetails
                .Where(r => r.DiffRectangles != null)
                .ToDictionary(
                    r => r.PdfPath,
                    r => r.DiffRectangles,
                    StringComparer.OrdinalIgnoreCase);

            // Final Comparison (Original V14 vs Modified V23)
            var finalResult = comparer2.CompareFolders(
                config.OriginalV14PDFFolderPath,
                config.ModifiedV23PDFFolder,
                restrictedTextCheck,
                config.OriginalWordTemplateFolderDir, priorDiffRects, showXYWInDiffReport: showXYZChar
            );

            return finalResult;
        }

        public static async Task<(int ExitCode, string Output)> ExecuteRobocopy(string source, string destination)
        {
            // 1. CRITICAL FIX: Remove trailing slashes so they don't escape the closing quotes
            source = source.TrimEnd('\\', '/');
            destination = destination.TrimEnd('\\', '/');

            // Quotes ensure paths with spaces don't break the command
            string arguments = $"\"{source}\" \"{destination}\" /E /Z /R:5 /W:5 /MT:16 /NP";

            ProcessStartInfo startInfo = new ProcessStartInfo
            {
                FileName = "robocopy.exe",
                Arguments = arguments,
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true // Good practice to capture this too
            };

            using (Process process = new Process { StartInfo = startInfo })
            {
                process.Start();

                // 2. CRITICAL FIX: Read the output asynchronously to prevent deadlocks 
                // and actually see what Robocopy is doing/failing on.
                Task<string> outputTask = process.StandardOutput.ReadToEndAsync();
                Task<string> errorTask = process.StandardError.ReadToEndAsync();

                await process.WaitForExitAsync();

                string output = await outputTask;
                string error = await errorTask;

                string fullLog = $"Output:\n{output}\nError:\n{error}";

                // Robocopy Exit Codes: 
                // 0-7 = Success (0 = no files copied, 1 = files copied, etc.)
                // 8+ = Failure

                return (process.ExitCode, fullLog);
            }
        }


        public static async Task RunExternalExeAsync(string exePath, string arguments)
        {
            string safePath = Path.GetFullPath(exePath);

            ProcessStartInfo startInfo = new ProcessStartInfo
            {
                FileName = safePath,
                Arguments = arguments,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,  // This captures the "thrown" error text
                CreateNoWindow = true,
                WorkingDirectory = Path.GetDirectoryName(safePath)
            };

            using (Process process = new Process { StartInfo = startInfo })
            {
                process.Start();

                // Start reading both streams simultaneously to avoid deadlocks
                Task<string> outputTask = process.StandardOutput.ReadToEndAsync();
                Task<string> errorTask = process.StandardError.ReadToEndAsync();

                await Task.Run(() => process.WaitForExit());

                string outputText = await outputTask;
                string errorText = await errorTask;

                if (process.ExitCode != 0 || !string.IsNullOrEmpty(errorText))
                {
                    // LOG THIS: This is where you'd write to your app's log file or database
                    string logMessage = $"EXE Failed! \nExit Code: {process.ExitCode}\nError: {errorText}";

                    // For example, using a simple local logger:
                    File.AppendAllText("app_log.txt", $"{DateTime.Now}: {logMessage}");

                    throw new Exception(logMessage);
                }

                // Optional: Log successful output too
                File.AppendAllText("app_log.txt", $"{DateTime.Now}: Success - {outputText}");
            }
        }

        private static string GetWin32LongPath(string path)
        {
            if (string.IsNullOrEmpty(path) || path.StartsWith(@"\\?\")) return path;
            if (path.StartsWith(@"\\")) return @"\\?\UNC\" + path.Substring(2);
            return @"\\?\" + path;
        }

        private static string RemoveLongPathPrefix(string path)
        {
            if (path.StartsWith(@"\\?\UNC\")) return @"\\" + path.Substring(8);
            if (path.StartsWith(@"\\?\")) return path.Substring(4);
            return path;
        }

        /// <summary>
        /// Converts multiple Word documents from a source folder to PDF using Aspose.Words version 23
        /// Creates a subfolder at destination with same name as source folder
        /// </summary>
        /// <param name="sourceFolderPath">Source folder containing .doc or .docx files</param>
        /// <param name="destinationBasePath">Base destination path where a subfolder will be created</param>
        /// <returns>Path to the created subfolder containing all PDFs</returns>
        public static string ConvertFolderWordToPdf_V23(string sourceFolderPath, string destinationBasePath)
        {
            try
            {
                EnsureEncodingProviderRegistered();
                // Validate inputs
                if (string.IsNullOrWhiteSpace(sourceFolderPath))
                    throw new ArgumentException("Source folder path cannot be null or empty.", nameof(sourceFolderPath));

                if (string.IsNullOrWhiteSpace(destinationBasePath))
                    throw new ArgumentException("Destination base path cannot be null or empty.", nameof(destinationBasePath));

                if (!Directory.Exists(sourceFolderPath))
                    throw new DirectoryNotFoundException($"Source folder not found: {sourceFolderPath}");

                // Get source folder name
                string sourceFolderName = new DirectoryInfo(sourceFolderPath).Name;

                // Create destination subfolder with same name as source folder
                string destinationFolderPath = destinationBasePath; // Path.Combine(destinationBasePath, sourceFolderName);
                //if (!Directory.Exists(destinationFolderPath))
                //    Directory.CreateDirectory(destinationFolderPath);

                // Get all .doc and .docx files
                var wordFiles = Directory.GetFiles(sourceFolderPath, "*.doc*")
                    .Where(f => f.EndsWith(".doc", StringComparison.OrdinalIgnoreCase) ||
                               f.EndsWith(".docx", StringComparison.OrdinalIgnoreCase))
                    .ToList();

                if (wordFiles.Count == 0)
                {
                    Console.WriteLine($"No .doc or .docx files found in {sourceFolderPath}");
                    return destinationFolderPath;
                }

                Console.WriteLine($"Found {wordFiles.Count} Word document(s) to convert...");

                int successCount = 0;
                int failCount = 0;

                foreach (string wordFile in wordFiles)
                {
                    try
                    {
                        // Load the Word document
                        var doc = new Aspose.Words.Document(wordFile);

                        // Generate output PDF path
                        string fileName = Path.GetFileNameWithoutExtension(wordFile);
                        string outputPath = Path.Combine(destinationFolderPath, fileName + ".pdf");

                        // Save as PDF
                        doc.Save(outputPath, Aspose.Words.SaveFormat.Pdf);

                        Console.WriteLine($"✓ Converted: {Path.GetFileName(wordFile)}");
                        successCount++;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"✗ Failed to convert {Path.GetFileName(wordFile)}: {ex.Message}");
                        failCount++;
                    }
                }

                Console.WriteLine($"\nConversion complete: {successCount} succeeded, {failCount} failed");
                Console.WriteLine($"Output folder: {destinationFolderPath}");

                return destinationFolderPath;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during batch conversion: {ex.Message}");
                throw;
            }
        }

        //Copies files passed in a string by separating the file by comma, semicolon or newline from source to destination
        //by preserving the hierachical structre of the source in destination
        //searches for the files recursively
        public static void CopyFiles(string sourceDir, string destDir, string fileListStr)
        {
            if(string.IsNullOrEmpty(fileListStr))
            {
                return;
            }
            // 1. Parse the input string into a HashSet for O(1) lookup performance
            var targetFileNames = new HashSet<string>(
                Regex.Split(fileListStr, @"[;,\n\r]+")
                     .Select(f => f.Trim())
                     .Where(f => !string.IsNullOrEmpty(f)),
                StringComparer.OrdinalIgnoreCase
            );

            if (targetFileNames.Count == 0) return;

            // 2. Enumerate EVERY file in the source directory recursively
            // This ensures that if 'data.log' exists in 5 different subfolders, we find all 5.
            var allSourceFiles = Directory.EnumerateFiles(sourceDir, "*.*", SearchOption.AllDirectories);

            int totalCopied = 0;

            foreach (string fullSrcPath in allSourceFiles)
            {
                string fileName = Path.GetFileName(fullSrcPath);

                // 3. If this specific file's name is in our target list, process it
                if (targetFileNames.Contains(fileName))
                {
                    // Determine the relative path (e.g., "SubFolder\Internal\file.txt")
                    string relativePath = Path.GetRelativePath(sourceDir, fullSrcPath);
                    string fullDestPath = Path.Combine(destDir, relativePath);

                    // 4. Create the directory branch in the destination if it doesn't exist
                    string destDirectory = Path.GetDirectoryName(fullDestPath);
                    if (!string.IsNullOrEmpty(destDirectory))
                    {
                        Directory.CreateDirectory(destDirectory);
                    }

                    // 5. Copy and Overwrite
                    // Even if the file exists in the destination, this replaces it with the source version
                    File.Copy(fullSrcPath, fullDestPath, overwrite: true);

                    Console.WriteLine($"Successfully copied: {relativePath}");
                    totalCopied++;
                }
            }

            Console.WriteLine($"\nProcess Complete. Total files copied/overwritten: {totalCopied}");
        }

    }
}