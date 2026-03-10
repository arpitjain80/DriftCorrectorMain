//extern alias AsposeWord14;
using Aspose.Words;
using System;
using System.IO;
using System.Linq;
using System.Text;



namespace DocumentConversion
{
    /// <summary>
    /// Provides methods to convert Word documents (.doc/.docx) to PDF using Aspose.Words
    /// Supports both version 14 and version 23 of Aspose.Words
    /// </summary>
    public partial class AsposeOldService
    {
        #region Aspose.Words Version 14 Methods

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
                            Console.WriteLine($"Failed to set license: {ex.ToString()}");
                            throw;
                        }
                    }
                }
            }
        }


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
                //throw; // Re-throw if you want the calling loop to handle the error count
            }
        }


        /// <summary>
        /// Registers the code pages encoding provider required by Aspose.Words
        /// This must be called before any Aspose.Words operations in .NET Core/.NET 5+
        /// </summary>
        private static void EnsureEncodingProviderRegistered()
        {
            //if (!_encodingProviderRegistered)
            //{
            //    lock (_lock)
            //    {
            //        if (!_encodingProviderRegistered)
            //        {
            //            // Register the code pages encoding provider
            //            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            //            _encodingProviderRegistered = true;
            //            Console.WriteLine("Encoding provider registered successfully.");
            //        }
            //    }
            //}
        }




        /// <summary>
        /// Converts a single Word document to PDF using Aspose.Words version 14
        /// </summary>
        /// <param name="sourceFilePath">Full path to the source .doc or .docx file</param>
        /// <param name="destinationFolderPath">Destination folder path where PDF will be saved</param>
        /// <returns>Full path to the created PDF file</returns>
        public static string ConvertWordToPdf_V14(string sourceFilePath, string destinationFolderPath)
        {
            try
            {
                EnsureEncodingProviderRegistered();
                // Validate inputs
                if (string.IsNullOrWhiteSpace(sourceFilePath))
                    throw new ArgumentException("Source file path cannot be null or empty.", nameof(sourceFilePath));

                if (string.IsNullOrWhiteSpace(destinationFolderPath))
                    throw new ArgumentException("Destination folder path cannot be null or empty.", nameof(destinationFolderPath));

                if (!File.Exists(sourceFilePath))
                    throw new FileNotFoundException("Source file not found.", sourceFilePath);

                // Validate file extension
                string extension = Path.GetExtension(sourceFilePath).ToLower();
                if (extension != ".doc" && extension != ".docx")
                    throw new ArgumentException("Source file must be a .doc or .docx file.", nameof(sourceFilePath));

                // Create destination folder if it doesn't exist
                if (!Directory.Exists(destinationFolderPath))
                    Directory.CreateDirectory(destinationFolderPath);

                // Load the Word document
                Aspose.Words.Document doc = new Aspose.Words.Document(sourceFilePath);

                // Generate output PDF path
                string fileName = Path.GetFileNameWithoutExtension(sourceFilePath);
                string outputPath = Path.Combine(destinationFolderPath, fileName + ".pdf");

                // Save as PDF
                doc.Save(outputPath, Aspose.Words.SaveFormat.Pdf);

                Console.WriteLine($"Successfully converted: {Path.GetFileName(sourceFilePath)} -> {Path.GetFileName(outputPath)}");
                return outputPath;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error converting {sourceFilePath}: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Converts multiple Word documents from a source folder to PDF using Aspose.Words version 14
        /// Creates a subfolder at destination with same name as source folder
        /// </summary>
        /// <param name="sourceFolderPath">Source folder containing .doc or .docx files</param>
        /// <param name="destinationBasePath">Base destination path where a subfolder will be created</param>
        /// <returns>Path to the created subfolder containing all PDFs</returns>
        public static string ConvertFolderWordToPdf_V14(string sourceFolderPath, string destinationBasePath)
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
                string destinationFolderPath = Path.Combine(destinationBasePath, sourceFolderName);
                if (!Directory.Exists(destinationFolderPath))
                    Directory.CreateDirectory(destinationFolderPath);

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
                        Aspose.Words.Document doc = new Aspose.Words.Document(wordFile);

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
                        Console.WriteLine($"✗ Failed to convert {Path.GetFileName(wordFile)}: {ex.ToString()}");
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

        #endregion

    }
}