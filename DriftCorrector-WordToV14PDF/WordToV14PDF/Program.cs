using DocumentConversion;
using System;
using System.IO;

namespace AsposeOldConsole
{
    internal class Program
    {
        // Restored internal license path
        private const string AsposeLicenseKeyPath = @"C:\Personal\Project\DriftCorrector\Files\Key\Aspose.Words.NET.lic";
        
        static void Main(string[] args)
        {
            if (args.Length == 0)
            {
                ShowUsage();
                Console.WriteLine("Running defaul conversion...");
                string wordFolderPath = @"C:\Personal\Project\DriftCorrector\Files\WordTemplates\Original";
                string V14PDFFolderPath = @"C:\Personal\Project\DriftCorrector\Files\PDF\Original\V14";
                RunConversion(wordFolderPath, V14PDFFolderPath);

                return;
            }

            string command = args[0].ToLower();

            try
            {
                switch (command)
                {
                    case "copy":
                        // Expected: copy [sourceDir] [destDir] [fileListPath]
                        if (args.Length < 4) { Console.WriteLine("Usage: copy <source> <dest> <fileListPath>"); return; }
                        string listContent = File.ReadAllText(args[3]);
                        CopyFilesFromList(args[1], args[2], listContent);
                        break;

                    case "convert":
                        // Expected: convert [sourceRoot] [destRoot]
                        if (args.Length < 3) { Console.WriteLine("Usage: convert <source> <dest>"); return; }
                        RunConversion(args[1], args[2]);
                        break;

                    default:
                        Console.WriteLine($"Unknown command: {command}");
                        ShowUsage();
                        break;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Critical Error: {ex.Message}");
            }
        }

        private static void RunConversion(string sourceRoot, string destRoot)
        {
            Console.WriteLine("Initializing Aspose License internally...");
            AsposeOldService.SetLicense(AsposeLicenseKeyPath);

            sourceRoot = Path.GetFullPath(sourceRoot).TrimEnd(Path.DirectorySeparatorChar);
            destRoot = Path.GetFullPath(destRoot).TrimEnd(Path.DirectorySeparatorChar);

            var files = Directory.EnumerateFiles(sourceRoot, "*.doc*", SearchOption.AllDirectories);

            foreach (string sourceFile in files)
            {
                string relativePath = sourceFile.Substring(sourceRoot.Length + 1);
                string destinationFile = Path.Combine(destRoot, Path.ChangeExtension(relativePath, ".pdf"));

                string destinationDir = Path.GetDirectoryName(destinationFile);
                if (!Directory.Exists(destinationDir)) Directory.CreateDirectory(destinationDir);

                AsposeOldService.ConvertDocToPdf(sourceFile, destinationFile);
                Console.WriteLine($"Converted: {relativePath}");
            }
            Console.WriteLine("Conversion Task Complete.");
        }

        public static void CopyFilesFromList(string sourceDir, string destDir, string fileList)
        {
            if (!Directory.Exists(sourceDir)) throw new DirectoryNotFoundException($"Source missing: {sourceDir}");
            if (!Directory.Exists(destDir)) Directory.CreateDirectory(destDir);

            string[] files = fileList.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.RemoveEmptyEntries);

            foreach (string fileName in files)
            {
                string cleanFileName = fileName.Trim();
                if (string.IsNullOrEmpty(cleanFileName)) continue;

                string sourcePath = Path.Combine(sourceDir, cleanFileName);
                string destPath = Path.Combine(destDir, cleanFileName);

                if (File.Exists(sourcePath))
                {
                    File.Copy(sourcePath, destPath, true);
                    Console.WriteLine($"Copied: {cleanFileName}");
                }
                else
                {
                    Console.WriteLine($"Not Found: {cleanFileName}");
                }
            }
        }

        private static void ShowUsage()
        {
            Console.WriteLine("\n--- Aspose Utility Usage ---");
            Console.WriteLine("1. Copy Files:");
            Console.WriteLine("   AsposeOldConsole.exe copy \"F:\\Source\" \"F:\\Dest\" \"C:\\list.txt\"");
            Console.WriteLine("\n2. Convert Folder:");
            Console.WriteLine("   AsposeOldConsole.exe convert \"F:\\Work\\Templates\" \"F:\\Work\\Output\"");
            Console.WriteLine("-----------Press any key to continue-----------------\n");
            Console.ReadLine();
        }
    }
}