// See https://aka.ms/new-console-template for more information
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Layout;
using Aspose.Words.Tables;
//using DocumentConversion;
using DocumentProcessor;

//DriftCorrector.CustomLogger.Clear();
//DriftCorrector.CustomLogger.Log("--- STARTING DRIFT CORRECTOR ---");
//ConvertWordToPDFV23();
 string licenseKeyPath = @"C:\Personal\Project\DriftCorrector\Files\Key\Aspose.Words.NET.lic";
DocumentService.SetLicense(licenseKeyPath);
  
ComparePDF();
//ComparePDF_Raw();
//CompareOnly();

static void ComparePDFCore()
{
    
    

    Console.WriteLine("Modified Comparison Report Generated, Modified DriftJSON Generated !");
    
}


static void ComparePDF()
{
    //Some Changes
    //string licenseKeyPath = @"C:\Personal\Project\DriftCorrector\Files\Key\Aspose.Words.NET.lic";
    //DocumentService.SetLicense(licenseKeyPath);
    string logPath = @"C:\Personal\Project\DriftCorrector\Files\Logs\logs.txt";
    //Approach 3
    string report = @"C:\Personal\Project\DriftCorrector\Files\Reports\Original";
    string OriginalV14PDFFolderPath = @"C:\Personal\Project\DriftCorrector\Files\PDF\Original\V14";
    string OriginalV23PDFFolderPath = @"C:\Personal\Project\DriftCorrector\Files\PDF\Original\V23";

    string wordTemplateFolderDir = @"C:\Personal\Project\DriftCorrector\Files\WordTemplates\Original";
    var comparer = new PdfComparer(false,
                             false,
                             false,
                             false,
                             true, logPath
                             );

    comparer.ReportBasePath = report;
    comparer.ReportFileName = "Report.html";

    List<string> restrictedTextCheck = new List<string>();

    (List<PdfComparer.FileComparisonResult> ComparisonDetails, List<string> ComparisonLogs, bool hasDifference, bool isTextPresent) = comparer.CompareFolders(OriginalV14PDFFolderPath, OriginalV23PDFFolderPath, restrictedTextCheck, wordTemplateFolderDir, showXYWInDiffReport: true);
    //MessageBox.Show("Comparison Report Generated, DriftJSON Generated !");

    //Apply Drifts
    string jsonDiffFolder = @"C:\Personal\Project\DriftCorrector\Files\Reports\Original\JSONDiff";
    //string driftReportFolder = @"C:\Personal\Project\DriftCorrector\Files\Reports\Original\JSONDiff\DriftReport";
    string modfiiedTemplateFolder = @"C:\Personal\Project\DriftCorrector\Files\WordTemplates\Modified";
    var corrector = new WordDriftCorrector(logPath);
    var results = corrector.ApplyDriftCorrections(wordTemplateFolderDir, jsonDiffFolder, modfiiedTemplateFolder);
    //string reportPathJsonDrift = corrector.GenerateCorrectionReport(results, driftReportFolder);

    //MessageBox.Show($"Drift Applied ! Json Drift Report Path : {reportPathJsonDrift}");
    //Convert new Word to PDF 
    string modfiedPDFFolder = @"C:\Personal\Project\DriftCorrector\Files\PDF\Modified\V23";
    DocumentService.ConvertWordFolderToPDF(modfiiedTemplateFolder, modfiedPDFFolder);
    //MessageBox.Show("Modified Wordtemplate's PDF Generated !");
    //Generate Comparison report with Modified PDF
    string report2 = @"C:\Personal\Project\DriftCorrector\Files\Reports\Modified";
    var comparer2 = new PdfComparer(false,
                            false,
                             false,
                             false,
                             true, logPath
                             );

    comparer2.ReportBasePath = report2;
    comparer2.ReportFileName = "Report.html";
    comparer2.IsModifiedComparison = true;   // Gate 3 active: unapplied drifts show as green


    // Build prior-diff-rect lookup keyed by the FULL baseline path (pdf1).
    // Using the full path (not just filename) avoids collisions when identically-named
    // files appear in different subfolders. Both comparisons share the same baseline
    // folder (OriginalV14PDFFolderPath), so file1Path is identical across both runs.
    var priorDiffRects = ComparisonDetails
        .Where(r => r.DiffRectangles != null)
        .ToDictionary(
            r => r.PdfPath,
            r => r.DiffRectangles,
            StringComparer.OrdinalIgnoreCase);

    (ComparisonDetails, ComparisonLogs, hasDifference, isTextPresent) = comparer2.CompareFolders(OriginalV14PDFFolderPath, modfiedPDFFolder, restrictedTextCheck, wordTemplateFolderDir, priorDiffRects, showXYWInDiffReport: false);
    Console.WriteLine("Modified Comparison Report Generated, Modified DriftJSON Generated !");
    
}

static void ComparePDF_Raw()
{
    //Some Changes
    //string licenseKeyPath = @"C:\Personal\Project\DriftCorrector\Files\Key\Aspose.Words.NET.lic";
    //DocumentService.SetLicense(licenseKeyPath);
    string logPath = @"C:\Personal\Project\DriftCorrector\Files\Logs\logs.txt";
    //Approach 3
    string report = @"C:\Personal\Project\DriftCorrector\Files\Reports\Original";
    string OriginalV14PDFFolderPath = @"C:\Personal\Project\DriftCorrector\Files\PDF\Original\V14";
    string OriginalV23PDFFolderPath = @"C:\Personal\Project\DriftCorrector\Files\PDF\Original\V23";
    //string jsonDiffFolder = @"C:\Personal\Project\DriftCorrector\Files\Reports\Original\JSONDiff";
    string modfiiedTemplateFolder = @"C:\Personal\Project\DriftCorrector\Files\WordTemplates\Modified";

    //string wordTemplateFolderDir = @"C:\Personal\Project\DriftCorrector\Files\WordTemplates\Original";
    var comparer = new PdfComparer(false,
                             false,
                             false,
                             false,
                             true, logPath
                             );

    comparer.ReportBasePath = report;
    comparer.ReportFileName = "Report.html";

    List<string> restrictedTextCheck = new List<string>();

    (List<PdfComparer.FileComparisonResult> ComparisonDetails, List<string> ComparisonLogs, bool hasDifference, bool isTextPresent) = comparer.CompareFolders(OriginalV14PDFFolderPath, OriginalV23PDFFolderPath, restrictedTextCheck, showXYWInDiffReport: true);
    //MessageBox.Show("Comparison Report Generated, DriftJSON Generated !");

    //Apply Drifts
    //var corrector = new WordDriftCorrector(logPath);
    //var results = corrector.ApplyDriftCorrections(wordTemplateFolderDir, jsonDiffFolder, modfiiedTemplateFolder);
    //string reportPathJsonDrift = corrector.GenerateCorrectionReport(results, driftReportFolder);

    //MessageBox.Show($"Drift Applied ! Json Drift Report Path : {reportPathJsonDrift}");
    //Convert new Word to PDF 
    string modfiedPDFFolder = @"C:\Personal\Project\DriftCorrector\Files\PDF\Modified\V23";
    DocumentService.ConvertWordFolderToPDF(modfiiedTemplateFolder, modfiedPDFFolder);
    //MessageBox.Show("Modified Wordtemplate's PDF Generated !");
    //Generate Comparison report with Modified PDF
    string report2 = @"C:\Personal\Project\DriftCorrector\Files\Reports\Modified";
    var comparer2 = new PdfComparer(false,
                            false,
                             false,
                             false,
                             true, logPath
                             );

    comparer2.ReportBasePath = report2;
    comparer2.ReportFileName = "Report.html";
    comparer2.IsModifiedComparison = true;   // Gate 3 active: unapplied drifts show as green


    // Build prior-diff-rect lookup keyed by the FULL baseline path (pdf1).
    // Using the full path (not just filename) avoids collisions when identically-named
    // files appear in different subfolders. Both comparisons share the same baseline
    // folder (OriginalV14PDFFolderPath), so file1Path is identical across both runs.
    var priorDiffRects = ComparisonDetails
        .Where(r => r.DiffRectangles != null)
        .ToDictionary(
            r => r.PdfPath,
            r => r.DiffRectangles,
            StringComparer.OrdinalIgnoreCase);

    (ComparisonDetails, ComparisonLogs, hasDifference, isTextPresent) = comparer2.CompareFolders(OriginalV14PDFFolderPath, modfiedPDFFolder, restrictedTextCheck, null, priorDiffRects, showXYWInDiffReport: false);
    Console.WriteLine("Modified Comparison Report Generated, Modified DriftJSON Generated !");
   
}


static void ConvertWordToPDFV23()
{
    //string licenseKeyPath = @"C:\Personal\Project\DriftCorrector\Files\Key\Aspose.Words.NET.lic";
    //DocumentService.SetLicense(licenseKeyPath);
    string policyFormsFolder = @"C:\Personal\Project\DriftCorrector\Files\WordTemplates\Original";
    string V23PDFFolder = @"C:\Personal\Project\DriftCorrector\Files\PDF\Original\V23";
    DocumentService.ConvertWordFolderToPDF(policyFormsFolder, V23PDFFolder);


}

static void CompareOnly()
{
    
             string ReportPath = @"C:\Personal\Project\DriftCorrector\Files\CompareOnly\Comp-V23-P-V23-A";
         string logPath = @"C:\Personal\Project\DriftCorrector\Files\Logs\logs.txt";
         string V23PPDFFolderPath = @"C:\Personal\Project\DriftCorrector\Files\CompareOnly\V23PDF-P";
         string V23APDFFolderPath = @"C:\Personal\Project\DriftCorrector\Files\CompareOnly\V23PDF-A";
         string V14PDFFolderPath = @"C:\Personal\Project\DriftCorrector\Files\CompareOnly\V14PDF";
         //string originalWordTemplateFolderDir = Program["AppSettings:DocumentService:originalWordTemplateFolderDir"]; //@"F:\Work\Aspose\Templates\ToTest";
         List<string> restrictedTextCheck = new List<string>();

         var comparer = new PdfComparer(false,
                     false,
                      false,
                      false,
                      true, logPath
                      );

         comparer.ReportBasePath = ReportPath;
         comparer.ReportFileName = "Report.html";

         (List<PdfComparer.FileComparisonResult> ComparisonDetails, List<string> ComparisonLogs, bool hasDifference, bool isTextPresent) = comparer.CompareFolders(V23PPDFFolderPath, V23APDFFolderPath, restrictedTextCheck);
         //Second COmpare
         ReportPath = @"C:\Personal\Project\DriftCorrector\Files\CompareOnly\Comp-V14-V23-P";
         comparer = new PdfComparer(false,
                     false,
                      false,
                      false,
                      true, logPath
                      );

         comparer.ReportBasePath = ReportPath;
         comparer.ReportFileName = "Report.html";

         (ComparisonDetails, ComparisonLogs, hasDifference, isTextPresent) = comparer.CompareFolders(V14PDFFolderPath, V23PPDFFolderPath, restrictedTextCheck);
         //Third COmpare
         ReportPath = @"C:\Personal\Project\DriftCorrector\Files\CompareOnly\Comp-V14-V23-A";
         comparer = new PdfComparer(false,
                     false,
                      false,
                      false,
                      true, logPath
                      );

         comparer.ReportBasePath = ReportPath;
         comparer.ReportFileName = "Report.html";

         (ComparisonDetails, ComparisonLogs, hasDifference, isTextPresent) = comparer.CompareFolders(V14PDFFolderPath, V23APDFFolderPath, restrictedTextCheck);
        




}