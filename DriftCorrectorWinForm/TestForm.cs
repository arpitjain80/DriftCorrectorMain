//using DocumentFormat.OpenXml.Bibliography;
//using DocumentProcessor;
using DocumentProcessor;
using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AsposeFormAdjustment
{
    public partial class TestForm : Form
    {
        private readonly IConfiguration Program;
        public TestForm(IConfiguration config)
        {
            InitializeComponent();
            Program = config;
        }

        private void btnConvertToV23PDF_Click(object sender, EventArgs e)
        {
            string policyFormsFolder = @"F:\Work\Aspose\Templates\ToTest\Set3";
            string V23PDFFolder = @"F:\Work\Aspose\TemplatesToPDF\V23\ToTest\Set3";

            DocumentService.ConvertWordFolderToPDF(policyFormsFolder, V23PDFFolder);

        }

        private void btnDiffNCompare_Click(object sender, EventArgs e)
        {
            //string originalReportPath = Program["AppSettings:DocumentService:originalReportPath"]; //@"F:\Work\Aspose\Reports\PDFComparison";
            //string logPath = Program["AppSettings:DocumentService:logPath"]; //@"F:\Work\Aspose\Logs";
            //string originalV14PDFFolderPath = Program["AppSettings:DocumentService:originalV14PDFFolderPath"];//@"F:\Work\Aspose\TemplatesToPDF\V14\ToTest";
            //string originalV23PDFFolderPath = Program["AppSettings:DocumentService:originalV23PDFFolderPath"]; //@"F:\Work\Aspose\TemplatesToPDF\V23\ToTest";
            //string originalWordTemplateFolderDir = Program["AppSettings:DocumentService:originalWordTemplateFolderDir"]; //@"F:\Work\Aspose\Templates\ToTest";
            //string originalJsonDiffFolder = Program["AppSettings:DocumentService:originalJsonDiffFolder"]; //@"F:\Work\Aspose\Reports\PDFComparison\JSONDiff";
            ////string originalDriftReportFolder = @"F:\Work\Aspose\Reports\PDFComparison\JSONDiff\DriftReport";
            //string modifiedWordTemplateFolder = Program["AppSettings:DocumentService:modifiedWordTemplateFolder"]; //@"F:\Work\Aspose\Templates\Modified";
            //string modfiedV23PDFFolder = Program["AppSettings:DocumentService:modfiedV23PDFFolder"]; //@"F:\Work\Aspose\TemplatesToPDF\Modified";
            //string modifiedReportPath = Program["AppSettings:DocumentService:modifiedReportPath"]; //@"F:\Work\Aspose\Reports\ModifiedPDFComparison";

            //var comparer = new PdfComparer(false,
            //                        false,
            //                         false,
            //                         false,
            //                         true, logPath
            //                         );

            //comparer.ReportBasePath = originalReportPath;
            //comparer.ReportFileName = "Report.html";

            //List<string> restrictedTextCheck = new List<string>();

            //(List<(string PdfPath, bool HasDifference, bool AllDiffsAreClean, bool isTextPresent, string ReportFileName)> ComparisonDetails, List<string> ComparisonLogs, bool hasDifference, bool isTextPresent) = comparer.CompareFolders(originalV14PDFFolderPath, originalV23PDFFolderPath, restrictedTextCheck, originalWordTemplateFolderDir);

            ////Apply Drifts
            //var corrector = new WordDriftCorrector(logPath);
            //var results = corrector.ApplyDriftCorrections(originalWordTemplateFolderDir, originalJsonDiffFolder, modifiedWordTemplateFolder);
            ////var results = corrector.ApplyDriftCorrections(hardenedWordTemplateFolderDir, jsonDiffFolder, modfiiedTemplateFolder);
            ////
            ////string reportPathJsonDrift = corrector.GenerateCorrectionReport(results, driftReportFolder);

            ////MessageBox.Show($"Drift Applied ! Json Drift Report Path : {reportPathJsonDrift}");
            ////Convert new Word to PDF 
            //DocumentService.ConvertWordFolderToPDF(modifiedWordTemplateFolder, modfiedV23PDFFolder);
            ////MessageBox.Show("Modified Wordtemplate's PDF Generated !");
            ////Generate Comparison report with Modified PDF
            //var comparer2 = new PdfComparer(false,
            //                        false,
            //                         false,
            //                         false,
            //                         true, logPath
            //                         );

            //comparer2.ReportBasePath = modifiedReportPath;
            //comparer2.ReportFileName = "Report.html";
            //comparer2.IsModifiedComparison = true;

            //(ComparisonDetails, ComparisonLogs, hasDifference, isTextPresent) = comparer2.CompareFolders(originalV14PDFFolderPath, modfiedV23PDFFolder, restrictedTextCheck, originalWordTemplateFolderDir);
            ////(ComparisonDetails, ComparisonLogs, hasDifference, isTextPresent) = comparer2.CompareFolders(txtAspone14.Text, modfiedPDFFolder, restrictedTextCheck, hardenedWordTemplateFolderDir);
            //MessageBox.Show("Modified Comparison Report Generated, Modified DriftJSON Generated !");


        }

        private void TestForm_Load(object sender, EventArgs e)
        {
            string AsposeLicenseKeyPath = Program["AppSettings:DocumentService:AsposeLicenseKeyPath"];
            //AsposeNewService.DiagnoseAsposeVersion("F:\\Work\\Aspose\\Diagnose23.txt");
            DocumentService.SetLicense(AsposeLicenseKeyPath);
            if (cmbCustomer.Items.Count > 0)
            {
                cmbCustomer.SelectedIndex = 0;
            }
        }

        private void btnOnlyCompare_Click(object sender, EventArgs e)
        {
            string ReportPath = "F:\\Work\\Aspose\\CompareOnly\\Comp-V23-P-V23-A";
            string logPath = Program["AppSettings:DocumentService:logPath"]; //@"F:\Work\Aspose\Logs";
            string V23PPDFFolderPath = @"F:\Work\Aspose\CompareOnly\V23PDF-P";
            string V23APDFFolderPath = @"F:\Work\Aspose\CompareOnly\V23PDF-A";
            string V14PDFFolderPath = @"F:\Work\Aspose\CompareOnly\V14PDF";
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
            ReportPath = @"F:\Work\Aspose\CompareOnly\Comp-V14-V23-P";
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
            ReportPath = @"F:\Work\Aspose\CompareOnly\Comp-V14-V23-A";
            comparer = new PdfComparer(false,
                        false,
                         false,
                         false,
                         true, logPath
                         );

            comparer.ReportBasePath = ReportPath;
            comparer.ReportFileName = "Report.html";
            // Build prior-diff-rect 
            //If this is passed then the difference between prior and current can be sorted more accurately
            var priorDiffRects = ComparisonDetails
                .Where(r => r.DiffRectangles != null)
                .ToDictionary(
                    r => r.PdfPath,
                    r => r.DiffRectangles,
                    StringComparer.OrdinalIgnoreCase);

            (ComparisonDetails, ComparisonLogs, hasDifference, isTextPresent) = comparer.CompareFolders(V14PDFFolderPath, V23APDFFolderPath, restrictedTextCheck, null, priorDiffRects, false);




            MessageBox.Show("Completed!");
        }

        private async void btnPolicyDocToV14PDF_Click(object sender, EventArgs e)
        {
            string customerName = "";
            if (cmbCustomer.SelectedIndex < 0)
            {
                MessageBox.Show("Please select a customer");
                return;
            }
            customerName = this.cmbCustomer.SelectedItem.ToString();
            btnPolicyDocToV14PDF.Enabled = false; // Prevent double-clicks

            try
            {

                string policyFormsWordToV14PDFConverterExePath = Program[$"AppSettings:DocumentService:WordToV14PDFConverterExePath"];
                string WorkDir = Program[$"AppSettings:DocumentService:Documents:Local:Customer:{customerName}:ConvertedDocsBaseFolderPath"];
                string policyFormsFolder = Program[$"AppSettings:DocumentService:Documents:Local:Customer:{customerName}:PolicyFormsBaseFolderPath"];
                string V14PDFFolder = Path.Combine(WorkDir, Program[$"AppSettings:DocumentService:Documents:Local:Customer:{customerName}:PolicyFormsPDFV14"]);
                string JsonProcessingStateFile = Path.Combine(WorkDir, Program[$"AppSettings:DocumentService:Documents:Local:Customer:{customerName}:ProcessingStateFilePathPDFV14"]);

                await DocumentService.RunExternalExeAsync(policyFormsWordToV14PDFConverterExePath, $"convert \"{policyFormsFolder}\" \"{V14PDFFolder}\" \"{JsonProcessingStateFile}\"");
                MessageBox.Show("Conversion Finished!");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.ToString()}");
            }
            finally
            {
                btnPolicyDocToV14PDF.Enabled = true;
            }
        }

        private void btnPolicyDocToV23PDF_Click(object sender, EventArgs e)
        {
            string customerName = "";
            if (cmbCustomer.SelectedIndex < 0)
            {
                MessageBox.Show("Please select a customer");
                return;
            }
            customerName = this.cmbCustomer.SelectedItem.ToString();

            string WorkDir = Program[$"AppSettings:DocumentService:Documents:Local:Customer:{customerName}:ConvertedDocsBaseFolderPath"];
            string policyFormsFolder = Program[$"AppSettings:DocumentService:Documents:Local:Customer:{customerName}:PolicyFormsBaseFolderPath"];
            string V23PDFFolder = Path.Combine(WorkDir, Program[$"AppSettings:DocumentService:Documents:Local:Customer:{customerName}:PolicyFormsPDFV23"]);
            string JsonProcessingStateFile = Path.Combine(WorkDir, Program[$"AppSettings:DocumentService:Documents:Local:Customer:{customerName}:ProcessingStateFilePathPDFV23"]);
            DocumentService.ConvertWordFolderToPDF(policyFormsFolder, V23PDFFolder, JsonProcessingStateFile);
            MessageBox.Show("Conversion Completed!");
        }

        private void btnRectifyPolicyDocs_Click(object sender, EventArgs e)
        {

            string customerName = "";
            if (cmbCustomer.SelectedIndex < 0)
            {
                MessageBox.Show("Please select a customer");
                return;
            }
            customerName = this.cmbCustomer.SelectedItem.ToString();

            string workDir = Program[$"AppSettings:DocumentService:Documents:Local:Customer:{customerName}:ConvertedDocsBaseFolderPath"];

            var myConfig = new PdfComparisonConfig
            {
                LogPath = Program["AppSettings:DocumentService:logPath"],
                OriginalWordTemplateFolderDir = Program[$"AppSettings:DocumentService:Documents:Local:Customer:{customerName}:PolicyFormsBaseFolderPath"],

                // Combining WorkDir with Relative Paths from Config
                OriginalReportPath = Path.Combine(workDir, Program[$"AppSettings:DocumentService:Documents:Local:Customer:{customerName}:ReportFolderPathOriginal"]),

                OriginalV14PDFFolderPath = Path.Combine(workDir, Program[$"AppSettings:DocumentService:Documents:Local:Customer:{customerName}:PolicyFormsPDFV14"]),

                OriginalV23PDFFolderPath = Path.Combine(workDir, Program[$"AppSettings:DocumentService:Documents:Local:Customer:{customerName}:PolicyFormsPDFV23"]),

                OriginalJsonDiffFolder = Path.Combine(workDir, Program[$"AppSettings:DocumentService:Documents:Local:Customer:{customerName}:JSONDiffFolderOriginal"]),

                ModifiedWordTemplateFolder = Path.Combine(workDir, Program[$"AppSettings:DocumentService:Documents:Local:Customer:{customerName}:PolicyFormsRectifiedFolderPath"]),

                ModifiedV23PDFFolder = Path.Combine(workDir, Program[$"AppSettings:DocumentService:Documents:Local:Customer:{customerName}:PolicyFormsPDFModified"]),

                ModifiedReportPath = Path.Combine(workDir, Program[$"AppSettings:DocumentService:Documents:Local:Customer:{customerName}:ReportFolderPathModified"])
            };

            DocumentService.ExecuteDocumentComparisonWorkflow(myConfig, true, true);
            MessageBox.Show("Completed!");
        }

        private async void btnCopyToDestination_Click(object sender, EventArgs e)
        {

            try
            {
                string customerName = "";
                if (cmbCustomer.SelectedIndex < 0)
                {
                    MessageBox.Show("Please select a customer");
                    return;
                }
                btnCopyToDestination.Enabled = false;
                customerName = this.cmbCustomer.SelectedItem.ToString();

                string localWorkDir = Program[$"AppSettings:DocumentService:Documents:Local:Customer:{customerName}:ConvertedDocsBaseFolderPath"];
                string DestinationFolder = Program[$"AppSettings:DocumentService:Documents:Local:Customer:{customerName}:FinalDestinationFolderPath"];
                //
                var result = await DocumentService.ExecuteRobocopy(localWorkDir, DestinationFolder);
                //MessageBox.Show($"Exit Code: {result.ExitCode} | Robocopy Log: {result.Output}");
                MessageBox.Show("Completed!");

            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error : {ex.ToString()}");
            }
            finally
            {
                btnCopyToDestination.Enabled = true;
            }

        }

        private void btnCopyPolicyTemplates_Click(object sender, EventArgs e)
        {
            try
            {
                string customerName = "";
                if (cmbCustomer.SelectedIndex < 0)
                {
                    MessageBox.Show("Please select a customer");
                    return;
                }
                btnCopyPolicyTemplates.Enabled = false;
                customerName = this.cmbCustomer.SelectedItem.ToString();

                string localWorkDir = Program[$"AppSettings:DocumentService:Documents:Local:Customer:{customerName}:PolicyFormsRepositoryFolderPath"];
                string DestinationFolder = Program[$"AppSettings:DocumentService:Documents:Local:Customer:{customerName}:PolicyFormsBaseFolderPath"];
                
                DocumentService.CopyFiles(localWorkDir, DestinationFolder, txtPolicyTemplateList.Text);
                //MessageBox.Show($"Exit Code: {result.ExitCode} | Robocopy Log: {result.Output}");
                MessageBox.Show("Completed!");

            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error : {ex.ToString()}");
            }
            finally
            {
                btnCopyPolicyTemplates.Enabled = true;
            }
        }
    }
}
