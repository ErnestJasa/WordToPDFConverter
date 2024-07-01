using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using System.Runtime.InteropServices;
using System.Xml.Linq;


namespace WordToPDFConverter
{
    internal static class Program
    {
        /// <summary>
        ///  The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            // To customize application configuration such as set high DPI settings or default font,
            // see https://aka.ms/applicationconfiguration.
            ApplicationConfiguration.Initialize();

            var fileContent = string.Empty;
            var filePath = string.Empty;

            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                bool isWindows = RuntimeInformation.IsOSPlatform(OSPlatform.Windows);
                if (isWindows)
                {

                string pcNameFull = System.Security.Principal.WindowsIdentity.GetCurrent().Name;
                string pcNameTrimmed = pcNameFull.Remove(0, pcNameFull.LastIndexOf(@"\") + 1);
                openFileDialog.InitialDirectory = @"c:\Users\" + pcNameTrimmed + @"\Desktop";
                }
                else
                {
                    openFileDialog.InitialDirectory = @"c:\";
                }
                openFileDialog.Filter = "Office files|*.docx;*.doc";
                openFileDialog.FilterIndex = 2;
                openFileDialog.RestoreDirectory = true;
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    filePath = openFileDialog.FileName;

                    var fileStream = openFileDialog.OpenFile();

                    using (StreamReader reader = new StreamReader(fileStream))
                    {
                        fileContent = reader.ReadToEnd();
                    }
                }
            }

            string outputPath = filePath.Remove(filePath.LastIndexOf(@"\") + 1);

            bool FileExists(string path) { 
                return File.Exists(path);
            }

            string NewFileName(string path) 
            {
                string fileNameNoExtention = filePath.Remove(filePath.LastIndexOf(@"."));
                string newFileName = fileNameNoExtention;

                int count = 1;
                bool exists = FileExists(fileNameNoExtention + ".pdf");
                while (exists)
                {
                    newFileName = fileNameNoExtention + "(" + count.ToString() + ")";
                    count++;
                    exists = FileExists(newFileName + ".pdf");
                }
                return Path.ChangeExtension(newFileName, ".pdf"); ;
            }

            string CreatePDF(string path, string exportDir)
            {
                Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();
                app.DisplayAlerts = WdAlertLevel.wdAlertsNone;
                app.Visible = false;

                var objPresSet = app.Documents;
                var objPres = objPresSet.Open(path, MsoTriState.msoTrue, MsoTriState.msoTrue, MsoTriState.msoFalse);
                


                var pdfFileName = NewFileName(path);
                var pdfPath = Path.Combine(exportDir, pdfFileName);

                try
                {
                    objPres.ExportAsFixedFormat(
                        pdfPath,
                        WdExportFormat.wdExportFormatPDF,
                        false,
                        WdExportOptimizeFor.wdExportOptimizeForPrint,
                        WdExportRange.wdExportAllDocument
                    );
                    Console.WriteLine("Successfully converted");
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error: ", ex.Message);
                    pdfPath = null;
                }
                finally
                {
                    objPres.Close();
                    app.Quit();
                }
                return pdfPath;
            }

            CreatePDF(filePath, outputPath);
            Console.ReadLine();
        }
    }
}