using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using System.Diagnostics;

namespace DocToDocxConverter
{
    public class Convert : IDisposable
    {
        private Word.Application wordApp;
        private Excel.Application excelApp;
        private PowerPoint.Application pptApp;
        public Convert()
        {
            wordApp = new Word.Application();
            wordApp.Visible = true;

            excelApp = new Excel.Application();
            excelApp.Visible = true;
            excelApp.DisplayAlerts = false; // Save Excel file without asking to overwrite it

            pptApp = new PowerPoint.Application();
            pptApp.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;
        }
        ~Convert()
        {
            this.Dispose();
        }
        public void ConvertFile(string fileFullPath)
        {
            if (!File.Exists(fileFullPath)) throw new Exception($"File {fileFullPath} does not exist.");
            fileFullPath = Path.GetFullPath(fileFullPath);
            var ext = Path.GetExtension(fileFullPath) ?? "";
            string folder = Path.GetDirectoryName(fileFullPath) ?? "";
            string name = Path.GetFileNameWithoutExtension(fileFullPath) ?? "";
            if (string.IsNullOrEmpty(folder) || string.IsNullOrEmpty(name)) throw new Exception("Assert failed.");

            switch (ext)
            {
                case ".doc":
                    {
                        string newName = Path.Combine(folder, name + ".docx");
                        if (File.Exists(newName)) RecycleBin.DeleteFile(newName);
                        var doc = wordApp.Documents.Open(fileFullPath, ReadOnly: true);
                        doc.SaveAs2(newName, Word.WdSaveFormat.wdFormatXMLDocument, CompatibilityMode: Word.WdCompatibilityMode.wdCurrent);
                        doc.Close();
                        break;
                    }
                case ".xls":
                    {
                        string newName = Path.Combine(folder, name + ".xlsx");
                        if (File.Exists(newName)) RecycleBin.DeleteFile(newName);
                        var xls = excelApp.Workbooks.Open(fileFullPath, ReadOnly: true);
                        xls.SaveAs(newName, Excel.XlFileFormat.xlWorkbookDefault);
                        xls.Close();
                        break;
                    }
                case ".ppt":
                    {
                        string newName = Path.Combine(folder, name + ".pptx");
                        if (File.Exists(newName)) RecycleBin.DeleteFile(newName);
                        // https://forum.uipath.com/t/powerpoint-com-interop/233326/5
                        var ppt = pptApp.Presentations.Open(fileFullPath, ReadOnly: Microsoft.Office.Core.MsoTriState.msoCTrue);
                        ppt.SaveAs(newName, PowerPoint.PpSaveAsFileType.ppSaveAsDefault);
                        ppt.Close();
                        break;
                    }
                default:
                    throw new Exception("Unexpected file extension.");
            }
        }

        public void Dispose()
        {
            try
            {
                if (excelApp != null) excelApp.Quit();
            }
            catch (Exception e) { Debug.WriteLine(e.Message); }
            try
            {

                if (wordApp != null) wordApp.Quit();
            }
            catch (Exception e) { Debug.WriteLine(e.Message); }
            try
            {
                if (pptApp != null) pptApp.Quit();
            }
            catch (Exception e) { Debug.WriteLine(e.Message); }
        }
    }
}
