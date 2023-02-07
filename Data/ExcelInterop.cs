using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace BlazorGenerateExcelReport.Data
{
    public class ExcelInterop
    {
        public void CopySheet()
        {
            Excel.Application excelApp;

            string sourceFileName = "";
            string tempFileName = "";
            string folderPath = @"";
            string sourceFilePath = System.IO.Path.Combine(folderPath, sourceFileName);
            string destinationFilePath = System.IO.Path.Combine(folderPath, tempFileName);

            System.IO.File.Copy(sourceFileName, destinationFilePath, true);

            excelApp = new Excel.Application();
            Excel.Workbook wbSource, wbTarget;
            Excel.Worksheet currentSheet;

            wbSource = excelApp.Workbooks.Open(sourceFilePath);
            wbTarget = excelApp.Workbooks.Open(destinationFilePath);

            currentSheet = (Excel.Worksheet)wbSource.Worksheets["Sheet1"];
            currentSheet.Name = "TempSheet";

            currentSheet.Copy(wbTarget.Worksheets[1]);
            wbSource.Close(false);
            wbTarget.Close(true);
            excelApp.Quit();

            System.IO.File.Delete(sourceFilePath);
            System.IO.File.Move(destinationFilePath, sourceFilePath);
        }
    }
}
