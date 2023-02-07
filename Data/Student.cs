using Microsoft.JSInterop;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace BlazorGenerateExcelReport.Data
{
    public class Student
    {
        public void GenerateExcel(IJSRuntime iJSRunTIme)
        {
            byte[] fileContents;
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using(var package= new ExcelPackage())
            {
                var workSheet = package.Workbook.Worksheets.Add("Sheet1");

                workSheet.Cells[1, 1].Value = "Student Name";
                workSheet.Cells[1, 1].Style.Font.Size = 12;
                workSheet.Cells[1, 1].Style.Font.Bold = true;
                workSheet.Cells[1, 1].Style.Border.Top.Style = ExcelBorderStyle.Hair;


                workSheet.Cells[1, 2].Value = "Student Roll";
                workSheet.Cells[1, 2].Style.Font.Size = 12;
                workSheet.Cells[1, 2].Style.Font.Bold = true;
                workSheet.Cells[1, 2].Style.Border.Top.Style = ExcelBorderStyle.Hair;



                workSheet.Cells[2, 1].Value = "New Student 1";
                workSheet.Cells[2, 1].Style.Font.Size = 12;
                workSheet.Cells[2, 1].Style.Border.Top.Style = ExcelBorderStyle.Hair;


                workSheet.Cells[2, 2].Value = "1000";
                workSheet.Cells[2, 2].Style.Font.Size = 12;
                workSheet.Cells[2, 2].Style.Border.Top.Style = ExcelBorderStyle.Hair;

                workSheet.Cells[3, 1].Value = "New Student 2";
                workSheet.Cells[3, 1].Style.Font.Size = 12;
                workSheet.Cells[3, 1].Style.Border.Top.Style = ExcelBorderStyle.Hair;


                workSheet.Cells[3, 2].Value = "1001";
                workSheet.Cells[3, 2].Style.Font.Size = 12;
                workSheet.Cells[3, 2].Style.Border.Top.Style = ExcelBorderStyle.Hair;

                fileContents = package.GetAsByteArray();

            }
            iJSRunTIme.InvokeAsync<Student>(
                "saveAsFile", "StudentList.xlsx",Convert.ToBase64String(fileContents)
                );
        }
    }
}
