using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using Excel = Microsoft.Office.Interop.Excel;

namespace DataChange
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Excel.Application excelApp = new Excel.Application();
            excelApp.Visible = true;
            string sourceFilePath = @"C:\Users\Nimap\Downloads\backups\Daily Transactions 2023 - Copy.xlsx";
            Excel.Workbook workbook = excelApp.Workbooks.Open(sourceFilePath);

            try
            {
                Worksheet sourceSheet = workbook.Worksheets["REF"];
                Excel.Range dateCell = sourceSheet.Cells[2, 1];

                var dateString = "10/24/2023";

                DateTime date = DateTime.ParseExact(dateString, "MM/dd/yyyy", System.Globalization.CultureInfo.InvariantCulture);
 
                if (date.DayOfWeek == DayOfWeek.Tuesday)
                {
                    dateCell.Value = date;
                    workbook.Save();
                }    

            }
            finally
            {
                workbook.Close(false, Type.Missing, Type.Missing);
                Marshal.ReleaseComObject(workbook);
                excelApp.Quit();
                Marshal.ReleaseComObject(excelApp);
            }



        }
    }
}

/*var dateString = "10/24/2023";

DateTime startDate = DateTime.ParseExact(dateString, "MM/dd/yyyy", System.Globalization.CultureInfo.InvariantCulture);
DateTime endDate = DateTime.Today;

for (DateTime currentDate = startDate; currentDate <= endDate; currentDate = currentDate.AddDays(1))
{
    if (currentDate.DayOfWeek == DayOfWeek.Tuesday)
    {
        dateCell.Value = currentDate;
        workbook.Save();
    }

}*/