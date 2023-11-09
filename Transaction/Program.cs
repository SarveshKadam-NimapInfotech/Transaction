using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using Excel = Microsoft.Office.Interop.Excel;

namespace Transaction
{
    internal class Program
    {
        static void Main(string[] args)
        {

            //opening the workbooks

            Excel.Application excelApp = new Excel.Application();
            excelApp.Visible = true;
            string filePath = @"C:\Users\Nimap\Downloads\backups\Daily sales - Copy.xlsx";
            Excel.Workbook workbook = excelApp.Workbooks.Open(filePath);

            Excel.Application targetexcelApp = new Excel.Application();
            targetexcelApp.Visible = true;
            string targetFilePath = @"C:\Users\Nimap\Downloads\backups\Daily Transactions 2023 - Copy.xlsx";
            Excel.Workbook targetworkbook = excelApp.Workbooks.Open(targetFilePath);

            Excel.Application storeApp = new Excel.Application();
            storeApp.Visible = true;
            string storeFilepath = @"C:\Users\Public\Documents\StoreList.xlsx";
            Excel.Workbook storeWorkBook = storeApp.Workbooks.Open(storeFilepath);


            try
            {

                //data transfer from one sales workbook to transaction workbook

                Worksheet sourceSheet = workbook.Worksheets["2023"];
                Worksheet targetSouth = targetworkbook.Worksheets["South 23"];
                Worksheet targetNorth = targetworkbook.Worksheets["North 23"];
                int southLastRow = targetSouth.Cells[targetSouth.Rows.Count, 2].End[Excel.XlDirection.xlUp].Row + 1;
                int northLastRow = targetNorth.Cells[targetNorth.Rows.Count, 2].End[Excel.XlDirection.xlUp].Row + 1;




                Range sourceRange = sourceSheet.Range[sourceSheet.Cells[1, 1], sourceSheet.Cells[1, sourceSheet.UsedRange.Column]];
                int sourceLastRow = sourceSheet.Cells[sourceSheet.Rows.Count, 2].End[Excel.XlDirection.xlUp].Row ;

               // var FilterDate = new object[]
               //{
               //  "10/23/2023"
               //};
                var SouthFilterList = new object[]
               {
                "D01",
                "D02",
                "D03",
                "D05",
                "D06",
                "D07",
                "D08",
                "D09",
                "D11"
               };
                //sourceRange.AutoFilter(3, FilterDate, XlAutoFilterOperator.xlFilterValues, Type.Missing, true);
                sourceRange.AutoFilter(7, SouthFilterList, XlAutoFilterOperator.xlFilterValues, Type.Missing, true);
                var rangeSouth = sourceSheet.Range["A3:X" + sourceLastRow];
                rangeSouth.Copy(Type.Missing);

               
                targetSouth.Cells[southLastRow, 1].PasteSpecial(XlPasteType.xlPasteValues, XlPasteSpecialOperation.xlPasteSpecialOperationNone, Type.Missing, Type.Missing);
                targetworkbook.Save();

                sourceSheet.AutoFilterMode = false;
                var NorthFilterList = new object[]
               {
                "CJ NORTH"
               };
                //sourceRange.AutoFilter(3, FilterDate, XlAutoFilterOperator.xlFilterValues, Type.Missing, true);
                sourceRange.AutoFilter(7, NorthFilterList, XlAutoFilterOperator.xlFilterValues, Type.Missing, true);
                var rangeNorth = sourceSheet.Range["A3:X" + sourceLastRow];
                rangeNorth.Copy(Type.Missing);

                targetNorth.Cells[northLastRow, 1].PasteSpecial(XlPasteType.xlPasteValues, XlPasteSpecialOperation.xlPasteSpecialOperationNone, Type.Missing, Type.Missing);
                targetworkbook.Save();


                //code for date change

                Worksheet targetSheet = targetworkbook.Worksheets["REF"];
                Excel.Range dateCell = targetSheet.Cells[2, 1];

                var dateString = "10/24/2023";

                DateTime date = DateTime.ParseExact(dateString, "MM/dd/yyyy", System.Globalization.CultureInfo.InvariantCulture);

                if (date.DayOfWeek == DayOfWeek.Tuesday)
                {
                    dateCell.Value = date;
                    targetworkbook.Save();
                }


                //copying data from storelist to transaction file

                Worksheet storeList = storeWorkBook.Worksheets["StoreList"];
                Worksheet destinationSheet = targetworkbook.Worksheets["Site List"];

                storeList.Copy(destinationSheet);

                storeWorkBook.Save();


            }
            finally
            {
                // Close and release Excel workbooks
                workbook?.Close(false, Type.Missing, Type.Missing);
                Marshal.ReleaseComObject(workbook);

                targetworkbook?.Close(false, Type.Missing, Type.Missing);
                Marshal.ReleaseComObject(targetworkbook);

                storeWorkBook?.Close(false, Type.Missing, Type.Missing);
                Marshal.ReleaseComObject(storeWorkBook);

                // Quit and release Excel applications
                excelApp?.Quit();
                Marshal.ReleaseComObject(excelApp);

                targetexcelApp?.Quit();
                Marshal.ReleaseComObject(targetexcelApp);

                storeApp?.Quit();
                Marshal.ReleaseComObject(storeApp);
            }
        }
    }
}
