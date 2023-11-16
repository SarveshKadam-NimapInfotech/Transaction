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
            string storeFilepath = @"C:\Users\Public\Documents\StoreList - Copy.xlsx";
            Excel.Workbook storeWorkBook = storeApp.Workbooks.Open(storeFilepath);


            try
            {

                /*//data transfer from one sales workbook to transaction workbook

                Worksheet sourceSheet = workbook.Worksheets["2023"];
                Worksheet targetSouth = targetworkbook.Worksheets["South 23"];
                Worksheet targetNorth = targetworkbook.Worksheets["North 23"];
                int southLastRow = targetSouth.Cells[targetSouth.Rows.Count, 2].End[Excel.XlDirection.xlUp].Row + 1;
                int northLastRow = targetNorth.Cells[targetNorth.Rows.Count, 2].End[Excel.XlDirection.xlUp].Row + 1;




                Range sourceRange = sourceSheet.Range[sourceSheet.Cells[1, 1], sourceSheet.Cells[1, sourceSheet.UsedRange.Column]];
                int sourceLastRow = sourceSheet.Cells[sourceSheet.Rows.Count, 2].End[Excel.XlDirection.xlUp].Row ;


                var date = "10/23/2023";
                var FilterDate = new object[]
               {
                 date
               };
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
                sourceRange.AutoFilter(3, FilterDate, XlAutoFilterOperator.xlFilterValues, Type.Missing, true);
                sourceRange.AutoFilter(7, SouthFilterList, XlAutoFilterOperator.xlFilterValues, Type.Missing, true);
                var rangeSouthFr = sourceSheet.Range["A3:E" + sourceLastRow];
                rangeSouthFr.Copy(Type.Missing);
                targetSouth.Cells[southLastRow, 1].PasteSpecial(XlPasteType.xlPasteValues, XlPasteSpecialOperation.xlPasteSpecialOperationNone, Type.Missing, Type.Missing);

                var rangeSouthSc = sourceSheet.Range["G3:H" + sourceLastRow];
                rangeSouthSc.Copy(Type.Missing);
                targetSouth.Cells[southLastRow, 6].PasteSpecial(XlPasteType.xlPasteValues, XlPasteSpecialOperation.xlPasteSpecialOperationNone, Type.Missing, Type.Missing);

                targetworkbook.Save();

                sourceSheet.AutoFilterMode = false;

                var northFilterlist = new object[]
               {
                 "cj north"
               };
                sourceRange.AutoFilter(3, FilterDate, XlAutoFilterOperator.xlFilterValues, Type.Missing, true);
                sourceRange.AutoFilter(7, northFilterlist, XlAutoFilterOperator.xlFilterValues, Type.Missing, true);
                var rangeNorthFr = sourceSheet.Range["A3:E" + sourceLastRow];
                rangeNorthFr.Copy(Type.Missing);
                targetNorth.Cells[northLastRow, 1].PasteSpecial(XlPasteType.xlPasteValues, XlPasteSpecialOperation.xlPasteSpecialOperationNone, Type.Missing, Type.Missing);

                var rangeNorthSc = sourceSheet.Range["G3:G" + sourceLastRow];
                rangeNorthSc.Copy(Type.Missing);
                targetNorth.Cells[northLastRow, 18].PasteSpecial(XlPasteType.xlPasteValues, XlPasteSpecialOperation.xlPasteSpecialOperationNone, Type.Missing, Type.Missing);

                var rangeNorthTh= sourceSheet.Range["P3:P" + sourceLastRow];
                rangeNorthTh.Copy(Type.Missing);
                targetNorth.Cells[northLastRow, 29].PasteSpecial(XlPasteType.xlPasteValues, XlPasteSpecialOperation.xlPasteSpecialOperationNone, Type.Missing, Type.Missing);

                targetworkbook.Save();

                sourceSheet.AutoFilterMode = false;


                //code for date change

                Worksheet targetSheet = targetworkbook.Worksheets["REF"];
                Excel.Range dateCell = targetSheet.Cells[2, 1];

                var dateString = "10/24/2023";

                DateTime dateObj = DateTime.ParseExact(dateString, "MM/dd/yyyy", System.Globalization.CultureInfo.InvariantCulture);

                if (dateObj.DayOfWeek == DayOfWeek.Tuesday)
                {
                    dateCell.Value = date;
                    targetworkbook.Save();
                }
                */


                //copying data from storelist to transaction file

                Worksheet storeList = storeWorkBook.Worksheets["StoreList"];
                Worksheet destinationSheet = targetworkbook.Worksheets["Site List"];

                destinationSheet.Cells.Clear();

                storeList.Cells.Copy(Type.Missing);

                Excel.Range destinationRange = destinationSheet.Cells;
                destinationRange.PasteSpecial(XlPasteType.xlPasteAll);

                targetworkbook.Save();


                //Dynamic transaction sheet cells change with reference to site list sheet 

                Worksheet siteList = targetworkbook.Worksheets["Site List"];
                Worksheet transaction = targetworkbook.Worksheets["Transactions 2023"];

                Excel.Range transactionCellValue = transaction.Cells[19, 2];
                var transvalue1 = transactionCellValue.Value;

                Excel.Range siteCellValue = siteList.Cells[9, 1];
                var distValue1 = siteCellValue.Value;

                var value1 = distValue1[5];
                var value2 = transvalue1[2];

                if(value1 != value2)
                {
                    char[] transvalueChars = transvalue1.ToCharArray();
                    transvalueChars[2] = value1;
                    transvalue1 = new string(transvalueChars);

                    transactionCellValue.Value = transvalue1;

                    targetworkbook.Save();
                }
                
                


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
