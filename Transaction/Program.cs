using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using static OfficeOpenXml.ExcelErrorValue;
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
            Excel.Workbook targetworkbook = targetexcelApp.Workbooks.Open(targetFilePath);

            Excel.Application storeApp = new Excel.Application();
            storeApp.Visible = true;
            string storeFilepath = @"C:\Users\Public\Documents\StoreList - Copy.xlsx";
            Excel.Workbook storeWorkBook = storeApp.Workbooks.Open(storeFilepath);


            try
            {

                //data transfer from one sales workbook to transaction workbook

                Worksheet sourceSheet = workbook.Worksheets["2023"];
                Worksheet targetSouth = targetworkbook.Worksheets["South 23"];
                Worksheet targetNorth = targetworkbook.Worksheets["North 23"];
                int southLastRow = targetSouth.Cells[targetSouth.Rows.Count, 2].End[Excel.XlDirection.xlUp].Row + 1;
                int northLastRow = targetNorth.Cells[targetNorth.Rows.Count, 2].End[Excel.XlDirection.xlUp].Row + 1;
                int lastFormulaRow = targetSouth.Cells[targetSouth.Rows.Count, 1].End[Excel.XlDirection.xlUp].Row;
                int lastFormulaColumn = targetSouth.Cells[lastFormulaRow, targetSouth.Columns.Count].End[Excel.XlDirection.xlToLeft].Column;
                int lastNFormulaRow = targetNorth.Cells[targetNorth.Rows.Count, 1].End[Excel.XlDirection.xlUp].Row;
                int lastNFormulaColumn = targetNorth.Cells[lastNFormulaRow, targetNorth.Columns.Count].End[Excel.XlDirection.xlToLeft].Column;





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

                sourceSheet.AutoFilterMode = false;

                Excel.Range formulaCopyRange = targetSouth.Cells[lastFormulaRow, 6].Resize[1, lastFormulaColumn - 5];
                formulaCopyRange.Copy(Type.Missing);
                int newLastRow = targetSouth.Cells[targetSouth.Rows.Count, 2].End[Excel.XlDirection.xlUp].Row;



                Excel.Range formulaRange = targetSouth.Range[$"F{southLastRow}:U" + newLastRow];
                formulaRange.PasteSpecial(XlPasteType.xlPasteFormulas, XlPasteSpecialOperation.xlPasteSpecialOperationNone, Type.Missing, Type.Missing);

                Excel.Range dateRange = targetSouth.Range[$"P{southLastRow}:P" + newLastRow];
                dateRange.NumberFormat = "MM/dd/yy ddd";
                dateRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                targetworkbook.Save();


                var northFilterlist = new object[]
               {
                 "cj north"
               };
                sourceRange.AutoFilter(3, FilterDate, XlAutoFilterOperator.xlFilterValues, Type.Missing, true);
                sourceRange.AutoFilter(7, northFilterlist, XlAutoFilterOperator.xlFilterValues, Type.Missing, true);
                var rangeNorthFr = sourceSheet.Range["A3:E" + sourceLastRow];
                rangeNorthFr.Copy(Type.Missing);
                targetNorth.Cells[northLastRow, 1].PasteSpecial(XlPasteType.xlPasteValues, XlPasteSpecialOperation.xlPasteSpecialOperationNone, Type.Missing, Type.Missing);

                sourceSheet.AutoFilterMode = false;

                Excel.Range formulaNCopyRange = targetNorth.Cells[lastNFormulaRow, 18].Resize[1, lastNFormulaColumn - 17];
                formulaNCopyRange.Copy(Type.Missing);
                int newNLastRow = targetNorth.Cells[targetNorth.Rows.Count, 2].End[Excel.XlDirection.xlUp].Row;



                Excel.Range formulaNRange = targetNorth.Range[$"R{northLastRow}:AG" + newNLastRow];
                formulaNRange.PasteSpecial(XlPasteType.xlPasteFormulas, XlPasteSpecialOperation.xlPasteSpecialOperationNone, Type.Missing, Type.Missing);


                targetworkbook.Save();



                //code for date change

                Worksheet targetSheet = targetworkbook.Worksheets["REF"];
                Excel.Range dateCell = targetSheet.Cells[2,1];

                var dateString = "10/10/2023";

                DateTime dateObj = DateTime.ParseExact(dateString, "MM/dd/yyyy", System.Globalization.CultureInfo.InvariantCulture);

                if (dateObj.DayOfWeek == DayOfWeek.Tuesday)
                {
                    dateCell.Value = dateString;
                    targetworkbook.Save();
                }


                //copying data from storelist to transaction file

                Worksheet storeList = storeWorkBook.Worksheets["StoreList"];
                Worksheet destinationSheet = targetworkbook.Worksheets["Site List"];

                Excel.Range clearRange = destinationSheet.Range["B1:O" + destinationSheet.Rows.Count];
                clearRange.Clear();
                //clearRange.ClearContents();

                Excel.Range printRange = storeList.Range["A1:N" + storeList.Rows.Count];
                printRange.Copy(Type.Missing);

                Excel.Range destinationRange = destinationSheet.Range["B1:O" + destinationSheet.Rows.Count];
                destinationRange.PasteSpecial(XlPasteType.xlPasteAll);

                targetworkbook.Save();


                //Dynamic transaction sheet cells change with reference to site list sheet 

                Worksheet siteList = targetworkbook.Worksheets["Site List"];
                Worksheet transaction = targetworkbook.Worksheets["Transactions 2023"];
              

                Excel.Range columnARange = siteList.Columns["W"]; 
                Excel.Range columnBRange = siteList.Columns["X"]; 

                Dictionary<string, List<string>> columnData = new Dictionary<string, List<string>>();

                bool breakLoop = false;

                int rowCount = siteList.Cells[siteList.Rows.Count, 2].End[Excel.XlDirection.xlUp].Row + 1;
                for (int i = 8; i <= rowCount; i++)
                {
                    string key = columnARange.Cells[i].Value?.ToString();
                    string value = columnBRange.Cells[i].Value?.ToString();




                    if (key != null)
                    {
                        if (!columnData.ContainsKey(key))
                        {
                            columnData[key] = new List<string>();
                        }

                        if (!string.IsNullOrEmpty(value) && !columnData[key].Contains(value))
                        {
                            columnData[key].Add(value);
                        }
                    }
                }

                int startTransactionRow = 19;
                int rowCounter = 0;
                int srNo = 1;
                int lastSrNo = 10;

                foreach (var kvp in columnData)
                {
                    foreach (var value in kvp.Value)
                    {
                        if (srNo == lastSrNo)
                        {
                            Excel.Range aboveRow = transaction.Rows[startTransactionRow + rowCounter - 1];
                            aboveRow.Copy(Type.Missing);
                            targetexcelApp.DisplayAlerts = false;

                            aboveRow.Insert(Excel.XlInsertShiftDirection.xlShiftDown);

                            aboveRow.PasteSpecial(XlPasteType.xlPasteAll);


                            srNo = lastSrNo; 
                            lastSrNo++;
                        }

                        transaction.Cells[startTransactionRow + rowCounter, "A"].Value = srNo;
                        transaction.Cells[startTransactionRow + rowCounter, "B"].Value = value;
                        srNo++;
                        rowCounter++;
                    }

                    transaction.Cells[startTransactionRow + rowCounter, "B"].Value = kvp.Key;
                    rowCounter++;
                }

                targetworkbook.Save();

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
