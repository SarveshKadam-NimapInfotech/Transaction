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

                //Worksheet storeList = storeWorkBook.Worksheets["StoreList"];
                //Worksheet destinationSheet = targetworkbook.Worksheets["Site List"];

                //Excel.Range clearRange = destinationSheet.Range["A1:O" + destinationSheet.Rows.Count];
                //clearRange.ClearContents();

                //Excel.Range usedRange = storeList.UsedRange;
                //usedRange.Copy(Type.Missing);

                //Excel.Range destinationRange = destinationSheet.Cells;
                //destinationRange.PasteSpecial(Excel.XlPasteType.xlPasteValues);

                //targetworkbook.Save();


                //Worksheet storeList = storeWorkBook.Worksheets["StoreList"];
                //Worksheet destinationSheet = targetworkbook.Worksheets["Site List"];

                //destinationSheet.Cells.Clear();

                //storeList.Cells.Copy(Type.Missing);

                //Excel.Range destinationRange = destinationSheet.Cells;
                //destinationRange.PasteSpecial(XlPasteType.xlPasteAll);

                //targetworkbook.Save();


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
                            Excel.Range currentRow = transaction.Rows[startTransactionRow + rowCounter];
                            currentRow.Insert(Excel.XlInsertShiftDirection.xlShiftDown); // Insert a new row

                            // Copy formulas from the above row (assuming the formulas are in columns A and B)
                            Excel.Range aboveRow = transaction.Rows[startTransactionRow + rowCounter - 1];
                            aboveRow.Copy(currentRow);
                            

                            srNo = lastSrNo; // Reset srNo after inserting the new row
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


                //foreach (var kvp in columnData)
                //{
                //    foreach (var value in kvp.Value)
                //    {
                //        //Console.WriteLine(value);
                //        transaction.Cells[startTransactionRow + rowCounter, "A"].Value = srNo;
                //        rowCounter++;
                //        srNo++;

                //        transaction.Cells[startTransactionRow + rowCounter, "B"].Value = value;
                //        rowCounter++;
                //    }
                //    transaction.Cells[startTransactionRow + rowCounter, "B"].Value = kvp.Key;
                //    rowCounter++;
                //    //Console.WriteLine(kvp.Key);

                //}




                //Excel.Range transactionCellValue1 = transaction.Cells[19, 2];
                //var transvalue1 = transactionCellValue1.Value;

                //Excel.Range siteCellValue1 = siteList.Cells[9, 1];
                //var distValue1 = siteCellValue1.Value;

                //var dValue1 = distValue1.Substring(5, 2);
                //var tValue1 = transvalue1.Substring(1, 2);

                //if (dValue1 != tValue1)
                //{
                //    char[] transvalueChars = transvalue1.ToCharArray();
                //    transvalueChars[1] = dValue1[0];
                //    transvalueChars[2] = dValue1[1];

                //    transvalue1 = new string(transvalueChars);

                //    transactionCellValue1.Value = transvalue1;

                //    targetworkbook.Save();
                //}


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
