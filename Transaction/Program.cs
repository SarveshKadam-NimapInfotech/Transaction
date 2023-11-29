using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Input;
using System.Windows.Media;
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

            //Excel.Application targetexcelApp = new Excel.Application();
            //targetexcelApp.Visible = true;
            string targetFilePath = @"C:\Users\Nimap\Downloads\backups\Daily Transactions 2023 - Copy.xlsx";
            Excel.Workbook targetworkbook = excelApp.Workbooks.Open(targetFilePath);

           // Excel.Application storeApp = new Excel.Application();
            //storeApp.Visible = true;
            string storeFilepath = @"C:\Users\Nimap\Downloads\backups\StoreList - Copy.xlsx";
            Excel.Workbook storeWorkBook = excelApp.Workbooks.Open(storeFilepath);


            try
            {
                //copying data from storelist to transaction file

                Worksheet storeList = storeWorkBook.Worksheets["StoreList"];
                Worksheet destinationSheet = targetworkbook.Worksheets["Site List"];

                Range clearRange = destinationSheet.Range["B1:AD" + destinationSheet.Rows.Count];
                clearRange.Clear();
                clearRange.ClearContents();

                Range printRange = storeList.Range["A1:N" + storeList.Rows.Count];
                printRange.Copy(Type.Missing);

                Range destinationRange = destinationSheet.Range["B1:O" + destinationSheet.Rows.Count];
                destinationRange.PasteSpecial(XlPasteType.xlPasteAll);

                destinationSheet.Range[$"P10:P{destinationSheet.UsedRange.Rows.Count}"].Formula = "=IF(ISNUMBER(B10),B10,\" \")";
                destinationSheet.Range[$"Q9:Q{destinationSheet.UsedRange.Rows.Count}"].Formula = "=IF(LEFT(B9,4)=\"Dist\",CONCAT(\"D\",TEXT(RIGHT(B9,2),\"00\")),Q8)";
                destinationSheet.Range[$"Q65:Q{destinationSheet.UsedRange.Rows.Count}"].Formula = "=IF(LEFT(B65,1)=\"D\",CONCAT(\"D\",TEXT(RIGHT(B65,2),\"00\")),Q64)";
                destinationSheet.Range[$"R8:R{destinationSheet.UsedRange.Rows.Count}"].Formula = "=IF(B8=\"Region 1\",\"R01\",IF(B8=\"Region 2\",\"R02\",R7))";
                destinationSheet.Range[$"S9:S{destinationSheet.UsedRange.Rows.Count}"].Formula = "=IFERROR(IF(LEFT(B9,4)=\"Dist\",CONCAT(\"D\",TEXT(RIGHT(B9,2),\"00\")),P9),\"\")";
                destinationSheet.Range[$"S65:S{destinationSheet.UsedRange.Rows.Count}"].Formula = "=IFERROR(IF(LEFT(B65,1)=\"D\",CONCAT(\"D\",TEXT(RIGHT(B65,2),\"00\")),P65),\"\")";

                destinationSheet.Range[$"T9:T{destinationSheet.UsedRange.Rows.Count}"].Formula = "=IFERROR(IF(J9=\"\",L9,T8),\"\")";
                destinationSheet.Range[$"V9:V{destinationSheet.UsedRange.Rows.Count}"].Formula = "=B9";
                destinationSheet.Range[$"W8:W{destinationSheet.UsedRange.Rows.Count}"].Formula = "=IF(B8=\"Region 1\",\"R01\",IF(B8=\"Region 2\",\"R02\",R7))";
                destinationSheet.Range[$"Y10:Y{destinationSheet.UsedRange.Rows.Count}"].Formula = "=P10";
                destinationSheet.Range[$"Z9:Z{destinationSheet.UsedRange.Rows.Count}"].Formula = "=Q9&Y9";

                destinationSheet.Range[$"P75:P{destinationSheet.UsedRange.Rows.Count}"].Formula = "=B75";
                destinationSheet.Range[$"Q75:Q{destinationSheet.UsedRange.Rows.Count}"].Formula = "=IF(ISNUMBER(SEARCH(\"CJ\",F75)),\"CJ NORTH\",\"\")";
                destinationSheet.Range[$"Y75:Y{destinationSheet.UsedRange.Rows.Count}"].Formula = "=P75";
                destinationSheet.Range[$"Z75:Z{destinationSheet.UsedRange.Rows.Count}"].Formula = "=Q75&Y75";

                destinationSheet.Calculate();

                //Worksheet storeList = storeWorkBook.Worksheets["StoreList"];
                //Worksheet destinationSheet = targetworkbook.Worksheets["Site List"];

                //Excel.Range clearRange = destinationSheet.Range["B1:AD" + destinationSheet.Rows.Count];
                //clearRange.Clear();
                ////clearRange.ClearContents();

                //Excel.Range printRange = storeList.Range["A1:N" + storeList.Rows.Count];
                //printRange.Copy(Type.Missing);

                //Excel.Range destinationRange = destinationSheet.Range["B1:O" + destinationSheet.Rows.Count];
                //destinationRange.PasteSpecial(XlPasteType.xlPasteAll);

                //Excel.Range printFormula = storeList.Range["O1:Y" + storeList.Rows.Count];
                //printFormula.Copy(Type.Missing);

                //Excel.Range destinationFormulaRange = destinationSheet.Range["P1:Z" + destinationSheet.Rows.Count];
                //destinationFormulaRange.PasteSpecial(XlPasteType.xlPasteFormulas, XlPasteSpecialOperation.xlPasteSpecialOperationNone, Type.Missing, Type.Missing);

                //targetworkbook.Save();


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

                 var date = "10/10/2023";
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
                 //sourceRange.AutoFilter(3, FilterDate, XlAutoFilterOperator.xlFilterValues, Type.Missing, true);
                 //sourceRange.AutoFilter(7, SouthFilterList, XlAutoFilterOperator.xlFilterValues, Type.Missing, true);
                 //var rangeSouthFr = sourceSheet.Range["A3:E" + sourceLastRow];
                 //rangeSouthFr.Copy(Type.Missing);


                 sourceRange.AutoFilter(3, FilterDate, XlAutoFilterOperator.xlFilterValues, Type.Missing, true);
                 var filteredByDateRange = sourceRange.SpecialCells(XlCellType.xlCellTypeVisible);


                 // sourceSheet.AutoFilterMode = false;

                 filteredByDateRange.AutoFilter(7, SouthFilterList, XlAutoFilterOperator.xlFilterValues, Type.Missing, true);
                 sourceSheet.Range["A3:E" + sourceLastRow].Copy(Type.Missing);
                 targetSouth.Cells[southLastRow, 1].PasteSpecial(XlPasteType.xlPasteValues, XlPasteSpecialOperation.xlPasteSpecialOperationNone, Type.Missing, Type.Missing);

                 //sourceSheet.AutoFilterMode = false;

                 Excel.Range formulaCopyRange = targetSouth.Cells[lastFormulaRow, 6].Resize[1, lastFormulaColumn - 5];
                 formulaCopyRange.Copy(Type.Missing);
                 int newLastRow = targetSouth.Cells[targetSouth.Rows.Count, 2].End[Excel.XlDirection.xlUp].Row;



                 Excel.Range formulaRange = targetSouth.Range[$"F{southLastRow}:U" + newLastRow];
                 formulaRange.PasteSpecial(XlPasteType.xlPasteFormulas, XlPasteSpecialOperation.xlPasteSpecialOperationNone, Type.Missing, Type.Missing);

                 Excel.Range dateRange = targetSouth.Range[$"P{southLastRow}:P" + newLastRow];
                 dateRange.NumberFormat = "MM/dd/yy ddd";
                 dateRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                 //targetworkbook.Save();


                 var northFilterlist = new object[]
                {
                  "cj north"
                };
                 //sourceRange.AutoFilter(3, FilterDate, XlAutoFilterOperator.xlFilterValues, Type.Missing, true);
                 //sourceRange.AutoFilter(7, northFilterlist, XlAutoFilterOperator.xlFilterValues, Type.Missing, true);
                 //var rangeNorthFr = sourceSheet.Range["A3:E" + sourceLastRow];
                 //rangeNorthFr.Copy(Type.Missing);

                 filteredByDateRange.AutoFilter(7, northFilterlist, XlAutoFilterOperator.xlFilterValues, Type.Missing, true);
                 sourceSheet.Range["A3:E" + sourceLastRow].Copy(Type.Missing);
                 targetNorth.Cells[northLastRow, 1].PasteSpecial(XlPasteType.xlPasteValues, XlPasteSpecialOperation.xlPasteSpecialOperationNone, Type.Missing, Type.Missing);

                 var rangeNorthSn = sourceSheet.Range["G3:G" + sourceLastRow];
                 rangeNorthSn.Copy(Type.Missing);
                 targetNorth.Cells[northLastRow, 18].PasteSpecial(XlPasteType.xlPasteValues, XlPasteSpecialOperation.xlPasteSpecialOperationNone, Type.Missing, Type.Missing);

                 //sourceSheet.AutoFilterMode = false;

                 Excel.Range formulaNCopyRange = targetNorth.Cells[lastNFormulaRow, 20].Resize[1, lastNFormulaColumn - 19];
                 formulaNCopyRange.Copy(Type.Missing);
                 int newNLastRow = targetNorth.Cells[targetNorth.Rows.Count, 2].End[Excel.XlDirection.xlUp].Row;



                 Excel.Range formulaNRange = targetNorth.Range[$"T{northLastRow}:AG" + newNLastRow];
                 formulaNRange.PasteSpecial(XlPasteType.xlPasteFormulas, XlPasteSpecialOperation.xlPasteSpecialOperationNone, Type.Missing, Type.Missing);


                 //targetworkbook.Save();

                 

                //code for date change

                Worksheet targetSheet = targetworkbook.Worksheets["REF"];
                Excel.Range dateCell = targetSheet.Cells[2,1];


                DateTime dateObj = DateTime.ParseExact(date, "MM/dd/yyyy", System.Globalization.CultureInfo.InvariantCulture);

                if (dateObj.DayOfWeek == DayOfWeek.Tuesday)
                {
                    dateCell.Value = date;
                    //targetworkbook.Save();
                }
            



                //Dynamic transaction sheet cells change with reference to site list sheet 

                Worksheet siteList = targetworkbook.Worksheets["Site List"];
                Worksheet transaction = targetworkbook.Worksheets["Transactions 2023"];
              

                Excel.Range columnARange = siteList.Columns["R"]; 
                Excel.Range columnBRange = siteList.Columns["V"]; 

                Dictionary<string, List<string>> columnData = new Dictionary<string, List<string>>();

                bool breakLoop = false;

                int rowCount = siteList.Cells[siteList.Rows.Count, 2].End[Excel.XlDirection.xlUp].Row + 1;

                // ... (previous code remains unchanged)

                for (int i = 9; i <= rowCount; i++)
                {
                    string key = columnARange.Cells[i].Value?.ToString();
                    string value = columnBRange.Cells[i].Value?.ToString();

                    if (key != null)
                    {
                        if (!columnData.ContainsKey(key))
                        {
                            columnData[key] = new List<string>();
                        }

                        if (!string.IsNullOrEmpty(value) && value != "CJ NORTH")
                        {
                            // Check if the value is not negative
                            if (value.StartsWith("Dist ") || value.StartsWith("D"))
                            {
                                string distNumber = value.Replace("Dist ", "").Replace("D", "");
                                int distValue;

                                if (int.TryParse(distNumber, out distValue))
                                {
                                    //string distCode = (distValue == 11) ? "D11" : "D" + distValue.ToString("00");

                                    string distCode = "D" + distValue.ToString("00");
                                    bool valueExistsInFirstKey = columnData.ContainsKey(columnARange.Cells[8].Value?.ToString()) &&
                                                                 columnData[columnARange.Cells[8].Value?.ToString()].Contains(distCode);

                                    if (!valueExistsInFirstKey && !columnData[key].Contains(distCode))
                                    {
                                        columnData[key].Add(distCode);
                                    }
                                }
                            }
                        }
                    }
                }



                //for (int i = 9; i <= rowCount; i++)
                //{
                //    string key = columnARange.Cells[i].Value?.ToString();
                //    string value = columnBRange.Cells[i].Value?.ToString();
                //    if (key != null)
                //    {
                //        if (!columnData.ContainsKey(key))
                //        {
                //            columnData[key] = new List<string>();
                //        }

                //        if (!string.IsNullOrEmpty(value) && value != "CJ NORTH")
                //        {
                //            // Check if the value is not negative
                //            if (!value.StartsWith("-"))
                //            {
                //                bool valueExistsInFirstKey = columnData.ContainsKey(columnARange.Cells[8].Value?.ToString()) &&
                //                                                columnData[columnARange.Cells[8].Value?.ToString()].Contains(value);

                //                if (!valueExistsInFirstKey && !columnData[key].Contains(value))
                //                {
                //                    columnData[key].Add(value);
                //                }
                //            }
                //        }
                //    }
                //}


                //if (key != null)
                //{
                //    if (!columnData.ContainsKey(key))
                //    {
                //        columnData[key] = new List<string>();
                //    }

                //    if (value != "CJ NORTH")
                //    {
                //        // Check if the value is not already associated with the first key
                //        bool valueExistsInFirstKey = columnData.ContainsKey(columnARange.Cells[8].Value?.ToString()) &&
                //                                     columnData[columnARange.Cells[8].Value?.ToString()].Contains(value);

                //        if (!valueExistsInFirstKey && !columnData[key].Contains(value) && !string.IsNullOrEmpty(value))
                //        {
                //            columnData[key].Add(value);
                //        }
                //    }

                //if (!string.IsNullOrEmpty(value) && !columnData[key].Contains(value) && value != "CJ NORTH")
                //{
                //    columnData[key].Add(value);
                //}


                foreach (var kvp in columnData)
                {
                    Console.WriteLine($"Key: {kvp.Key}");
                    Console.WriteLine("Values:");
                    foreach (var value in kvp.Value)
                    {
                        Console.WriteLine(value);
                    }
                    Console.WriteLine("------");
                }

                int startTransactionRow = 19;
                int rowCounter = 0;
                int srNo = 1;
                //int regionCounter = 29;

                foreach (var kvp in columnData)
                {
                    foreach (var value in kvp.Value)
                    {
                        if (transaction.Cells[startTransactionRow + rowCounter, "B"].Value == "R01" || transaction.Cells[startTransactionRow + rowCounter, "B"].Value == "R02")
                        {
                            Excel.Range aboveRow = transaction.Rows[startTransactionRow + rowCounter - 1];
                            aboveRow.Copy(Type.Missing);
                            excelApp.DisplayAlerts = false;

                            aboveRow.Insert(Excel.XlInsertShiftDirection.xlShiftDown);

                            aboveRow.PasteSpecial(XlPasteType.xlPasteAll);

                        }

                        //transaction.Rows[startTransactionRow + rowCounter].Interior.Color = XlRgbColor.rgbWhite; 
                        transaction.Cells[startTransactionRow + rowCounter, "A"].Value = srNo;
                        transaction.Cells[startTransactionRow + rowCounter, "B"].Value = value;
                        srNo++;
                        rowCounter++;

                        
                    }
                    if (transaction.Cells[startTransactionRow + rowCounter, "B"].Value != "R01")
                    {
                        if (transaction.Cells[startTransactionRow + rowCounter, "B"].Value != "R02")
                        {
                            Excel.Range regionRow = transaction.Rows[startTransactionRow + rowCounter];
                            excelApp.DisplayAlerts = false;

                            regionRow.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);


                        }

                    }
                    

                    //transaction.Rows[startTransactionRow + rowCounter].Interior.Color = XlRgbColor.rgbYellow;
                    //transaction.Cells[startTransactionRow + rowCounter, "A"].Value = null;
                    //transaction.Cells[startTransactionRow + rowCounter, "C"].Value = null;
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

              
            }
        }
    }
}
