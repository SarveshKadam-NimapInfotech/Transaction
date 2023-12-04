using System;
using System.Collections.Generic;
using System.ComponentModel;
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
    internal class RowData
    {
        public string Region { get; set; }
        public DateTime BusDate { get; set; }
        public double Store { get; set; }
        public string Source { get; set; }
        public double Sales { get; set; }
        public double Transaction { get; set; }
        public string Dist { get; set; }
    }

    internal class Program
    {

        static void Main(string[] args)
        {



            string filePath = @"C:\Users\Nimap\source\repos\Test\Test\test.xlsx";

            string targetFilePath = @"C:\Users\Nimap\source\repos\Test\Test\Daily.xlsx";

            string storeFilepath = @"C:\Users\Nimap\source\repos\Test\Test\siteList.xlsx";




            try
            {

                //opening the workbooks

                using (var excelPackage = new ExcelPackage(new System.IO.FileInfo(filePath)))
                {
                    var dailyWorkbook = excelPackage.Workbook;
                    var worksheet = dailyWorkbook.Worksheets["2023"];
                    var excelPackage2 = new ExcelPackage(new System.IO.FileInfo(targetFilePath));
                    var excelPakckage3 = new ExcelPackage(new System.IO.FileInfo(storeFilepath));
                    var transactionWorkbook = excelPackage2.Workbook;
                    var northWorksheet = transactionWorkbook.Worksheets["North 23"];
                    var southWorksheet = transactionWorkbook.Worksheets["South 23"];
                    var storeListWorkbook = excelPakckage3.Workbook;

                    var destinationSheet = transactionWorkbook.Worksheets["Site List"];
                    var storeListWorksheet = storeListWorkbook.Worksheets["Sheet1"];

                    int busDate = 3;
                    int dist = 7;
                    int lastRow = worksheet.Cells[worksheet.Dimension.Start.Row, 1, worksheet.Dimension.End.Row, 1]
                                    .Reverse()
                                    .FirstOrDefault(cell => !string.IsNullOrWhiteSpace(cell.Text))
                                    ?.Start.Row ?? 0;


                    int lastRowForSouth = southWorksheet.Cells[southWorksheet.Dimension.Start.Row, 1, southWorksheet.Dimension.End.Row, 1]
                                         .Reverse()
                                         .FirstOrDefault(cell => !string.IsNullOrWhiteSpace(cell.Text))
                                         ?.Start.Row ?? 0;
                    int lastRowForNorth = northWorksheet.Cells[northWorksheet.Dimension.Start.Row, 1, northWorksheet.Dimension.End.Row, 1]
         .Reverse()
         .FirstOrDefault(cell => !string.IsNullOrWhiteSpace(cell.Text))
         ?.Start.Row ?? 0;

                    int lastRowForSiteListCol2 = destinationSheet.Cells[destinationSheet.Dimension.Start.Row, 2, destinationSheet.Dimension.End.Row, 2]
                                    .Reverse()
                                    .FirstOrDefault(cell => !string.IsNullOrWhiteSpace(cell.Text))
                                    ?.Start.Row ?? 0;
                    int lastRowForStoreListCol2 = storeListWorksheet.Cells[storeListWorksheet.Dimension.Start.Row, 1, storeListWorksheet.Dimension.End.Row, 1]
                                    .Reverse()
                                    .FirstOrDefault(cell => !string.IsNullOrWhiteSpace(cell.Text))
                                    ?.Start.Row ?? 0;

                    destinationSheet.Cells["B1:AD" + lastRowForSiteListCol2].Clear();
                    destinationSheet.Cells["P10:P" + lastRowForSiteListCol2].Formula = "=IF(ISNUMBER(B10),B10,\" \")";
                    destinationSheet.Cells["Q9:Q" + lastRowForSiteListCol2].Formula = "=IF(LEFT(B9,4)=\"Dist\",CONCAT(\"D\",TEXT(RIGHT(B9,2),\"00\")),Q8)";
                    destinationSheet.Cells["Q65:Q" + lastRowForSiteListCol2].Formula = "=IF(LEFT(B65,1)=\"D\",CONCAT(\"D\",TEXT(RIGHT(B65,2),\"00\")),Q64)";
                    destinationSheet.Cells["R8:R" + lastRowForSiteListCol2].Formula = "=IF(B8=\"Region 1\",\"R01\",IF(B8=\"Region 2\",\"R02\",R7))";
                    destinationSheet.Cells["S9:S" + lastRowForSiteListCol2].Formula = "=IFERROR(IF(LEFT(B9,4)=\"Dist\",CONCAT(\"D\",TEXT(RIGHT(B9,2),\"00\")),P9),\"\")";
                    destinationSheet.Cells["S65:S" + lastRowForSiteListCol2].Formula = "=IFERROR(IF(LEFT(B65,1)=\"D\",CONCAT(\"D\",TEXT(RIGHT(B65,2),\"00\")),P65),\"\")";
                    destinationSheet.Cells["T9:T" + lastRowForSiteListCol2].Formula = "=IFERROR(IF(J9=\"\",L9,T8),\"\")";
                    destinationSheet.Cells["V9:V" + lastRowForSiteListCol2].Formula = "=B9";
                    destinationSheet.Cells["W8:W" + lastRowForSiteListCol2].Formula = "=IF(B8=\"Region 1\",\"R01\",IF(B8=\"Region 2\",\"R02\",R7))";
                    destinationSheet.Cells["Y10:Y" + lastRowForSiteListCol2].Formula = "=P10";
                    destinationSheet.Cells["Z9:Z" + lastRowForSiteListCol2].Formula = "=Q9&Y9";

                    destinationSheet.Cells["P76:P" + lastRowForSiteListCol2].Formula = "=NUMBERVALUE(B76)";
                    destinationSheet.Cells["Q75:Q" + lastRowForSiteListCol2].Formula = "=IF(ISNUMBER(SEARCH(\"CJ\",F75)),\"CJ NORTH\",\"\")";
                    destinationSheet.Cells["Y75:Y" + lastRowForSiteListCol2].Formula = "=P75";
                    destinationSheet.Cells["Z75:Z" + lastRowForSiteListCol2].Formula = "=Q75&Y75";

                    destinationSheet.Calculate();

                    // Copying data from StoreList to DestinationSheet
                    storeListWorksheet.Cells["A1:N" + lastRowForStoreListCol2].Copy(destinationSheet.Cells["B1:O" + lastRowForSiteListCol2]);
                    var refSheet = transactionWorkbook.Worksheets["REF"];
                    var dateCell = refSheet.Cells[2, 1];
                    var date = "11/20/2023";

                    DateTime dateObj = DateTime.ParseExact(date, "MM/dd/yyyy", System.Globalization.CultureInfo.InvariantCulture);

                    if (dateObj.DayOfWeek == DayOfWeek.Tuesday)
                    {
                        dateCell.Value = date;

                    }


                    var transactionSheet = transactionWorkbook.Worksheets["Transactions 2023"];


                    List<string> region1 = new List<string> { };
                    List<string> region2 = new List<string> { };
                    Dictionary<string, List<string>> columnData = new Dictionary<string, List<string>>();

                    bool isRegion1 = false;
                    bool isRegion2 = false;
                    string regKey = "";


                    for (int i = 1; i <= lastRowForSiteListCol2; i++)
                    {
                        var key = destinationSheet.Cells[i, 2].Value;
                        var value = destinationSheet.Cells[i, 2].Value;

                        if (key is string stringValue)
                        {
                            if (stringValue.Contains("Region 1"))
                            {
                                isRegion1 = true;

                                // Extract numeric part
                                string numericPart = new string(stringValue.Where(char.IsDigit).ToArray());

                                // Convert to integer
                                if (int.TryParse(numericPart, out int regionNumber))
                                {
                                    // Format as R01
                                    regKey = $"R{regionNumber:D2}";
                                }
                                columnData[regKey] = new List<string>();

                                continue;
                            }

                            if (isRegion1 && !isRegion2 && !stringValue.Contains("Region 2"))
                            {
                                string numericPart = new string(stringValue.Where(char.IsDigit).ToArray());

                                // Convert to integer
                                if (int.TryParse(numericPart, out int districtNumber))
                                {
                                    // Format as D01 or D10
                                    string formattedDistrict = $"D{districtNumber:D2}";

                                    // Add to the list
                                    columnData[regKey].Add(formattedDistrict);
                                }
                            }

                            if (stringValue.Contains("Region 2"))
                            {
                                isRegion1 = false;
                                isRegion2 = true;
                                // Extract numeric part
                                string numericPart = new string(stringValue.Where(char.IsDigit).ToArray());

                                // Convert to integer
                                if (int.TryParse(numericPart, out int regionNumber))
                                {
                                    // Format as R01
                                    regKey = $"R{regionNumber:D2}";
                                }
                                columnData[regKey] = new List<string>();

                                continue;
                            }

                            if (isRegion2 && !isRegion1 && !stringValue.Contains("North"))
                            {
                                string numericPart = new string(stringValue.Where(char.IsDigit).ToArray());

                                // Convert to integer
                                if (int.TryParse(numericPart, out int districtNumber))
                                {
                                    // Format as D01 or D10
                                    string formattedDistrict = $"D{districtNumber:D2}";

                                    // Add to the list
                                    columnData[regKey].Add(formattedDistrict);
                                }
                            }
                            if (stringValue.Contains("North"))
                            {
                                break;
                            }
                        }
                    }


                    int startTransactionRow = 19;
                    int rowCounter = 0;
                    int srNo = 1;

                    foreach (var kvp in columnData)
                    {
                        foreach (var value in kvp.Value)
                        {
                            if (transactionSheet.Cells[startTransactionRow + rowCounter, 2].Text == "R01" || transactionSheet.Cells[startTransactionRow + rowCounter, 2].Text == "R02")
                            {
                                transactionSheet.InsertRow(startTransactionRow + rowCounter + 1, 1);
                                transactionSheet.Cells[startTransactionRow + rowCounter, 1, startTransactionRow + rowCounter, transactionSheet.Dimension.End.Column].Copy(transactionSheet.Cells[startTransactionRow + rowCounter + 1, 1]);
                            }

                            transactionSheet.Cells[startTransactionRow + rowCounter, 1].Value = srNo;
                            transactionSheet.Cells[startTransactionRow + rowCounter, 2].Value = value;
                            srNo++;
                            rowCounter++;
                        }

                        if (transactionSheet.Cells[startTransactionRow + rowCounter, 2].Text != "R01" && transactionSheet.Cells[startTransactionRow + rowCounter, 2].Text != "R02")
                        {
                            transactionSheet.DeleteRow(startTransactionRow + rowCounter, 1);
                        }

                        transactionSheet.Cells[startTransactionRow + rowCounter, 2].Value = kvp.Key;
                        rowCounter++;
                    }

                    List<RowData> northList = new List<RowData>();
                    List<RowData> southList = new List<RowData>();

                    for (int i = 1; i <= lastRow; i++)
                    {
                        var cellDate = worksheet.Cells[i, busDate].Value;
                        if (cellDate is DateTime dataValue)
                        {
                            string formattedDate = dataValue.ToString("MM/dd/yyyy");
                            if (formattedDate == date)
                            {
                                string region = worksheet.Cells[i, dist].Text;
                                string source = worksheet.Cells[i, 1].Text;
                                double store = Convert.ToDouble(worksheet.Cells[i, 2].Text);
                                double sales = Convert.ToDouble(worksheet.Cells[i, 4].Text);
                                double trans = Convert.ToDouble(worksheet.Cells[i, 5].Text);

                                RowData rowData = new RowData
                                {
                                    Region = region,
                                    Source = source,
                                    BusDate = dataValue,
                                    Store = store,
                                    Sales = sales,
                                    Transaction = trans
                                };

                                if (region.Contains("CJ"))
                                {
                                    northList.Add(rowData);
                                }
                                else
                                {
                                    southList.Add(rowData);
                                }
                            }
                        }
                    }

                    int columnRangeForSouth = 1;
                    int columnRangeForNorth = 1;
                    for (int i = 0; i < southList.Count; i++)
                    {
                        int row = columnRangeForSouth + lastRowForSouth;
                        southWorksheet.Cells[row, 1].Value = southList[i].Source;
                        southWorksheet.Cells[row, 2].Value = southList[i].Store;
                        southWorksheet.Cells[row, 3].Value = southList[i].BusDate;
                        southWorksheet.Cells[row, 4].Value = southList[i].Sales;
                        southWorksheet.Cells[row, 5].Value = southList[i].Transaction;
                        southWorksheet.Cells[row, 6].Formula = $"=VLOOKUP(B{row},'Site List'!P:Q,2,0)";
                        southWorksheet.Cells[row, 7].Formula = $"=VLOOKUP(B{row},'Site List'!P:W,8,0)";
                        southWorksheet.Cells[row, 8].Formula = $"=MONTH(C{row})&YEAR(C{row})";
                        southWorksheet.Cells[row, 9].Formula = $"=F{row}&H{row}";
                        southWorksheet.Cells[row, 10].Formula = $"=G{row}&H{row}";
                        southWorksheet.Cells[row, 11].Formula = $"=IFERROR(VLOOKUP(C{row},REF!A:B,2,0),\"\")";
                        southWorksheet.Cells[row, 12].Formula = $"=IF(K{row}=\"WTD\",F{row}&C{row},\"\")";
                        southWorksheet.Cells[row, 13].Formula = $"=IF(K{row}=\"WTD\",G{row}&C{row},\"\")";
                        southWorksheet.Cells[row, 14].Formula = $"=F{row}&K{row}";
                        southWorksheet.Cells[row, 15].Formula = $"=G{row}&K{row}";
                        southWorksheet.Cells[row, 16].Formula = $"=C{row}-365+1";
                        southWorksheet.Cells[row, 17].Formula = $"=SUMIFS('2022-2023'!E:E,'2022-2023'!B:B,B{row},'2022-2023'!C:C,P{row})";
                        southWorksheet.Cells[row, 18].Formula = $"=IFERROR(E{row}-Q{row},\"\")";
                        southWorksheet.Cells[row, 20].Formula = $"=YEAR(C{row})";
                        southWorksheet.Cells[row, 21].Formula = $"=MONTH(P{row})";
                        columnRangeForSouth++;


                    }

                    for (int i = 0; i < northList.Count; i++)
                    {
                        int row = columnRangeForNorth + lastRowForNorth;
                        northWorksheet.Cells[row, 1].Value = northList[i].Source;
                        northWorksheet.Cells[row, 2].Value = northList[i].Store;
                        northWorksheet.Cells[row, 3].Value = northList[i].BusDate;
                        northWorksheet.Cells[row, 4].Value = northList[i].Sales;
                        northWorksheet.Cells[row, 5].Value = northList[i].Transaction;
                        northWorksheet.Cells[row, 18].Formula = $"=VLOOKUP(B{row},'Site List'!P:Q,2,0)";
                        northWorksheet.Cells[row, 20].Formula = $"=MONTH(C{row})&YEAR(C{row})";
                        northWorksheet.Cells[row, 21].Formula = $"=R{row}&T{row}";
                        northWorksheet.Cells[row, 22].Formula = $"=S{row}&T{row}";
                        northWorksheet.Cells[row, 23].Formula = $"=IFERROR(VLOOKUP(C{row},REF!A:B,2,0),\"\")";
                        northWorksheet.Cells[row, 24].Formula = $"=IF(W{row}=\"WTD\",R{row}&C{row},\"\")";
                        northWorksheet.Cells[row, 25].Formula = $"=IF(W{row}=\"WTD\",S{row}&C31,\"\")";
                        northWorksheet.Cells[row, 26].Formula = $"=R{row}&W{row}";
                        northWorksheet.Cells[row, 27].Formula = $"=S{row}&W{row}";
                        northWorksheet.Cells[row, 28].Formula = $"=C{row}-365+1";
                        northWorksheet.Cells[row, 29].Formula = $"=SUMIFS('2022-2023'!E:E,'2022-2023'!B:B,B{row},'2022-2023'!C:C,AB{row})";
                        northWorksheet.Cells[row, 30].Formula = $"=IFERROR(E{row}-AC{row},\"\")";
                        northWorksheet.Cells[row, 33].Formula = $"=YEAR(C{row})";
                        columnRangeForNorth++;
                    }

                    excelPackage2.SaveAs(@"C:\Users\Nimap\source\repos\Test\Test\test3.xlsx");
                }
                Console.WriteLine("Closed");




            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

        }
    }
}
