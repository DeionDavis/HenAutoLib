using System;
using System.Configuration;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Data;
using System.Diagnostics;
using Spire.Xls;
using System.Drawing;
using System.IO;
using CommonLibrary.DataDrivenTesting;
using Microsoft.Win32;
using System.Globalization;

namespace TestExexution
{
    [TestClass]
    public class TestStart
    {
        public string allBatchReport = DateTime.Now.Day.ToString() + "_" + DateTime.Now.Month.ToString() + "_" + DateTime.Now.Year.ToString();
        public string path = ConfigurationManager.AppSettings["BatchReport"];
        string Init = ConfigurationManager.AppSettings["Initialization"];
        public static string BatchReportFile = string.Empty;

        /// <summary>
        /// This will create the report file based on the date and number of run. this will invoked when we run the test from the batch file
        /// This will create registry key name under "SOFTWARE\Testing" and store report file name, and test summery report will right in to that only.
        /// </summary>
        [TestMethod]
        public void StartExecution()
        {
            DataTable BatchTable = new DataTable();
            string directory = ConfigurationManager.AppSettings["BatchOptions"];
            BatchTable = DataReader.ExcelData(directory);
            Directory.CreateDirectory(path);
            int count = Directory.GetFiles(path, "Henkel_AllBatch_" + allBatchReport + "*.xlsx", SearchOption.AllDirectories).Length;
            BatchReportFile = path + "\\Henkel_AllBatch_" + allBatchReport + "RUN" + ++count + ".xlsx";
            RegistryKey key = Registry.CurrentUser.CreateSubKey(@"SOFTWARE\Testing");
            key.SetValue("AutomationReport", BatchReportFile);
            foreach (DataRow Batch in BatchTable.Rows)
            {
                if (Batch["Execution"].ToString() == "Yes")
                {
                    if (Batch["BatchName"].ToString() != "GR_StockReceipt" &&  Batch["BatchName"].ToString() != "Dispatch" && Batch["BatchName"].ToString() != "Manufacture" && Batch["BatchName"].ToString() != "Staging")
                    {
                        ReportFileGen(Batch["BatchName"].ToString());
                    }
                    DataTable TestTable = new DataTable();
                    TestTable = DataReader.ExcelData(Batch["TestcaseInformationRefference"].ToString());
                    foreach (DataRow TestCase in TestTable.Rows)
                    {
                        if (Batch["BatchName"].ToString() == "GR_StockReceipt" || Batch["BatchName"].ToString() == "Dispatch" || Batch["BatchName"].ToString() == "Manufacture" || Batch["BatchName"].ToString() == "Staging")
                        {
                            //if (Batch["Execution"].ToString() == "Yes")
                            if (TestCase["Execution Options"].ToString() == "Yes")
                            {
                                ReportFileGen(TestCase["TestCase Name"].ToString());
                            }
                        }
                        string testcase = TestCase["TestCase Name"].ToString();
                        string option = TestCase["Execution Options"].ToString();
                        string refference = TestCase["TestCase Refference"].ToString();
                        Process proc = new Process();
                        proc.StartInfo.WindowStyle = ProcessWindowStyle.Maximized;
                        proc.StartInfo.FileName = @"D:\HENKEL_AUTOMATION\Henkel_Controls\TestcaseDriver.bat";
                        proc.StartInfo.Arguments = String.Format("{0} {1} {2}", testcase, option, refference);
                        proc.Start();
                        proc.WaitForExit();
                        proc.Close();
                    }
                }
            }
            WriteTestCountSummary(BatchReportFile);
            Process Rproc = new Process();
            Rproc.StartInfo.WindowStyle = ProcessWindowStyle.Maximized;
            Rproc.StartInfo.FileName = @"C:\HybridFrameworkAutomation\Controls\TestcaseDriver.bat";
            Rproc.StartInfo.Arguments = String.Format("{0} {1} {2}", "ReportSending", "Yes", @"D:\HENKEL_AUTOMATION\Henkel_Projects\LibraryandExecution\AllReport\AllReport\bin\Debug\AllReport.dll");
            Rproc.Start();
            Rproc.WaitForExit();
            Rproc.Close();
            key.DeleteValue("AutomationReport");
        }

        /// <summary>
        /// this will create the file in the specified folder, and sheet with batch name.
        /// </summary>
        /// <param name="batchName"></param>
        public void ReportFileGen(string batchName)
        {
            Directory.CreateDirectory(path);
            ExcelDataTable.PopulateInCollection(Init + "\\GlobalElements.xlsx"); 
            bool availability = File.Exists(BatchReportFile);
            if (!availability)
            {
                CreateFile(BatchReportFile, batchName);
            }
            else
            {
                AddSheetFile(BatchReportFile, batchName);
            }
        }

        public void CreateFile(string filepath, string Batch)
        {
            Workbook BatchWorkBook = new Workbook();
            BatchWorkBook.ActiveSheet.Name = Batch;

            SummaryReportGen(filepath,Batch,BatchWorkBook);

            BatchWorkBook.LoadFromFile(filepath, ExcelVersion.Version2016);
            BatchWorkBook.ActiveSheet.Name = Batch;
            int rowIndex = 1;
            Worksheet sheet1 = BatchWorkBook.Worksheets[Batch];
            sheet1.Clear();
            sheet1.SetText(1, rowIndex, "ProjectName");
            sheet1.SetColumnWidth(rowIndex, 15);
            rowIndex++;
            sheet1.SetText(1, rowIndex, "StartTime");
            sheet1.SetColumnWidth(rowIndex, 15);
            rowIndex++;
            sheet1.SetText(1, rowIndex, "EndTime");
            sheet1.SetColumnWidth(rowIndex, 15);
            rowIndex++;
            sheet1.SetText(1, rowIndex, "MachineName");
            sheet1.SetColumnWidth(rowIndex, 15);
            rowIndex++;
            sheet1.SetText(1, rowIndex, "IP_Address");
            sheet1.SetColumnWidth(rowIndex, 15);
            rowIndex++;
            sheet1.SetText(1, rowIndex, "TestResult");
            sheet1.SetColumnWidth(rowIndex, 15);
            rowIndex++;
            sheet1.SetText(1, rowIndex, "ErrorScreens");
            sheet1.SetColumnWidth(rowIndex, 15);
            rowIndex++;
            sheet1.SetText(1, rowIndex, "FailureReason");
            sheet1.SetColumnWidth(rowIndex, 15);
            rowIndex++;
            sheet1.SetText(1, rowIndex, "File Name : Tab Name : Step No");
            sheet1.SetColumnWidth(rowIndex, 20);
            rowIndex++;
            sheet1.SetText(1, rowIndex, "Warning Screen");
            sheet1.SetColumnWidth(rowIndex, 15);
            rowIndex++;
            sheet1.SetText(1, rowIndex, "Warning Message");
            sheet1.SetColumnWidth(rowIndex, 15);
            rowIndex++;
            sheet1.SetText(1, rowIndex, "Duration");
            sheet1.SetColumnWidth(rowIndex, 15);


            BatchWorkBook.Worksheets[0].Activate();
            
            BatchWorkBook.SaveToFile(filepath, ExcelVersion.Version2016);
        }

        public void AddSheetFile(string filepath, string Batch)
        {
            Workbook BatchWorkBook = new Workbook();
            BatchWorkBook.ActiveSheet.Name = Batch;

            SummaryReportGen(filepath, Batch, BatchWorkBook);

            BatchWorkBook.LoadFromFile(filepath, ExcelVersion.Version2016);
            BatchWorkBook.ActiveSheet.Name = Batch;
            int rowIndex = 1;
            Worksheet sheet1 = BatchWorkBook.Worksheets[Batch];
            sheet1.Clear();
            sheet1.SetText(1, rowIndex, "ProjectName");
            sheet1.SetColumnWidth(rowIndex, 15);
            rowIndex++;
            sheet1.SetText(1, rowIndex, "StartTime");
            sheet1.SetColumnWidth(rowIndex, 15);
            rowIndex++;
            sheet1.SetText(1, rowIndex, "EndTime");
            sheet1.SetColumnWidth(rowIndex, 15);
            rowIndex++;
            sheet1.SetText(1, rowIndex, "MachineName");
            sheet1.SetColumnWidth(rowIndex, 15);
            rowIndex++;
            sheet1.SetText(1, rowIndex, "IP_Address");
            sheet1.SetColumnWidth(rowIndex, 15);
            rowIndex++;
            sheet1.SetText(1, rowIndex, "TestResult");
            sheet1.SetColumnWidth(rowIndex, 15);
            rowIndex++;
            sheet1.SetText(1, rowIndex, "ErrorScreens");
            sheet1.SetColumnWidth(rowIndex, 15);
            rowIndex++;
            sheet1.SetText(1, rowIndex, "FailureReason");
            sheet1.SetColumnWidth(rowIndex, 15);
            rowIndex++;
            sheet1.SetText(1, rowIndex, "File Name : Tab Name : Step No");
            sheet1.SetColumnWidth(rowIndex, 20);
            rowIndex++;
            sheet1.SetText(1, rowIndex, "Warning Screen");
            sheet1.SetColumnWidth(rowIndex, 15);
            rowIndex++;
            sheet1.SetText(1, rowIndex, "Warning Message");
            sheet1.SetColumnWidth(rowIndex, 15);
            rowIndex++;
            sheet1.SetText(1, rowIndex, "Duration");
            sheet1.SetColumnWidth(rowIndex, 15);

            BatchWorkBook.Worksheets[Batch].Activate();
        
            BatchWorkBook.SaveToFile(filepath, ExcelVersion.Version2016);
        }

        public void SummaryReportGen(string filepath, string batch, Workbook Workbook)
        {
            #region new
            ExcelFont heading = Workbook.CreateFont();
            heading.IsBold = true;
            heading.Size = 12;
            heading.FontName = "Calibri";

            bool availability = File.Exists(filepath);
            if (availability)
            {
                Workbook.LoadFromFile(filepath, ExcelVersion.Version2016);
                Worksheet sheet1 = Workbook.Worksheets["Summary Report"];
                if (sheet1 != null)
                {
                    WriteBatchDetails(filepath, batch, Workbook);
                }
                else
                {
                    throw new Exception("Summary Report Page not generated...");
                }
            }
            else
            {
                Workbook.ActiveSheet.Name = "Summary Report";
                Worksheet sheet1 = Workbook.Worksheets["Summary Report"];
                int rowIndex = 1;
                CellRange range = sheet1.Range[1, rowIndex];
                RichText texthead = sheet1.Range[1, rowIndex].RichText;
                texthead.Text = "Batch";
                texthead.SetFont(0, texthead.Text.Length - 1, heading);
                range.Style.HorizontalAlignment = HorizontalAlignType.Center;
                range.Style.Color = Color.AliceBlue;
                range.BorderAround(LineStyleType.Thin, Color.Black);
                sheet1.SetColumnWidth(rowIndex, 40);
                rowIndex++;
                texthead = sheet1.Range[1, rowIndex].RichText;
                texthead.Text = "Total Test case";
                texthead.SetFont(0, texthead.Text.Length - 1, heading);
                range = sheet1.Range[1, rowIndex];
                range.Style.HorizontalAlignment = HorizontalAlignType.Center;
                range.Style.Color = Color.AliceBlue;
                range.BorderAround(LineStyleType.Thin, Color.Black);
                sheet1.SetColumnWidth(rowIndex, 20);
                rowIndex++;
                texthead = sheet1.Range[1, rowIndex].RichText;
                texthead.Text = "Passed";
                texthead.SetFont(0, texthead.Text.Length - 1, heading);
                range = sheet1.Range[1, rowIndex];
                range.Style.HorizontalAlignment = HorizontalAlignType.Center;
                range.Style.Color = Color.AliceBlue;
                range.BorderAround(LineStyleType.Thin, Color.Black);
                sheet1.SetColumnWidth(rowIndex, 10);
                rowIndex++;
                texthead = sheet1.Range[1, rowIndex].RichText;
                texthead.Text = "Failed";
                texthead.SetFont(0, texthead.Text.Length - 1, heading);
                range = sheet1.Range[1, rowIndex];
                range.Style.HorizontalAlignment = HorizontalAlignType.Center;
                range.Style.Color = Color.AliceBlue;
                range.BorderAround(LineStyleType.Thin, Color.Black);
                sheet1.SetColumnWidth(rowIndex, 10);
                rowIndex++;
                /////////
                texthead = sheet1.Range[1, rowIndex].RichText;
                texthead.Text = "Total Duration ";
                texthead.SetFont(0, texthead.Text.Length - 1, heading);
                range = sheet1.Range[1, rowIndex];
                range.Style.HorizontalAlignment = HorizontalAlignType.Center;
                range.Style.Color = Color.AliceBlue;
                range.BorderAround(LineStyleType.Thin, Color.Black);
                sheet1.SetColumnWidth(rowIndex, 20);
                rowIndex++;
                //////////
                texthead = sheet1.Range[1, rowIndex].RichText;
                texthead.Text = "Comments";
                texthead.SetFont(0, texthead.Text.Length - 1, heading);
                range = sheet1.Range[1, rowIndex];
                range.Style.HorizontalAlignment = HorizontalAlignType.Center;
                range.Style.Color = Color.AliceBlue;
                range.BorderAround(LineStyleType.Thin, Color.Black);
                sheet1.SetColumnWidth(rowIndex, 40);

                Worksheet sheet2 = Workbook.Worksheets[1];
                sheet2.Remove();

                Worksheet sheet3 = Workbook.Worksheets[1];
                sheet3.Remove();

                Workbook.Worksheets[0].Activate();
                Workbook.SaveToFile(filepath, ExcelVersion.Version2016);;
                WriteBatchDetails(filepath, batch, Workbook);
            }
            #endregion
        }

        public void WriteBatchDetails(string filepath,string batchname, Workbook workbook)
        {
            DataTable dt = new DataTable();
            workbook.LoadFromFile(filepath, ExcelVersion.Version2016);
            Worksheet Sheet = workbook.Worksheets["Summary Report"];

            ExcelFont heading = workbook.CreateFont();
            heading.IsBold = false;
            heading.Size = 12;
            heading.FontName = "Calibri";
            heading.Color = Color.Black;

            ExcelFont Fail = workbook.CreateFont();
            Fail.IsBold = false;
            Fail.Size = 12;
            Fail.FontName = "Calibri";
            Fail.Color = Color.Red;

            if (Sheet != null)
            {
                dt = Sheet.ExportDataTable();
                int rowIndex = 0;
                rowIndex = dt.Rows.Count + 2;
                int columeIndex = 1;

                CellRange range = Sheet.Range[rowIndex, columeIndex];

                HyperLink link = Sheet.HyperLinks.Add(range);
                link.Type = HyperLinkType.Workbook;
                link.Address = "#" + batchname + "!" + "B1";

                RichText textData = Sheet.Range[rowIndex,columeIndex].RichText;
                textData.Text = batchname;
                textData.SetFont(0, textData.Text.Length - 1, heading);
                range.BorderAround(LineStyleType.Thin, Color.Black);
                columeIndex++;
                range = Sheet.Range[rowIndex, columeIndex];
                range.Style.HorizontalAlignment = HorizontalAlignType.Center;
                textData = Sheet.Range[rowIndex, columeIndex].RichText;
                textData.Text = "0";
                textData.SetFont(0, textData.Text.Length - 1, heading);
                range.BorderAround(LineStyleType.Thin, Color.Black);
                columeIndex++;
                range = Sheet.Range[rowIndex, columeIndex];
                range.Style.HorizontalAlignment = HorizontalAlignType.Center;
                textData = Sheet.Range[rowIndex, columeIndex].RichText;
                textData.Text = "0";
                textData.SetFont(0, textData.Text.Length - 1, heading);
                range.BorderAround(LineStyleType.Thin, Color.Black);
                columeIndex++;
                range = Sheet.Range[rowIndex, columeIndex];
                range.Style.HorizontalAlignment = HorizontalAlignType.Center;
                textData = Sheet.Range[rowIndex, columeIndex].RichText;
                textData.Text = "0";
                textData.SetFont(0, textData.Text.Length - 1, Fail);
                range.BorderAround(LineStyleType.Thin, Color.Black);
                columeIndex++;
                ////////////////
                range = Sheet.Range[rowIndex, columeIndex];
                range.Style.HorizontalAlignment = HorizontalAlignType.Center;
                textData = Sheet.Range[rowIndex, columeIndex].RichText;
                textData.Text = "00:00:00";
                textData.SetFont(0, textData.Text.Length - 1, heading);
                range.BorderAround(LineStyleType.Thin, Color.Black);
                columeIndex++;
                ///////////////////////
                range = Sheet.Range[rowIndex, columeIndex];
                range.Style.HorizontalAlignment = HorizontalAlignType.Center;
                textData = Sheet.Range[rowIndex, columeIndex].RichText;
                textData.Text = "";
                range.BorderAround(LineStyleType.Thin, Color.Black);

                workbook.Worksheets[0].Activate();
                workbook.SaveToFile(filepath, ExcelVersion.Version2016);
            }
            else
            {
                throw new Exception("Summary Report Sheet Not Generated");
            }
        }

        public void WriteTestCountSummary(string filepath)
        {
            DataTable dt = new DataTable();
            Workbook workbook = new Workbook();
            workbook.LoadFromFile(filepath, ExcelVersion.Version2016);
            Worksheet Sheet = workbook.Worksheets["Summary Report"];

            int totalPTest = 0;
            int totalFTest = 0;
            int totalTTest = 0;
            string  totalDuration;
            TimeSpan gradndTotal = TimeSpan.Parse("00:00:00");

            ExcelFont heading = workbook.CreateFont();
            heading.IsBold = false;
            heading.Size = 12;
            heading.FontName = "Calibri";
            heading.Color = Color.Black;

            ExcelFont Fail = workbook.CreateFont();
            Fail.IsBold = false;
            Fail.Size = 12;
            Fail.FontName = "Calibri";
            Fail.Color = Color.Red;
            if (Sheet != null)
            {
                dt = Sheet.ExportDataTable();
                int rowIndex = 0;
                rowIndex = dt.Rows.Count + 2;
                int columeIndex = 1;

                CellRange range = Sheet.Range[rowIndex, columeIndex];

                RichText textData = Sheet.Range[rowIndex, columeIndex].RichText;
                textData.Text = "Total";
                textData.SetFont(0, textData.Text.Length - 1, heading);
                range.BorderAround(LineStyleType.Thin, Color.Black);
                columeIndex++;

                for(int i=2; i<=rowIndex; i++)
                {
                    totalTTest = totalTTest + Convert.ToInt32(Sheet.GetText(i, 2));
                    totalPTest = totalPTest + Convert.ToInt32(Sheet.GetText(i, 3));
                    totalFTest = totalFTest + Convert.ToInt32(Sheet.GetText(i, 4));     
                }
                for (int i = 2; i < rowIndex; i++)
                {
                    totalDuration = Sheet.GetText(i, 5).ToString();
                    int Secs = Convert.ToInt16(totalDuration.Split(':')[2]);
                    int min = Convert.ToInt16(totalDuration.Split(':')[1]);
                    int hr = Convert.ToInt16(totalDuration.Split(':')[0]);
                    TimeSpan ts = new TimeSpan(hr, min, Secs);
                    gradndTotal = gradndTotal.Add(ts);
                }

                range = Sheet.Range[rowIndex, columeIndex];
                range.Style.HorizontalAlignment = HorizontalAlignType.Center;
                textData = Sheet.Range[rowIndex, columeIndex].RichText;
                textData.Text = totalTTest.ToString();
                textData.SetFont(0, textData.Text.Length - 1, heading);
                range.BorderAround(LineStyleType.Thin, Color.Black);
                columeIndex++;

                range = Sheet.Range[rowIndex, columeIndex];
                range.Style.HorizontalAlignment = HorizontalAlignType.Center;
                textData = Sheet.Range[rowIndex, columeIndex].RichText;
                textData.Text = totalPTest.ToString();
                textData.SetFont(0, textData.Text.Length - 1, heading);
                range.BorderAround(LineStyleType.Thin, Color.Black);
                columeIndex++;

                range = Sheet.Range[rowIndex, columeIndex];
                range.Style.HorizontalAlignment = HorizontalAlignType.Center;
                textData = Sheet.Range[rowIndex, columeIndex].RichText;
                textData.Text = totalFTest.ToString();
                textData.SetFont(0, textData.Text.Length - 1, Fail);
                range.BorderAround(LineStyleType.Thin, Color.Black);
                columeIndex++;




                range = Sheet.Range[rowIndex, columeIndex];
                range.Style.HorizontalAlignment = HorizontalAlignType.Center;
                textData = Sheet.Range[rowIndex, columeIndex].RichText;
                textData.Text = string.Format("{0:D2}:{1:D2}:{2:D2}", gradndTotal.Hours, gradndTotal.Minutes, gradndTotal.Seconds);
                textData.SetFont(0, textData.Text.Length - 1, heading);
                range.BorderAround(LineStyleType.Thin, Color.Black);
                columeIndex++;

                range = Sheet.Range[rowIndex, columeIndex];
                range.Style.HorizontalAlignment = HorizontalAlignType.Center;
                textData = Sheet.Range[rowIndex, columeIndex].RichText;
                textData.Text = "";
                range.BorderAround(LineStyleType.Thin, Color.Black);


                workbook.Worksheets[0].Activate();
                workbook.SaveToFile(filepath, ExcelVersion.Version2016);

            }
            else
            {
                throw new Exception("Summary Report Sheet Not Generated");
            }

        }
    }
}
