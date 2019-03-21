using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Spire.Xls;
using System.Configuration;
using Microsoft.Win32;
using System.Globalization;

namespace OperationLibrary.Report
{
    public class Report
    {
        UploadScreenshot uploadScreenshot = new UploadScreenshot();
        public string SingleReport = DateTime.Now.Day.ToString() + "_" + DateTime.Now.Month.ToString() + "_" + DateTime.Now.Year.ToString() + "@" + DateTime.Now.ToString("HH:mm:ss").Replace(':', '_');
        string GradndTotal;
        /// <summary>
        /// This will generate the report summery and rite each report in the specified batch
        /// </summary>
        /// <param name="ProjectName">Name of the project or test case.</param>
        /// <param name="StartTime">Start time of the  test</param>
        /// <param name="EndTime">End time of the test</param>
        /// <param name="result">Result of the test, that pass or fail.</param>
        /// <param name="HostName">This will get the name of the PC that you running</param>
        /// <param name="ipaddress">This will get the IP of the PC that you running</param>
        /// <param name="Batch">Batch name of the test case.</param>
        public void ReportGrneration(string ProjectName, string StartTime, string EndTime, string result, string HostName, string ipaddress, string Batch)
        {
            string filepath = string.Empty;
            RegistryKey key = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Testing");
            if (key.GetValue("AutomationReport") != null)
            {
                filepath = key.GetValue("AutomationReport").ToString();
            }
            if (filepath != string.Empty)
            {
                bool availability = File.Exists(filepath);
                if (availability)
                {
                    Workbook WorkBook = new Workbook();
                    WorkBook.LoadFromFile(filepath, ExcelVersion.Version2016);
                    Worksheet Sheet = WorkBook.Worksheets[Batch];
                    Worksheet SummarySheet = WorkBook.Worksheets["Summary Report"];
                    if (SummarySheet != null)
                    {
                        string totalDuration = DateTime.Parse(EndTime).Subtract(DateTime.Parse(StartTime)).ToString();
                        updateSummary(SummarySheet, WorkBook, filepath,result,totalDuration);
                    }
                    if (Sheet != null)
                    {
                        SetFirstData(Sheet, WorkBook, filepath, ProjectName, StartTime, EndTime, result, HostName, ipaddress);
                    }
                    else
                    { Console.WriteLine("Sheet name is not proper"); }
                }
                else
                {
                    Console.WriteLine("Report File not found");
                }
            }
            else { }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sheet">reference to the sheet</param>
        /// <param name="book">reference to the excel sheet</param>
        /// <param name="filepath">path of the report file</param>
        /// <param name="ProjectName">Name of the project or test case.</param>
        /// <param name="StartTime">Start time of the  test</param>
        /// <param name="EndTime">End time of the test</param>
        /// <param name="result">Result of the test, that pass or fail.</param>
        /// <param name="HostName">This will get the name of the PC that you running</param>
        /// <param name="ipaddress">This will get the IP of the PC that you running</param>
        public void SetFirstData(Worksheet sheet, Workbook book, string filepath, string ProjectName, string StartTime, string EndTime, string result, string HostName, string ipaddress)
        {
            DataTable dt = new DataTable();
          
            dt = sheet.ExportDataTable();
            int rowIndex = 0;
            rowIndex = dt.Rows.Count + 2;
            int columeIndex = 1;
            int errorcount = Operation.ErrorScreenPath.Count();
            sheet.SetText(rowIndex, columeIndex, ProjectName);
            columeIndex++;
            sheet.SetText(rowIndex, columeIndex, StartTime);
            columeIndex++;
            sheet.SetText(rowIndex, columeIndex, EndTime);
            columeIndex++;
            sheet.SetText(rowIndex, columeIndex, HostName);
            columeIndex++;
            sheet.SetText(rowIndex, columeIndex, ipaddress);
            columeIndex++;
            sheet.SetText(rowIndex, columeIndex, result);
            columeIndex++;
            //Setting Error Screen and message if available.
            if (Operation.ErrorScreenPath != " ")
            {
                string ViewLink = uploadScreenshot.UploadImage(Operation.ErrorScreenPath);
                CellRange Range = sheet.Range[rowIndex, columeIndex];
                HyperLink link = sheet.HyperLinks.Add(Range);
                link.Type = HyperLinkType.File;
                link.Address = ViewLink;
               
            }
            columeIndex++;
            sheet.SetText(rowIndex, columeIndex, Operation.FailerReason);
            columeIndex++;
            /// Writting Failed Sheet name, tab name and step number
            if (result == "Failed")
            {
                /// Writting Failed Sheet name, tab name and step number
                if (Operation.CurrentOperation == "Operations")
                {
                    sheet.SetValue(rowIndex, columeIndex, Operation.ForeignSheetName + " : " + Operation.ForeigntabName + " : " + Operation.ForeignStepNumber);
                    columeIndex++;
                }
                /// Writting Failed tab name and step number
                else
                {
                    sheet.SetValue(rowIndex, columeIndex, Operation.tabName + " : " + Operation.StepNumber);
                    columeIndex++;
                }
            }
            else
            {
                columeIndex++;
            }
           
            //Setting Warning Screen and message if available.
            if (Operation.warningMessage != " ")
            {
                string ViewLink = uploadScreenshot.UploadImage(Operation.warningScreenPath);
               CellRange Range = sheet.Range[rowIndex, columeIndex];
               HyperLink link = sheet.HyperLinks.Add(Range);
                link.Type = HyperLinkType.File;
                link.Address = ViewLink;
            }
            columeIndex++;
            sheet.SetText(rowIndex, columeIndex, Operation.warningMessage);
            columeIndex++;
            string totalDuration = DateTime.Parse(EndTime).Subtract(DateTime.Parse(StartTime)).ToString();
            sheet.SetValue(rowIndex, columeIndex, totalDuration);
            columeIndex++;
            book.ActiveSheetIndex.Equals(0);
            book.Worksheets[0].Activate();
            book.SaveToFile(filepath, ExcelVersion.Version2016);
        }



        public void updateSummary(Worksheet sheet, Workbook book, string filepath, string result,string totalDuration)
        {
            DataTable dt = new DataTable();
            dt = sheet.ExportDataTable();
            int rowIndex = 0;
            rowIndex = dt.Rows.Count + 1;
            int columeIndex = 2;
            int totaltest = Convert.ToInt32(sheet.GetText(rowIndex, columeIndex));
            totaltest = totaltest + 1;
            sheet.SetText(rowIndex, columeIndex, totaltest.ToString());
            columeIndex++;
            int passedtest = Convert.ToInt32(sheet.GetText(rowIndex, columeIndex));
            if (result == "Passed")
            {
                passedtest = passedtest + 1;
            }
            sheet.SetText(rowIndex, columeIndex, passedtest.ToString());
            columeIndex++;
            int FailedTest = Convert.ToInt32(sheet.GetText(rowIndex, columeIndex));
            if (result == "Failed")
            {
                FailedTest = FailedTest + 1;
            }
            sheet.SetText(rowIndex, columeIndex, FailedTest.ToString());
            columeIndex++;
            TimeSpan BDuration = TimeSpan.Parse(sheet.GetText(rowIndex, columeIndex));
            if (totalDuration != null)
            {
                CultureInfo us = new CultureInfo("en-US");
                int Secs = Convert.ToInt16(totalDuration.Split(':')[2]);
                int min = Convert.ToInt16(totalDuration.Split(':')[1]);
                int hr = Convert.ToInt16(totalDuration.Split(':')[0]);
                TimeSpan ts = new TimeSpan(hr, min, Secs);
                BDuration = BDuration.Add(ts);
                GradndTotal = string.Format("{0:D2}:{1:D2}:{2:D2}", BDuration.Hours, BDuration.Minutes, BDuration.Seconds);
            }
            sheet.SetText(rowIndex, columeIndex, GradndTotal);
            columeIndex++;
            book.ActiveSheetIndex.Equals(0);
            book.Worksheets[0].Activate();
            book.SaveToFile(filepath, ExcelVersion.Version2016);
            
        }
    }
}
 