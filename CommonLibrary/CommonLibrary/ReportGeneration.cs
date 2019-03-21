using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using System.IO;
using System.Reflection;
using ExcelLibrary.SpreadSheet;
using System.Data;
using CommonLibrary.KeywordDrivenTesting;
using CommonLibrary.Log;

namespace CommonLibrary.Reports
{
    public class ReportGeneration
    {
        public string ReportDate = DateTime.Now.Day.ToString() + "_" + DateTime.Now.Month.ToString() + "_" + DateTime.Now.Year.ToString() + "@" + DateTime.Now.ToString("HH:mm:ss").Replace(':', '_');
        public string path = ConfigurationManager.AppSettings["Report"];
        public static string DetailedReportFilePath = string.Empty;

        public void Reports(string stepno, string description, string keywords, bool status, string batch, string DetaildReportStatus, string recordeddata)
        {
            Directory.CreateDirectory(path);
            if (DetaildReportStatus == "Yes")
            {
                string filename = LoginOperatrion.ProjectName + "_" + ReportDate + ".xls";
                DetailedReportFilePath = path + "\\" + filename;
                bool availability = File.Exists(DetailedReportFilePath);
                string statusTxt = string.Empty;
                if (status)
                {
                    statusTxt = "Pass";
                    if (!availability)
                    {
                        CreateFile(stepno, description, keywords, statusTxt, batch, DetailedReportFilePath, recordeddata);
                    }
                    else
                    {
                        UpdateData(stepno, description, keywords, statusTxt, batch, DetailedReportFilePath, recordeddata);
                    }
                }
                else
                {
                    statusTxt = "Fail";
                    RecordFailReport(stepno, description, keywords, statusTxt, batch, DetailedReportFilePath, recordeddata);
                }
            }
        }
        public void CreateFile(string Stepno, string Description, string Keywords, string Status, string Batch, string Filepath, string RecordedData)
        {
            try
            {
                int columnNum = 0;
                Workbook wBook = new Workbook();
                Worksheet wSheet = new Worksheet(Batch + "_" +  LoginOperatrion.requiredPlant);
                wBook.Worksheets.Add(wSheet);

                wSheet.Cells[0, columnNum] = new Cell("Step No.");
                columnNum++;
                wSheet.Cells[0, columnNum] = new Cell("Description");
                columnNum++;
                wSheet.Cells[0, columnNum] = new Cell("Keyword");
                columnNum++;
                wSheet.Cells[0, columnNum] = new Cell("Status");
                columnNum++;
                wSheet.Cells[0, columnNum] = new Cell("Recorded Data");
                columnNum++;

                wSheet.Cells.ColumnWidth[0, (ushort)columnNum] = 5000;
                wSheet.Cells.ColumnWidth[0, (ushort)columnNum] = 5000;
                wSheet.Cells.ColumnWidth[0, (ushort)columnNum] = 5000;
                wSheet.Cells.ColumnWidth[0, (ushort)columnNum] = 5000;

                SetFirstData(wSheet, wBook, Stepno, Description, Keywords, Status, Batch, Filepath, RecordedData);
            }
            catch (Exception e) { }
        }

        public void AddNewSheet(Workbook wBook, Worksheet wSheet, string Batch, string _filepath)
        {
            try
            {
                int columnNum = 0;
                wBook.Worksheets.Add(wSheet);

                wSheet.Cells[0, columnNum] = new Cell("Step No.");
                columnNum++;
                wSheet.Cells[0, columnNum] = new Cell("Description");
                columnNum++;
                wSheet.Cells[0, columnNum] = new Cell("Keyword");
                columnNum++;
                wSheet.Cells[0, columnNum] = new Cell("Status");
                columnNum++;
                wSheet.Cells[0, columnNum] = new Cell("Recorded Data");
                columnNum++;

                wSheet.Cells.ColumnWidth[0, (ushort)columnNum] = 5000;
                wSheet.Cells.ColumnWidth[0, (ushort)columnNum] = 5000;
                wSheet.Cells.ColumnWidth[0, (ushort)columnNum] = 5000;
                wSheet.Cells.ColumnWidth[0, (ushort)columnNum] = 5000;
                wBook.Save(_filepath);


            }
            catch (Exception e) { }
        }

        public void SetFirstData(Worksheet wSheet, Workbook wBook, string _Stepno, string _Description, string _Keywords, string _Status, string _Batch, string _Filepath, string _RecordedData)
        {
            try
            {
                int columnIndex = 0;
                int rowNum = 1;

                columnIndex = 0;
                wSheet.Cells[rowNum, columnIndex] = new Cell(_Stepno);
                columnIndex++;
                wSheet.Cells[rowNum, columnIndex] = new Cell(_Description);
                columnIndex++;
                wSheet.Cells[rowNum, columnIndex] = new Cell(_Keywords);
                columnIndex++;
                wSheet.Cells[rowNum, columnIndex] = new Cell(_Status);
                columnIndex++;
                wSheet.Cells[rowNum, columnIndex] = new Cell(_RecordedData);
                columnIndex++;

                wBook.Save(_Filepath);
                
            }
            catch (Exception e) { }
        }

        public void UpdateData(string _Stepno, string _Description, string _Keywords, string _Status, string _Batch, string _Filepath, string _RecordedData)
        {
            try
            {
                Workbook loadBook = Workbook.Load(_Filepath);
                Worksheet loadSheet = new Worksheet(_Batch + "_" + LoginOperatrion.requiredPlant);

                DataTable dtReport = new DataTable();
                dtReport.Clear();
                dtReport = ExcelKeywordTable.ExcelData(_Filepath, _Batch + "_" + LoginOperatrion.requiredPlant);
                
                if (dtReport == null)
                {
                    AddNewSheet(loadBook,loadSheet,_Batch,_Filepath);
                    dtReport = ExcelKeywordTable.ExcelData(_Filepath, _Batch + "_" + LoginOperatrion.requiredPlant);
                }
                int e = loadBook.Worksheets.Count-1;

                loadSheet = loadBook.Worksheets[e];

                int rowindex = dtReport.Rows.Count;
                int columindex = 0;
                loadSheet.Cells[rowindex + 1, columindex] = new Cell(_Stepno);
                columindex++;
                loadSheet.Cells[rowindex + 1, columindex] = new Cell(_Description);
                columindex++;
                loadSheet.Cells[rowindex + 1, columindex] = new Cell(_Keywords);
                columindex++;
                loadSheet.Cells[rowindex + 1, columindex] = new Cell(_Status);
                columindex++;
                loadSheet.Cells[rowindex + 1, columindex] = new Cell(_RecordedData);


                loadBook.Save(_Filepath);

            }
            catch (Exception e) { }
        }

        public void RecordFailReport(string _Stepno, string _Description, string _Keywords, string _Status, string _Batch, string _Filepath, string _RecordedData)
        {
            try
            {
                Workbook loadBook = Workbook.Load(_Filepath);
                Worksheet loadSheet = new Worksheet(_Batch + "_" + LoginOperatrion.requiredPlant);
                DataTable dtReport = new DataTable();
                dtReport.Clear();
                dtReport = ExcelKeywordTable.ExcelData(_Filepath, _Batch + "_" + LoginOperatrion.requiredPlant);

                if (dtReport == null)
                {
                    AddNewSheet(loadBook, loadSheet, _Batch, _Filepath);
                    dtReport = ExcelKeywordTable.ExcelData(_Filepath, _Batch + "_" + LoginOperatrion.requiredPlant);
                }
                int e = loadBook.Worksheets.Count - 1;
                loadSheet = loadBook.Worksheets[e];

                int rowindex = dtReport.Rows.Count;
                int columindex = 0;

                loadSheet.Cells[rowindex + 1, columindex] = new Cell(_Stepno);
                columindex++;
                loadSheet.Cells[rowindex + 1, columindex] = new Cell(_Description);
                columindex++;
                loadSheet.Cells[rowindex + 1, columindex] = new Cell(_Keywords);
                columindex++;
                loadSheet.Cells[rowindex + 1, columindex] = new Cell(_Status);
                columindex++;
                loadSheet.Cells[rowindex + 1, columindex] = new Cell(_RecordedData);

                if (_Keywords != "Wait" && _Keywords != "SearchDataOperation")
                {
                    if (rowindex <= 80)
                    {
                        int rows = 100 - rowindex;
                        for (int i = 0; i < rows; i++)
                        {
                            columindex = 25;
                            loadSheet.Cells[rowindex, columindex] = new Cell(" ");
                            rowindex++;
                        }
                    }
                }
                loadBook.Save(_Filepath);

            }
            catch (Exception e) { }
        }

        public void FileCorreptionCheck()
        {
            try
            {
                if(new System.IO.FileInfo(DetailedReportFilePath).Length < 8192)
                {
                    DataTable dtReport = new DataTable();
                    dtReport.Clear();
                    dtReport = ExcelKeywordTable.ExcelData(DetailedReportFilePath, LoginOperatrion.batchforReport+"_"+LoginOperatrion.requiredPlant);

                    Workbook loadBook = Workbook.Load(DetailedReportFilePath);
                    Worksheet loadSheet = loadBook.Worksheets[0];

                    int rowindex = dtReport.Rows.Count;
                    int columindex = 0;

                    if (rowindex <= 80)
                    {
                        int rows = 100 - rowindex;
                        for (int i = 0; i < rows; i++)
                        {
                            columindex = 25;
                            loadSheet.Cells[rowindex, columindex] = new Cell(" ");
                            rowindex++;
                        }
                    }
                    loadBook.Save(DetailedReportFilePath);
                }
            }
            catch (Exception e) { }
        }
    }
}
