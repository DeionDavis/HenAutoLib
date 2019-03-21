using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using System.IO;
using ExcelLibrary.SpreadSheet;
using System.Data;
using OperationLibrary;
using CommonLibrary.KeywordDrivenTesting;
using CommonLibrary.DataDrivenTesting;


namespace CommonLibrary.Writedata
{
    public class WriteAndReadData
    {
        public string DataPath = ConfigurationManager.AppSettings["ExcelDataFile"] + "\\DynamicData";
        public static string DataFilePath = string.Empty;
        public string writeRowIndex = string.Empty;
        public int writecolumeIndex = 0;
        public static int refcount = 0;
        Workbook wBook = new Workbook();
        Worksheet wSheet = new Worksheet("sheet1");

        public void WriteExcel(string ColumeHead, string WriteData, string Operation, string referenceData)
        {
            Directory.CreateDirectory(DataPath);
            string Datafilename = ConfigurationManager.AppSettings["Batch"] + ".xls";
            DataFilePath = DataPath + "\\" + Datafilename;
            bool availability = File.Exists(DataFilePath);

            if (!availability)
            {
                CreateDynamicFile(ColumeHead);
                AddData(ColumeHead,WriteData);
            }
            else
            {
                if(Operation== "SetRef_Data")
                {
                    AddReferenceData(ColumeHead, WriteData);
                }
                else if(Operation == "ReadData")
                {
                    FileCorreptionCheck();
                    AddDataForReference(DataFilePath, ColumeHead, WriteData, referenceData);
                }
            }
        }
        public void CreateDynamicFile(string _ColumeHead)
        { 
            try
            {
                int columnNum = 0;
                wBook.Worksheets.Add(wSheet);
                wSheet.Cells[0, columnNum] = new Cell(_ColumeHead);;
                wSheet.Cells[0, columnNum + 1] = new Cell("Order");
                wSheet.Cells[1, columnNum + 1] = new Cell("0");
                wSheet.Cells.ColumnWidth[0, (ushort)columnNum] = 5000;
                wBook.Save(DataFilePath);
            }
            catch(Exception e)
            {
                throw new Exception("Error While Creating File For Writing Reference Data");
            }
        }
        public void AddData(string ColumHead,string WriteData)
        {
            DataTable dtReport = new DataTable();
            dtReport.Clear();
            dtReport = ExcelKeywordTable.ExcelData(DataFilePath, "sheet1");
            getColumeIndexData(dtReport.Columns.Count, ColumHead);
            wSheet.Cells[1, writecolumeIndex] = new Cell(WriteData);
            refcount++;
            wBook.Save(DataFilePath);
        }
        public void AddReferenceData(string _ColumeHead, string _writedata)
        {
            DataTable dtReport = new DataTable();
            int columindex = 0;
            try
            {
                if (refcount == 0)
                {
                    
                    dtReport.Clear();
                    dtReport = ExcelKeywordTable.ExcelData(DataFilePath, "sheet1");
                    if (dtReport.Rows.Count >= 1)
                    {
                        CreateDynamicFile(_ColumeHead);
                        AddData(_ColumeHead, _writedata);
                    }
                    refcount++;
                }
                else
                {
                    dtReport.Clear();
                    dtReport = ExcelKeywordTable.ExcelData(DataFilePath, "sheet1");
                    int rowindex = dtReport.Rows.Count;
                    wSheet.Cells[rowindex + 1, columindex] = new Cell(_writedata);
                    wSheet.Cells[rowindex + 1, columindex + 1] = new Cell("0");
                    wBook.Save(DataFilePath);
                }
            }
            catch (Exception e)
            {
                throw new Exception("Error while writing reference data to the File.");
            }
        }
        public void AddDataForReference(string _Filepath, string _columhead, string _writeData, string referencedata)
        {
            try
            {
                DataTable dtReport = new DataTable();
                dtReport.Clear();
                dtReport = ExcelKeywordTable.ExcelData(_Filepath, "sheet1");
                int columeindex = dtReport.Columns.Count;
                int rowindex = dtReport.Rows.Count;
                int referenceDataCount = 0;
                string material = string.Empty;
                var Wbook = Workbook.Load(DataFilePath);
                var Wsheet = Wbook.Worksheets[0];

                #region RR
                if (referencedata.Contains(','))
                {
                    string[] MatchData = referencedata.Split(',');
                    if (MatchData.Count() > 1)
                    {
                        if (MatchData[0].StartsWith("Rec_"))
                        {
                            material = Operation.recordedData[MatchData[0].Replace("Rec_", string.Empty)].ToString();
                        }
                        else if (MatchData[0].StartsWith("Order"))
                        {
                            material = ExcelDataTable.ReadRefferenceInfoByOrder(MatchData[0].Split('=')[1],"Material");
                        }
                        else
                        {
                            material = ExcelDataTable.ReadData(1, MatchData[0]);
                        }

                        if (MatchData[1].Contains('+'))
                        {
                            string Store = ExcelDataTable.ReadData(1, MatchData[1].Split('+')[0]);
                            referencedata = material + ':' + ExcelDataTable.ReadRefferenceInfoByMaterialWithStore(Store, material, MatchData[1].Split('=')[1]);
                        }
                        else if (MatchData[1].StartsWith("Store"))
                        {
                            referencedata =  ExcelDataTable.ReadData(1, MatchData[1].Split('=')[0]) + ':' + material; 
                        }
                        getRefferenceIndexData(rowindex, referencedata);
                    }
                    else
                    {
                        referencedata = ExcelDataTable.ReadData(1, MatchData[0]);
                    }
                }
                #endregion

                getColumeIndexData(columeindex, _columhead);
                if (writecolumeIndex == 0)
                {
                    Wsheet.Cells[0, columeindex] = new Cell(_columhead);
                    Wbook.Save(_Filepath);
                    getColumeIndexData(columeindex, _columhead);
                    getRefferenceIndexData(rowindex, referencedata);
                    if (writeRowIndex.Contains(','))
                    {
                        string[] rowcount = writeRowIndex.Split(',');
                        foreach (var row in rowcount)
                        {
                            if (row != string.Empty)
                            {
                                Wsheet.Cells[Convert.ToInt16(row), columeindex] = new Cell(_writeData);
                            }
                        }
                    }
                    else
                    {
                        Wsheet.Cells[Convert.ToInt16(writeRowIndex), columeindex] = new Cell(_writeData);
                    }
                }
                else
                {
                    getRefferenceIndexData(rowindex, referencedata);
                    if (referencedata.Contains(':'))
                    {
                        if (writeRowIndex.Contains(','))
                        {
                            string[] rowcount = writeRowIndex.Split(',');
                            foreach (var row in rowcount)
                            {
                                if (row != string.Empty)
                                {
                                    Wsheet.Cells[Convert.ToInt16(row), writecolumeIndex] = new Cell(_writeData);
                                }
                            }
                        }
                    }
                    else
                    {
                        for (int a = 1; a <= rowindex; a++)
                        {
                            if (Wsheet.Cells[a, 0].StringValue == referencedata)
                            {
                                referenceDataCount++;
                            }
                        }
                        if (referenceDataCount > 1)
                        {
                            Wsheet.Cells[Convert.ToInt16(writeRowIndex) + 1, writecolumeIndex] = new Cell(_writeData);
                        }
                        else
                        {
                            Wsheet.Cells[Convert.ToInt16(writeRowIndex), writecolumeIndex] = new Cell(_writeData);
                        }
                    }
                }
                Wsheet.Cells.ColumnWidth[0, (ushort)columeindex] = 5000;
                Wbook.Save(_Filepath);
            }
            catch (Exception e)
            {
                throw new Exception("Error while writing data for the reference data.");
            }
        }
        public void getRefferenceIndexData(int rowcount, string refdata)
        {
            var Wbook = Workbook.Load(DataFilePath);
            var Wsheet = Wbook.Worksheets[0];
            writeRowIndex = string.Empty;
            for (int i = 1; i <= rowcount; i++)
            {
                if (refdata.Contains(':'))
                {
                    for (int j = 0; j < 10; j++)
                    {
                        string val = Wsheet.Cells[i, j].StringValue.ToString();
                        if (Wsheet.Cells[i, j].StringValue.ToString() == refdata.Split(':')[0])
                        {
                            int requiredRowIndex = i;
                            while (j < 10)
                            {
                                j++;
                                if (Wsheet.Cells[requiredRowIndex, j].StringValue.ToString() == refdata.Split(':')[1])
                                {
                                    writeRowIndex = writeRowIndex + i.ToString() + ',';
                                    break;
                                }
                            }
                        }
                    }
                }
                else if (refdata.Contains("Order"))
                {
                    string val = Wsheet.Cells[i, 1].StringValue.ToString();
                    if (Wsheet.Cells[i, 1].StringValue.ToString() == refdata.Split('=')[1])
                    {
                        writeRowIndex = i.ToString();
                        break;
                    }
                }
                else
                {
                    string val = Wsheet.Cells[i, 0].StringValue.ToString();
                    if (Wsheet.Cells[i, 0].StringValue.ToString() == refdata)
                    {
                        writeRowIndex = i.ToString();
                        break;
                    }
                }
            }
            writeRowIndex.Remove(writeRowIndex.Length - 1);
        }
        public void getColumeIndexData(int columeCount, string _headColume)
        {
            var Wbook = Workbook.Load(DataFilePath);
            var Wsheet = Wbook.Worksheets[0];
            for (int j = 0; j < columeCount; j++)
            {
                string val = Wsheet.Cells[0, j].StringValue.ToString();
                if (Wsheet.Cells[0, j].StringValue.ToString() == _headColume)
                {
                    writecolumeIndex = j;
                    break;
                }
                else
                {
                    writecolumeIndex = 0;
                }
            }
        }
        public void FileCorreptionCheck()
        {
            try
            {
                if (new System.IO.FileInfo(DataFilePath).Length < 8192)
                {
                    DataTable dtReport = new DataTable();
                    dtReport.Clear();
                    dtReport = ExcelKeywordTable.ExcelData(DataFilePath, "sheet1");
                    int rowindex = dtReport.Rows.Count;
                    int columindex = 0;

                    if (rowindex <= 80)
                    {
                        int rows = 150 - rowindex;
                        for (int i = 0; i < rows; i++)
                        {
                            columindex = 0;
                            wSheet.Cells[rowindex + 1, columindex] = new Cell("");
                            rowindex++;
                        }
                    }
                    wBook.Save(DataFilePath);
                }
            }
            catch (Exception e) { }
        }
    }
}
