using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.IO;
using ExcelDataReader;


namespace CommonLibrary.DataDrivenTesting
{
    public class ExcelDataTable
    {
        public static DataTable ExcelData(string filePath)
        {
            FileStream stream = File.Open(filePath, FileMode.Open, FileAccess.Read);
            IExcelDataReader excelreader = ExcelReaderFactory.CreateReader(stream);
            ////Here set the First Row as a column
            DataSet result = excelreader.AsDataSet(new ExcelDataSetConfiguration()
            {
                UseColumnDataType = true,
                ConfigureDataTable = (tablereader) => new ExcelDataTableConfiguration()
                {
                    EmptyColumnNamePrefix = "Colume",
                    UseHeaderRow = true,
                }
            });
            DataTableCollection table = result.Tables;
            DataTable resultTable = table["sheet1"];
            excelreader.Close();
            return resultTable;
        }

        public static List<DataCollection> datacol = new List<DataCollection>();

        public static void PopulateInCollection(string filePath)
        {
            DataTable table = ExcelData(filePath);
            for (int row = 1; row <= table.Rows.Count; row++)
            {
                for (int col = 0; col <= table.Columns.Count - 1; col++)
                {
                    DataCollection dtTable = new DataCollection()
                    {
                        rowNumber = row,
                        colName = table.Columns[col].ColumnName,
                        colValue = table.Rows[row - 1][col].ToString()
                    };
                    datacol.Add(dtTable);
                }
            }
            table.Clear();
        }
        public static string ReadData(int rowNumber, string columnName)
        {
            try
            {
                string data = (from colData in datacol
                               where colData.colName == columnName && colData.rowNumber == rowNumber
                               select colData.colValue).LastOrDefault();
                return data.ToString();
            }
            catch (Exception e)
            {
                return columnName;
            }
        }

        #region Get Batch Wise Data for each batch
        public static List<Data> batchData = new List<Data>();

        public static void PopulateBatchWiseData(string filePath)
        {
            DataTable ControlTable = ExcelData(filePath);
            for (int row = 1; row <= ControlTable.Rows.Count; row++)
            {
                for (int col = 0; col < ControlTable.Columns.Count - 1; col++)
                {
                    Data Table = new Data()
                    {
                        _BatchInfo = ControlTable.Rows[row - 1]["BatchName"].ToString(),
                        _ColumHeadInfo = ControlTable.Columns[col + 1].ColumnName,
                        _ColumData = ControlTable.Rows[row - 1][col + 1].ToString()
                    };
                    batchData.Add(Table);
                }
            }
            ControlTable.Clear();
        }

        public static string ReadBatchData(string batchName, string columnData)
        {
            try
            {
                string data = (from bData in batchData
                               where bData._BatchInfo == batchName && bData._ColumHeadInfo == columnData
                               select bData._ColumData).LastOrDefault();
                return data.ToString();
            }
            catch (Exception e)
            {
                return null;
            }
        }
        #endregion

        #region Get Login Details Based on Plant
        public static List<PlantData> PlantLog = new List<PlantData>();
        public static void PopulatePlantWiseData(string filePath)
        {
            DataTable ControlTable = ExcelData(filePath);
            for (int row = 1; row <= ControlTable.Rows.Count; row++)
            {
                for (int col = 0; col < ControlTable.Columns.Count - 1; col++)
                {
                    PlantData Table = new PlantData()
                    {
                        _PlantInfo = ControlTable.Rows[row - 1]["Plant Name"].ToString(),
                        _ColumHeadInfo = ControlTable.Columns[col + 1].ColumnName,
                        _ColumData = ControlTable.Rows[row - 1][col + 1].ToString()
                    };
                    PlantLog.Add(Table);
                }
            }
            ControlTable.Clear();
        }
        public static string ReadPlantLoginInfo(string PlantName, string columnData)
        {
            try
            {
                string data = (from bData in PlantLog
                               where bData._PlantInfo == PlantName && bData._ColumHeadInfo == columnData
                               select bData._ColumData).LastOrDefault();
                return data.ToString();
            }
            catch (Exception e)
            {
                return null;
            }
        }
        #endregion

        #region Get Reference Data Details
        public static List<ReferenceData> RefResourceData = new List<ReferenceData>();
        public static void PopulateRecordData(string filePath)
        {
            DataTable ControlTable = ExcelData(filePath);

            if (ControlTable.Columns.Contains("Material"))
            {
                for (int row = 1; row <= ControlTable.Rows.Count; row++)
                {
                    for (int col = 0; col <= ControlTable.Columns.Count - 1; col++)
                    {
                        ReferenceData Table = new ReferenceData()
                        {
                            _RefData = ControlTable.Rows[row - 1][0].ToString(),
                            _ReferenceOrder = ControlTable.Rows[row - 1]["Order"].ToString(),
                            _Material = ControlTable.Rows[row - 1]["Material"].ToString(),
                            _RefDataHead = ControlTable.Columns[col].ColumnName,
                            _ReferenceInfo = ControlTable.Rows[row - 1][col].ToString()
                        };
                        RefResourceData.Add(Table);
                    }
                }
            }
            else if (!ControlTable.Columns.Contains("Material"))
            {
                for (int row = 1; row <= ControlTable.Rows.Count; row++)
                {
                    for (int col = 0; col <= ControlTable.Columns.Count - 1; col++)
                    {
                        ReferenceData Table = new ReferenceData()
                        {
                            _RefData = ControlTable.Rows[row - 1][0].ToString(),
                            _ReferenceOrder = ControlTable.Rows[row - 1]["Order"].ToString(),
                            _RefDataHead = ControlTable.Columns[col].ColumnName,
                            _ReferenceInfo = ControlTable.Rows[row - 1][col].ToString()
                        };
                        RefResourceData.Add(Table);
                    }
                }
            }
            ControlTable.Clear();
        }
        public static string ReadRefferenceInfo(string ReferenceData, string RequiredData)
        {
            try
            {
                string data = (from bData in RefResourceData
                               where bData._RefData == ReferenceData && bData._RefDataHead == RequiredData
                               select bData._ReferenceInfo).LastOrDefault();
                if (data == string.Empty)
                {
                    data = (from bData in RefResourceData
                            where bData._RefData == ReferenceData && bData._RefDataHead == RequiredData
                            select bData._ReferenceInfo).FirstOrDefault();
                }
                return data.ToString();
            }
            catch (Exception e)
            {
                throw new Exception("No data found in the recorded file");
            }
        }

        public static string ReadRefferenceInfoByOrder(string order, string RequiredData)
        {
            try
            {
                string data = (from bData in RefResourceData
                               where bData._ReferenceOrder == order && bData._RefDataHead == RequiredData
                               select bData._ReferenceInfo).LastOrDefault();
                return data.ToString();
            }
            catch (Exception e)
            {
                throw new Exception("No data found in the recorded file");
            }
        }

        public static string ReadRefferenceInfoByMaterial(string material, string RequiredData)
        {
            try
            {
                string data = (from bData in RefResourceData
                               where bData._Material == material && bData._RefDataHead == RequiredData
                               select bData._ReferenceInfo).LastOrDefault();
                return data.ToString();
            }
            catch (Exception e)
            {
                throw new Exception("No data found in the recorded file");
            }
        }

        public static string ReadRefferenceInfoByMaterialWithStore(string store, string material, string RequiredData)
        {
            try
            {
                string data = (from bData in RefResourceData
                               where bData._RefData == store && bData._Material == material && bData._RefDataHead==RequiredData
                               select bData._ReferenceInfo).LastOrDefault();
                return data.ToString();
            }
            catch (Exception e)
            {
                throw new Exception("No data found in the recorded file");
            }
        }
        #endregion
    }

    public class DataCollection
    {
        public int rowNumber { get; set; }
        public string colName { get; set; }
        public string colValue { get; set; }
    }

    public class Data
    {
        public string _BatchInfo { get; set; }
        public string _ColumHeadInfo { get; set; }
        public string _ColumData { get; set; }
    }
    public class PlantData
    {
        public string _PlantInfo { get; set; }
        public string _ColumHeadInfo { get; set; }
        public string _ColumData { get; set; }
    }

    public class ReferenceData
    {
        public string _ReferenceInfo { get; set; }
        public string _ReferenceOrder { get; set; }
        public string _RefDataHead { get; set; }
        public string _RefData { get; set; }
        public string _Material { get; set; }
    }
}
