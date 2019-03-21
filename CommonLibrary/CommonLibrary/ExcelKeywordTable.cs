using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CommonLibrary.Exceptions;

namespace CommonLibrary.KeywordDrivenTesting
{
    public class ExcelKeywordTable
    {
     
        public static DataSet ExcelSheet = new DataSet();
        public static DataTable ExcelData(string filePath, string sheet)
        {
            FileStream stream = File.Open(filePath, FileMode.Open, FileAccess.Read);
            IExcelDataReader excelreader = ExcelReaderFactory.CreateReader(stream);
            //Here set the First Row as a column
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
            DataTable resultTable = table[sheet];
            excelreader.Close();
            return resultTable;
        }

        public static int GetTableCount(string FilePath)
        {
            int tablecount = 0;
            FileStream stream = File.Open(FilePath, FileMode.Open, FileAccess.Read);
            IExcelDataReader excelreader = ExcelReaderFactory.CreateReader(stream);
            ExcelSheet = excelreader.AsDataSet(new ExcelDataSetConfiguration()
            {
                UseColumnDataType = true,
                ConfigureDataTable = (tablereader) => new ExcelDataTableConfiguration()
                {
                    EmptyColumnNamePrefix = "Colume",
                    UseHeaderRow = true,
                }
            });
            DataTableCollection table = ExcelSheet.Tables;
            if (table.Contains("Conditional Execute"))
            {
                tablecount = ExcelSheet.Tables.Count - 2;
            }
            else if (table.Contains("Keyword Data"))
            {
                tablecount = ExcelSheet.Tables.Count - 1;
            }
            else
            {
                tablecount = ExcelSheet.Tables.Count;
            }
            excelreader.Close();
            return tablecount;
        }
        public static DataTable getTableData(int count)
        {
            DataTableCollection table = ExcelSheet.Tables;
        
            DataTable resultTable = table[count];
            return resultTable;
        }

        public static List<Data> keywordData = new List<Data>();
        public static void getKewordData(string filePath, string sheet)
        {
            DataTable table = ExcelData(filePath, sheet);
            for (int row = 1; row <= table.Rows.Count; row++)
            {
                for (int col = 0; col <= table.Columns.Count - 1; col++)
                {
                    Data Table = new Data()
                    {
                        _ControlKeyword = table.Rows[row - 1][table.Columns.Count - 1].ToString(),
                        _ColumeHead = table.Columns[col].ColumnName,
                        _ColumHeadValue = table.Rows[row - 1][col].ToString()
                    };
                    keywordData.Add(Table);
                }
            }
            table.Clear();
        }

        public static string ReadKeywordData(string Keyword, string Head)
        {
            try
            {
                string data = (from KeyData in keywordData
                               where KeyData._ControlKeyword == Keyword && KeyData._ColumeHead == Head
                               select KeyData._ColumHeadValue).LastOrDefault();
                return data.ToString();
            }
            catch (Exception e)
            {
                throw new NoSuchModuleFound("The module " + Keyword + "Not available, please check the data given");
            }
        }

    }

    public class Data
    {
        public string _ControlKeyword { get; set; }
        public string _ColumeHead { get; set; }
        public string _ColumHeadValue { get; set; }
    }
}
