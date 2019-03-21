using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDataReader;
using System.Data;
using System.IO;

namespace TestExexution
{
    public class DataReader
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

        public static List<BatchCollection> datacol = new List<BatchCollection>();

        public static void PopulateInCollection(string filePath)
        {
            DataTable table = ExcelData(filePath);
            for (int row = 0; row < table.Rows.Count; row++)
            {
                for (int col = 0; col <= table.Columns.Count - 1; col++)
                {
                    BatchCollection dtTable = new BatchCollection()
                    {
                        _Testcase = table.Rows[row]["TestcaseName"].ToString(),
                        _Option = table.Rows[row]["Option"].ToString(),
                        _Path = table.Rows[row]["Refference"].ToString()
                    };
                    datacol.Add(dtTable);
                }
            }
        }
        public static string ReadKeywordMessage(string Keyword)
        {
            try
            {
                string data = (from KeyData in datacol
                               where KeyData._Testcase == Keyword
                               select KeyData._Option).LastOrDefault();
                return data.ToString();
            }
            catch (Exception e)
            {
                return null;
            }
        }
    }

    public class BatchCollection
    {
        public string _Testcase { get; set; }
        public string _Option { get; set; }
        public string _Path { get; set; }
    }
}
