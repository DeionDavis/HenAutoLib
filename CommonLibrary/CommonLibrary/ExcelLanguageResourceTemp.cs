using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CommonLibrary.Exceptions;

namespace CommonLibrary.LanguageTemp
{
    public class ExcelLanguageResourceTemp
    {
        public static DataTable ExcelData(string filePath, string sheet)
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
            DataTable resultTable = table[sheet];
            excelreader.Close();
            return resultTable;
        }

        public static List<LanguageKeyword> keywordData = new List<LanguageKeyword>();
        public static void PopulateInCollection(string filePath, string sheet)
        {
            DataTable table = ExcelData(filePath, sheet);
            for (int row = 1; row <= table.Rows.Count; row++)
            {
                for (int col = 0; col <= table.Columns.Count - 1; col++)
                {
                    LanguageKeyword dtLangTable = new LanguageKeyword()
                    {
                        _FileKeyword = table.Rows[row - 1][col].ToString(),
                        _LangFileHead = table.Columns[table.Columns.Count - 1].ColumnName,
                        _FileName = table.Rows[row - 1][table.Columns.Count - 1].ToString()
                    };
                    keywordData.Add(dtLangTable);
                }
            }
            table.Clear();

        }

        public static string ReadKeywordData(string Keyword)
        {
            try
            {
                string data = (from KeyData in keywordData
                               where KeyData._FileKeyword == Keyword
                               select KeyData._FileName).LastOrDefault();
                return data.ToString();
            }
            catch (Exception e)
            {
                throw new NoSuchControlTypeFound("The module " + Keyword + "Not available, please check the data given");
            }
        }
    }

    public class LanguageKeyword
    {
        public string _FileKeyword { get; set; }
        public string _LangFileHead { get; set; }
        public string _FileName { get; set; }
    }
}
