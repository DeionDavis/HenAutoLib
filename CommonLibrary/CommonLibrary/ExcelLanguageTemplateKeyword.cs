using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Data;
using ExcelDataReader;
using CommonLibrary.Exceptions;

namespace CommonLibrary.LanguageTemplate
{
    public class ExcelLanguageTemplateKeyword
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
        public static List<LanguageResource> keywordData = new List<LanguageResource>();

        public static void getKewordData(string filePath, string sheet)
        {
            DataTable table = ExcelData(filePath, sheet);
            for (int row = 0; row < table.Rows.Count; row++)
            {
                for (int col = 0; col <= table.Columns.Count - 1; col++)
                {
                    LanguageResource Table = new LanguageResource()
                    {
                        _LanguageKey = table.Rows[row]["Name"].ToString(),
                        _LanguageValue = table.Rows[row]["Value"].ToString(),
                    };
                    keywordData.Add(Table);
                }
            }
            table.Clear();
        }
        public static string ReadKeywordMessage(string Keyword)
        {
            try
            {
                string data = (from KeyData in keywordData
                               where KeyData._LanguageKey == Keyword
                               select KeyData._LanguageValue).LastOrDefault();
                return data.ToString();
            }
            catch (Exception e)
            {
                throw new NoSuchModuleFound("The module " + Keyword + " Not available, please check the data given");
            }
        }
    }
    public class LanguageResource
    {
        public string _LanguageKey { get; set; }
        public string _LanguageValue { get; set; }
    }
}
