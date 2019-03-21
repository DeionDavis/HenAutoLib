using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using OperationLibrary;
using CommonLibrary.LanguageTemplate;
using CommonLibrary.LanguageTemp;
using CommonLibrary.Operations;
using CommonLibrary.Exceptions;
using System.Text.RegularExpressions;
using CommonLibrary.DataDrivenTesting;

namespace CommonLibrary.CommonLanguageReader
{
    public class CommonLanguageTemplateReader
    {
        static string ResourceFilePath = ConfigurationManager.AppSettings["ResourceFile"];
        public static PerformOperation pops= new PerformOperation();

        static string resMsg_GetTemplateMessage = string.Empty;
        public string Msg_GetTemplateMessage
        {
            get { return resMsg_GetTemplateMessage; }
            set { resMsg_GetTemplateMessage = value; }
        }

        /// <summary>
        /// This will get the data from the resource file excel, based on the language specified in the login data file.
        /// if we specify a variable with "Rec_" followed by '+' this will get the stored values and validate with the data.
        /// </summary>
        /// <param name="lang">this will accept the language code.</param>
        /// <param name="keyword">This will be the reference to the resource file and the reference to the message to the resource file data seperated by ':'</param>
        public static void Message(string lang, string keyword)
        {
            string LogPath = ConfigurationManager.AppSettings["LogOperation"];
            string[] KeyMessages = keyword.Split(',');
            for (int i = 0; i < KeyMessages.Count(); i++)
            {
                string[] data = KeyMessages[i].Split(':');
                if(data.Count()>1)
                {
                    ExcelLanguageResourceTemp.PopulateInCollection(LogPath + "\\KeywordDrivenData.xlsx", "LanguageResource");
                    if (lang == "zh-CN")
                    {
                        ExcelLanguageTemplateKeyword.getKewordData(ResourceFilePath + "\\" + ExcelLanguageResourceTemp.ReadKeywordData(data[0]) + ".xlsx", "LanguageRsourceCN");
                        ResourceMessages(keyword);
                    }
                    else if (lang == "th-TH")
                    {
                        ExcelLanguageTemplateKeyword.getKewordData(ResourceFilePath + "\\" + ExcelLanguageResourceTemp.ReadKeywordData(data[0]) + ".xlsx", "LanguageRsourceTH");
                        ResourceMessages(keyword);
                    }
                    else if (lang == "vi-VN")
                    {
                        ExcelLanguageTemplateKeyword.getKewordData(ResourceFilePath + "\\" + ExcelLanguageResourceTemp.ReadKeywordData(data[0]) + ".xlsx", "LanguageRsourceVN");
                        ResourceMessages(keyword);
                    }
                    else if (lang == "ko-KR")
                    {
                        ExcelLanguageTemplateKeyword.getKewordData(ResourceFilePath + "\\" + ExcelLanguageResourceTemp.ReadKeywordData(data[0]) + ".xlsx", "LanguageRsourceKR");
                        ResourceMessages(keyword);
                    }
                    else if (lang == "ja-JP")
                    {
                        ExcelLanguageTemplateKeyword.getKewordData(ResourceFilePath + "\\" + ExcelLanguageResourceTemp.ReadKeywordData(data[0]) + ".xlsx", "LanguageRsourceJP");
                        ResourceMessages(keyword);
                    }
                    else if (lang == "id-ID")
                    {
                        ExcelLanguageTemplateKeyword.getKewordData(ResourceFilePath + "\\" + ExcelLanguageResourceTemp.ReadKeywordData(data[0]) + ".xlsx", "LanguageRsourceID");
                        ResourceMessages(keyword);
                    }
                    else
                    {
                        ExcelLanguageTemplateKeyword.getKewordData(ResourceFilePath + "\\" + ExcelLanguageResourceTemp.ReadKeywordData(data[0]) + ".xlsx", "LanguageRsourceEN");
                        ResourceMessages(keyword);
                    }
                }
                else
                {
                    ResourceMessages(keyword);
                }
            }
        }
        public static void ResourceMessages(string datareferenceKeyword)
        {
            string[] KeyMessages;
            if (datareferenceKeyword.Contains('|'))
            {
                KeyMessages = datareferenceKeyword.Split('|');
            }
            else
            {
                KeyMessages = datareferenceKeyword.Split(',');
            }

            resMsg_GetTemplateMessage = string.Empty;
            string resMsg_GetTemplateMessageStore = string.Empty;
            for (int i = 0; i < KeyMessages.Count(); i++)
            {
                string[] getResource = KeyMessages[i].Split('+');
                if (getResource.Count() > 1)
                {
                    string resMsgTemplateMessage = ExcelLanguageTemplateKeyword.ReadKeywordMessage(getResource[0].Split(':')[1]);
                    for (int j = 1; j <= getResource.Count() - 1; j++)
                    {
                        if (j == 1)
                        {
                            if (getResource[j].StartsWith("Rec_"))
                            {
                                resMsg_GetTemplateMessageStore = resMsgTemplateMessage.Replace("[X" + j + "]", Operation.recordedData[getResource[j].Replace("Rec_", string.Empty)]).ToString();
                            }
                            else if (getResource[j].StartsWith("Order"))
                            {
                                resMsg_GetTemplateMessageStore = resMsgTemplateMessage.Replace("[X" + j + "]", pops.getDataFromDynamicExcel(getResource[j]));
                            }
                            else if (getResource[j].StartsWith("Material"))
                            {
                                resMsg_GetTemplateMessageStore = resMsgTemplateMessage.Replace("[X" + j + "]", pops.getDataFromDynamicExcel(getResource[j]));
                            }
                            else if (getResource[j].StartsWith("Reference"))
                            {
                                resMsg_GetTemplateMessageStore = resMsgTemplateMessage.Replace("[X" + j + "]", pops.getDataFromDynamicExcel(getResource[j]));
                            }
                            else
                            {
                                resMsg_GetTemplateMessageStore = resMsgTemplateMessage.Replace("[X" + j + "]", ExcelDataTable.ReadData(1, getResource[j])).ToString();
                            }
                        }
                        else
                        {
                            if (getResource[j].StartsWith("Rec_"))
                            {
                                resMsg_GetTemplateMessageStore = resMsg_GetTemplateMessageStore.Replace("[X" + j + "]", Operation.recordedData[getResource[j].Replace("Rec_", string.Empty)]).ToString();
                            }
                            else if (getResource[j].StartsWith("Order"))
                            {
                                resMsg_GetTemplateMessageStore = resMsg_GetTemplateMessageStore.Replace("[X" + j + "]", pops.getDataFromDynamicExcel(getResource[j]));
                            }
                            else if (getResource[j].StartsWith("Material"))
                            {
                                resMsg_GetTemplateMessageStore = resMsg_GetTemplateMessageStore.Replace("[X" + j + "]", pops.getDataFromDynamicExcel(getResource[j]));
                            }
                            else if (getResource[j].StartsWith("Reference"))
                            {
                                resMsg_GetTemplateMessageStore = resMsg_GetTemplateMessageStore.Replace("[X" + j + "]", pops.getDataFromDynamicExcel(getResource[j]));
                            }
                            else
                            {
                                resMsg_GetTemplateMessageStore = resMsg_GetTemplateMessageStore.Replace("[X" + j + "]", ExcelDataTable.ReadData(1, getResource[j])).ToString();
                            }
                        }
                    }
                    if (datareferenceKeyword.Contains('|'))
                    {
                        resMsg_GetTemplateMessage = resMsg_GetTemplateMessage + resMsg_GetTemplateMessageStore + '|';
                    }
                    else
                    {
                        resMsg_GetTemplateMessage = resMsg_GetTemplateMessage + resMsg_GetTemplateMessageStore;
                    }
                }
                else
                {
                    string[] data = KeyMessages[i].Split(':');
                    if (data.Count() > 1)
                    {
                        resMsg_GetTemplateMessage = resMsg_GetTemplateMessage + ExcelLanguageTemplateKeyword.ReadKeywordMessage(data[1]);
                    }
                    else
                    {
                        resMsg_GetTemplateMessage = ExcelDataTable.ReadData(1, data[i].ToString());
                    }
                }
            }
        }
    }
}
