using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using System.Data;
using AutoItX3Lib;
using System.Threading;
using System.Reflection;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using CommonLibrary.Reports;
using CommonLibrary.Log;
using CommonLibrary.Operations;
using CommonLibrary.Exceptions;
using CommonLibrary.Writedata;
using CommonLibrary.DataDrivenTesting;
using CommonLibrary.KeywordDrivenTesting;
using CommonLibrary.CommonLanguageReader;
using CommonLibrary.FileOperation;
using System.Text.RegularExpressions;
using Microsoft.VisualStudio.TestTools.UITesting;

namespace OperationLibrary
{
    public class Operation
    {
        CommonLanguageTemplateReader languageResource = new CommonLanguageTemplateReader();
        LoginOperatrion log = new LoginOperatrion();
        PerformOperation pop = new PerformOperation();
        ReportGeneration repgen = new ReportGeneration();
        WriteAndReadData writeData = new WriteAndReadData();
        FileOpearions file = new FileOpearions();
        AutoItX3 auto = new AutoItX3();
        DataTable dt = new DataTable();
        public static string lang = string.Empty;

        public string screenShotName = "";
        public static string ErrorScreenPath = " ";
        public static string FailerReason = " ";

        public string warningScreen="";
        public static string warningScreenPath = " ";
        public static string warningMessage = " ";

        public static string BatchName = string.Empty;
        public static string ControlData = string.Empty;
        public string OperationPath = ConfigurationManager.AppSettings["BasicOperation"];
        public int numberofExecution = 0;
        public string Message = string.Empty;
        int min = 0;
        int mid = 0;
        int max = 0;
        public static Dictionary<string, string> recordedData = new Dictionary<string, string>();
        public static string tabName = string.Empty;
        public static string StepNumber = string.Empty;
        public static string CurrentOperation = string.Empty;
        public static string ForeigntabName = string.Empty;
        public static string ForeignStepNumber = string.Empty;
        public static string ForeignSheetName = string.Empty;

        public void PerformOperation(int i)
          {
            LoginOperatrion.LogPath = ConfigurationManager.AppSettings["LogOperation"];
            BatchName = ConfigurationManager.AppSettings["Batch"];
            lang = ExcelDataTable.ReadBatchData(BatchName, "languageCode");
            int totalExecution = Convert.ToInt32(ExcelDataTable.ReadBatchData(BatchName, "Execution"));
            min = Convert.ToInt16(ExcelDataTable.ReadData(1, "Minimum"));
            mid = Convert.ToInt16(ExcelDataTable.ReadData(1, "Medium"));
            max = Convert.ToInt16(ExcelDataTable.ReadData(1, "Maximum"));
            int TestStepCount = ExcelKeywordTable.GetTableCount(OperationPath);
            string assemblyName = Assembly.GetCallingAssembly().GetName().Name;
            numberofExecution = i;
            for (int j = 0; j < TestStepCount; j++)
            {
                dt.Clear();
                dt = ExcelKeywordTable.getTableData(j);
                //get excel tab name
                tabName = dt.TableName;
                foreach (DataRow count in dt.Rows)
                {
                    //get current step and Current keyword using
                    StepNumber = count["Step No"].ToString();
                    CurrentOperation = count["Keyword"].ToString();
                    if (count.IsNull("Step No"))
                    {
                        repgen.UpdateData("No Step Found", count["Description"].ToString(), count["Keyword"].ToString(), "Skipped", LoginOperatrion.batchforReport, ReportGeneration.DetailedReportFilePath, "");
                        continue;
                    }
                    else
                    {

                        string operation = count["Keyword"].ToString();
                        switch (operation)
                        {
                            case "Initialize":
                                log.Initialize(Assembly.GetCallingAssembly().GetName().Name);
                                auto.Send("{F5}");
                                Thread.Sleep(max);
                                break;
                            case "Navigate2CDA":
                            case "Navigate2CRA":
                            case "Navigate2CPA":
                                log.SubModuleNavigationPack(operation);
                                break;
                            case "SyncPO":
                            case "SyncDO":
                                log.SyncOrder(operation);
                                break;
                            case "Operations":
                                PreformForigneOperations(count["ControlKeywordValue"].ToString(), count["DataRefferencekeyword"].ToString(), assemblyName, numberofExecution);
                                break;
                            case "Logout":
                                log.Logout();
                                break;
                            case "Login":
                                log.login(operation);
                                break;
                            case "PlantLogin":
                                log.PlantLogin(operation, numberofExecution);
                                break;
                            case "RegionalLogin":
                                log.login(operation);
                                break;
                            case "SuperAdminLogin":
                                log.SuperLogin();
                                break;
                            case "PlantRegionalLogin":
                                log.PlantLogin(operation, numberofExecution);
                                break;
                            case "PageNavigation":
                                log.NavigationPage(count["ControlKeywordValue"].ToString(), count["Step No"].ToString(), count["Description"].ToString());
                                break;
                            case "ZeroTotalQtyValidate":
                                log.ValiateZero(count["Step No"].ToString(), count["Description"].ToString(), count["ControlKeywordValue"].ToString());
                                break;
                            case "ChangLanguage":
                                log.ChangLanguage(lang);
                                break;
                            case "ModuleNavigation":
                                log.ModuleNavigation(count["ControlKeywordValue"].ToString());
                                break;
                            case "SubModuleNavigation":
                                log.SubModuleNavigation(count["ControlKeywordValue"].ToString());
                                break;
                            case "LinkNavigation":
                                log.LinkNavigation(count["ControlKeywordValue"].ToString());
                                break;
                            case "EnterText":
                                string[] Testdata = count["DataRefferencekeyword"].ToString().Split(':');
                                if (Testdata.Count() > 1)
                                {
                                    if(Testdata[0]!= "BatchData")
                                    {
                                        pop.OperationStart(count["Keyword"].ToString(), count["TypeOfControl"].ToString(), count["ControlKeyword"].ToString(), count["ControlKeywordValue"].ToString(), ExcelDataTable.ReadData(Convert.ToInt16(Testdata[1]), Testdata[0]), count["TypeOfWindow"].ToString(), count["Step No"].ToString(), count["Description"].ToString());
                                    }
                                    else
                                    {
                                        pop.OperationStart(count["Keyword"].ToString(), count["TypeOfControl"].ToString(), count["ControlKeyword"].ToString(), count["ControlKeywordValue"].ToString(), ExcelDataTable.ReadBatchData(BatchName, Testdata[1]), count["TypeOfWindow"].ToString(), count["Step No"].ToString(), count["Description"].ToString());
                                    }
                                    
                                }
                                else
                                {
                                    pop.OperationStart(count["Keyword"].ToString(), count["TypeOfControl"].ToString(), count["ControlKeyword"].ToString(), count["ControlKeywordValue"].ToString(), ExcelDataTable.ReadData(1, Testdata[0]), count["TypeOfWindow"].ToString(), count["Step No"].ToString(), count["Description"].ToString());
                                }
                                break;
                            case "EnterLanguageText":
                                CommonLanguageTemplateReader.Message(lang, count["DataRefferencekeyword"].ToString());
                                pop.OperationStart(count["Keyword"].ToString(), count["TypeOfControl"].ToString(), count["ControlKeyword"].ToString(), count["ControlKeywordValue"].ToString(), languageResource.Msg_GetTemplateMessage, count["TypeOfWindow"].ToString(), count["Step No"].ToString(), count["Description"].ToString());
                                break;
                            case "Click":
                            case "RightClick":
                                pop.OperationStart(count["Keyword"].ToString(), count["TypeOfControl"].ToString(), count["ControlKeyword"].ToString(), count["ControlKeywordValue"].ToString(), count["DataRefferencekeyword"].ToString(), count["TypeOfWindow"].ToString(), count["Step No"].ToString(), count["Description"].ToString());
                                break;
                            case "ScriptClick":
                                pop.ScriptClick(count["Keyword"].ToString(), count["ControlKeyword"].ToString(), count["ControlKeywordValue"].ToString(), count["Step No"].ToString(), count["Description"].ToString());
                                repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                break;
                            case "ValidateControlAvailable":
                            case "ValidateControlNotAvailable":
                                if (count["ControlKeywordValue"].ToString().Contains('='))
                                {
                                    WriteAndReadData.DataFilePath = ConfigurationManager.AppSettings["ExcelDataFile"] + "\\DynamicData" + "\\" + BatchName + ".xls";
                                    ExcelDataTable.PopulateRecordData(WriteAndReadData.DataFilePath);
                                }
                                bool controlavailability = pop.ControlAvailability(count["TypeOfControl"].ToString(), count["ControlKeyword"].ToString(), count["ControlKeywordValue"].ToString(), count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString());
                                if(operation== "ValidateControlAvailable")
                                {
                                    if (!controlavailability)
                                    {
                                        screenShotName = count["AssertionScreenshotName"].ToString();
                                        ErrorScreenPath = log.screenShot(screenShotName);
                                        FailerReason = count["AssertionMessage"].ToString();
                                        repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                        repgen.FileCorreptionCheck();
                                        Assert.Fail(FailerReason);
                                    }
                                }
                                else
                                {
                                    if (controlavailability)
                                    {
                                        screenShotName = count["AssertionScreenshotName"].ToString();
                                        ErrorScreenPath = log.screenShot(screenShotName);
                                        FailerReason = count["AssertionMessage"].ToString();
                                        repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                        repgen.FileCorreptionCheck();
                                        Assert.Fail(FailerReason);
                                    }
                                }
                                repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                break;
                            case "ValidateLog":
                                bool status = log.validateLOGcheckFail(count["DataRefferencekeyword"].ToString());
                                if (!status)
                                {
                                    screenShotName = count["AssertionScreenshotName"].ToString();
                                    ErrorScreenPath = log.screenShot(screenShotName);
                                    FailerReason = count["AssertionMessage"].ToString();
                                    log.CloseFile();
                                    repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                    repgen.FileCorreptionCheck();
                                    Assert.Fail(FailerReason);
                                }
                                log.CloseFile();
                                repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                break;
                            case "Validate":
                            case "Validate&Continue":
                                string validationMessage = pop.WebGetControlData(count["TypeOfControl"].ToString(), count["ControlKeyword"].ToString(), count["ControlKeywordValue"].ToString(), count["AssertionScreenshotName"].ToString(), count["AssertionMessage"].ToString(), count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString()).Trim();
                                if (count["DataRefferencekeyword"].ToString().Contains('='))
                                {
                                    WriteAndReadData.DataFilePath = ConfigurationManager.AppSettings["ExcelDataFile"] + "\\DynamicData" + "\\" + BatchName + ".xls";
                                    ExcelDataTable.PopulateRecordData(WriteAndReadData.DataFilePath);
                                }
                                if (validationMessage != string.Empty)
                                {
                                    validationMessage = Regex.Replace(validationMessage, "<.*?>", String.Empty).Trim();
                                    Message = System.Web.HttpUtility.HtmlDecode(validationMessage);
                                    validationMessage = Message.Replace(" ", string.Empty).Trim();
                                    CommonLanguageTemplateReader.Message(lang, count["DataRefferencekeyword"].ToString());
                                    if (operation == "Validate")
                                    {
                                        if (languageResource.Msg_GetTemplateMessage.Replace(" ", string.Empty) != validationMessage.Trim())
                                        {
                                            screenShotName = count["AssertionScreenshotName"].ToString();
                                            ErrorScreenPath = log.screenShot(screenShotName);
                                            FailerReason = count["AssertionMessage"].ToString();
                                            repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                            repgen.FileCorreptionCheck();
                                        }
                                        Assert.AreEqual(languageResource.Msg_GetTemplateMessage.Replace(" ", string.Empty), validationMessage, FailerReason);
                                        repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                    }
                                    else if (operation == "Validate&Continue")
                                    {
                                        if (languageResource.Msg_GetTemplateMessage.Replace(" ", string.Empty) != validationMessage.Trim())
                                        {
                                            warningScreen = count["AssertionScreenshotName"].ToString();
                                            warningScreenPath = log.screenShot(warningScreen);
                                            warningMessage = count["AssertionMessage"].ToString();
                                            repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                        }
                                    }
                                }
                                break;
                            case "ValidateOR":
                                bool validationstatus = false;
                                if (count["DataRefferencekeyword"].ToString().Contains('='))
                                {
                                    WriteAndReadData.DataFilePath = ConfigurationManager.AppSettings["ExcelDataFile"] + "\\DynamicData" + "\\" + BatchName + ".xls";
                                    ExcelDataTable.PopulateRecordData(WriteAndReadData.DataFilePath);
                                }
                                CommonLanguageTemplateReader.Message(lang, count["DataRefferencekeyword"].ToString());
                                string[] availableMsg = languageResource.Msg_GetTemplateMessage.ToString().Split('|');
                                string[] getMessage = count["ControlKeywordValue"].ToString().Split('|');
                                for (i = 0; i < getMessage.Count(); i++)
                                {
                                    validationMessage = pop.WebGetControlData(count["TypeOfControl"].ToString(), count["ControlKeyword"].ToString(), getMessage[i], count["AssertionScreenshotName"].ToString(), count["AssertionMessage"].ToString(), count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString()).Trim();
                                    validationMessage = Regex.Replace(validationMessage, "<.*?>", String.Empty).Trim();
                                    Message = System.Web.HttpUtility.HtmlDecode(validationMessage);
                                    validationMessage = Message.Replace(" ", string.Empty).Trim();
                                    if (availableMsg[i].Replace(" ", string.Empty) == validationMessage)
                                    {
                                        validationstatus = true;
                                        break;
                                    }
                                    else
                                    {
                                        validationstatus = false;
                                    }
                                }
                                if (!validationstatus)
                                {
                                    screenShotName = count["AssertionScreenshotName"].ToString();
                                    ErrorScreenPath = log.screenShot(screenShotName);
                                    FailerReason = count["AssertionMessage"].ToString();
                                    repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                    repgen.FileCorreptionCheck();
                                    Assert.Fail(FailerReason);
                                }
                                repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                break;
                            case "ValidateNotEqual":
                                string validationMessagenotEqual = pop.WebGetControlData(count["TypeOfControl"].ToString(), count["ControlKeyword"].ToString(), count["ControlKeywordValue"].ToString(), count["AssertionScreenshotName"].ToString(), count["AssertionMessage"].ToString(), count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString()).Trim();
                                validationMessage = Regex.Replace(validationMessagenotEqual, "<.*?>", String.Empty).Trim();
                                Message = System.Web.HttpUtility.HtmlDecode(validationMessage);
                                validationMessage = Message.Replace(" ", string.Empty).Trim();
                                CommonLanguageTemplateReader.Message(lang, count["DataRefferencekeyword"].ToString());
                                if (languageResource.Msg_GetTemplateMessage.Replace(" ", string.Empty) == validationMessage)
                                {
                                    screenShotName = count["AssertionScreenshotName"].ToString();
                                    ErrorScreenPath = log.screenShot(screenShotName);
                                    FailerReason = count["AssertionMessage"].ToString();
                                    repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                    repgen.FileCorreptionCheck();
                                }
                                Assert.AreNotEqual(languageResource.Msg_GetTemplateMessage.Replace(" ", string.Empty), validationMessage, FailerReason);
                                repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                break;
                            case "ClickDelete":
                            case "OprRightClick":
                                Thread.Sleep(max * 2);
                                if (count["ControlKeywordValue"].ToString().Contains('='))
                                {
                                    WriteAndReadData.DataFilePath = ConfigurationManager.AppSettings["ExcelDataFile"] + "\\DynamicData" + "\\" + BatchName + ".xls";
                                    ExcelDataTable.PopulateRecordData(WriteAndReadData.DataFilePath);
                                    validationMessage = pop.getDataFromDynamicExcel(count["ControlKeywordValue"].ToString());
                                    pop.OperationStart(count["Keyword"].ToString(), count["TypeOfControl"].ToString(), count["ControlKeyword"].ToString(), validationMessage, count["DataRefferencekeyword"].ToString(), count["TypeOfWindow"].ToString(), count["Step No"].ToString(), count["Description"].ToString(), new int[] { Convert.ToInt16(count["XCodinate"]), Convert.ToInt16(count["Ycodinate"]) });
                                }
                                else
                                {
                                    pop.OperationStart(count["Keyword"].ToString(), count["TypeOfControl"].ToString(), count["ControlKeyword"].ToString(), ExcelDataTable.ReadData(1, count["ControlKeywordValue"].ToString()), count["DataRefferencekeyword"].ToString(), count["TypeOfWindow"].ToString(), count["Step No"].ToString(), count["Description"].ToString(), new int[] { Convert.ToInt16(count["XCodinate"]), Convert.ToInt16(count["Ycodinate"]) });
                                }
                                break;
                            case "ClickEdit":
                            case "DClickEdit":
                            case "DClickDelete":
                                Thread.Sleep(max * 2);
                                if (count["ControlKeywordValue"].ToString().Contains('='))
                                {
                                    WriteAndReadData.DataFilePath = ConfigurationManager.AppSettings["ExcelDataFile"] + "\\DynamicData" + "\\" + BatchName + ".xls";
                                    ExcelDataTable.PopulateRecordData(WriteAndReadData.DataFilePath);
                                    validationMessage =  pop.getDataFromDynamicExcel(count["ControlKeywordValue"].ToString());
                                    pop.OperationStart(count["Keyword"].ToString(), count["TypeOfControl"].ToString(), count["ControlKeyword"].ToString(), validationMessage, count["DataRefferencekeyword"].ToString(), count["TypeOfWindow"].ToString(), count["Step No"].ToString(), count["Description"].ToString(), new int[] { Convert.ToInt16(count["XCodinate"]), Convert.ToInt16(count["Ycodinate"]) });
                                }
                                else
                                {
                                    pop.OperationStart(count["Keyword"].ToString(), count["TypeOfControl"].ToString(), count["ControlKeyword"].ToString(), ExcelDataTable.ReadData(1, count["ControlKeywordValue"].ToString()), count["DataRefferencekeyword"].ToString(), count["TypeOfWindow"].ToString(), count["Step No"].ToString(), count["Description"].ToString(), new int[] { Convert.ToInt16(count["XCodinate"]), Convert.ToInt16(count["Ycodinate"]) });
                                }
                                break;
                            case "MouseMove":
                            case "MouseMoveRightClick":
                                Thread.Sleep(max * 2);
                                string[] data = count["ControlKeywordValue"].ToString().Split(':');
                                if (data.Count() <= 1)
                                {
                                    pop.OperationStart(count["Keyword"].ToString(), count["TypeOfControl"].ToString(), count["ControlKeyword"].ToString(), count["ControlKeywordValue"].ToString(), count["DataRefferencekeyword"].ToString(), count["TypeOfWindow"].ToString(), count["Step No"].ToString(), count["Description"].ToString(), new int[] { Convert.ToInt16(count["XCodinate"]), Convert.ToInt16(count["Ycodinate"]) });
                                }
                                else
                                {
                                    CommonLanguageTemplateReader.Message(lang, count["ControlKeywordValue"].ToString());
                                    pop.OperationStart(count["Keyword"].ToString(), count["TypeOfControl"].ToString(), count["ControlKeyword"].ToString(), languageResource.Msg_GetTemplateMessage, count["DataRefferencekeyword"].ToString(), count["TypeOfWindow"].ToString(), count["Step No"].ToString(), count["Description"].ToString(), new int[] { Convert.ToInt16(count["XCodinate"]), Convert.ToInt16(count["Ycodinate"]) });
                                }
                                break;
                            case "Clear":
                                pop.OperationStart(count["Keyword"].ToString(), count["TypeOfControl"].ToString(), count["ControlKeyword"].ToString(), count["ControlKeywordValue"].ToString(), count["DataRefferencekeyword"].ToString(), count["TypeOfWindow"].ToString(), count["Step No"].ToString(), count["Description"].ToString());
                                break;
                            case "Wait":
                            case "WaitPropertySet":
                                pop.OperationStart(count["Keyword"].ToString(), count["TypeOfControl"].ToString(), count["ControlKeyword"].ToString(), count["ControlKeywordValue"].ToString(), count["DataRefferencekeyword"].ToString(), count["TypeOfWindow"].ToString(), count["Step No"].ToString(), count["Description"].ToString());
                                break;
                            case "Sleep":
                                pop.GotoSleep(count["DataRefferencekeyword"].ToString());
                                repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                break;
                            case "ClickLanguageText":
                                CommonLanguageTemplateReader.Message(lang, count["ControlKeywordValue"].ToString());
                                pop.OperationStart(count["Keyword"].ToString(), count["TypeOfControl"].ToString(), count["ControlKeyword"].ToString(), languageResource.Msg_GetTemplateMessage, count["DataRefferencekeyword"].ToString(), count["TypeOfWindow"].ToString(), count["Step No"].ToString(), count["Description"].ToString());
                                repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                break;
                            case "CloseFile":
                                log.CloseFile();
                                repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                break;
                            case "SendKey":
                            case "SendText":
                                if (operation == "SendKey")
                                {
                                    log.SendKey(count["ControlKeywordValue"].ToString());
                                }
                                else
                                {
                                    if(count["DataRefferencekeyword"].ToString().Contains('='))
                                    {
                                        WriteAndReadData.DataFilePath = ConfigurationManager.AppSettings["ExcelDataFile"] + "\\DynamicData" + "\\" + BatchName + ".xls";
                                        ExcelDataTable.PopulateRecordData(WriteAndReadData.DataFilePath);
                                        string[] refferenceArray = count["DataRefferencekeyword"].ToString().Split('=');
                                        log.Sendtext(pop.getDataFromDynamicExcel(count["DataRefferencekeyword"].ToString()));
                                    }
                                    else
                                    {
                                        log.Sendtext(ExcelDataTable.ReadData(1, count["DataRefferencekeyword"].ToString()).ToString());
                                    }
                                }
                                repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                break;
                            case "OpenFile":
                                file.OpenFile(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), count["AssertionScreenshotName"].ToString(), count["AssertionMessage"].ToString());
                                break;
                            case "SaveOpenFile":
                                file.SaveOpenFile(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), count["AssertionScreenshotName"].ToString(), count["AssertionMessage"].ToString());
                                break;
                            case "CancelDownload":
                                file.CancelDownload(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString());
                                break;
                            case "SetValue":
                                if (!count["DataRefferencekeyword"].ToString().Contains('='))
                                {
                                    Testdata = count["DataRefferencekeyword"].ToString().Split(':');
                                    if (Testdata.Count() > 1)
                                    {
                                        pop.ScriptExcuite(count["Keyword"].ToString(), count["ControlKeyword"].ToString(), count["ControlKeywordValue"].ToString(), ExcelDataTable.ReadData(Convert.ToInt16(Testdata[1]), Testdata[0]), count["Step No"].ToString(), count["Description"].ToString());
                                    }
                                    else
                                    {
                                        pop.ScriptExcuite(count["Keyword"].ToString(), count["ControlKeyword"].ToString(), count["ControlKeywordValue"].ToString(), ExcelDataTable.ReadData(1, Testdata[0]), count["Step No"].ToString(), count["Description"].ToString());
                                    }
                                }
                                else
                                {
                                    WriteAndReadData.DataFilePath = ConfigurationManager.AppSettings["ExcelDataFile"] + "\\DynamicData" + "\\" + BatchName + ".xls";
                                    ExcelDataTable.PopulateRecordData(WriteAndReadData.DataFilePath);
                                    string recData = pop.getDataFromDynamicExcel(count["DataRefferencekeyword"].ToString());
                                    pop.ScriptExcuite(count["Keyword"].ToString(), count["ControlKeyword"].ToString(), count["ControlKeywordValue"].ToString(), recData, count["Step No"].ToString(), count["Description"].ToString());
                                }
                                repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                break;
                            case "GetValueMatch":
                            case "GetValueNoMatch":
                            case "GetValueContains":
                                string valueGEt = pop.ScriptGetExcuite(count["Keyword"].ToString(), count["ControlKeyword"].ToString(), count["ControlKeywordValue"].ToString(), count["Step No"].ToString(), count["Description"].ToString());
                                string valueMatch = string.Empty;
                                string[] RefResources = { };
                                if (count["DataRefferencekeyword"].ToString().Contains('='))
                                {
                                    WriteAndReadData.DataFilePath = ConfigurationManager.AppSettings["ExcelDataFile"] + "\\DynamicData" + "\\" + BatchName + ".xls";
                                    ExcelDataTable.PopulateRecordData(WriteAndReadData.DataFilePath);
                                    string referenceData = count["DataRefferencekeyword"].ToString();
                                    if (referenceData.Contains('+') || referenceData.Contains('-') || referenceData.Contains('*'))
                                    {
                                        if(referenceData.Contains('+'))
                                        {
                                            RefResources = referenceData.Split('+');
                                            valueMatch = pop.getDataFromDynamicExcel(RefResources[0]);
                                            if (RefResources[1].ToString().Contains('='))
                                            {
                                                valueMatch = pop.calculatedata(valueMatch + '+' + pop.getDataFromDynamicExcel(RefResources[1])).ToString();
                                            }
                                            else
                                            {
                                                valueMatch = pop.calculatedata(valueMatch + '+' + ExcelDataTable.ReadData(1, RefResources[1])).ToString();
                                            }
                                        }
                                        else if (referenceData.Contains('-'))
                                        {
                                            RefResources = referenceData.Split('-');
                                            valueMatch = pop.getDataFromDynamicExcel(RefResources[0]);
                                            if (RefResources[1].ToString().Contains('='))
                                            {
                                                valueMatch = pop.calculatedata(valueMatch + '-' + pop.getDataFromDynamicExcel(RefResources[1])).ToString();
                                            }
                                            else
                                            {
                                                valueMatch = pop.calculatedata(valueMatch + '-' + ExcelDataTable.ReadData(1, RefResources[1])).ToString();
                                            }
                                        }
                                        else if (referenceData.Contains('*'))
                                        {
                                            RefResources = referenceData.Split('*');
                                            valueMatch = pop.getDataFromDynamicExcel(RefResources[0]);
                                            if (RefResources[1].ToString().Contains('='))
                                            {
                                                valueMatch = pop.calculatedata(valueMatch + '*' + pop.getDataFromDynamicExcel(RefResources[1])).ToString();
                                            }
                                            else
                                            {
                                                valueMatch = pop.calculatedata(valueMatch + '*' + ExcelDataTable.ReadData(1, RefResources[1])).ToString();
                                            }
                                        }
                                    }
                                    else
                                    {
                                        valueMatch = pop.getDataFromDynamicExcel(count["DataRefferencekeyword"].ToString());
                                    }
                                }
                                else
                                {
                                    if (count["DataRefferencekeyword"].ToString().Contains('+') || count["DataRefferencekeyword"].ToString().Contains('-') || count["DataRefferencekeyword"].ToString().Contains('*'))
                                    {
                                        if (count["DataRefferencekeyword"].ToString().Contains('+'))
                                        {
                                            RefResources = count["DataRefferencekeyword"].ToString().Split('+');
                                            valueMatch = pop.calculatedata(ExcelDataTable.ReadData(1, RefResources[0]) + '+' + ExcelDataTable.ReadData(1, RefResources[1])).ToString();
                                        }
                                        else if(count["DataRefferencekeyword"].ToString().Contains('-'))
                                        {
                                            RefResources = count["DataRefferencekeyword"].ToString().Split('-');
                                            valueMatch = pop.calculatedata(ExcelDataTable.ReadData(1, RefResources[0]) + '-' + ExcelDataTable.ReadData(1, RefResources[1])).ToString();
                                        }
                                        else if (count["DataRefferencekeyword"].ToString().Contains('*'))
                                        {
                                            RefResources = count["DataRefferencekeyword"].ToString().Split('*');
                                            valueMatch = pop.calculatedata(ExcelDataTable.ReadData(1, RefResources[0]) + '*' + ExcelDataTable.ReadData(1, RefResources[1])).ToString();
                                        }
                                        else
                                        {
                                            throw new Exception("No Calculation Symbol Not Found..Calculation is not possible");
                                        }
                                    }
                                    else
                                    {
                                        Testdata = count["DataRefferencekeyword"].ToString().Split(':');
                                        if (Testdata.Count() > 1)
                                        {
                                            int number;
                                            if (Int32.TryParse(Testdata[1], out number))
                                            {
                                                valueMatch = ExcelDataTable.ReadData(Convert.ToInt16(Testdata[1]), Testdata[0]);
                                            }
                                            else
                                            {
                                                CommonLanguageTemplateReader.Message(lang, count["DataRefferencekeyword"].ToString());
                                                valueMatch = languageResource.Msg_GetTemplateMessage;
                                            }
                                        }
                                        else
                                        {
                                            valueMatch = ExcelDataTable.ReadData(1, Testdata[0]);
                                        }
                                    }
                                }
                                if (operation == "GetValueMatch")
                                {
                                    if (valueGEt != valueMatch)
                                    {
                                        screenShotName = count["AssertionScreenshotName"].ToString();
                                        ErrorScreenPath = log.screenShot(screenShotName);
                                        FailerReason = count["AssertionMessage"].ToString();
                                        repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                        repgen.FileCorreptionCheck();
                                    }
                                    Assert.AreEqual(valueMatch, valueGEt, FailerReason);
                                }
                                else if (operation == "GetValueNoMatch")
                                {
                                    if (valueGEt == valueMatch)
                                    {
                                        screenShotName = count["AssertionScreenshotName"].ToString();
                                        ErrorScreenPath = log.screenShot(screenShotName);
                                        FailerReason = count["AssertionMessage"].ToString();
                                        repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                        repgen.FileCorreptionCheck();
                                    }
                                    Assert.AreNotEqual(valueMatch, valueGEt, FailerReason);
                                }
                                else if(operation == "GetValueContains")
                                {
                                    if (count["DataRefferencekeyword"].ToString().Contains("Container Number"))
                                    {
                                        if (count["DataRefferencekeyword"].ToString().Split(':')[1] == "Container Number")
                                        {
                                            valueMatch = valueMatch.Substring(2, valueMatch.Length - 2);
                                        }
                                    }
                                    if (!valueGEt.Contains(valueMatch))
                                    {
                                        screenShotName = count["AssertionScreenshotName"].ToString();
                                        ErrorScreenPath = log.screenShot(screenShotName);
                                        FailerReason = count["AssertionMessage"].ToString();
                                        repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                        repgen.FileCorreptionCheck();
                                    }
                                    StringAssert.Contains(valueGEt, valueMatch, FailerReason);
                                }
                                repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                break;
                            case "ValidatePOStatus":
                                string POStatus = log.ValidateCalenderStatus(count["ControlKeywordValue"].ToString(), count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString());
                                CommonLanguageTemplateReader.Message(lang, count["DataRefferencekeyword"].ToString());
                                if (languageResource.Msg_GetTemplateMessage.Replace(" ", string.Empty) != POStatus.Trim())
                                {
                                    screenShotName = count["AssertionScreenshotName"].ToString();
                                    ErrorScreenPath= log.screenShot(screenShotName);
                                    FailerReason = count["AssertionMessage"].ToString();
                                    repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                    repgen.FileCorreptionCheck();
                                }
                                Assert.AreEqual(languageResource.Msg_GetTemplateMessage.Replace(" ", string.Empty), POStatus, FailerReason);
                                repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                break;
                            case "ValidateEmptyValueEqual":
                            case "ValidateEmptyValueNotEqual":
                                string EmptyvalueGEt = pop.ScriptGetExcuite(count["Keyword"].ToString(), count["ControlKeyword"].ToString(), count["ControlKeywordValue"].ToString(), count["Step No"].ToString(), count["Description"].ToString());
                                validationMessage = Regex.Replace(EmptyvalueGEt, "<.*?>", String.Empty).Trim();
                                Message = System.Web.HttpUtility.HtmlDecode(validationMessage);
                                validationMessage = Message.Replace(" ", string.Empty).Trim();
                                if (operation == "ValidateEmptyValueEqual")
                                {
                                    if (validationMessage != "")
                                    {
                                        screenShotName = count["AssertionScreenshotName"].ToString();
                                        ErrorScreenPath = log.screenShot(screenShotName);
                                        FailerReason = count["AssertionMessage"].ToString();
                                        repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                        repgen.FileCorreptionCheck();
                                    }
                                    Assert.AreEqual("", validationMessage, FailerReason);
                                }
                                else if (operation == "ValidateEmptyValueNotEqual")
                                {
                                    if (validationMessage == "")
                                    {
                                        screenShotName = count["AssertionScreenshotName"].ToString();
                                        ErrorScreenPath = log.screenShot(screenShotName);
                                        FailerReason = count["AssertionMessage"].ToString();
                                        repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                        repgen.FileCorreptionCheck();
                                    }
                                    Assert.AreNotEqual("", validationMessage, FailerReason);
                                }
                                repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                break;
                            case "ValidateGRStatus":
                                string GRStatus = log.ValidateGRStatus(ExcelDataTable.ReadData(1, count["ControlKeywordValue"].ToString()), count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString());
                                validationMessage = GRStatus.Replace(" ", string.Empty).Trim();
                                CommonLanguageTemplateReader.Message(lang, count["DataRefferencekeyword"].ToString());
                                if (languageResource.Msg_GetTemplateMessage.Replace(" ", string.Empty) != validationMessage.Trim())
                                {
                                    screenShotName = count["AssertionScreenshotName"].ToString();
                                    ErrorScreenPath = log.screenShot(screenShotName);
                                    FailerReason = count["AssertionMessage"].ToString();
                                    repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                    repgen.FileCorreptionCheck();
                                }
                                Assert.AreEqual(languageResource.Msg_GetTemplateMessage.Replace(" ", string.Empty), validationMessage, FailerReason);
                                repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                break;
                            case "RecordData":
                                RecordData(count["DataRefferencekeyword"].ToString(), pop.WebGetControlData(count["TypeOfControl"].ToString(), count["ControlKeyword"].ToString(), count["ControlKeywordValue"].ToString(), count["AssertionScreenshotName"].ToString(), count["AssertionMessage"].ToString(), count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString()).Trim());
                                repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, recordedData[count["DataRefferencekeyword"].ToString()]);
                                break;
                            case "GetRecordData":
                                pop.OperationStart("EnterText", count["TypeOfControl"].ToString(), count["ControlKeyword"].ToString(), count["ControlKeywordValue"].ToString(), recordedData[count["DataRefferencekeyword"].ToString()], count["TypeOfWindow"].ToString(), count["Step No"].ToString(), count["Description"].ToString());
                                repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, recordedData[count["DataRefferencekeyword"].ToString()]);
                                break;
                            case "assignPropsedGate":
                                CommonLanguageTemplateReader.Message(lang, count["DataRefferencekeyword"].ToString());
                                string proposedgate = log.getProposedGate(count["ControlKeywordValue"].ToString(), languageResource.Msg_GetTemplateMessage, count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString());
                                if (proposedgate == "")
                                {
                                    screenShotName = count["AssertionScreenshotName"].ToString();
                                    ErrorScreenPath = log.screenShot(screenShotName);
                                    FailerReason = count["AssertionMessage"].ToString();
                                    repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                    repgen.FileCorreptionCheck();
                                    Assert.Fail(FailerReason);
                                }
                                else
                                {
                                    Thread.Sleep(max * 2);
                                    pop.ScriptClick("ScriptClick", "ID", proposedgate, count["Step No"].ToString(), count["Description"].ToString());
                                }
                                repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                break;
                            case "TruckDetailsValidate":
                                string truckAddstatus = string.Empty;
                                if (count["DataRefferencekeyword"].ToString().Contains(':'))
                                {
                                    WriteAndReadData.DataFilePath = ConfigurationManager.AppSettings["ExcelDataFile"] + "\\DynamicData" + "\\" + BatchName + ".xls";
                                    ExcelDataTable.PopulateRecordData(WriteAndReadData.DataFilePath);
                                    truckAddstatus = log.ValidateTruckDetails(count["ControlKeywordValue"].ToString(), count["DataRefferencekeyword"].ToString(), count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString());
                                }
                                else if(count["DataRefferencekeyword"].ToString().StartsWith("Rec_"))
                                {
                                    truckAddstatus = log.ValidateTruckDetails(count["ControlKeywordValue"].ToString(), recordedData[count["DataRefferencekeyword"].ToString()], count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString());
                                }
                                string[] truckStatus = truckAddstatus.Split(':');
                                if (truckStatus[1] == "false")
                                {
                                    screenShotName = count["AssertionScreenshotName"].ToString();
                                    ErrorScreenPath = log.screenShot(screenShotName);
                                    FailerReason = count["AssertionMessage"].ToString();
                                    repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                    repgen.FileCorreptionCheck();
                                    Assert.Fail(FailerReason);
                                }
                                writeData.WriteExcel("Gate Number", truckStatus[0], "ReadData", count["DataRefferencekeyword"].ToString().Split('=')[1].Split(':')[0]);
                                repgen.Reports(count["Step No"].ToString(), "Record Assigned Gate Number for the Truck", "ReadData", true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, truckStatus[0]);
                                repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                break;
                            case "ScrollToControl":
                                string ScrollStatus = pop.ScriptGetExcuite(count["Keyword"].ToString(), count["ControlKeyword"].ToString(), count["ControlKeywordValue"].ToString(), count["Step No"].ToString(), count["Description"].ToString());
                                if (ScrollStatus == "true")
                                {
                                    repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                }
                                else
                                {
                                    repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                }
                                break;
                            case "ValidateControlCount":
                                int ControlCount = Convert.ToInt16(pop.ScriptGetExcuite(count["Keyword"].ToString(), count["ControlKeyword"].ToString(), count["ControlKeywordValue"].ToString(), count["Step No"].ToString(), count["Description"].ToString()));
                                if (ControlCount <= 0)
                                {
                                    screenShotName = count["AssertionScreenshotName"].ToString();
                                    ErrorScreenPath = log.screenShot(screenShotName);
                                    FailerReason = count["AssertionMessage"].ToString();
                                    repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                    repgen.FileCorreptionCheck();
                                    Assert.Fail(FailerReason);
                                }
                                repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                break;
                            case "SaveQualityResult":
                                log.EnterResultForQuality(count["ControlKeywordValue"].ToString(), ExcelDataTable.ReadData(1, count["DataRefferencekeyword"].ToString()), count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString());
                                repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                break;
                            case "CheckQualityResult":
                                string qualityStatus = log.CheckQualityResult(count["ControlKeywordValue"].ToString(), count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString());
                                CommonLanguageTemplateReader.Message(lang, count["DataRefferencekeyword"].ToString());
                                if (languageResource.Msg_GetTemplateMessage.Replace(" ", string.Empty) != qualityStatus.Trim())
                                {
                                    screenShotName = count["AssertionScreenshotName"].ToString();
                                    ErrorScreenPath = log.screenShot(screenShotName);
                                    FailerReason = count["AssertionMessage"].ToString();
                                    repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                    repgen.FileCorreptionCheck();
                                }
                                Assert.AreEqual(languageResource.Msg_GetTemplateMessage.Replace(" ", string.Empty), qualityStatus, FailerReason);
                                repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                break;
                            case "SamplingDataCheck":
                                qualityStatus = log.ValidateSamplingInstruction(count["ControlKeywordValue"].ToString(), count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString());
                                if (qualityStatus == "false")
                                {
                                    screenShotName = count["AssertionScreenshotName"].ToString();
                                    ErrorScreenPath = log.screenShot(screenShotName);
                                    FailerReason = count["AssertionMessage"].ToString();
                                    repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                    repgen.FileCorreptionCheck();
                                    Assert.Fail(FailerReason);
                                }
                                repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                break;
                            case "ValidateIsJIT":
                                qualityStatus = log.validateIsJit(count["ControlKeywordValue"].ToString(), count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString());
                                if (qualityStatus == "false")
                                {
                                    screenShotName = count["AssertionScreenshotName"].ToString();
                                    ErrorScreenPath = log.screenShot(screenShotName);
                                    FailerReason = count["AssertionMessage"].ToString();
                                    repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                    repgen.FileCorreptionCheck();
                                    Assert.Fail(FailerReason);
                                }
                                repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                break;
                            case "ValidateAvailMetrial":
                                string materialStatus = string.Empty;
                                if (count["DataRefferencekeyword"].ToString().Contains('='))
                                {
                                    WriteAndReadData.DataFilePath = ConfigurationManager.AppSettings["ExcelDataFile"] + "\\DynamicData" + "\\" + BatchName + ".xls";
                                    ExcelDataTable.PopulateRecordData(WriteAndReadData.DataFilePath);
                                    string value = pop.getDataFromDynamicExcel(count["DataRefferencekeyword"].ToString());
                                    materialStatus = log.ValidateMetrialAvailable(count["ControlKeywordValue"].ToString(), value, count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString());
                                }
                                else
                                {
                                    materialStatus = log.ValidateMetrialAvailable(count["ControlKeywordValue"].ToString(), ExcelDataTable.ReadData(1, count["DataRefferencekeyword"].ToString()), count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString());
                                }
                                if (materialStatus == "false")
                                {
                                    screenShotName = count["AssertionScreenshotName"].ToString();
                                    ErrorScreenPath = log.screenShot(screenShotName);
                                    FailerReason = count["AssertionMessage"].ToString();
                                    repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                    repgen.FileCorreptionCheck();
                                    Assert.Fail(FailerReason);
                                }
                                repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                break;
                            case "IsReadOnly":
                            case "IsNotReadOnly":
                            case "IsChecked":
                                ScrollStatus = pop.ScriptGetExcuite(count["Keyword"].ToString(), count["ControlKeyword"].ToString(), count["ControlKeywordValue"].ToString(), count["Step No"].ToString(), count["Description"].ToString());
                                if(operation== "IsNotReadOnly")
                                {
                                    if (ScrollStatus == "True")
                                    {
                                        screenShotName = count["AssertionScreenshotName"].ToString();
                                        ErrorScreenPath = log.screenShot(screenShotName);
                                        FailerReason = count["AssertionMessage"].ToString();
                                        repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                        repgen.FileCorreptionCheck();
                                        Assert.Fail(FailerReason);
                                    }
                                }
                                else
                                {
                                    if (ScrollStatus == "False")
                                    {
                                        screenShotName = count["AssertionScreenshotName"].ToString();
                                        ErrorScreenPath = log.screenShot(screenShotName);
                                        FailerReason = count["AssertionMessage"].ToString();
                                        repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                        repgen.FileCorreptionCheck();
                                        Assert.Fail(FailerReason);
                                    }
                                }
                                repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                break;
                            case "WaitPageLoad":
                                status = pop.PageLoadWait(count["TypeOfControl"].ToString(), count["ControlKeyword"].ToString(), count["ControlKeywordValue"].ToString(), count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString());
                                if (!status)
                                {
                                    repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                }
                                repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                break;
                            case "DownloadUpload":
                                Thread.Sleep(max);
                                log.ClickDownloadUpload(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString());
                                repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                break;
                            case "Check":
                                pop.OperationStart(count["Keyword"].ToString(), count["TypeOfControl"].ToString(), count["ControlKeyword"].ToString(), count["ControlKeywordValue"].ToString(), ExcelDataTable.ReadData(1, count["DataRefferencekeyword"].ToString()), count["TypeOfWindow"].ToString(), count["Step No"].ToString(), count["Description"].ToString());
                                break;
                            case "ValidateCustomerInst":
                                qualityStatus = log.ValidateCustomerInst(count["ControlKeywordValue"].ToString(), count["DataRefferencekeyword"].ToString(), count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString());
                                if (qualityStatus == "false")
                                {
                                    screenShotName = count["AssertionScreenshotName"].ToString();
                                    ErrorScreenPath = log.screenShot(screenShotName);
                                    FailerReason = count["AssertionMessage"].ToString();
                                    repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                    repgen.FileCorreptionCheck();
                                    Assert.Fail(FailerReason);
                                }
                                repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                break;
                            case "SelectPPE":
                                qualityStatus = log.SelectPPEInfo(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), count["DataRefferencekeyword"].ToString());
                                if (qualityStatus == "False")
                                {
                                    screenShotName = count["AssertionScreenshotName"].ToString();
                                    ErrorScreenPath = log.screenShot(screenShotName);
                                    FailerReason = count["AssertionMessage"].ToString();
                                    repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                    repgen.FileCorreptionCheck();
                                    Assert.Fail(FailerReason);
                                }
                                repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                break;
                            case "SelectDGClass":
                                qualityStatus = log.SelectDGInfo(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), count["DataRefferencekeyword"].ToString());
                                if (qualityStatus == "False")
                                {
                                    screenShotName = count["AssertionScreenshotName"].ToString();
                                    ErrorScreenPath = log.screenShot(screenShotName);
                                    FailerReason = count["AssertionMessage"].ToString();
                                    repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                    repgen.FileCorreptionCheck();
                                    Assert.Fail(FailerReason);
                                }
                                repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                break;
                            case "ValidateDGMetrilalStatus":
                                string MetrialDGtatus = log.ValidateDG_PPE(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), count["DataRefferencekeyword"].ToString());
                                string[] DGStatus = MetrialDGtatus.Split(':');
                                if (DGStatus[0] == "false")
                                {
                                    screenShotName = count["AssertionScreenshotName"].ToString();
                                    ErrorScreenPath = log.screenShot(screenShotName);
                                    FailerReason = count["AssertionMessage"].ToString();
                                    repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                    repgen.FileCorreptionCheck();
                                    Assert.Fail(FailerReason);
                                }
                                repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                break;
                            case "ValidatePPEStatus":
                                bool Metrialstatus = log.ValidatePPE_Avail(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString());
                                if (!Metrialstatus)
                                {
                                    screenShotName = count["AssertionScreenshotName"].ToString();
                                    ErrorScreenPath = log.screenShot(screenShotName);
                                    FailerReason = count["AssertionMessage"].ToString();
                                    repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                    repgen.FileCorreptionCheck();
                                    Assert.Fail(FailerReason);
                                }
                                repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                break;
                            case "AssignedOrderValidate":
                            case "ValidatePickStatus":
                            case "ValidatePackStatus":
                                CommonLanguageTemplateReader.Message(lang, count["DataRefferencekeyword"].ToString());
                                string message = languageResource.Msg_GetTemplateMessage;
                                if (operation == "AssignedOrderValidate")
                                {
                                    qualityStatus = log.AssignedOrderValidate(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), count["ControlKeywordValue"].ToString(), message);
                                }
                                else if (operation == "ValidatePackStatus")
                                {
                                    qualityStatus = log.MobilePackOrderStatus(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), ExcelDataTable.ReadData(1, count["ControlKeywordValue"].ToString()), message);
                                }
                                else
                                {
                                    qualityStatus = log.MobilePickOrderStatus(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), ExcelDataTable.ReadData(1, count["ControlKeywordValue"].ToString()), message);
                                }
                                if (qualityStatus == "False")
                                {
                                    screenShotName = count["AssertionScreenshotName"].ToString();
                                    ErrorScreenPath = log.screenShot(screenShotName);
                                    FailerReason = count["AssertionMessage"].ToString();
                                    repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                    repgen.FileCorreptionCheck();
                                    Assert.Fail(FailerReason);
                                }
                                repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                break;
                            case "GetDispatchBinLocation":
                            case "GetSAPStock_Batch":
                                string Dispatchbin = string.Empty;
                                if (operation == "GetDispatchBinLocation")
                                {
                                    Dispatchbin = log.GettingBinLocation(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), count["ControlKeywordValue"].ToString());
                                }
                                else if (operation == "GetSAPStock_Batch")
                                {
                                    Dispatchbin = log.GettingSAP_BatchNumber(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), count["ControlKeywordValue"].ToString());
                                }
                                if (Dispatchbin == string.Empty)
                                {
                                    screenShotName = count["AssertionScreenshotName"].ToString();
                                    ErrorScreenPath = log.screenShot(screenShotName);
                                    FailerReason = count["AssertionMessage"].ToString();
                                    repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                    repgen.FileCorreptionCheck();
                                    Assert.Fail(FailerReason);
                                }
                                writeData.WriteExcel(count["DataRefferencekeyword"].ToString().Split(':')[1], Dispatchbin.Split(':')[0], "ReadData", ExcelDataTable.ReadData(1, count["DataRefferencekeyword"].ToString().Split(':')[0]).ToString());
                                if(operation== "GetDispatchBinLocation")
                                {
                                    writeData.WriteExcel("Available SAP Quantity", Dispatchbin.Split(':')[1], "ReadData", ExcelDataTable.ReadData(1, count["DataRefferencekeyword"].ToString().Split(':')[0]).ToString());
                                }
                                writeData.FileCorreptionCheck();
                                repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, Dispatchbin);
                                break;
                            case "DispatchPPEValidation":
                            case "DispatchPickInstValidation":
                                if (operation == "DispatchPPEValidation")
                                {
                                    status = log.DispatchMobilePPEValidation(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString());
                                }
                                else
                                {
                                    status = log.DispatchMobilePickInstruction(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString());
                                }
                                if (!status)
                                {
                                    screenShotName = count["AssertionScreenshotName"].ToString();
                                    ErrorScreenPath = log.screenShot(screenShotName);
                                    FailerReason = count["AssertionMessage"].ToString();
                                    repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                    repgen.FileCorreptionCheck();
                                    Assert.Fail(FailerReason);
                                }
                                repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                break;
                            case "ValidatePickOrderStatus":
                                qualityStatus = log.DispatchMobPickingOrderStatus(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), ExcelDataTable.ReadData(1, count["ControlKeywordValue"].ToString()));
                                CommonLanguageTemplateReader.Message(lang, count["DataRefferencekeyword"].ToString());
                                if (languageResource.Msg_GetTemplateMessage.Trim() != qualityStatus.Trim())
                                {
                                    screenShotName = count["AssertionScreenshotName"].ToString();
                                    ErrorScreenPath = log.screenShot(screenShotName);
                                    FailerReason = count["AssertionMessage"].ToString();
                                    repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                    repgen.FileCorreptionCheck();
                                }
                                Assert.AreEqual(languageResource.Msg_GetTemplateMessage.Trim(), qualityStatus, FailerReason);
                                repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                break;
                            case "ValidateOrderAvailStatus":
                                status = log.CheckOrderAvailable(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), ExcelDataTable.ReadData(1, count["ControlKeywordValue"].ToString()));
                                if (!status)
                                {
                                    screenShotName = count["AssertionScreenshotName"].ToString();
                                    ErrorScreenPath = log.screenShot(screenShotName);
                                    FailerReason = count["AssertionMessage"].ToString();
                                    repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                    repgen.FileCorreptionCheck();
                                    Assert.Fail(FailerReason);
                                }
                                repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                break;
                            case "ValidatePackInstructionAvail":
                                status = log.PackingInstructions(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString());
                                if (!status)
                                {
                                    screenShotName = count["AssertionScreenshotName"].ToString();
                                    ErrorScreenPath = log.screenShot(screenShotName);
                                    FailerReason = count["AssertionMessage"].ToString();
                                    repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                    repgen.FileCorreptionCheck();
                                    Assert.Fail(FailerReason);
                                }
                                repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                break;
                            case "ValidatePartialPickSatus":
                                qualityStatus = pop.ScriptGetExcuite(count["Keyword"].ToString(), count["ControlKeyword"].ToString(), count["ControlKeywordValue"].ToString(), count["Step No"].ToString(), count["Description"].ToString());
                                CommonLanguageTemplateReader.Message(lang, count["DataRefferencekeyword"].ToString());
                                if (languageResource.Msg_GetTemplateMessage.Trim() != qualityStatus.Trim())
                                {
                                    screenShotName = count["AssertionScreenshotName"].ToString();
                                    ErrorScreenPath = log.screenShot(screenShotName);
                                    FailerReason = count["AssertionMessage"].ToString();
                                    repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                    repgen.FileCorreptionCheck();
                                }
                                Assert.AreEqual(languageResource.Msg_GetTemplateMessage.Trim(), qualityStatus, FailerReason);
                                repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                break;
                            case "MatchPackageDetails":
                                status = log.PackagingDetailsValidation(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), count["ControlKeywordValue"].ToString(), count["DataRefferencekeyword"].ToString());
                                if (!status)
                                {
                                    screenShotName = count["AssertionScreenshotName"].ToString();
                                    ErrorScreenPath = log.screenShot(screenShotName);
                                    FailerReason = count["AssertionMessage"].ToString();
                                    repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                    repgen.FileCorreptionCheck();
                                    Assert.Fail(FailerReason);
                                }
                                repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                break;
                            case "LinkPPE":
                                qualityStatus = log.PPELink(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), count["ControlKeywordValue"].ToString(), ExcelDataTable.ReadData(1, count["DataRefferencekeyword"].ToString()));
                                if (qualityStatus == "False")
                                {
                                    screenShotName = count["AssertionScreenshotName"].ToString();
                                    ErrorScreenPath = log.screenShot(screenShotName);
                                    FailerReason = count["AssertionMessage"].ToString();
                                    repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                    repgen.FileCorreptionCheck();
                                    Assert.Fail(FailerReason);
                                }
                                repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                break;
                            case "SearchDataOperation":
                                log.NavigationPageDeleteData(count["Step No"].ToString(), count["Description"].ToString(), count["ControlKeywordValue"].ToString(), ExcelDataTable.ReadData(1, count["DataRefferencekeyword"].ToString()), count["TypeOfControl"].ToString(), new int[] { Convert.ToInt16(count["XCodinate"]), Convert.ToInt16(count["Ycodinate"]) });
                                break;
                            case "GetDeviceInfoID":
                                qualityStatus = log.GETDeviceID(count["Step No"].ToString(), count["Description"].ToString(), count["DataRefferencekeyword"].ToString(), count["ControlKeywordValue"].ToString());
                                if (qualityStatus == "0")
                                {
                                    screenShotName = count["AssertionScreenshotName"].ToString();
                                    ErrorScreenPath = log.screenShot(screenShotName);
                                    FailerReason = count["AssertionMessage"].ToString();
                                    repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                    repgen.FileCorreptionCheck();
                                    Assert.Fail(FailerReason);
                                }
                                RecordData("DeviceInfoID", qualityStatus);
                                repgen.Reports(count["Step No"].ToString(), "Record Device ID For the Give details.", "RecordData", true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, recordedData["DeviceInfoID"]);
                                break;
                            case "GetDeviceIDValue":
                                valueGEt = pop.ScriptGetExcuite(count["Keyword"].ToString(), count["ControlKeyword"].ToString(), count["ControlKeywordValue"].ToString(), count["Step No"].ToString(), count["Description"].ToString());
                                RecordData("DeviceID", valueGEt);
                                repgen.Reports(count["Step No"].ToString(), "Record Device ID For the Give details.", "RecordData", true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, recordedData["DeviceID"]);
                                break;
                            case "BrowserTabVal":
                                status = pop.BrowserTabValidation(count["ControlKeywordValue"].ToString());
                                if (!status)
                                {
                                    screenShotName = count["AssertionScreenshotName"].ToString();
                                    ErrorScreenPath = log.screenShot(screenShotName);
                                    FailerReason = count["AssertionMessage"].ToString();
                                    repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                    repgen.FileCorreptionCheck();
                                    Assert.Fail(FailerReason);
                                }
                                repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                break;
                            case "ValidateDropdown":
                                validationMessage = pop.WebGetControlData(count["TypeOfControl"].ToString(), count["ControlKeyword"].ToString(), count["ControlKeywordValue"].ToString(), count["AssertionScreenshotName"].ToString(), count["AssertionMessage"].ToString(), count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString()).Trim();
                                if (!validationMessage.Contains(ExcelDataTable.ReadData(1, count["DataRefferencekeyword"].ToString())))
                                {
                                    screenShotName = count["AssertionScreenshotName"].ToString();
                                    ErrorScreenPath = log.screenShot(screenShotName);
                                    FailerReason = count["AssertionMessage"].ToString();
                                    repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                    repgen.FileCorreptionCheck();
                                    Assert.Fail(FailerReason);
                                }
                                StringAssert.Contains(validationMessage, ExcelDataTable.ReadData(1, count["DataRefferencekeyword"].ToString()), FailerReason);
                                repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                break;
                            case "GetRemainingPickQty":
                                string remainingQty = log.GetRemainingQuantity(count["ControlKeywordValue"].ToString(), ExcelDataTable.ReadData(1, count["DataRefferencekeyword"].ToString()), count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString()).Trim();
                                RecordData("RemainingOrderQTY", remainingQty);
                                repgen.Reports(count["Step No"].ToString(), "Record Remaining Quantity for order.", "RecordData", true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, recordedData["RemainingOrderQTY"]);
                                break;
                            case "GetMaterialBatchData":
                                string BatchData = log.GetBatchReqQuantitry(count["ControlKeywordValue"].ToString(),count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString()).Trim();

                                //Recording Batch Number of the given material.
                                string referenceName = count["ControlKeywordValue"].ToString().Split(':')[1] + "_Batch";
                                writeData.WriteExcel("Batch Number", BatchData.Split(':')[0], "ReadData", ExcelDataTable.ReadData(1,count["ControlKeywordValue"].ToString().Split(':')[1]).ToString());
                                repgen.Reports(count["Step No"].ToString(), "Write down Batch Number for specified "+ count["ControlKeywordValue"].ToString().Split(':')[1] + "", "ReadData", true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, BatchData.Split(':')[0]);

                                //Recording Required quantity of the given material.
                                referenceName = count["ControlKeywordValue"].ToString().Split(':')[1] + "_ReqQty";
                                writeData.WriteExcel("Required Quantity", BatchData.Split(':')[1], "ReadData", ExcelDataTable.ReadData(1, count["ControlKeywordValue"].ToString().Split(':')[1]).ToString());
                                repgen.Reports(count["Step No"].ToString(), "Write down Required Quantity for specified " + count["ControlKeywordValue"].ToString().Split(':')[1] + "", "ReadData", true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, BatchData.Split(':')[1]);
                                break;
                            case "SetRef_Data":
                                string inputdata = string.Empty;
                                if (count["DataRefferencekeyword"].ToString().StartsWith("Rec_"))
                                {
                                    inputdata = recordedData[count["DataRefferencekeyword"].ToString().Replace("Rec_", string.Empty)];
                                }
                                else
                                {
                                    inputdata = ExcelDataTable.ReadData(1, count["DataRefferencekeyword"].ToString());
                                }
                                if (inputdata != count["DataRefferencekeyword"].ToString())
                                {
                                    writeData.WriteExcel(count["ControlKeywordValue"].ToString(), inputdata, count["Keyword"].ToString(), "");
                                    repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                }
                                else
                                {
                                    if (count["DataRefferencekeyword"].ToString().StartsWith("Data"))
                                    {
                                        writeData.WriteExcel(count["ControlKeywordValue"].ToString(), inputdata.Split(':')[1], count["Keyword"].ToString(), "");
                                        repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                    }
                                    else
                                    {
                                        repgen.Reports(count["Step No"].ToString(), "No Data found in " + count["DataRefferencekeyword"].ToString() + "in test data", count["Keyword"].ToString(), false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                        Assert.Fail("No Data found in " + count["DataRefferencekeyword"].ToString() + "in test data");
                                    }
                                }
                                break;
                            case "WriteData":
                                if (WriteAndReadData.DataFilePath == string.Empty)
                                {
                                    WriteAndReadData.DataFilePath = ConfigurationManager.AppSettings["ExcelDataFile"] + "\\DynamicData" + "\\" + BatchName + ".xls";
                                } 
                                pop.GetAndWriteDataforReff(count["ControlKeyword"].ToString(),count["TypeOfControl"].ToString(), count["TypeOfWindow"].ToString(), count["ControlKeywordValue"].ToString(), count["DataRefferencekeyword"].ToString(), count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString());
                                break;
                            case "ReadData":
                                if (WriteAndReadData.DataFilePath == string.Empty)
                                {
                                    WriteAndReadData.DataFilePath = ConfigurationManager.AppSettings["ExcelDataFile"] + "\\DynamicData" + "\\" + BatchName + ".xls";
                                }
                                if (count["ControlKeywordValue"].ToString().StartsWith("Rec_")) 
                                {
                                    validationMessagenotEqual = recordedData[count["ControlKeywordValue"].ToString().Replace("Rec_",string.Empty)];
                                }
                                else
                                {
                                    validationMessagenotEqual = pop.WebGetControlData(count["TypeOfControl"].ToString(), count["ControlKeyword"].ToString(), count["ControlKeywordValue"].ToString(), count["AssertionScreenshotName"].ToString(), count["AssertionMessage"].ToString(), count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString()).Trim();
                                }
                                writeData.WriteExcel(count["DataRefferencekeyword"].ToString().Split(':')[1], validationMessagenotEqual, "ReadData", ExcelDataTable.ReadData(1, count["DataRefferencekeyword"].ToString().Split(':')[0]).ToString());
                                repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, validationMessagenotEqual);
                                break;
                            case "ReadStockInfo":
                                if (WriteAndReadData.DataFilePath == string.Empty)
                                {
                                    WriteAndReadData.DataFilePath = ConfigurationManager.AppSettings["ExcelDataFile"] + "\\DynamicData" + "\\" + BatchName + ".xls";
                                }
                                validationMessage = pop.WebGetControlData(count["TypeOfControl"].ToString(), count["ControlKeyword"].ToString(), count["ControlKeywordValue"].ToString().Split(':')[0], count["AssertionScreenshotName"].ToString(), count["AssertionMessage"].ToString(), count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString()).Trim();
                                validationMessage = Regex.Replace(validationMessage, "<.*?>", String.Empty).Trim();
                                Message = System.Web.HttpUtility.HtmlDecode(validationMessage);
                                if (Message == "No records found")
                                {
                                    writeData.WriteExcel(count["DataRefferencekeyword"].ToString().Split(':')[1], "0", "ReadData", ExcelDataTable.ReadData(1, count["DataRefferencekeyword"].ToString().Split(':')[0]).ToString());
                                    repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "0");
                                }
                                else
                                {
                                    validationMessagenotEqual = pop.WebGetControlData(count["TypeOfControl"].ToString(), count["ControlKeyword"].ToString(), count["ControlKeywordValue"].ToString().Split(':')[1], count["AssertionScreenshotName"].ToString(), count["AssertionMessage"].ToString(), count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString()).Trim();
                                    writeData.WriteExcel(count["DataRefferencekeyword"].ToString().Split(':')[1], validationMessagenotEqual, "ReadData", ExcelDataTable.ReadData(1, count["DataRefferencekeyword"].ToString().Split(':')[0]).ToString());
                                    repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, validationMessagenotEqual);
                                }
                                break;
                            case "WriteMaterialOrder":
                                string MaterialOrder=pop.WebGetControlData(count["TypeOfControl"].ToString(), count["ControlKeyword"].ToString(), count["ControlKeywordValue"].ToString(), count["AssertionScreenshotName"].ToString(), count["AssertionMessage"].ToString(), count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString()).Trim();
                                if(MaterialOrder != string.Empty)
                                {
                                    if (count["DataRefferencekeyword"].ToString().Contains('+'))
                                    {
                                        string[] metrialOrder = count["DataRefferencekeyword"].ToString().Split('+');
                                        writeData.WriteExcel(metrialOrder[1].ToString().Split(':')[0], metrialOrder[1].ToString().Split(':')[1], "ReadData", ExcelDataTable.ReadData(1, metrialOrder[0].Split('=')[1]) + ":" + MaterialOrder);
                                    }
                                    else
                                    {
                                        writeData.WriteExcel(count["DataRefferencekeyword"].ToString().Split(':')[0], count["DataRefferencekeyword"].ToString().Split(':')[1], "ReadData", MaterialOrder);
                                    }
                                    repgen.Reports(count["Step No"].ToString(), "Write down Material order based on the picking Screen", operation, true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                }
                                else
                                {
                                    repgen.Reports(count["Step No"].ToString(), "Write down Material order based on the picking Screen", operation, false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                    Assert.Fail("Write down Material order based on the picking Screen Failed");
                                }
                                break;
                            case "EnterSelectedBin":
                                MaterialOrder = log.EnterProposedBin(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), count["ControlKeyword"].ToString(), count["TypeOfControl"].ToString(), count["TypeOfWindow"].ToString(),count["ControlKeywordValue"].ToString(), count["DataRefferencekeyword"].ToString());
                                repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), operation, true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                RecordData("PickedBin", MaterialOrder.Split(':')[0].ToString());
                                repgen.Reports(count["Step No"].ToString(), "Record Picked bin for Material.", "RecordData", true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, recordedData["PickedBin"]);
                                RecordData("PickedBinAvailableQTY", MaterialOrder.Split(':')[1].ToString());
                                repgen.Reports(count["Step No"].ToString(), "Record Picked bin for Material.", "RecordData", true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, recordedData["PickedBinAvailableQTY"]);
                                break;
                            case "ValidateReturnWarehouseInfo":
                                status = log.ValidateMaterialInfoINReturnToWarehouse(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), count["ControlKeywordValue"].ToString(), count["DataRefferencekeyword"].ToString());
                                if (!status)
                                {
                                    screenShotName = count["AssertionScreenshotName"].ToString();
                                    ErrorScreenPath = log.screenShot(screenShotName);
                                    FailerReason = count["AssertionMessage"].ToString();
                                    repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                    repgen.FileCorreptionCheck();
                                    Assert.Fail(FailerReason);
                                }
                                repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), operation, true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                break;
                            case "GetMaterialBatchStoreID":
                                BatchData = log.GetMaterialDetialsWithStoreID(count["ControlKeywordValue"].ToString(), count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString()).Trim();
                                string[] materialData = BatchData.Split(',');
                                foreach (var item in materialData)
                                {
                                    if (item != string.Empty)
                                    {
                                        string[] details = item.Split('-');
                                        writeData.WriteExcel("Material", details[1].Trim(), "ReadData", details[0].Split(':')[1].Trim());
                                        writeData.WriteExcel("Batch Number", details[2].Trim(), "ReadData", details[0].Split(':')[1].Trim());
                                        writeData.WriteExcel("Required Quantity", details[3].Trim(), "ReadData", details[0].Split(':')[1].Trim());
                                    }
                                }
                                writeData.FileCorreptionCheck();
                                repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), operation, true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                break;
                            case "ValidatePickedBinStatus":
                                ExcelDataTable.PopulateRecordData(WriteAndReadData.DataFilePath);
                                BatchData = pop.getDataFromDynamicExcel(count["ControlKeywordValue"].ToString());
                                validationMessage = log.GetPickedBinStatusColor(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), BatchData);
                                if (validationMessage != "True")
                                {
                                    validationMessage = Regex.Replace(validationMessage, "<.*?>", String.Empty).Trim();
                                    Message = System.Web.HttpUtility.HtmlDecode(validationMessage);
                                    validationMessage = Message.Replace(" ", string.Empty).Trim();
                                    CommonLanguageTemplateReader.Message(lang, count["DataRefferencekeyword"].ToString());
                                    if (languageResource.Msg_GetTemplateMessage.Replace(" ", string.Empty) != validationMessage.Trim())
                                    {
                                        screenShotName = count["AssertionScreenshotName"].ToString();
                                        ErrorScreenPath = log.screenShot(screenShotName);
                                        FailerReason = count["AssertionMessage"].ToString();
                                        repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                        repgen.FileCorreptionCheck();
                                    }
                                    Assert.AreEqual(languageResource.Msg_GetTemplateMessage.Replace(" ", string.Empty), validationMessage, FailerReason);
                                    repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                }
                                else
                                {
                                    repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "Picked Pin Not Found");
                                }
                                break;
                            case "ValidateMatchDataStatus":
                                if (WriteAndReadData.DataFilePath != string.Empty)
                                {
                                    ExcelDataTable.PopulateRecordData(WriteAndReadData.DataFilePath);
                                }
                                validationMessage = log.GetMatchTableDataStatusColor(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), count["ControlKeywordValue"].ToString());
                                validationMessage = Regex.Replace(validationMessage, "<.*?>", String.Empty).Trim();
                                Message = System.Web.HttpUtility.HtmlDecode(validationMessage);
                                validationMessage = Message.Replace(" ", string.Empty).Trim();
                                CommonLanguageTemplateReader.Message(lang, count["DataRefferencekeyword"].ToString());
                                if (languageResource.Msg_GetTemplateMessage.Replace(" ", string.Empty) != validationMessage.Trim())
                                {
                                    screenShotName = count["AssertionScreenshotName"].ToString();
                                    ErrorScreenPath = log.screenShot(screenShotName);
                                    FailerReason = count["AssertionMessage"].ToString();
                                    repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                    repgen.FileCorreptionCheck();
                                }
                                Assert.AreEqual(languageResource.Msg_GetTemplateMessage.Replace(" ", string.Empty), validationMessage, FailerReason);
                                repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                break;
                            case "WarehouseStockValidation":
                                status = log.ValidateTransferStockMaterialAndQty(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), count["ControlKeywordValue"].ToString(), count["DataRefferencekeyword"].ToString());
                                if (!status)
                                {
                                    screenShotName = count["AssertionScreenshotName"].ToString();
                                    ErrorScreenPath = log.screenShot(screenShotName);
                                    FailerReason = count["AssertionMessage"].ToString();
                                    repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                    repgen.FileCorreptionCheck();
                                    Assert.Fail(FailerReason);
                                }
                                repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                break;
                            case "LaunchDataBase":
                                log.sqldeveloperApplicationLaunch();
                                repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                break;
                            case "ConditionEquals":
                            case "ConditionNotEquals":
                                validationMessage = pop.WebGetControlData(count["TypeOfControl"].ToString(), count["ControlKeyword"].ToString(), count["ControlKeywordValue"].ToString(), count["AssertionScreenshotName"].ToString(), count["AssertionMessage"].ToString(), count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString()).Trim();
                                validationMessage = Regex.Replace(validationMessage, "<.*?>", String.Empty).Trim();
                                Message = System.Web.HttpUtility.HtmlDecode(validationMessage);
                                validationMessage = Message.Replace(" ", string.Empty).Trim();
                                CommonLanguageTemplateReader.Message(lang, count["DataRefferencekeyword"].ToString());
                                if (operation== "ConditionEquals")
                                {
                                    if (languageResource.Msg_GetTemplateMessage.Replace(" ", string.Empty) == validationMessage.Trim())
                                    {
                                        repgen.Reports("Given Condition Is Matched Starting Conditional Steps", "", "", true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                        PreformForigneOperations("Conditional Execute", OperationPath, assemblyName, numberofExecution);
                                        repgen.Reports("Executing Conditional Steps Completed", "", "", true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                    }
                                }
                                else if(operation == "ConditionNotEquals")
                                {
                                    if (languageResource.Msg_GetTemplateMessage.Replace(" ", string.Empty) != validationMessage.Trim())
                                    {
                                        repgen.Reports("Given Condition Is Matched Starting Conditional Steps", "", "", true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                        PreformForigneOperations("Conditional Execute", OperationPath, assemblyName, numberofExecution);
                                        repgen.Reports("Executing Conditional Steps Completed", "", "", true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                    }
                                }
                                break;
                            case "GetReferenceNumber":
                            case "GetSTONumber":
                                ScrollStatus = pop.ScriptGetExcuite(count["Keyword"].ToString(), count["ControlKeyword"].ToString(), count["ControlKeywordValue"].ToString(), count["Step No"].ToString(), count["Description"].ToString());
                                if (ScrollStatus != string.Empty)
                                {
                                    writeData.WriteExcel(count["DataRefferencekeyword"].ToString().Split(':')[1], ScrollStatus, "ReadData", ExcelDataTable.ReadData(1, count["DataRefferencekeyword"].ToString().Split(':')[0]).ToString());
                                }
                                repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                break;
                            case "GetDeliveryOrdReport":
                                ScrollStatus = log.GetDispatchReportAvailability(count["ControlKeywordValue"].ToString(), count["DataRefferencekeyword"].ToString().Split(':')[0],count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString());
                                if (ScrollStatus == "0")
                                {
                                    screenShotName = count["AssertionScreenshotName"].ToString();
                                    ErrorScreenPath = log.screenShot(screenShotName);
                                    FailerReason = count["AssertionMessage"].ToString();
                                    repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                    repgen.FileCorreptionCheck();
                                    Assert.Fail(FailerReason);
                                }
                                else
                                {
                                    writeData.WriteExcel("Dispatch Report Order",ScrollStatus,"ReadData", ExcelDataTable.ReadData(1,count["DataRefferencekeyword"].ToString().Split(':')[1]));
                                    RecordData("DispatchReportOrder", ScrollStatus);
                                    repgen.Reports(count["Step No"].ToString(), "Record the order count for the specified order", "RecordData", true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, recordedData["DispatchReportOrder"]);
                                }
                                repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                break;
                            case "DateTimeReportValidate":
                                status = log.getdispatchDateTimeValidation(count["ControlKeywordValue"].ToString(), count["DataRefferencekeyword"].ToString(), count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString());
                                if (!status)
                                {
                                    screenShotName = count["AssertionScreenshotName"].ToString();
                                    ErrorScreenPath = log.screenShot(screenShotName);
                                    FailerReason = count["AssertionMessage"].ToString();
                                    repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                    repgen.FileCorreptionCheck();
                                    Assert.Fail(FailerReason);
                                }
                                repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                break;
                            case "B2BSAPValidation":
                                status = log.B2BDropValidation(count["ControlKeywordValue"].ToString(), count["DataRefferencekeyword"].ToString(), count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString());
                                if (!status)
                                {
                                    screenShotName = count["AssertionScreenshotName"].ToString();
                                    ErrorScreenPath = log.screenShot(screenShotName);
                                    FailerReason = count["AssertionMessage"].ToString();
                                    repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                    repgen.FileCorreptionCheck();
                                    Assert.Fail(FailerReason);
                                }
                                repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                break;
                            case "GetSystemDate":
                                //Get's the current System Date.
                                string date = log.getSystemDate(ExcelDataTable.ReadData(1, "DateFormat").ToString());
                                if (date != string.Empty)
                                {
                                    pop.ScriptExcuite("SetValue", count["ControlKeyword"].ToString(), count["ControlKeywordValue"].ToString(), date, count["Step No"].ToString(), count["Description"].ToString());
                                    repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                }
                                else
                                {
                                    repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                    repgen.FileCorreptionCheck();
                                    Assert.Fail(FailerReason);
                                }
                                break;
                            case "EnterRandomNumber":
                                //"Reference Data in Dynamic Excel"="Number or Digit":"Column Head in Dynamic Excel"
                                materialData = count["DataRefferencekeyword"].ToString().Split('=');
                                date = log.RamdomBatch(Convert.ToInt16(materialData[1].Split(':')[0]));
                                if (date != string.Empty)
                                {
                                    writeData.WriteExcel(count["DataRefferencekeyword"].ToString().Split(':')[1], date, "ReadData", materialData[0]);
                                    pop.OperationStart(count["Keyword"].ToString(), count["TypeOfControl"].ToString(), count["ControlKeyword"].ToString(), count["ControlKeywordValue"].ToString(), date, count["TypeOfWindow"].ToString(), count["Step No"].ToString(), count["Description"].ToString());
                                    repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, date);
                                }
                                break;
                            case "SetAttribute":
                                // set specified attribute for a control.
                                pop.ScriptSetAttribute(count["Keyword"].ToString(), count["ControlKeyword"].ToString(), count["ControlKeywordValue"].ToString(), count["DataRefferencekeyword"].ToString(), count["Step No"].ToString(), count["Description"].ToString());
                                repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                break;
                            case "GetDate":
                                date = log.getDate(ExcelDataTable.ReadData(1, "DateFormat").ToString(), count["DataRefferencekeyword"].ToString());
                                if (date != string.Empty)
                                {
                                    pop.ScriptExcuite("SetValue", count["ControlKeyword"].ToString(), count["ControlKeywordValue"].ToString(), date, count["Step No"].ToString(), count["Description"].ToString());
                                    repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                }
                                else
                                {
                                    repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                    repgen.FileCorreptionCheck();
                                    Assert.Fail("Number of days you want to ADD or SUBTRACT is not proper");
                                }
                                break;
                            case "SetDefault":
                                date = log.GetDefaultData(count["ControlKeywordValue"].ToString(), count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString());
                                writeData.WriteExcel(count["DataRefferencekeyword"].ToString().Split(':')[1], date, "ReadData", count["DataRefferencekeyword"].ToString().Split(':')[0]);
                                repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, date);
                                break;
                            case "PartBagValidation":
                                CommonLanguageTemplateReader.Message(lang, count["DataRefferencekeyword"].ToString().Split('=')[0]);
                                int PartBagcount = log.getPartBagCount(count["ControlKeywordValue"].ToString(), languageResource.Msg_GetTemplateMessage, count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString());
                                int numberofpartbag = Convert.ToInt16(ExcelDataTable.ReadData(1, count["DataRefferencekeyword"].ToString().Split('=')[1]));
                                if(PartBagcount!= numberofpartbag)
                                {
                                    screenShotName = count["AssertionScreenshotName"].ToString();
                                    ErrorScreenPath = log.screenShot(screenShotName);
                                    FailerReason = count["AssertionMessage"].ToString();
                                    repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                    repgen.FileCorreptionCheck();
                                    Assert.Fail(FailerReason);
                                }
                                repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                break;
                            case "ValidatePOAvailable":
                                log.POAvailabilityCalender(count["ControlKeyword"].ToString(), count["Keyword"].ToString(), count["ControlKeywordValue"].ToString(), count["Step No"].ToString(), count["Description"].ToString());
                                break;
                            default:
                                FailerReason = "No " + operation + " keyword found. Please check the Keyword.";
                                throw new NoSuchOperationFound("No " + operation + " keyword found. Please check the Keyword.");
                        }
                    }
                }
            }
            if (totalExecution == 1)
            {
                repgen.FileCorreptionCheck();
            }
            dt.Clear();
        }
        /// <summary>
        /// This will invoke by the "Operations" keyword. if we are referring automated test with in another test
        /// </summary>
        /// <param name="ControlKeywordValue">This contains the sheet name of the referenced excel.</param>
        /// <param name="DataRefferencekeyword">This will accept the path of the excel.</param>
        /// <param name="ProjectName">name of the project.</param>
        public void PreformForigneOperations(string ControlKeywordValue, string DataRefferencekeyword, string ProjectName, int iteration)
        {
            DataTable OperationTable = new DataTable();
            OperationTable.Clear();
            OperationTable = ExcelKeywordTable.ExcelData(DataRefferencekeyword, ControlKeywordValue);
            ForeigntabName = OperationTable.TableName;
            int lengthCount = DataRefferencekeyword.LastIndexOf("\\") + 1;
            //get excel sheet name
            ForeignSheetName = DataRefferencekeyword.Substring(lengthCount, DataRefferencekeyword.Length - lengthCount);
            foreach (DataRow count in OperationTable.Rows)
            {
                //get current step number
                ForeignStepNumber = count["Step No"].ToString();
                if (count.IsNull("Step No"))
                {
                    repgen.UpdateData("No Step Found", count["Description"].ToString(), count["Keyword"].ToString(), "Skipped", LoginOperatrion.batchforReport, ReportGeneration.DetailedReportFilePath, "");
                    continue;
                }
                else
                {
                    string operation = count["Keyword"].ToString();
                    switch (operation)
                    {
                        case "Initialize":
                            log.Initialize(ProjectName);
                            auto.Send("{F5}");
                            Thread.Sleep(max);
                            break;
                        case "Navigate2CDA":
                        case "Navigate2CRA":
                        case "Navigate2CPA":
                            log.SubModuleNavigationPack(operation);
                            break;
                        case "SyncPO":
                        case "SyncDO":
                            log.SyncOrder(operation);
                            break;
                        case "Operations":
                            PreformForigneOperations(count["ControlKeywordValue"].ToString(), count["DataRefferencekeyword"].ToString(), ProjectName, numberofExecution);
                            break;
                        case "Logout":
                            log.Logout();
                            break;
                        case "Login":
                            log.login(operation);
                            break;
                        case "PlantLogin":
                            log.PlantLogin(operation, iteration);
                            break;
                        case "RegionalLogin":
                            log.login(operation);
                            break;
                        case "SuperAdminLogin":
                            log.SuperLogin();
                            break;
                        case "PlantRegionalLogin":
                            log.PlantLogin(operation, iteration);
                            break;
                        case "PageNavigation":
                            log.NavigationPage(count["ControlKeywordValue"].ToString(), count["Step No"].ToString(), count["Description"].ToString());
                            break;
                        case "ChangLanguage":
                            log.ChangLanguage(lang);
                            break;
                        case "ModuleNavigation":
                            log.ModuleNavigation(count["ControlKeywordValue"].ToString());
                            break;
                        case "SubModuleNavigation":
                            log.SubModuleNavigation(count["ControlKeywordValue"].ToString());
                            break;
                        case "LinkNavigation":
                            log.LinkNavigation(count["ControlKeywordValue"].ToString());
                            break;
                        case "EnterText":
                            string[] Testdata = count["DataRefferencekeyword"].ToString().Split(':');
                            if (Testdata.Count() > 1)
                            {
                                if (Testdata[0] != "BatchData")
                                {
                                    pop.OperationStart(count["Keyword"].ToString(), count["TypeOfControl"].ToString(), count["ControlKeyword"].ToString(), count["ControlKeywordValue"].ToString(), ExcelDataTable.ReadData(Convert.ToInt16(Testdata[1]), Testdata[0]), count["TypeOfWindow"].ToString(), count["Step No"].ToString(), count["Description"].ToString());
                                }
                                else
                                {
                                    pop.OperationStart(count["Keyword"].ToString(), count["TypeOfControl"].ToString(), count["ControlKeyword"].ToString(), count["ControlKeywordValue"].ToString(), ExcelDataTable.ReadBatchData(BatchName, Testdata[1]), count["TypeOfWindow"].ToString(), count["Step No"].ToString(), count["Description"].ToString());
                                }
                            }
                            else
                            {
                                pop.OperationStart(count["Keyword"].ToString(), count["TypeOfControl"].ToString(), count["ControlKeyword"].ToString(), count["ControlKeywordValue"].ToString(), ExcelDataTable.ReadData(1, Testdata[0]), count["TypeOfWindow"].ToString(), count["Step No"].ToString(), count["Description"].ToString());
                            }
                            break;
                        case "EnterLanguageText":
                            CommonLanguageTemplateReader.Message(lang, count["DataRefferencekeyword"].ToString());
                            pop.OperationStart(count["Keyword"].ToString(), count["TypeOfControl"].ToString(), count["ControlKeyword"].ToString(), count["ControlKeywordValue"].ToString(), languageResource.Msg_GetTemplateMessage, count["TypeOfWindow"].ToString(), count["Step No"].ToString(), count["Description"].ToString());
                            break;
                        case "Click":
                        case "RightClick":
                            pop.OperationStart(count["Keyword"].ToString(), count["TypeOfControl"].ToString(), count["ControlKeyword"].ToString(), count["ControlKeywordValue"].ToString(), count["DataRefferencekeyword"].ToString(), count["TypeOfWindow"].ToString(), count["Step No"].ToString(), count["Description"].ToString());
                            break;
                        case "ScriptClick":
                            pop.ScriptClick(count["Keyword"].ToString(), count["ControlKeyword"].ToString(), count["ControlKeywordValue"].ToString(), count["Step No"].ToString(), count["Description"].ToString());
                            repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                            break;
                        case "ValidateControlAvailable":
                        case "ValidateControlNotAvailable":
                            if (count["ControlKeywordValue"].ToString().Contains('='))
                            {
                                WriteAndReadData.DataFilePath = ConfigurationManager.AppSettings["ExcelDataFile"] + "\\DynamicData" + "\\" + BatchName + ".xls";
                                ExcelDataTable.PopulateRecordData(WriteAndReadData.DataFilePath);
                            }
                            bool controlavailability = pop.ControlAvailability(count["TypeOfControl"].ToString(), count["ControlKeyword"].ToString(), count["ControlKeywordValue"].ToString(), count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString());
                            if (operation == "ValidateControlAvailable")
                            {
                                if (!controlavailability)
                                {
                                    screenShotName = count["AssertionScreenshotName"].ToString();
                                    ErrorScreenPath = log.screenShot(screenShotName);
                                    FailerReason = count["AssertionMessage"].ToString();
                                    repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                    repgen.FileCorreptionCheck();
                                    Assert.Fail(FailerReason);
                                }
                            }
                            else
                            {
                                if (controlavailability)
                                {
                                    screenShotName = count["AssertionScreenshotName"].ToString();
                                    ErrorScreenPath = log.screenShot(screenShotName);
                                    FailerReason = count["AssertionMessage"].ToString();
                                    repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                    repgen.FileCorreptionCheck();
                                    Assert.Fail(FailerReason);
                                }
                            }
                            repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                            break;
                        case "ValidateLog":
                            bool status = log.validateLOGcheckFail(count["DataRefferencekeyword"].ToString());
                            if (!status)
                            {
                                screenShotName = count["AssertionScreenshotName"].ToString();
                                ErrorScreenPath = log.screenShot(screenShotName);
                                FailerReason = count["AssertionMessage"].ToString();
                                log.CloseFile();
                                repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                repgen.FileCorreptionCheck();
                                Assert.Fail(FailerReason);
                            }
                            log.CloseFile();
                            repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                            break;
                        case "Validate":
                            string validationMessage = pop.WebGetControlData(count["TypeOfControl"].ToString(), count["ControlKeyword"].ToString(), count["ControlKeywordValue"].ToString(), count["AssertionScreenshotName"].ToString(), count["AssertionMessage"].ToString(), count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString()).Trim();
                            if (count["DataRefferencekeyword"].ToString().Contains('='))
                            {
                                WriteAndReadData.DataFilePath = ConfigurationManager.AppSettings["ExcelDataFile"] + "\\DynamicData" + "\\" + BatchName + ".xls";
                                ExcelDataTable.PopulateRecordData(WriteAndReadData.DataFilePath);
                            }
                            if (validationMessage != string.Empty)
                            {
                                validationMessage = Regex.Replace(validationMessage, "<.*?>", String.Empty).Trim();
                                Message = System.Web.HttpUtility.HtmlDecode(validationMessage);
                                validationMessage = Message.Replace(" ", string.Empty).Trim();
                                CommonLanguageTemplateReader.Message(lang, count["DataRefferencekeyword"].ToString());
                                if (operation == "Validate")
                                {
                                    if (languageResource.Msg_GetTemplateMessage.Replace(" ", string.Empty) != validationMessage.Trim())
                                    {
                                        screenShotName = count["AssertionScreenshotName"].ToString();
                                        ErrorScreenPath = log.screenShot(screenShotName);
                                        FailerReason = count["AssertionMessage"].ToString();
                                        repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                        repgen.FileCorreptionCheck();
                                    }
                                    Assert.AreEqual(languageResource.Msg_GetTemplateMessage.Replace(" ", string.Empty), validationMessage, FailerReason);
                                    repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                }
                               
                            }
                            break;
                        case "Validate&Continue":
                            validationMessage = pop.WebGetControlData(count["TypeOfControl"].ToString(), count["ControlKeyword"].ToString(), count["ControlKeywordValue"].ToString(), count["AssertionScreenshotName"].ToString(), count["AssertionMessage"].ToString(), count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString()).Trim();
                            if (count["DataRefferencekeyword"].ToString().Contains('='))
                            {
                                WriteAndReadData.DataFilePath = ConfigurationManager.AppSettings["ExcelDataFile"] + "\\DynamicData" + "\\" + BatchName + ".xls";
                                ExcelDataTable.PopulateRecordData(WriteAndReadData.DataFilePath);
                            }
                            if (validationMessage != string.Empty)
                            {
                                validationMessage = Regex.Replace(validationMessage, "<.*?>", String.Empty).Trim();
                                Message = System.Web.HttpUtility.HtmlDecode(validationMessage);
                                validationMessage = Message.Replace(" ", string.Empty).Trim();
                                CommonLanguageTemplateReader.Message(lang, count["DataRefferencekeyword"].ToString());
                               if (languageResource.Msg_GetTemplateMessage.Replace(" ", string.Empty) != validationMessage.Trim())
                               {
                                        warningScreen = count["AssertionScreenshotName"].ToString();
                                        warningScreenPath = log.screenShot(warningScreen);
                                        warningMessage = count["AssertionMessage"].ToString();
                                        repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                               }
                             }
                            break;
                        case "ValidateOR":
                            bool validationstatus = false;
                            if (count["DataRefferencekeyword"].ToString().Contains('='))
                            {
                                WriteAndReadData.DataFilePath = ConfigurationManager.AppSettings["ExcelDataFile"] + "\\DynamicData" + "\\" + BatchName + ".xls";
                                ExcelDataTable.PopulateRecordData(WriteAndReadData.DataFilePath);
                            }
                            CommonLanguageTemplateReader.Message(lang, count["DataRefferencekeyword"].ToString());
                            string[] availableMsg = languageResource.Msg_GetTemplateMessage.ToString().Split('|');
                            string[] getMessage = count["ControlKeywordValue"].ToString().Split('|');
                            for (int i = 0; i < getMessage.Count(); i++)
                            {
                                validationMessage = pop.WebGetControlData(count["TypeOfControl"].ToString(), count["ControlKeyword"].ToString(), getMessage[i], count["AssertionScreenshotName"].ToString(), count["AssertionMessage"].ToString(), count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString()).Trim();
                                validationMessage = Regex.Replace(validationMessage, "<.*?>", String.Empty).Trim();
                                Message = System.Web.HttpUtility.HtmlDecode(validationMessage);
                                validationMessage = Message.Replace(" ", string.Empty).Trim();
                                if (availableMsg[i].Replace(" ", string.Empty) == validationMessage)
                                {
                                    validationstatus = true;
                                    break;
                                }
                                else
                                {
                                    validationstatus = false;
                                }
                            }
                            if (!validationstatus)
                            {
                                screenShotName = count["AssertionScreenshotName"].ToString();
                                ErrorScreenPath = log.screenShot(screenShotName);
                                FailerReason = count["AssertionMessage"].ToString();
                                repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                repgen.FileCorreptionCheck();
                                Assert.Fail(FailerReason);
                            }
                            repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                            break;
                        case "ValidateNotEqual":
                            string validationMessagenotEqual = pop.WebGetControlData(count["TypeOfControl"].ToString(), count["ControlKeyword"].ToString(), count["ControlKeywordValue"].ToString(), count["AssertionScreenshotName"].ToString(), count["AssertionMessage"].ToString(), count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString()).Trim();
                            validationMessage = Regex.Replace(validationMessagenotEqual, "<.*?>", String.Empty).Trim();
                            Message = System.Web.HttpUtility.HtmlDecode(validationMessage);
                            validationMessage = Message.Replace(" ", string.Empty).Trim();
                            CommonLanguageTemplateReader.Message(lang, count["DataRefferencekeyword"].ToString());
                            if (languageResource.Msg_GetTemplateMessage.Replace(" ", string.Empty) == validationMessage)
                            {
                                screenShotName = count["AssertionScreenshotName"].ToString();
                                ErrorScreenPath = log.screenShot(screenShotName);
                                FailerReason = count["AssertionMessage"].ToString();
                                repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                repgen.FileCorreptionCheck();
                            }
                            Assert.AreNotEqual(languageResource.Msg_GetTemplateMessage.Replace(" ", string.Empty), validationMessage, FailerReason);
                            repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                            break;
                        case "ClickDelete":
                        case "OprRightClick":
                            Thread.Sleep(max * 2);
                            if (count["ControlKeywordValue"].ToString().Contains('='))
                            {
                                WriteAndReadData.DataFilePath = ConfigurationManager.AppSettings["ExcelDataFile"] + "\\DynamicData" + "\\" + BatchName + ".xls";
                                ExcelDataTable.PopulateRecordData(WriteAndReadData.DataFilePath);
                                validationMessage = pop.getDataFromDynamicExcel(count["ControlKeywordValue"].ToString());
                                pop.OperationStart(count["Keyword"].ToString(), count["TypeOfControl"].ToString(), count["ControlKeyword"].ToString(), validationMessage, count["DataRefferencekeyword"].ToString(), count["TypeOfWindow"].ToString(), count["Step No"].ToString(), count["Description"].ToString(), new int[] { Convert.ToInt16(count["XCodinate"]), Convert.ToInt16(count["Ycodinate"]) });
                            }
                            else
                            {
                                pop.OperationStart(count["Keyword"].ToString(), count["TypeOfControl"].ToString(), count["ControlKeyword"].ToString(), ExcelDataTable.ReadData(1, count["ControlKeywordValue"].ToString()), count["DataRefferencekeyword"].ToString(), count["TypeOfWindow"].ToString(), count["Step No"].ToString(), count["Description"].ToString(), new int[] { Convert.ToInt16(count["XCodinate"]), Convert.ToInt16(count["Ycodinate"]) });
                            }
                            break;
                        case "ClickEdit":
                        case "DClickEdit":
                        case "DClickDelete":
                            Thread.Sleep(max * 2);
                            if (count["ControlKeywordValue"].ToString().Contains('='))
                            {
                                WriteAndReadData.DataFilePath = ConfigurationManager.AppSettings["ExcelDataFile"] + "\\DynamicData" + "\\" + BatchName + ".xls";
                                ExcelDataTable.PopulateRecordData(WriteAndReadData.DataFilePath);
                                validationMessage = pop.getDataFromDynamicExcel(count["ControlKeywordValue"].ToString());
                                pop.OperationStart(count["Keyword"].ToString(), count["TypeOfControl"].ToString(), count["ControlKeyword"].ToString(), validationMessage, count["DataRefferencekeyword"].ToString(), count["TypeOfWindow"].ToString(), count["Step No"].ToString(), count["Description"].ToString(), new int[] { Convert.ToInt16(count["XCodinate"]), Convert.ToInt16(count["Ycodinate"]) });
                            }
                            else
                            {
                                pop.OperationStart(count["Keyword"].ToString(), count["TypeOfControl"].ToString(), count["ControlKeyword"].ToString(), ExcelDataTable.ReadData(1, count["ControlKeywordValue"].ToString()), count["DataRefferencekeyword"].ToString(), count["TypeOfWindow"].ToString(), count["Step No"].ToString(), count["Description"].ToString(), new int[] { Convert.ToInt16(count["XCodinate"]), Convert.ToInt16(count["Ycodinate"]) });
                            }
                            break;
                        case "MouseMove":
                        case "MouseMoveRightClick":
                            Thread.Sleep(max * 2);
                            string[] data = count["ControlKeywordValue"].ToString().Split(':');
                            if (data.Count() <= 1)
                            {
                                pop.OperationStart(count["Keyword"].ToString(), count["TypeOfControl"].ToString(), count["ControlKeyword"].ToString(), count["ControlKeywordValue"].ToString(), count["DataRefferencekeyword"].ToString(), count["TypeOfWindow"].ToString(), count["Step No"].ToString(), count["Description"].ToString(), new int[] { Convert.ToInt16(count["XCodinate"]), Convert.ToInt16(count["Ycodinate"]) });
                            }
                            else
                            {
                                CommonLanguageTemplateReader.Message(lang, count["ControlKeywordValue"].ToString());
                                pop.OperationStart(count["Keyword"].ToString(), count["TypeOfControl"].ToString(), count["ControlKeyword"].ToString(), languageResource.Msg_GetTemplateMessage, count["DataRefferencekeyword"].ToString(), count["TypeOfWindow"].ToString(), count["Step No"].ToString(), count["Description"].ToString(), new int[] { Convert.ToInt16(count["XCodinate"]), Convert.ToInt16(count["Ycodinate"]) });
                            }
                            break;
                        case "Clear":
                            pop.OperationStart(count["Keyword"].ToString(), count["TypeOfControl"].ToString(), count["ControlKeyword"].ToString(), count["ControlKeywordValue"].ToString(), count["DataRefferencekeyword"].ToString(), count["TypeOfWindow"].ToString(), count["Step No"].ToString(), count["Description"].ToString());
                            break;
                        case "Wait":
                        case "WaitPropertySet":
                            pop.OperationStart(count["Keyword"].ToString(), count["TypeOfControl"].ToString(), count["ControlKeyword"].ToString(), count["ControlKeywordValue"].ToString(), count["DataRefferencekeyword"].ToString(), count["TypeOfWindow"].ToString(), count["Step No"].ToString(), count["Description"].ToString());
                            break;
                        case "Sleep":
                            pop.GotoSleep(count["DataRefferencekeyword"].ToString());
                            repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                            break;
                        case "ClickLanguageText":
                            CommonLanguageTemplateReader.Message(lang, count["ControlKeywordValue"].ToString());
                            pop.OperationStart(count["Keyword"].ToString(), count["TypeOfControl"].ToString(), count["ControlKeyword"].ToString(), languageResource.Msg_GetTemplateMessage, count["DataRefferencekeyword"].ToString(), count["TypeOfWindow"].ToString(), count["Step No"].ToString(), count["Description"].ToString());
                            repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                            break;
                        case "CloseFile":
                            log.CloseFile();
                            repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                            break;
                        case "SendKey":
                        case "SendText":
                            if (operation == "SendKey")
                            {
                                log.SendKey(count["ControlKeywordValue"].ToString());
                            }
                            else
                            {
                                if (count["DataRefferencekeyword"].ToString().Contains('='))
                                {
                                    WriteAndReadData.DataFilePath = ConfigurationManager.AppSettings["ExcelDataFile"] + "\\DynamicData" + "\\" + BatchName + ".xls";
                                    ExcelDataTable.PopulateRecordData(WriteAndReadData.DataFilePath);
                                    string[] refferenceArray = count["DataRefferencekeyword"].ToString().Split('=');
                                    log.Sendtext(pop.getDataFromDynamicExcel(count["DataRefferencekeyword"].ToString()));
                                }
                                else
                                {
                                    log.Sendtext(ExcelDataTable.ReadData(1, count["DataRefferencekeyword"].ToString()).ToString());
                                }
                            }
                            repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                            break;
                        case "OpenFile":
                            file.OpenFile(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), count["AssertionScreenshotName"].ToString(), count["AssertionMessage"].ToString());
                            break;
                        case "SaveOpenFile":
                            file.SaveOpenFile(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), count["AssertionScreenshotName"].ToString(), count["AssertionMessage"].ToString());
                            break;
                        case "CancelDownload":
                            file.CancelDownload(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString());
                            break;
                        case "SetValue":
                            if (!count["DataRefferencekeyword"].ToString().Contains('='))
                            {
                                Testdata = count["DataRefferencekeyword"].ToString().Split(':');
                                if (Testdata.Count() > 1)
                                {
                                    pop.ScriptExcuite(count["Keyword"].ToString(), count["ControlKeyword"].ToString(), count["ControlKeywordValue"].ToString(), ExcelDataTable.ReadData(Convert.ToInt16(Testdata[1]), Testdata[0]), count["Step No"].ToString(), count["Description"].ToString());
                                }
                                else
                                {
                                    pop.ScriptExcuite(count["Keyword"].ToString(), count["ControlKeyword"].ToString(), count["ControlKeywordValue"].ToString(), ExcelDataTable.ReadData(1, Testdata[0]), count["Step No"].ToString(), count["Description"].ToString());
                                }
                            }
                            else
                            {
                                WriteAndReadData.DataFilePath = ConfigurationManager.AppSettings["ExcelDataFile"] + "\\DynamicData" + "\\" + BatchName + ".xls";
                                ExcelDataTable.PopulateRecordData(WriteAndReadData.DataFilePath);
                                string recData= pop.getDataFromDynamicExcel(count["DataRefferencekeyword"].ToString());
                                pop.ScriptExcuite(count["Keyword"].ToString(), count["ControlKeyword"].ToString(), count["ControlKeywordValue"].ToString(), recData, count["Step No"].ToString(), count["Description"].ToString());
                            }
                            repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                            break;
                        case "GetValueMatch":
                        case "GetValueNoMatch":
                        case "GetValueContains":
                            string valueGEt = pop.ScriptGetExcuite(count["Keyword"].ToString(), count["ControlKeyword"].ToString(), count["ControlKeywordValue"].ToString(), count["Step No"].ToString(), count["Description"].ToString());
                            string valueMatch = string.Empty;
                            string[] RefResources = { };
                            if (count["DataRefferencekeyword"].ToString().Contains('='))
                            {
                                WriteAndReadData.DataFilePath = ConfigurationManager.AppSettings["ExcelDataFile"] + "\\DynamicData" + "\\" + BatchName + ".xls";
                                ExcelDataTable.PopulateRecordData(WriteAndReadData.DataFilePath);
                                string referenceData = count["DataRefferencekeyword"].ToString();
                                if (referenceData.Contains('+') || referenceData.Contains('-') || referenceData.Contains('*'))
                                {
                                    if (referenceData.Contains('+'))
                                    {
                                        RefResources = referenceData.Split('+');
                                        valueMatch = pop.getDataFromDynamicExcel(RefResources[0]);
                                        if (RefResources[1].ToString().Contains('='))
                                        {
                                            valueMatch = pop.calculatedata(valueMatch + '+' + pop.getDataFromDynamicExcel(RefResources[1])).ToString();
                                        }
                                        else
                                        {
                                            valueMatch = pop.calculatedata(valueMatch + '+' + ExcelDataTable.ReadData(1, RefResources[1])).ToString();
                                        }
                                    }
                                    else if (referenceData.Contains('-'))
                                    {
                                        RefResources = referenceData.Split('-');
                                        valueMatch = pop.getDataFromDynamicExcel(RefResources[0]);
                                        if (RefResources[1].ToString().Contains('='))
                                        {
                                            valueMatch = pop.calculatedata(valueMatch + '-' + pop.getDataFromDynamicExcel(RefResources[1])).ToString();
                                        }
                                        else
                                        {
                                            valueMatch = pop.calculatedata(valueMatch + '-' + ExcelDataTable.ReadData(1, RefResources[1])).ToString();
                                        }
                                    }
                                    else if (referenceData.Contains('*'))
                                    {
                                        RefResources = referenceData.Split('*');
                                        valueMatch = pop.getDataFromDynamicExcel(RefResources[0]);
                                        if (RefResources[1].ToString().Contains('='))
                                        {
                                            valueMatch = pop.calculatedata(valueMatch + '*' + pop.getDataFromDynamicExcel(RefResources[1])).ToString();
                                        }
                                        else
                                        {
                                            valueMatch = pop.calculatedata(valueMatch + '*' + ExcelDataTable.ReadData(1, RefResources[1])).ToString();
                                        }
                                    }
                                }
                                else
                                {
                                    valueMatch = pop.getDataFromDynamicExcel(count["DataRefferencekeyword"].ToString());
                                }
                            }
                            else
                            {
                                if (count["DataRefferencekeyword"].ToString().Contains('+') || count["DataRefferencekeyword"].ToString().Contains('-') || count["DataRefferencekeyword"].ToString().Contains('*'))
                                {
                                    if (count["DataRefferencekeyword"].ToString().Contains('+'))
                                    {
                                        RefResources = count["DataRefferencekeyword"].ToString().Split('+');
                                        valueMatch = pop.calculatedata(ExcelDataTable.ReadData(1, RefResources[0]) + '+' + ExcelDataTable.ReadData(1, RefResources[1])).ToString();
                                    }
                                    else if (count["DataRefferencekeyword"].ToString().Contains('-'))
                                    {
                                        RefResources = count["DataRefferencekeyword"].ToString().Split('-');
                                        valueMatch = pop.calculatedata(ExcelDataTable.ReadData(1, RefResources[0]) + '-' + ExcelDataTable.ReadData(1, RefResources[1])).ToString();
                                    }
                                    else if (count["DataRefferencekeyword"].ToString().Contains('*'))
                                    {
                                        RefResources = count["DataRefferencekeyword"].ToString().Split('*');
                                        valueMatch = pop.calculatedata(ExcelDataTable.ReadData(1, RefResources[0]) + '*' + ExcelDataTable.ReadData(1, RefResources[1])).ToString();
                                    }
                                    else
                                    {
                                        throw new Exception("No Calculation Symbol Not Found..Calculation is not possible");
                                    }
                                }
                                else
                                {
                                    Testdata = count["DataRefferencekeyword"].ToString().Split(':');
                                    if (Testdata.Count() > 1)
                                    {
                                        int number;
                                        if (Int32.TryParse(Testdata[1], out number))
                                        {
                                            valueMatch = ExcelDataTable.ReadData(Convert.ToInt16(Testdata[1]), Testdata[0]);
                                        }
                                        else
                                        {
                                            CommonLanguageTemplateReader.Message(lang, count["DataRefferencekeyword"].ToString());
                                            valueMatch = languageResource.Msg_GetTemplateMessage;
                                        }
                                    }
                                    else
                                    {
                                        valueMatch = ExcelDataTable.ReadData(1, Testdata[0]);
                                    }
                                }
                            }
                            if (operation == "GetValueMatch")
                            {
                                if (valueGEt != valueMatch)
                                {
                                    screenShotName = count["AssertionScreenshotName"].ToString();
                                    ErrorScreenPath = log.screenShot(screenShotName);
                                    FailerReason = count["AssertionMessage"].ToString();
                                    repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                    repgen.FileCorreptionCheck();
                                }
                                Assert.AreEqual(valueMatch, valueGEt, FailerReason);
                            }
                            else if (operation == "GetValueNoMatch")
                            {
                                if (valueGEt == valueMatch)
                                {
                                    screenShotName = count["AssertionScreenshotName"].ToString();
                                    ErrorScreenPath = log.screenShot(screenShotName);
                                    FailerReason = count["AssertionMessage"].ToString();
                                    repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                    repgen.FileCorreptionCheck();
                                }
                                Assert.AreNotEqual(valueMatch, valueGEt, FailerReason);
                            }
                            else if (operation == "GetValueContains")
                            {
                                if (count["DataRefferencekeyword"].ToString().Contains("Container Number"))
                                {
                                    if (count["DataRefferencekeyword"].ToString().Split(':')[1] == "Container Number")
                                    {
                                        valueMatch = valueMatch.Substring(2, valueMatch.Length - 2);
                                    }
                                }
                                if (!valueGEt.Contains(valueMatch))
                                {
                                    screenShotName = count["AssertionScreenshotName"].ToString();
                                    ErrorScreenPath = log.screenShot(screenShotName);
                                    FailerReason = count["AssertionMessage"].ToString();
                                    repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                    repgen.FileCorreptionCheck();
                                }
                                StringAssert.Contains(valueGEt, valueMatch, FailerReason);
                            }
                            repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                            break;
                        case "ValidatePOStatus":
                            string POStatus = log.ValidateCalenderStatus(count["ControlKeywordValue"].ToString(), count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString());
                            CommonLanguageTemplateReader.Message(lang, count["DataRefferencekeyword"].ToString());
                            if (languageResource.Msg_GetTemplateMessage.Replace(" ", string.Empty) != POStatus.Trim())
                            {
                                screenShotName = count["AssertionScreenshotName"].ToString();
                                ErrorScreenPath = log.screenShot(screenShotName);
                                FailerReason = count["AssertionMessage"].ToString();
                                repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                repgen.FileCorreptionCheck();
                            }
                            Assert.AreEqual(languageResource.Msg_GetTemplateMessage.Replace(" ", string.Empty), POStatus, FailerReason);
                            repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                            break;
                        case "ValidateEmptyValueEqual":
                        case "ValidateEmptyValueNotEqual":
                            string EmptyvalueGEt = pop.ScriptGetExcuite(count["Keyword"].ToString(), count["ControlKeyword"].ToString(), count["ControlKeywordValue"].ToString(), count["Step No"].ToString(), count["Description"].ToString());
                            validationMessage = Regex.Replace(EmptyvalueGEt, "<.*?>", String.Empty).Trim();
                            Message = System.Web.HttpUtility.HtmlDecode(validationMessage);
                            validationMessage = Message.Replace(" ", string.Empty).Trim();
                            if (operation == "ValidateEmptyValueEqual")
                            {
                                if (validationMessage != "")
                                {
                                    screenShotName = count["AssertionScreenshotName"].ToString();
                                    ErrorScreenPath = log.screenShot(screenShotName);
                                    FailerReason = count["AssertionMessage"].ToString();
                                    repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                    repgen.FileCorreptionCheck();
                                }
                                Assert.AreEqual("", validationMessage, FailerReason);
                            }
                            else if (operation == "ValidateEmptyValueNotEqual")
                            {
                                if (validationMessage == "")
                                {
                                    screenShotName = count["AssertionScreenshotName"].ToString();
                                    ErrorScreenPath = log.screenShot(screenShotName);
                                    FailerReason = count["AssertionMessage"].ToString();
                                    repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                    repgen.FileCorreptionCheck();
                                }
                                Assert.AreNotEqual("", validationMessage, FailerReason);
                            }
                            repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                            break;
                        case "ValidateGRStatus":
                            string GRStatus = log.ValidateGRStatus(ExcelDataTable.ReadData(1, count["ControlKeywordValue"].ToString()), count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString());
                            validationMessage = GRStatus.Replace(" ", string.Empty).Trim();
                            CommonLanguageTemplateReader.Message(lang, count["DataRefferencekeyword"].ToString());
                            if (languageResource.Msg_GetTemplateMessage.Replace(" ", string.Empty) != validationMessage.Trim())
                            {
                                screenShotName = count["AssertionScreenshotName"].ToString();
                                ErrorScreenPath = log.screenShot(screenShotName);
                                FailerReason = count["AssertionMessage"].ToString();
                                repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                repgen.FileCorreptionCheck();
                            }
                            Assert.AreEqual(languageResource.Msg_GetTemplateMessage.Replace(" ", string.Empty), validationMessage, FailerReason);
                            repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                            break;
                        case "RecordData":
                            RecordData(count["DataRefferencekeyword"].ToString(), pop.WebGetControlData(count["TypeOfControl"].ToString(), count["ControlKeyword"].ToString(), count["ControlKeywordValue"].ToString(), count["AssertionScreenshotName"].ToString(), count["AssertionMessage"].ToString(), count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString()).Trim());
                            repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, recordedData[count["DataRefferencekeyword"].ToString()]);
                            break;
                        case "GetRecordData":
                            pop.OperationStart("EnterText", count["TypeOfControl"].ToString(), count["ControlKeyword"].ToString(), count["ControlKeywordValue"].ToString(), recordedData[count["DataRefferencekeyword"].ToString()], count["TypeOfWindow"].ToString(), count["Step No"].ToString(), count["Description"].ToString());
                            repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, recordedData[count["DataRefferencekeyword"].ToString()]);
                            break;
                        case "assignPropsedGate":
                            CommonLanguageTemplateReader.Message(lang, count["DataRefferencekeyword"].ToString());
                            string proposedgate = log.getProposedGate(count["ControlKeywordValue"].ToString(), languageResource.Msg_GetTemplateMessage, count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString());
                            if (proposedgate == "")
                            {
                                screenShotName = count["AssertionScreenshotName"].ToString();
                                ErrorScreenPath = log.screenShot(screenShotName);
                                FailerReason = count["AssertionMessage"].ToString();
                                repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                repgen.FileCorreptionCheck();
                                Assert.Fail(FailerReason);
                            }
                            else
                            {
                                Thread.Sleep(max * 2);
                                pop.ScriptClick("ScriptClick", "ID", proposedgate, count["Step No"].ToString(), count["Description"].ToString());
                            }
                            repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                            break;
                        case "TruckDetailsValidate":
                            string truckAddstatus = string.Empty;
                            if (count["DataRefferencekeyword"].ToString().Contains(':'))
                            {
                                WriteAndReadData.DataFilePath = ConfigurationManager.AppSettings["ExcelDataFile"] + "\\DynamicData" + "\\" + BatchName + ".xls";
                                ExcelDataTable.PopulateRecordData(WriteAndReadData.DataFilePath);
                                truckAddstatus = log.ValidateTruckDetails(count["ControlKeywordValue"].ToString(), count["DataRefferencekeyword"].ToString(), count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString());
                            }
                            else if (count["DataRefferencekeyword"].ToString().StartsWith("Rec_"))
                            {
                                truckAddstatus = log.ValidateTruckDetails(count["ControlKeywordValue"].ToString(), recordedData[count["DataRefferencekeyword"].ToString()], count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString());
                            }
                            string[] truckStatus = truckAddstatus.Split(':');
                            if (truckStatus[1] == "false")
                            {
                                screenShotName = count["AssertionScreenshotName"].ToString();
                                ErrorScreenPath = log.screenShot(screenShotName);
                                FailerReason = count["AssertionMessage"].ToString();
                                repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                repgen.FileCorreptionCheck();
                                Assert.Fail(FailerReason);
                            }
                            writeData.WriteExcel("Gate Number", truckStatus[0], "ReadData", count["DataRefferencekeyword"].ToString().Split('=')[1].Split(':')[0]);
                            repgen.Reports(count["Step No"].ToString(), "Record Assigned Gate Number for the Truck", "ReadData", true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, truckStatus[0]);
                            repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                            break;
                        case "ScrollToControl":
                            string ScrollStatus = pop.ScriptGetExcuite(count["Keyword"].ToString(), count["ControlKeyword"].ToString(), count["ControlKeywordValue"].ToString(), count["Step No"].ToString(), count["Description"].ToString());
                            if (ScrollStatus == "true")
                            {
                                repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                            }
                            else
                            {
                                repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                            }
                            break;
                        case "IsReadOnly":
                        case "IsNotReadOnly":
                        case "IsChecked":
                            ScrollStatus = pop.ScriptGetExcuite(count["Keyword"].ToString(), count["ControlKeyword"].ToString(), count["ControlKeywordValue"].ToString(), count["Step No"].ToString(), count["Description"].ToString());
                            if (operation == "IsNotReadOnly")
                            {
                                if (ScrollStatus == "True")
                                {
                                    screenShotName = count["AssertionScreenshotName"].ToString();
                                    ErrorScreenPath = log.screenShot(screenShotName);
                                    FailerReason = count["AssertionMessage"].ToString();
                                    repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                    repgen.FileCorreptionCheck();
                                    Assert.Fail(FailerReason);
                                }
                            }
                            else
                            {
                                if (ScrollStatus == "False")
                                {
                                    screenShotName = count["AssertionScreenshotName"].ToString();
                                    ErrorScreenPath = log.screenShot(screenShotName);
                                    FailerReason = count["AssertionMessage"].ToString();
                                    repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                    repgen.FileCorreptionCheck();
                                    Assert.Fail(FailerReason);
                                }
                            }
                            repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                            break;
                        case "ValidateControlCount":
                            int ControlCount = Convert.ToInt16(pop.ScriptGetExcuite(count["Keyword"].ToString(), count["ControlKeyword"].ToString(), count["ControlKeywordValue"].ToString(), count["Step No"].ToString(), count["Description"].ToString()));
                            if (ControlCount <= 0)
                            {
                                screenShotName = count["AssertionScreenshotName"].ToString();
                                ErrorScreenPath = log.screenShot(screenShotName);
                                FailerReason = count["AssertionMessage"].ToString();
                                repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                repgen.FileCorreptionCheck();
                                Assert.Fail(FailerReason);
                            }
                            repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                            break;
                        case "SaveQualityResult":
                            log.EnterResultForQuality(count["ControlKeywordValue"].ToString(), ExcelDataTable.ReadData(1, count["DataRefferencekeyword"].ToString()), count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString());
                            repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                            break;
                        case "CheckQualityResult":
                            string qualityStatus = log.CheckQualityResult(count["ControlKeywordValue"].ToString(), count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString());
                            CommonLanguageTemplateReader.Message(lang, count["DataRefferencekeyword"].ToString());
                            if (languageResource.Msg_GetTemplateMessage.Replace(" ", string.Empty) != qualityStatus.Trim())
                            {
                                screenShotName = count["AssertionScreenshotName"].ToString();
                                ErrorScreenPath = log.screenShot(screenShotName);
                                FailerReason = count["AssertionMessage"].ToString();
                                repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                repgen.FileCorreptionCheck();
                            }
                            Assert.AreEqual(languageResource.Msg_GetTemplateMessage.Replace(" ", string.Empty), qualityStatus, FailerReason);
                            repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                            break;
                        case "SamplingDataCheck":
                            qualityStatus = log.ValidateSamplingInstruction(count["ControlKeywordValue"].ToString(), count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString());
                            if (qualityStatus == "false")
                            {
                                screenShotName = count["AssertionScreenshotName"].ToString();
                                ErrorScreenPath = log.screenShot(screenShotName);
                                FailerReason = count["AssertionMessage"].ToString();
                                repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                repgen.FileCorreptionCheck();
                                Assert.Fail(FailerReason);
                            }
                            repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                            break;
                        case "ValidateIsJIT":
                            qualityStatus = log.validateIsJit(count["ControlKeywordValue"].ToString(), count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString());
                            if (qualityStatus == "false")
                            {
                                screenShotName = count["AssertionScreenshotName"].ToString();
                                ErrorScreenPath = log.screenShot(screenShotName);
                                FailerReason = count["AssertionMessage"].ToString();
                                repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                repgen.FileCorreptionCheck();
                                Assert.Fail(FailerReason);
                            }
                            repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                            break;
                        case "ValidateAvailMetrial":
                            string materialStatus = string.Empty;
                            if (count["DataRefferencekeyword"].ToString().Contains('='))
                            {
                                WriteAndReadData.DataFilePath = ConfigurationManager.AppSettings["ExcelDataFile"] + "\\DynamicData" + "\\" + BatchName + ".xls";
                                ExcelDataTable.PopulateRecordData(WriteAndReadData.DataFilePath);
                                string value= pop.getDataFromDynamicExcel(count["DataRefferencekeyword"].ToString());
                                materialStatus = log.ValidateMetrialAvailable(count["ControlKeywordValue"].ToString(), value, count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString());
                            }
                            else
                            {
                                materialStatus = log.ValidateMetrialAvailable(count["ControlKeywordValue"].ToString(), ExcelDataTable.ReadData(1, count["DataRefferencekeyword"].ToString()), count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString());
                            }
                            if (materialStatus == "false")
                            {
                                screenShotName = count["AssertionScreenshotName"].ToString();
                                ErrorScreenPath = log.screenShot(screenShotName);
                                FailerReason = count["AssertionMessage"].ToString();
                                repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                repgen.FileCorreptionCheck();
                                Assert.Fail(FailerReason);
                            }
                            repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                            break;
                        case "WaitPageLoad":
                            status = pop.PageLoadWait(count["TypeOfControl"].ToString(), count["ControlKeyword"].ToString(), count["ControlKeywordValue"].ToString(), count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString());
                            if (!status)
                            {
                                repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                            }
                            repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                            break;
                        case "DownloadUpload":
                            log.ClickDownloadUpload(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString());
                            repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                            break;
                        case "Check":
                            pop.OperationStart(count["Keyword"].ToString(), count["TypeOfControl"].ToString(), count["ControlKeyword"].ToString(), count["ControlKeywordValue"].ToString(), ExcelDataTable.ReadData(1, count["DataRefferencekeyword"].ToString()), count["TypeOfWindow"].ToString(), count["Step No"].ToString(), count["Description"].ToString());
                            break;
                        case "ValidateCustomerInst":
                            qualityStatus = log.ValidateCustomerInst(count["ControlKeywordValue"].ToString(), count["DataRefferencekeyword"].ToString(), count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString());
                            if (qualityStatus == "false")
                            {
                                screenShotName = count["AssertionScreenshotName"].ToString();
                                ErrorScreenPath = log.screenShot(screenShotName);
                                FailerReason = count["AssertionMessage"].ToString();
                                repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                repgen.FileCorreptionCheck();
                                Assert.Fail(FailerReason);
                            }
                            repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                            break;
                        case "SelectPPE":
                            qualityStatus = log.SelectPPEInfo(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), count["DataRefferencekeyword"].ToString());
                            if (qualityStatus == "false")
                            {
                                screenShotName = count["AssertionScreenshotName"].ToString();
                                ErrorScreenPath = log.screenShot(screenShotName);
                                FailerReason = count["AssertionMessage"].ToString();
                                repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                repgen.FileCorreptionCheck();
                                Assert.Fail(FailerReason);
                            }
                            repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                            break;
                        case "SelectDGClass":
                            qualityStatus = log.SelectDGInfo(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), count["DataRefferencekeyword"].ToString());
                            if (qualityStatus == "False")
                            {
                                screenShotName = count["AssertionScreenshotName"].ToString();
                                ErrorScreenPath = log.screenShot(screenShotName);
                                FailerReason = count["AssertionMessage"].ToString();
                                repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                repgen.FileCorreptionCheck();
                                Assert.Fail(FailerReason);
                            }
                            repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                            break;
                        case "ValidateDGMetrilalStatus":
                            string MetrialDGtatus = log.ValidateDG_PPE(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), count["DataRefferencekeyword"].ToString());
                            string[] DGStatus = MetrialDGtatus.Split(':');
                            if (DGStatus[0] == "false")
                            {
                                screenShotName = count["AssertionScreenshotName"].ToString();
                                ErrorScreenPath = log.screenShot(screenShotName);
                                FailerReason = count["AssertionMessage"].ToString();
                                repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                repgen.FileCorreptionCheck();
                                Assert.Fail(FailerReason);
                            }
                            repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                            break;
                        case "ValidatePPEStatus":
                            bool Metrialstatus = log.ValidatePPE_Avail(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString());
                            if (!Metrialstatus)
                            {
                                screenShotName = count["AssertionScreenshotName"].ToString();
                                ErrorScreenPath = log.screenShot(screenShotName);
                                FailerReason = count["AssertionMessage"].ToString();
                                repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                repgen.FileCorreptionCheck();
                                Assert.Fail(FailerReason);
                            }
                            repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                            break;
                        case "AssignedOrderValidate":
                        case "ValidatePickStatus":
                        case "ValidatePackStatus":
                            CommonLanguageTemplateReader.Message(lang, count["DataRefferencekeyword"].ToString());
                            string message = languageResource.Msg_GetTemplateMessage;
                            if (operation == "AssignedOrderValidate")
                            {
                                qualityStatus = log.AssignedOrderValidate(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), count["ControlKeywordValue"].ToString(), message);
                            }
                            else if (operation == "ValidatePackStatus")
                            {
                                qualityStatus = log.MobilePackOrderStatus(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), ExcelDataTable.ReadData(1, count["ControlKeywordValue"].ToString()), message);
                            }
                            else
                            {
                                qualityStatus = log.MobilePickOrderStatus(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), ExcelDataTable.ReadData(1, count["ControlKeywordValue"].ToString()), message);
                            }
                            if (qualityStatus == "False")
                            {
                                screenShotName = count["AssertionScreenshotName"].ToString();
                                ErrorScreenPath = log.screenShot(screenShotName);
                                FailerReason = count["AssertionMessage"].ToString();
                                repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                repgen.FileCorreptionCheck();
                                Assert.Fail(FailerReason);
                            }
                            repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                            break;
                        case "GetDispatchBinLocation":
                        case "GetSAPStock_Batch":
                            string Dispatchbin = string.Empty;
                            if (operation == "GetDispatchBinLocation")
                            {
                                Dispatchbin = log.GettingBinLocation(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), count["ControlKeywordValue"].ToString());
                            }
                            else if (operation == "GetSAPStock_Batch")
                            {
                                Dispatchbin = log.GettingSAP_BatchNumber(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), count["ControlKeywordValue"].ToString());
                            }
                            if (Dispatchbin == string.Empty)
                            {
                                screenShotName = count["AssertionScreenshotName"].ToString();
                                ErrorScreenPath = log.screenShot(screenShotName);
                                FailerReason = count["AssertionMessage"].ToString();
                                repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                repgen.FileCorreptionCheck();
                                Assert.Fail(FailerReason);
                            }
                            writeData.WriteExcel(count["DataRefferencekeyword"].ToString().Split(':')[1], Dispatchbin.Split(':')[0], "ReadData", ExcelDataTable.ReadData(1, count["DataRefferencekeyword"].ToString().Split(':')[0]).ToString());
                            if (operation == "GetDispatchBinLocation")
                            {
                                writeData.WriteExcel("Available SAP Quantity", Dispatchbin.Split(':')[1], "ReadData", ExcelDataTable.ReadData(1, count["DataRefferencekeyword"].ToString().Split(':')[0]).ToString());
                            }
                            writeData.FileCorreptionCheck();
                            repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, Dispatchbin);
                            break;
                        case "DispatchPPEValidation":
                        case "DispatchPickInstValidation":
                            if (operation == "DispatchPPEValidation")
                            {
                                status = log.DispatchMobilePPEValidation(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString());
                            }
                            else
                            {
                                status = log.DispatchMobilePickInstruction(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString());
                            }
                            if (!status)
                            {
                                screenShotName = count["AssertionScreenshotName"].ToString();
                                ErrorScreenPath = log.screenShot(screenShotName);
                                FailerReason = count["AssertionMessage"].ToString();
                                repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                repgen.FileCorreptionCheck();
                                Assert.Fail(FailerReason);
                            }
                            repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                            break;
                        case "ValidatePickOrderStatus":
                            qualityStatus = log.DispatchMobPickingOrderStatus(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), ExcelDataTable.ReadData(1, count["ControlKeywordValue"].ToString()));
                            CommonLanguageTemplateReader.Message(lang, count["DataRefferencekeyword"].ToString());
                            if (languageResource.Msg_GetTemplateMessage.Trim() != qualityStatus.Trim())
                            {
                                screenShotName = count["AssertionScreenshotName"].ToString();
                                ErrorScreenPath = log.screenShot(screenShotName);
                                FailerReason = count["AssertionMessage"].ToString();
                                repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                repgen.FileCorreptionCheck();
                            }
                            Assert.AreEqual(languageResource.Msg_GetTemplateMessage.Replace(" ", string.Empty), qualityStatus, FailerReason);
                            repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                            break;
                        case "ValidateOrderAvailStatus":
                            status = log.CheckOrderAvailable(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), ExcelDataTable.ReadData(1, count["ControlKeywordValue"].ToString()));
                            if (!status)
                            {
                                screenShotName = count["AssertionScreenshotName"].ToString();
                                ErrorScreenPath = log.screenShot(screenShotName);
                                FailerReason = count["AssertionMessage"].ToString();
                                repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                repgen.FileCorreptionCheck();
                                Assert.Fail(FailerReason);
                            }
                            repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                            break;
                        case "ValidatePackInstructionAvail":
                            status = log.PackingInstructions(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString());
                            if (!status)
                            {
                                screenShotName = count["AssertionScreenshotName"].ToString();
                                ErrorScreenPath = log.screenShot(screenShotName);
                                FailerReason = count["AssertionMessage"].ToString();
                                repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                repgen.FileCorreptionCheck();
                                Assert.Fail(FailerReason);
                            }
                            repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                            break;
                        case "ValidatePartialPickSatus":
                            qualityStatus = pop.ScriptGetExcuite(count["Keyword"].ToString(), count["ControlKeyword"].ToString(), count["ControlKeywordValue"].ToString(), count["Step No"].ToString(), count["Description"].ToString());
                            CommonLanguageTemplateReader.Message(lang, count["DataRefferencekeyword"].ToString());
                            if (languageResource.Msg_GetTemplateMessage.Trim() != qualityStatus.Trim())
                            {
                                screenShotName = count["AssertionScreenshotName"].ToString();
                                ErrorScreenPath = log.screenShot(screenShotName);
                                FailerReason = count["AssertionMessage"].ToString();
                                repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                repgen.FileCorreptionCheck();
                            }
                            Assert.AreEqual(languageResource.Msg_GetTemplateMessage.Trim(), qualityStatus, FailerReason);
                            repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                            break;
                        case "MatchPackageDetails":
                            status = log.PackagingDetailsValidation(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), count["ControlKeywordValue"].ToString(), count["DataRefferencekeyword"].ToString());
                            if (!status)
                            {
                                screenShotName = count["AssertionScreenshotName"].ToString();
                                ErrorScreenPath = log.screenShot(screenShotName);
                                FailerReason = count["AssertionMessage"].ToString();
                                repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                repgen.FileCorreptionCheck();
                                Assert.Fail(FailerReason);
                            }
                            repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                            break;
                        case "LinkPPE":
                            qualityStatus = log.PPELink(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), count["ControlKeywordValue"].ToString(), ExcelDataTable.ReadData(1, count["DataRefferencekeyword"].ToString()));
                            if (qualityStatus == "False")
                            {
                                screenShotName = count["AssertionScreenshotName"].ToString();
                                ErrorScreenPath = log.screenShot(screenShotName);
                                FailerReason = count["AssertionMessage"].ToString();
                                repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                repgen.FileCorreptionCheck();
                                Assert.Fail(FailerReason);
                            }
                            repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                            break;
                        case "SearchDataOperation":
                            log.NavigationPageDeleteData(count["Step No"].ToString(), count["Description"].ToString(), count["ControlKeywordValue"].ToString(), ExcelDataTable.ReadData(1, count["DataRefferencekeyword"].ToString()), count["TypeOfControl"].ToString(), new int[] { Convert.ToInt16(count["XCodinate"]), Convert.ToInt16(count["Ycodinate"]) });
                            break;
                        case "GetDeviceInfoID":
                            qualityStatus = log.GETDeviceID(count["Step No"].ToString(), count["Description"].ToString(), count["DataRefferencekeyword"].ToString(), count["ControlKeywordValue"].ToString());
                            if (qualityStatus == "0")
                            {
                                screenShotName = count["AssertionScreenshotName"].ToString();
                                ErrorScreenPath = log.screenShot(screenShotName);
                                FailerReason = count["AssertionMessage"].ToString();
                                repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                repgen.FileCorreptionCheck();
                                Assert.Fail(FailerReason);
                            }
                            RecordData("DeviceInfoID", qualityStatus);
                            repgen.Reports(count["Step No"].ToString(), "Record Device ID For the Give details.", "RecordData", true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, recordedData["DeviceInfoID"]);
                            break;
                        case "GetDeviceIDValue":
                            valueGEt = pop.ScriptGetExcuite(count["Keyword"].ToString(), count["ControlKeyword"].ToString(), count["ControlKeywordValue"].ToString(), count["Step No"].ToString(), count["Description"].ToString());
                            RecordData("DeviceID", valueGEt);
                            repgen.Reports(count["Step No"].ToString(), "Record Device ID For the Give details.", "RecordData", true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, recordedData["DeviceID"]);
                            break;
                        case "BrowserTabVal":
                            status = pop.BrowserTabValidation(count["ControlKeywordValue"].ToString());
                            if (!status)
                            {
                                screenShotName = count["AssertionScreenshotName"].ToString();
                                ErrorScreenPath = log.screenShot(screenShotName);
                                FailerReason = count["AssertionMessage"].ToString();
                                repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                repgen.FileCorreptionCheck();
                                Assert.Fail(FailerReason);
                            }
                            repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                            break;
                        case "ValidateDropdown":
                            validationMessage = pop.WebGetControlData(count["TypeOfControl"].ToString(), count["ControlKeyword"].ToString(), count["ControlKeywordValue"].ToString(), count["AssertionScreenshotName"].ToString(), count["AssertionMessage"].ToString(), count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString()).Trim();
                            if (!validationMessage.Contains(ExcelDataTable.ReadData(1, count["DataRefferencekeyword"].ToString())))
                            {
                                screenShotName = count["AssertionScreenshotName"].ToString();
                                ErrorScreenPath = log.screenShot(screenShotName);
                                FailerReason = count["AssertionMessage"].ToString();
                                repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                repgen.FileCorreptionCheck();
                                Assert.Fail(FailerReason);
                            }
                            StringAssert.Contains(validationMessage, ExcelDataTable.ReadData(1, count["DataRefferencekeyword"].ToString()), FailerReason);
                            repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                            break;
                        case "GetRemainingPickQty":
                            string remainingQty = log.GetRemainingQuantity(count["ControlKeywordValue"].ToString(), ExcelDataTable.ReadData(1, count["DataRefferencekeyword"].ToString()), count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString()).Trim();
                            if(remainingQty == "Null")
                            {
                                screenShotName = count["AssertionScreenshotName"].ToString();
                                ErrorScreenPath = log.screenShot(screenShotName);
                                FailerReason = count["AssertionMessage"].ToString();
                                repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                repgen.FileCorreptionCheck();
                                Assert.Fail(FailerReason);
                            }
                            RecordData("RemainingOrderQTY", remainingQty);
                            repgen.Reports(count["Step No"].ToString(), "Record Remaining Quantity for order.", "RecordData", true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, recordedData["RemainingOrderQTY"]);
                            break;
                        case "GetMaterialBatchData":
                            string BatchData = log.GetBatchReqQuantitry(count["ControlKeywordValue"].ToString(), count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString()).Trim();

                            //Recording Batch Number of the given material.
                            string referenceName = count["ControlKeywordValue"].ToString().Split(':')[1] + "_Batch";
                            writeData.WriteExcel("Batch Number", BatchData.Split(':')[0], "ReadData", ExcelDataTable.ReadData(1, count["ControlKeywordValue"].ToString().Split(':')[1]).ToString());
                            repgen.Reports(count["Step No"].ToString(), "Write down Batch Number for specified " + count["ControlKeywordValue"].ToString().Split(':')[1] + "", "ReadData", true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, BatchData.Split(':')[0]);

                            //Recording Required quantity of the given material.
                            referenceName = count["ControlKeywordValue"].ToString().Split(':')[1] + "_ReqQty";
                            writeData.WriteExcel("Required Quantity", BatchData.Split(':')[1], "ReadData", ExcelDataTable.ReadData(1, count["ControlKeywordValue"].ToString().Split(':')[1]).ToString());
                            repgen.Reports(count["Step No"].ToString(), "Write down Required Quantity for specified " + count["ControlKeywordValue"].ToString().Split(':')[1] + "", "ReadData", true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, BatchData.Split(':')[1]);
                            break;
                        case "SetRef_Data":
                            string inputdata = string.Empty;
                            if (count["DataRefferencekeyword"].ToString().StartsWith("Rec_"))
                            {
                                inputdata = recordedData[count["DataRefferencekeyword"].ToString().Replace("Rec_", string.Empty)];
                            }
                            else
                            {
                                inputdata = ExcelDataTable.ReadData(1, count["DataRefferencekeyword"].ToString());
                            }
                            if (inputdata != count["DataRefferencekeyword"].ToString())
                            {
                                writeData.WriteExcel(count["ControlKeywordValue"].ToString(), inputdata, count["Keyword"].ToString(), "");
                                repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                            }
                            else
                            {
                                if (count["DataRefferencekeyword"].ToString().StartsWith("Data"))
                                {
                                    writeData.WriteExcel(count["ControlKeywordValue"].ToString(), inputdata.Split(':')[1], count["Keyword"].ToString(), "");
                                    repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                }
                                else
                                {
                                    repgen.Reports(count["Step No"].ToString(), "No Data found in " + count["DataRefferencekeyword"].ToString() + "in test data", count["Keyword"].ToString(), false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                    Assert.Fail("No Data found in " + count["DataRefferencekeyword"].ToString() + "in test data");
                                }
                            }
                            break;
                        case "WriteData":
                            if (WriteAndReadData.DataFilePath == string.Empty)
                            {
                                WriteAndReadData.DataFilePath = ConfigurationManager.AppSettings["ExcelDataFile"] + "\\DynamicData" + "\\" + BatchName + ".xls";
                            }
                            pop.GetAndWriteDataforReff(count["ControlKeyword"].ToString(), count["TypeOfControl"].ToString(), count["TypeOfWindow"].ToString(), count["ControlKeywordValue"].ToString(), count["DataRefferencekeyword"].ToString(), count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString());
                            break;
                        case "ReadData":
                            if (WriteAndReadData.DataFilePath == string.Empty)
                            {
                                WriteAndReadData.DataFilePath = ConfigurationManager.AppSettings["ExcelDataFile"] + "\\DynamicData" + "\\" + BatchName + ".xls";
                            }
                            if (count["ControlKeywordValue"].ToString().StartsWith("Rec_"))
                            {
                                validationMessagenotEqual = recordedData[count["ControlKeywordValue"].ToString().Replace("Rec_", string.Empty)];
                            }
                            else
                            {
                                validationMessagenotEqual = pop.WebGetControlData(count["TypeOfControl"].ToString(), count["ControlKeyword"].ToString(), count["ControlKeywordValue"].ToString(), count["AssertionScreenshotName"].ToString(), count["AssertionMessage"].ToString(), count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString()).Trim();
                            }
                            writeData.WriteExcel(count["DataRefferencekeyword"].ToString().Split(':')[1], validationMessagenotEqual, "ReadData", ExcelDataTable.ReadData(1, count["DataRefferencekeyword"].ToString().Split(':')[0]).ToString());
                            repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, validationMessagenotEqual);
                            break;
                        case "ReadStockInfo":
                            if (WriteAndReadData.DataFilePath == string.Empty)
                            {
                                WriteAndReadData.DataFilePath = ConfigurationManager.AppSettings["ExcelDataFile"] + "\\DynamicData" + "\\" + BatchName + ".xls";
                            }
                            validationMessage = pop.WebGetControlData(count["TypeOfControl"].ToString(), count["ControlKeyword"].ToString(), count["ControlKeywordValue"].ToString().Split(':')[0], count["AssertionScreenshotName"].ToString(), count["AssertionMessage"].ToString(), count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString()).Trim();
                            validationMessage = Regex.Replace(validationMessage, "<.*?>", String.Empty).Trim();
                            Message = System.Web.HttpUtility.HtmlDecode(validationMessage);
                            if (Message == "No records found")
                            {
                                writeData.WriteExcel(count["DataRefferencekeyword"].ToString().Split(':')[1], "0", "ReadData", ExcelDataTable.ReadData(1, count["DataRefferencekeyword"].ToString().Split(':')[0]).ToString());
                                repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "0");
                            }
                            else
                            {
                                validationMessagenotEqual = pop.WebGetControlData(count["TypeOfControl"].ToString(), count["ControlKeyword"].ToString(), count["ControlKeywordValue"].ToString().Split(':')[1], count["AssertionScreenshotName"].ToString(), count["AssertionMessage"].ToString(), count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString()).Trim();
                                writeData.WriteExcel(count["DataRefferencekeyword"].ToString().Split(':')[1], validationMessagenotEqual, "ReadData", ExcelDataTable.ReadData(1, count["DataRefferencekeyword"].ToString().Split(':')[0]).ToString());
                                repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, validationMessagenotEqual);
                            }
                            break;
                        case "WriteMaterialOrder":
                            string MaterialOrder = pop.WebGetControlData(count["TypeOfControl"].ToString(), count["ControlKeyword"].ToString(), count["ControlKeywordValue"].ToString(), count["AssertionScreenshotName"].ToString(), count["AssertionMessage"].ToString(), count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString()).Trim();
                            if (MaterialOrder != string.Empty)
                            {
                                if (count["DataRefferencekeyword"].ToString().Contains('+'))
                                {
                                    string[] metrialOrder = count["DataRefferencekeyword"].ToString().Split('+');
                                    writeData.WriteExcel(metrialOrder[1].ToString().Split(':')[0], metrialOrder[1].ToString().Split(':')[1], "ReadData", ExcelDataTable.ReadData(1, metrialOrder[0].Split('=')[1]) + ":" + MaterialOrder);
                                }
                                else
                                {
                                    writeData.WriteExcel(count["DataRefferencekeyword"].ToString().Split(':')[0], count["DataRefferencekeyword"].ToString().Split(':')[1], "ReadData", MaterialOrder);
                                }
                                repgen.Reports(count["Step No"].ToString(), "Write down Material order based on the picking Screen", operation, true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                            }
                            else
                            {
                                repgen.Reports(count["Step No"].ToString(), "Write down Material order based on the picking Screen", operation, false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                Assert.Fail("Write down Material order based on the picking Screen Failed");
                            }
                            break;
                        case "EnterSelectedBin":
                            MaterialOrder = log.EnterProposedBin(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), count["ControlKeyword"].ToString(), count["TypeOfControl"].ToString(), count["TypeOfWindow"].ToString(), count["ControlKeywordValue"].ToString(), count["DataRefferencekeyword"].ToString());
                            repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), operation, true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                            RecordData("PickedBin", MaterialOrder.Split(':')[0].ToString());
                            repgen.Reports(count["Step No"].ToString(), "Record Picked bin for Material.", "RecordData", true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, recordedData["PickedBin"]);
                            RecordData("PickedBinAvailableQTY", MaterialOrder.Split(':')[1].ToString());
                            repgen.Reports(count["Step No"].ToString(), "Record Picked bin for Material.", "RecordData", true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, recordedData["PickedBinAvailableQTY"]);
                            break;
                        case "ValidateReturnWarehouseInfo":
                            status = log.ValidateMaterialInfoINReturnToWarehouse(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), count["ControlKeywordValue"].ToString(), count["DataRefferencekeyword"].ToString());
                            if (!status)
                            {
                                screenShotName = count["AssertionScreenshotName"].ToString();
                                ErrorScreenPath = log.screenShot(screenShotName);
                                FailerReason = count["AssertionMessage"].ToString();
                                repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                repgen.FileCorreptionCheck();
                                Assert.Fail(FailerReason);
                            }
                            repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), operation, true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                            break;
                        case "ZeroTotalQtyValidate":
                            log.ValiateZero(count["Step No"].ToString(), count["Description"].ToString(), count["ControlKeywordValue"].ToString());
                            break;
                        case "GetMaterialBatchStoreID":
                            BatchData = log.GetMaterialDetialsWithStoreID(count["ControlKeywordValue"].ToString(), count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString()).Trim();
                            string[] materialData = BatchData.Split(',');
                            foreach (var item in materialData)
                            {
                                if (item != string.Empty)
                                {
                                    string[] details = item.Split('-');
                                    writeData.WriteExcel("Material", details[1].Trim(), "ReadData", details[0].Split(':')[1].Trim());
                                    writeData.WriteExcel("Batch Number", details[2].Trim(), "ReadData", details[0].Split(':')[1].Trim());
                                    writeData.WriteExcel("Required Quantity", details[3].Trim(), "ReadData", details[0].Split(':')[1].Trim());
                                }
                            }
                            writeData.FileCorreptionCheck();
                            repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), operation, true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                            break;
                        case "ValidatePickedBinStatus":
                            ExcelDataTable.PopulateRecordData(WriteAndReadData.DataFilePath);
                            BatchData = pop.getDataFromDynamicExcel(count["ControlKeywordValue"].ToString());
                            validationMessage = log.GetPickedBinStatusColor(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), BatchData);
                            if (validationMessage != "True")
                            {
                                validationMessage = Regex.Replace(validationMessage, "<.*?>", String.Empty).Trim();
                                Message = System.Web.HttpUtility.HtmlDecode(validationMessage);
                                validationMessage = Message.Replace(" ", string.Empty).Trim();
                                CommonLanguageTemplateReader.Message(lang, count["DataRefferencekeyword"].ToString());
                                if (languageResource.Msg_GetTemplateMessage.Replace(" ", string.Empty) != validationMessage.Trim())
                                {
                                    screenShotName = count["AssertionScreenshotName"].ToString();
                                    ErrorScreenPath = log.screenShot(screenShotName);
                                    FailerReason = count["AssertionMessage"].ToString();
                                    repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                    repgen.FileCorreptionCheck();
                                }
                                Assert.AreEqual(languageResource.Msg_GetTemplateMessage.Replace(" ", string.Empty), validationMessage, FailerReason);
                                repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                            }
                            else
                            {
                                repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "Picked Pin Not Found");
                            }
                            break;
                        case "ValidateMatchDataStatus":
                            if (WriteAndReadData.DataFilePath != string.Empty)
                            {
                                ExcelDataTable.PopulateRecordData(WriteAndReadData.DataFilePath);
                            }
                            validationMessage = log.GetMatchTableDataStatusColor(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), count["ControlKeywordValue"].ToString());
                            validationMessage = Regex.Replace(validationMessage, "<.*?>", String.Empty).Trim();
                            Message = System.Web.HttpUtility.HtmlDecode(validationMessage);
                            validationMessage = Message.Replace(" ", string.Empty).Trim();
                            CommonLanguageTemplateReader.Message(lang, count["DataRefferencekeyword"].ToString());
                            if (languageResource.Msg_GetTemplateMessage.Replace(" ", string.Empty) != validationMessage.Trim())
                            {
                                screenShotName = count["AssertionScreenshotName"].ToString();
                                ErrorScreenPath = log.screenShot(screenShotName);
                                FailerReason = count["AssertionMessage"].ToString();
                                repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                repgen.FileCorreptionCheck();
                            }
                            Assert.AreEqual(languageResource.Msg_GetTemplateMessage.Replace(" ", string.Empty), validationMessage, FailerReason);
                            repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                            break;
                        case "WarehouseStockValidation":
                            status = log.ValidateTransferStockMaterialAndQty(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), count["ControlKeywordValue"].ToString(), count["DataRefferencekeyword"].ToString());
                            if (!status)
                            {
                                screenShotName = count["AssertionScreenshotName"].ToString();
                                ErrorScreenPath = log.screenShot(screenShotName);
                                FailerReason = count["AssertionMessage"].ToString();
                                repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                repgen.FileCorreptionCheck();
                                Assert.Fail(FailerReason);
                            }
                            repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                            break;
                        case "LaunchDataBase":
                            log.sqldeveloperApplicationLaunch();
                            repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                            break;
                        case "ConditionEquals":
                        case "ConditionNotEquals":
                            validationMessage = pop.WebGetControlData(count["TypeOfControl"].ToString(), count["ControlKeyword"].ToString(), count["ControlKeywordValue"].ToString(), count["AssertionScreenshotName"].ToString(), count["AssertionMessage"].ToString(), count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString()).Trim();
                            validationMessage = Regex.Replace(validationMessage, "<.*?>", String.Empty).Trim();
                            Message = System.Web.HttpUtility.HtmlDecode(validationMessage);
                            validationMessage = Message.Replace(" ", string.Empty).Trim();
                            CommonLanguageTemplateReader.Message(lang, count["DataRefferencekeyword"].ToString());
                            if (operation == "ConditionEquals")
                            {
                                if (languageResource.Msg_GetTemplateMessage.Replace(" ", string.Empty) == validationMessage.Trim())
                                {
                                    repgen.Reports("Given Condition Is Matched Starting Conditional Steps", "", "", true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                    PreformForigneOperations("Conditional Execute", OperationPath, ProjectName, numberofExecution);
                                    repgen.Reports("Executing Conditional Steps Completed", "", "", true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                }
                            }
                            else if (operation == "ConditionNotEquals")
                            {
                                if (languageResource.Msg_GetTemplateMessage.Replace(" ", string.Empty) != validationMessage.Trim())
                                {
                                    repgen.Reports("Given Condition Is Matched Starting Conditional Steps", "", "", true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                    PreformForigneOperations("Conditional Execute", OperationPath, ProjectName, numberofExecution);
                                    repgen.Reports("Executing Conditional Steps Completed", "", "", true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                }
                            }
                            break;
                        case "GetReferenceNumber":
                        case "GetSTONumber":
                            ScrollStatus = pop.ScriptGetExcuite(count["Keyword"].ToString(), count["ControlKeyword"].ToString(), count["ControlKeywordValue"].ToString(), count["Step No"].ToString(), count["Description"].ToString());
                            if (ScrollStatus != string.Empty)
                            {
                                writeData.WriteExcel(count["DataRefferencekeyword"].ToString().Split(':')[1], ScrollStatus, "ReadData", ExcelDataTable.ReadData(1, count["DataRefferencekeyword"].ToString().Split(':')[0]).ToString());
                            }
                            repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                            break;
                        case "GetDeliveryOrdReport":
                            ScrollStatus = log.GetDispatchReportAvailability(count["ControlKeywordValue"].ToString(), count["DataRefferencekeyword"].ToString().Split(':')[0], count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString());
                            if (ScrollStatus == "0")
                            {
                                screenShotName = count["AssertionScreenshotName"].ToString();
                                ErrorScreenPath = log.screenShot(screenShotName);
                                FailerReason = count["AssertionMessage"].ToString();
                                repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                repgen.FileCorreptionCheck();
                                Assert.Fail(FailerReason);
                            }
                            else
                            {
                                writeData.WriteExcel("Dispatch Report Order", ScrollStatus, "ReadData", ExcelDataTable.ReadData(1, count["DataRefferencekeyword"].ToString().Split(':')[1]));
                                RecordData("DispatchReportOrder", ScrollStatus);
                                repgen.Reports(count["Step No"].ToString(), "Record the order count for the specified order", "RecordData", true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, recordedData["DispatchReportOrder"]);
                            }
                            repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                            break;
                        case "DateTimeReportValidate":
                            status = log.getdispatchDateTimeValidation(count["ControlKeywordValue"].ToString(), count["DataRefferencekeyword"].ToString(), count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString());
                            if (!status)
                            {
                                screenShotName = count["AssertionScreenshotName"].ToString();
                                ErrorScreenPath = log.screenShot(screenShotName);
                                FailerReason = count["AssertionMessage"].ToString();
                                repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                repgen.FileCorreptionCheck();
                                Assert.Fail(FailerReason);
                            }
                            repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                            break;
                        case "B2BSAPValidation":
                            status = log.B2BDropValidation(count["ControlKeywordValue"].ToString(), count["DataRefferencekeyword"].ToString(), count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString());
                            if (!status)
                            {
                                screenShotName = count["AssertionScreenshotName"].ToString();
                                ErrorScreenPath = log.screenShot(screenShotName);
                                FailerReason = count["AssertionMessage"].ToString();
                                repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                repgen.FileCorreptionCheck();
                                Assert.Fail(FailerReason);
                            }
                            repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                            break;
                        case "GetSystemDate":
                            //Get's the current System Date.
                            string date = log.getSystemDate(ExcelDataTable.ReadData(1, "DateFormat").ToString());
                            if (date != string.Empty)
                            {
                                pop.ScriptExcuite("SetValue", count["ControlKeyword"].ToString(), count["ControlKeywordValue"].ToString(), date, count["Step No"].ToString(), count["Description"].ToString());
                                repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                            }
                            else
                            {
                                repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                repgen.FileCorreptionCheck();
                                Assert.Fail(FailerReason);
                            }
                            break;
                        case "EnterRandomNumber":
                            //"Reference Data in Dynamic Excel"="Number or Digit":"Column Head in Dynamic Excel"
                            materialData = count["DataRefferencekeyword"].ToString().Split('=');
                            date = log.RamdomBatch(Convert.ToInt16(materialData[1].Split(':')[0]));
                            if (date != string.Empty)
                            {
                                writeData.WriteExcel(count["DataRefferencekeyword"].ToString().Split(':')[1], date, "ReadData", materialData[0]);
                                pop.OperationStart(count["Keyword"].ToString(), count["TypeOfControl"].ToString(), count["ControlKeyword"].ToString(), count["ControlKeywordValue"].ToString(), date, count["TypeOfWindow"].ToString(), count["Step No"].ToString(), count["Description"].ToString());
                                repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, date);
                            }
                            break;
                        case "SetAttribute":
                            // set specified attribute for a control.
                            pop.ScriptSetAttribute(count["Keyword"].ToString(), count["ControlKeyword"].ToString(), count["ControlKeywordValue"].ToString(), count["DataRefferencekeyword"].ToString(), count["Step No"].ToString(), count["Description"].ToString());
                            repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                            break;
                        case "GetDate":
                            date = log.getDate(ExcelDataTable.ReadData(1, "DateFormat").ToString(), count["DataRefferencekeyword"].ToString());
                            if (date != string.Empty)
                            {
                                pop.ScriptExcuite("SetValue", count["ControlKeyword"].ToString(), count["ControlKeywordValue"].ToString(), date, count["Step No"].ToString(), count["Description"].ToString());
                                repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                            }
                            else
                            {
                                repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                repgen.FileCorreptionCheck();
                                Assert.Fail("Number of days you want to ADD or SUBTRACT is not proper");
                            }
                            break;
                        case "SetDefault":
                            date = log.GetDefaultData(count["ControlKeywordValue"].ToString(), count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString());
                            writeData.WriteExcel(count["DataRefferencekeyword"].ToString().Split(':')[1], date, "ReadData", count["DataRefferencekeyword"].ToString().Split(':')[0]);
                            repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, date);
                            break;
                        case "PartBagValidation":
                            CommonLanguageTemplateReader.Message(lang, count["DataRefferencekeyword"].ToString().Split('=')[0]);
                            int PartBagcount = log.getPartBagCount(count["ControlKeywordValue"].ToString(), languageResource.Msg_GetTemplateMessage, count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString());
                            int numberofpartbag = Convert.ToInt16(ExcelDataTable.ReadData(1, count["DataRefferencekeyword"].ToString().Split('=')[1]));
                            if (PartBagcount != numberofpartbag)
                            {
                                screenShotName = count["AssertionScreenshotName"].ToString();
                                ErrorScreenPath = log.screenShot(screenShotName);
                                FailerReason = count["AssertionMessage"].ToString();
                                repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                                repgen.FileCorreptionCheck();
                                Assert.Fail(FailerReason);
                            }
                            repgen.Reports(count["Step No"].ToString(), count["Description"].ToString(), count["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                            break;
                        case "ValidatePOAvailable":
                            log.POAvailabilityCalender(count["ControlKeyword"].ToString(), count["Keyword"].ToString(), count["ControlKeywordValue"].ToString(), count["Step No"].ToString(), count["Description"].ToString());
                            break;
                        default:
                            FailerReason = "No " + operation + " keyword found. Please check the Keyword.";
                            throw new NoSuchOperationFound("No " + operation + " keyword found. Please check the Keyword.");
                    }
                }
            }
            OperationTable.Clear();
        }

        /// <summary>
        /// This will setup the log generation settings.The settings is contained in the Global Elements excel
        /// 1. Errors Snapshot (take snapshot for error screens).
        /// 2. No Log Generation (Logs won't be generated.)
        /// 3. All Action Snapshot (This will generate logs and take screen shot for all screens).
        /// </summary>
        public void SetLogSettings()
        {
            string LogSettings = ExcelDataTable.ReadData(1, "LogGenerationSettings");
            if (LogSettings == "All Action Snapshot")
            {
                Playback.PlaybackSettings.LoggerOverrideState = HtmlLoggerState.AllActionSnapshot;
            }
            else if (LogSettings == "No Log Generation")
            {
                Playback.PlaybackSettings.LoggerOverrideState = HtmlLoggerState.Disabled;
            }
            else
            {
                Playback.PlaybackSettings.LoggerOverrideState = HtmlLoggerState.ErrorAndWarningOnlySnapshot;
            }
        }

        /// <summary>
        /// This will create the dynamic variable and store the value in that variable
        /// </summary>
        /// <param name="refference">This will accept the variable name to store the value.</param>
        /// <param name="data">This will accept the value to be stored in the specified variable</param>
        public void RecordData(string refference, string data)
        {
            try
            {
                if(recordedData.ContainsKey(refference))
                {
                    recordedData.Remove(refference);
                    recordedData.Add(refference, data);
                }
                else
                {
                    recordedData.Add(refference, data);
                }
            }
            catch (NullReferenceException e)
            {
                throw new Exception(refference + "Variable don't have any data to store");
            }
            catch (Exception e) { throw new Exception(refference + "Variable is already declared, please give another.."); }

            

        }

    }
}
