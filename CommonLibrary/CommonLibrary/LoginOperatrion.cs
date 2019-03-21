using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Configuration;
using Microsoft.VisualStudio.TestTools.UITesting;
using System.Threading;
using AutoItX3Lib;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;
using System.Drawing.Imaging;
using System.Drawing;
using CommonLibrary.Exceptions;
using CommonLibrary.DataDrivenTesting;
using CommonLibrary.KeywordDrivenTesting;
using CommonLibrary.Operations;
using CommonLibrary.Reports;
using CommonLibrary.Writedata;
using System.Diagnostics;
using OperationLibrary;
using Microsoft.VisualStudio.TestTools.UITest.Extension;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;
using System.Globalization;

namespace CommonLibrary.Log
{
    public class LoginOperatrion
    {
        private string ConfigPath = string.Empty;
        private string Url = string.Empty;
        string path = string.Empty;
        string url = string.Empty;
        public static int len = 0, min = 0, mid = 0, max = 0;
        DataTable dt = new DataTable();
        public static string ProjectName;
        public static string LogPath = string.Empty;
        public static string SecurityCode = string.Empty;
        public static string WareHouseNo = string.Empty;
        public static string DetaildReportStatus = string.Empty;
        public static string batchforReport = string.Empty;
        public static string requiredPlant = string.Empty;
        public string UsertypeLogin = string.Empty;
        public int numberOfIteration = 0;
        public string Init = string.Empty;
        private static string _numbers = "0123456789";
        Random random = new Random();
        AutoItX3 auto = new AutoItX3();
        LogLanguageTemplete lang = new LogLanguageTemplete();
        PerformOperation Pop = new PerformOperation();
        ReportGeneration genreport = new ReportGeneration();
        WriteAndReadData datawrite = new WriteAndReadData();

        #region InititalOperations
        /// <summary>
        /// This will  launch the browser with specified URL.
        /// </summary>
        /// <param name="project"></param>
        public void Initialize(string project)
        {
            ProjectName = project;
            BrowserWindow window = new BrowserWindow();
            window.SearchProperties[BrowserWindow.PropertyNames.ClassName] = "IEFrame";
            UITestControlCollection wndcollection = window.FindMatchingControls();
            foreach (UITestControl control in wndcollection)
            {
                if (control is BrowserWindow)
                {
                    ((BrowserWindow)control).Close();
                }
            }
            //Reading the data from the excel file from the path given in the project.
            path = ConfigurationManager.AppSettings["LoginDetails"];
            Init = ConfigurationManager.AppSettings["Initialization"];
            ExcelDataTable.PopulateBatchWiseData(path);
            ExcelDataTable.PopulateInCollection(Init + "\\GlobalElements.xlsx");
            batchforReport = ConfigurationManager.AppSettings["Batch"];
            len = Convert.ToInt16(ExcelDataTable.ReadBatchData(batchforReport, "Execution"));
            DetaildReportStatus = ExcelDataTable.ReadData(1, "DetailedReportGeneration");
            for (int i = 0; i <= len; i++)
            {
                url = ExcelDataTable.ReadData(1, "Url");
            }
            BrowserWindow.Launch(url);
        }
        #endregion

        #region Login & Logout
        /// <summary>
        /// This will perform the Login operations that specified in the Common keyword drivers. for login refer Sheet Name:"Login" 
        /// </summary>
        /// <param name="Usertype"></param>
        public void login(string Usertype)
        {
            min = Convert.ToInt16(ExcelDataTable.ReadData(1, "Minimum"));
            mid = Convert.ToInt16(ExcelDataTable.ReadData(1, "Medium"));
            max = Convert.ToInt16(ExcelDataTable.ReadData(1, "Maximum"));
            LogLanguageTemplete.messageResource(ExcelDataTable.ReadBatchData(batchforReport, "languageCode"));
            requiredPlant = ExcelDataTable.ReadBatchData(batchforReport, "Plant");
            //Reading the excel file for common operations like Login, Logout, Manage language, Navigation.
            LogPath = ConfigurationManager.AppSettings["LogOperation"];
            Thread.Sleep(max);
            UsertypeLogin = Usertype;
            dt.Clear();
            if (Usertype == "Login")
            {
                dt = ExcelKeywordTable.ExcelData(LogPath + "\\KeywordDrivenData.xlsx", "Login");
            }
            else
            {
                dt = ExcelKeywordTable.ExcelData(LogPath + "\\KeywordDrivenData.xlsx", "RegionalLogin");
            }
            Thread.Sleep(max);
            BrowserWindow window = new BrowserWindow();
            window.Maximized = true;
            Thread.Sleep(max);
            Pop.OperationStart(dt.Rows[0]["Keyword"].ToString(), dt.Rows[0]["TypeOfControl"].ToString(), dt.Rows[0]["ControlKeyword"].ToString(), dt.Rows[0]["ControlKeywordValue"].ToString(), ExcelDataTable.ReadBatchData(batchforReport, dt.Rows[0]["DataRefferencekeyword"].ToString()), dt.Rows[0]["TypeOfWindow"].ToString(), dt.Rows[0]["Step No"].ToString(), dt.Rows[0]["Description"].ToString());
            Pop.OperationStart(dt.Rows[1]["Keyword"].ToString(), dt.Rows[1]["TypeOfControl"].ToString(), dt.Rows[1]["ControlKeyword"].ToString(), dt.Rows[1]["ControlKeywordValue"].ToString(), ExcelDataTable.ReadBatchData(batchforReport, dt.Rows[1]["DataRefferencekeyword"].ToString()), dt.Rows[1]["TypeOfWindow"].ToString(), dt.Rows[1]["Step No"].ToString(), dt.Rows[1]["Description"].ToString());
            Pop.OperationStart(dt.Rows[2]["Keyword"].ToString(), dt.Rows[2]["TypeOfControl"].ToString(), dt.Rows[2]["ControlKeyword"].ToString(), dt.Rows[2]["ControlKeywordValue"].ToString(), dt.Rows[2]["DataRefferencekeyword"].ToString(), dt.Rows[2]["TypeOfWindow"].ToString(), dt.Rows[2]["Step No"].ToString(), dt.Rows[2]["Description"].ToString());
            Thread.Sleep(max * 2);
            string Login = window.ExecuteScript("var data = document.getElementById('ContentPlaceHolder_lblErrorMessage');return data;").ToString();
            if (Login != string.Empty)
            {
                string LoginValidation = window.ExecuteScript("var data = document.getElementById('ContentPlaceHolder_lblErrorMessage').innerHTML;return data;").ToString();
                if (LoginValidation != "")
                {
                    Thread.Sleep(max);
                    Pop.OperationStart(dt.Rows[1]["Keyword"].ToString(), dt.Rows[1]["TypeOfControl"].ToString(), dt.Rows[1]["ControlKeyword"].ToString(), dt.Rows[1]["ControlKeywordValue"].ToString(), ExcelDataTable.ReadBatchData(batchforReport, dt.Rows[1]["DataRefferencekeyword"].ToString()), dt.Rows[1]["TypeOfWindow"].ToString(), dt.Rows[1]["Step No"].ToString(), dt.Rows[1]["Description"].ToString());
                    Thread.Sleep(mid);
                    Pop.OperationStart(dt.Rows[2]["Keyword"].ToString(), dt.Rows[2]["TypeOfControl"].ToString(), dt.Rows[2]["ControlKeyword"].ToString(), dt.Rows[2]["ControlKeywordValue"].ToString(), dt.Rows[2]["DataRefferencekeyword"].ToString(), dt.Rows[2]["TypeOfWindow"].ToString(), dt.Rows[2]["Step No"].ToString(), dt.Rows[2]["Description"].ToString());
                    Thread.Sleep(max * 2);
                    Login = window.ExecuteScript("var data = document.getElementById('ContentPlaceHolder_lblErrorMessage');return data;").ToString();
                    if (Login != string.Empty)
                    {
                        LoginValidation = window.ExecuteScript("var data = document.getElementById('ContentPlaceHolder_lblErrorMessage').innerHTML;return data;").ToString();
                        if (LoginValidation != "")
                        {
                            string screenShotName = "LoginFailed";
                            Operation.ErrorScreenPath = screenShot(screenShotName);
                            Operation.FailerReason = "Login Failed";
                            genreport.FileCorreptionCheck();
                        }
                        Assert.AreEqual("", LoginValidation, "Login Failed");
                    }
                    else
                    {
                        Thread.Sleep(max);
                        auto.Send("{F5}");
                        Pop.OperationStart(dt.Rows[3]["Keyword"].ToString(), dt.Rows[3]["TypeOfControl"].ToString(), dt.Rows[3]["ControlKeyword"].ToString(), dt.Rows[3]["ControlKeywordValue"].ToString(), dt.Rows[3]["DataRefferencekeyword"].ToString(), dt.Rows[3]["TypeOfWindow"].ToString(), dt.Rows[3]["Step No"].ToString(), dt.Rows[3]["Description"].ToString());
                    }
                }
            }
            else
            {
                Thread.Sleep(max);
                auto.Send("{F5}");
                Pop.OperationStart(dt.Rows[3]["Keyword"].ToString(), dt.Rows[3]["TypeOfControl"].ToString(), dt.Rows[3]["ControlKeyword"].ToString(), dt.Rows[3]["ControlKeywordValue"].ToString(), dt.Rows[3]["DataRefferencekeyword"].ToString(), dt.Rows[3]["TypeOfWindow"].ToString(), dt.Rows[3]["Step No"].ToString(), dt.Rows[3]["Description"].ToString());
            }
            
        }

        /// <summary>
        /// This will help to login to super admin.
        /// </summary>
        public void SuperLogin()
        {
            min = Convert.ToInt16(ExcelDataTable.ReadData(1, "Minimum"));
            mid = Convert.ToInt16(ExcelDataTable.ReadData(1, "Medium"));
            max = Convert.ToInt16(ExcelDataTable.ReadData(1, "Maximum"));
            LogLanguageTemplete.messageResource(ExcelDataTable.ReadBatchData(batchforReport, "languageCode"));
            requiredPlant = ExcelDataTable.ReadBatchData(batchforReport, "Plant");
            LogPath = ConfigurationManager.AppSettings["LogOperation"];
            Thread.Sleep(max);
            dt.Clear();

            dt = ExcelKeywordTable.ExcelData(LogPath + "\\KeywordDrivenData.xlsx", "SuperAdminLogin");

            Thread.Sleep(max);
            BrowserWindow window = new BrowserWindow();
            window.Maximized = true;
            Thread.Sleep(max);
            Pop.OperationStart(dt.Rows[0]["Keyword"].ToString(), dt.Rows[0]["TypeOfControl"].ToString(), dt.Rows[0]["ControlKeyword"].ToString(), dt.Rows[0]["ControlKeywordValue"].ToString(), ExcelDataTable.ReadBatchData(batchforReport, dt.Rows[0]["DataRefferencekeyword"].ToString()), dt.Rows[0]["TypeOfWindow"].ToString(), dt.Rows[0]["Step No"].ToString(), dt.Rows[0]["Description"].ToString());
            Pop.OperationStart(dt.Rows[1]["Keyword"].ToString(), dt.Rows[1]["TypeOfControl"].ToString(), dt.Rows[1]["ControlKeyword"].ToString(), dt.Rows[1]["ControlKeywordValue"].ToString(), ExcelDataTable.ReadBatchData(batchforReport, dt.Rows[1]["DataRefferencekeyword"].ToString()), dt.Rows[1]["TypeOfWindow"].ToString(), dt.Rows[1]["Step No"].ToString(), dt.Rows[1]["Description"].ToString());
            Pop.OperationStart(dt.Rows[2]["Keyword"].ToString(), dt.Rows[2]["TypeOfControl"].ToString(), dt.Rows[2]["ControlKeyword"].ToString(), dt.Rows[2]["ControlKeywordValue"].ToString(), dt.Rows[2]["DataRefferencekeyword"].ToString(), dt.Rows[2]["TypeOfWindow"].ToString(), dt.Rows[2]["Step No"].ToString(), dt.Rows[2]["Description"].ToString());
            Thread.Sleep(max * 2);
            string Login = window.ExecuteScript("var data = document.getElementById('ContentPlaceHolder_lblErrorMessage');return data;").ToString();
            if (Login != string.Empty)
            {
                string LoginValidation = window.ExecuteScript("var data = document.getElementById('ContentPlaceHolder_lblErrorMessage').innerHTML;return data;").ToString();
                if (LoginValidation != "")
                {
                    Thread.Sleep(max);
                    Pop.OperationStart(dt.Rows[1]["Keyword"].ToString(), dt.Rows[1]["TypeOfControl"].ToString(), dt.Rows[1]["ControlKeyword"].ToString(), dt.Rows[1]["ControlKeywordValue"].ToString(), ExcelDataTable.ReadBatchData(batchforReport, dt.Rows[1]["DataRefferencekeyword"].ToString()), dt.Rows[1]["TypeOfWindow"].ToString(), dt.Rows[1]["Step No"].ToString(), dt.Rows[1]["Description"].ToString());
                    Thread.Sleep(mid);
                    Pop.OperationStart(dt.Rows[2]["Keyword"].ToString(), dt.Rows[2]["TypeOfControl"].ToString(), dt.Rows[2]["ControlKeyword"].ToString(), dt.Rows[2]["ControlKeywordValue"].ToString(), dt.Rows[2]["DataRefferencekeyword"].ToString(), dt.Rows[2]["TypeOfWindow"].ToString(), dt.Rows[2]["Step No"].ToString(), dt.Rows[2]["Description"].ToString());
                    Thread.Sleep(max * 2);
                    Login = window.ExecuteScript("var data = document.getElementById('ContentPlaceHolder_lblErrorMessage');return data;").ToString();
                    if (Login != string.Empty)
                    {
                        LoginValidation = window.ExecuteScript("var data = document.getElementById('ContentPlaceHolder_lblErrorMessage').innerHTML;return data;").ToString();
                        if (LoginValidation != "")
                        {
                            string screenShotName = "LoginFailed";
                            Operation.ErrorScreenPath = screenShot(screenShotName);
                            Operation.FailerReason = "Login Failed";
                            genreport.FileCorreptionCheck();
                            Assert.Fail("Login Failed");
                        }
                        else
                        {
                            Thread.Sleep(max);
                            Pop.OperationStart(dt.Rows[3]["Keyword"].ToString(), dt.Rows[3]["TypeOfControl"].ToString(), dt.Rows[3]["ControlKeyword"].ToString(), dt.Rows[3]["ControlKeywordValue"].ToString(), dt.Rows[3]["DataRefferencekeyword"].ToString(), dt.Rows[3]["TypeOfWindow"].ToString(), dt.Rows[3]["Step No"].ToString(), dt.Rows[3]["Description"].ToString());
                            generateSecurityCode(ExcelDataTable.ReadBatchData(batchforReport, dt.Rows[0]["DataRefferencekeyword"].ToString()));
                            Pop.OperationStart("EnterText", dt.Rows[4]["TypeOfControl"].ToString(), dt.Rows[4]["ControlKeyword"].ToString(), dt.Rows[4]["ControlKeywordValue"].ToString(), SecurityCode, dt.Rows[4]["TypeOfWindow"].ToString(), dt.Rows[4]["Step No"].ToString(), dt.Rows[4]["Description"].ToString());
                            Pop.OperationStart(dt.Rows[5]["Keyword"].ToString(), dt.Rows[5]["TypeOfControl"].ToString(), dt.Rows[5]["ControlKeyword"].ToString(), dt.Rows[5]["ControlKeywordValue"].ToString(), dt.Rows[5]["DataRefferencekeyword"].ToString(), dt.Rows[5]["TypeOfWindow"].ToString(), dt.Rows[5]["Step No"].ToString(), dt.Rows[5]["Description"].ToString());
                            Thread.Sleep(max * 2);
                            string SecurityMsg = window.ExecuteScript("var data = document.getElementById('spnSecurityCodeErrorMsg');return data;").ToString();
                            if (SecurityMsg != string.Empty)
                            {
                                string SecurityMsgValidation = window.ExecuteScript("var data = document.getElementById('spnSecurityCodeErrorMsg').innerHTML;return data;").ToString();
                                if (SecurityMsgValidation == lang.Msg_WrongSecurityCode)
                                {
                                    string screenShotName = "WrongSecurityCodeGenerated";
                                    Operation.ErrorScreenPath = screenShot(screenShotName);
                                    Operation.FailerReason = "Wrong Security Code Generated, Login Failed";
                                    genreport.FileCorreptionCheck();
                                }
                                Assert.AreEqual("", SecurityMsgValidation, "Wrong Security Code Generated, Login Failed");
                            }
                            Pop.OperationStart(dt.Rows[6]["Keyword"].ToString(), dt.Rows[6]["TypeOfControl"].ToString(), dt.Rows[6]["ControlKeyword"].ToString(), dt.Rows[6]["ControlKeywordValue"].ToString(), dt.Rows[6]["DataRefferencekeyword"].ToString(), dt.Rows[6]["TypeOfWindow"].ToString(), dt.Rows[6]["Step No"].ToString(), dt.Rows[6]["Description"].ToString());
                        }
                    }
                }
                else
                {
                    Thread.Sleep(max);
                    Pop.OperationStart(dt.Rows[3]["Keyword"].ToString(), dt.Rows[3]["TypeOfControl"].ToString(), dt.Rows[3]["ControlKeyword"].ToString(), dt.Rows[3]["ControlKeywordValue"].ToString(), dt.Rows[3]["DataRefferencekeyword"].ToString(), dt.Rows[3]["TypeOfWindow"].ToString(), dt.Rows[3]["Step No"].ToString(), dt.Rows[3]["Description"].ToString());
                    generateSecurityCode(ExcelDataTable.ReadBatchData(batchforReport, dt.Rows[0]["DataRefferencekeyword"].ToString()));
                    Pop.OperationStart("EnterText", dt.Rows[4]["TypeOfControl"].ToString(), dt.Rows[4]["ControlKeyword"].ToString(), dt.Rows[4]["ControlKeywordValue"].ToString(), SecurityCode, dt.Rows[4]["TypeOfWindow"].ToString(), dt.Rows[4]["Step No"].ToString(), dt.Rows[4]["Description"].ToString());
                    Pop.OperationStart(dt.Rows[5]["Keyword"].ToString(), dt.Rows[5]["TypeOfControl"].ToString(), dt.Rows[5]["ControlKeyword"].ToString(), dt.Rows[5]["ControlKeywordValue"].ToString(), dt.Rows[5]["DataRefferencekeyword"].ToString(), dt.Rows[5]["TypeOfWindow"].ToString(), dt.Rows[5]["Step No"].ToString(), dt.Rows[5]["Description"].ToString());
                    Thread.Sleep(max * 2);
                    string SecurityMsg = window.ExecuteScript("var data = document.getElementById('spnSecurityCodeErrorMsg');return data;").ToString();
                    if (SecurityMsg != string.Empty)
                    {
                        string SecurityMsgValidation = window.ExecuteScript("var data = document.getElementById('spnSecurityCodeErrorMsg').innerHTML;return data;").ToString();
                        if (SecurityMsgValidation == lang.Msg_WrongSecurityCode)
                        {
                            string screenShotName = "WrongSecurityCodeGenerated";
                            Operation.ErrorScreenPath = screenShot(screenShotName);
                            Operation.FailerReason = "Wrong Security Code Generated, Login Failed";
                            genreport.FileCorreptionCheck();
                        }
                        Assert.AreEqual("", SecurityMsgValidation, "Wrong Security Code Generated, Login Failed");
                    }
                    Pop.OperationStart(dt.Rows[6]["Keyword"].ToString(), dt.Rows[6]["TypeOfControl"].ToString(), dt.Rows[6]["ControlKeyword"].ToString(), dt.Rows[6]["ControlKeywordValue"].ToString(), dt.Rows[6]["DataRefferencekeyword"].ToString(), dt.Rows[6]["TypeOfWindow"].ToString(), dt.Rows[6]["Step No"].ToString(), dt.Rows[6]["Description"].ToString());
                }
            }
        }
        /// <summary>
        /// This will perform the Login operations that specified in the Common keyword drivers. for login refer Sheet Name:"Logout" 
        /// </summary>
        public void Logout()
        {
            BrowserWindow window = new BrowserWindow();
            Thread.Sleep(mid);
            auto.Send("{F5}");
            Thread.Sleep(mid);
            dt.Clear();
            dt = ExcelKeywordTable.ExcelData(LogPath + "\\KeywordDrivenData.xlsx", "Logout");
            Thread.Sleep(max);
            Pop.OperationStart(dt.Rows[0]["Keyword"].ToString(), dt.Rows[0]["TypeOfControl"].ToString(), dt.Rows[0]["ControlKeyword"].ToString(), dt.Rows[0]["ControlKeywordValue"].ToString(), dt.Rows[0]["DataRefferencekeyword"].ToString(), dt.Rows[0]["TypeOfWindow"].ToString(), dt.Rows[0]["Step No"].ToString(), dt.Rows[0]["Description"].ToString());
            Thread.Sleep(max);
            try
            {
                Pop.OperationStart(dt.Rows[1]["Keyword"].ToString(), dt.Rows[1]["TypeOfControl"].ToString(), dt.Rows[1]["ControlKeyword"].ToString(), dt.Rows[1]["ControlKeywordValue"].ToString(), dt.Rows[1]["DataRefferencekeyword"].ToString(), dt.Rows[1]["TypeOfWindow"].ToString(), dt.Rows[1]["Step No"].ToString(), dt.Rows[1]["Description"].ToString());
            }
            catch (Exception e) { }
            Thread.Sleep(mid * 2);

            Pop.OperationStart(dt.Rows[2]["Keyword"].ToString(), dt.Rows[2]["TypeOfControl"].ToString(), dt.Rows[2]["ControlKeyword"].ToString(), dt.Rows[2]["ControlKeywordValue"].ToString(), dt.Rows[2]["DataRefferencekeyword"].ToString(), dt.Rows[2]["TypeOfWindow"].ToString(), dt.Rows[2]["Step No"].ToString(), dt.Rows[2]["Description"].ToString());

            string AddingSuccess = Pop.WebGetControlData(dt.Rows[3]["TypeOfControl"].ToString(), dt.Rows[3]["ControlKeyword"].ToString(), dt.Rows[3]["ControlKeywordValue"].ToString(), "LoginPageNotOpen", "Logout Failed", dt.Rows[3]["Step No"].ToString(), dt.Rows[3]["Description"].ToString(), dt.Rows[3]["Keyword"].ToString()).Trim();
            if (AddingSuccess != null)
            {
                string[] msg = AddingSuccess.Split('.');
                AddingSuccess = msg[0];
            }
            if (AddingSuccess != lang.Msg_LogoutSuccessMessage)
            {
                string screenShotName = "LoginPageNotOpen";
                Operation.ErrorScreenPath = screenShot(screenShotName);
                Operation.FailerReason = "Logout Failed";
                genreport.FileCorreptionCheck();
            }
            Assert.AreEqual(lang.Msg_LogoutSuccessMessage, AddingSuccess, "Logout Failed");

            Thread.Sleep(min);
            auto.Send("{F5}");
            Thread.Sleep(mid);
            Pop.OperationStart(dt.Rows[4]["Keyword"].ToString(), dt.Rows[4]["TypeOfControl"].ToString(), dt.Rows[4]["ControlKeyword"].ToString(), dt.Rows[4]["ControlKeywordValue"].ToString(), dt.Rows[4]["DataRefferencekeyword"].ToString(), dt.Rows[4]["TypeOfWindow"].ToString(), dt.Rows[4]["Step No"].ToString(), dt.Rows[4]["Description"].ToString());
            genreport.FileCorreptionCheck();
        }

        public void generateSecurityCode(string username)
        {
            SecurityCode = DateTime.Today.ToString("MMM") + DateTime.Today.Day.ToString("00") + username.Substring(0, 3);
        }
        #endregion

        #region Plant wise Functions
        /// <summary>
        /// This will perform the Login operations based on Plant this is used in Plant Specific Batch only that specified in the Common keyword drivers. for login refer Sheet Name:"Login" 
        /// </summary>
        /// <param name="Usertype">Regional User or Plant User</param>
        /// <param name="Iteration">Number of execution of the test case</param>
        public void PlantLogin(string Usertype, int Iteration)
        {
            min = Convert.ToInt16(ExcelDataTable.ReadData(1, "Minimum"));
            mid = Convert.ToInt16(ExcelDataTable.ReadData(1, "Medium"));
            max = Convert.ToInt16(ExcelDataTable.ReadData(1, "Maximum"));
            string PlantPath = ConfigurationManager.AppSettings["PlantLoginDetails"];
            ExcelDataTable.PopulatePlantWiseData(PlantPath);
            numberOfIteration = Iteration;
            LogLanguageTemplete.messageResource(ExcelDataTable.ReadBatchData(batchforReport, "languageCode"));
            requiredPlant = ExcelDataTable.ReadBatchData(batchforReport, "Plant_" + Iteration);
            //Reading the excel file for common operations like Login, Logout, Manage language, Navigation.
            LogPath = ConfigurationManager.AppSettings["LogOperation"];
            Thread.Sleep(max);
            UsertypeLogin = Usertype;
            dt.Clear();
            if (Usertype == "PlantLogin")
            {
                dt = ExcelKeywordTable.ExcelData(LogPath + "\\KeywordDrivenData.xlsx", "Login");
            }
            else if (Usertype == "PlantRegionalLogin")
            {
                dt = ExcelKeywordTable.ExcelData(LogPath + "\\KeywordDrivenData.xlsx", "RegionalLogin");
            }
            Thread.Sleep(max);
            BrowserWindow window = new BrowserWindow();
            window.Maximized = true;
            Thread.Sleep(max);
            Pop.OperationStart(dt.Rows[0]["Keyword"].ToString(), dt.Rows[0]["TypeOfControl"].ToString(), dt.Rows[0]["ControlKeyword"].ToString(), dt.Rows[0]["ControlKeywordValue"].ToString(), ExcelDataTable.ReadPlantLoginInfo(requiredPlant, dt.Rows[0]["DataRefferencekeyword"].ToString()), dt.Rows[0]["TypeOfWindow"].ToString(), dt.Rows[0]["Step No"].ToString(), dt.Rows[0]["Description"].ToString());
            Pop.OperationStart(dt.Rows[1]["Keyword"].ToString(), dt.Rows[1]["TypeOfControl"].ToString(), dt.Rows[1]["ControlKeyword"].ToString(), dt.Rows[1]["ControlKeywordValue"].ToString(), ExcelDataTable.ReadPlantLoginInfo(requiredPlant, dt.Rows[1]["DataRefferencekeyword"].ToString()), dt.Rows[1]["TypeOfWindow"].ToString(), dt.Rows[1]["Step No"].ToString(), dt.Rows[1]["Description"].ToString());
            Pop.OperationStart(dt.Rows[2]["Keyword"].ToString(), dt.Rows[2]["TypeOfControl"].ToString(), dt.Rows[2]["ControlKeyword"].ToString(), dt.Rows[2]["ControlKeywordValue"].ToString(), dt.Rows[2]["DataRefferencekeyword"].ToString(), dt.Rows[2]["TypeOfWindow"].ToString(), dt.Rows[2]["Step No"].ToString(), dt.Rows[2]["Description"].ToString());
            Thread.Sleep(max * 2);
            string Login = window.ExecuteScript("var data = document.getElementById('ContentPlaceHolder_lblErrorMessage');return data;").ToString();
            if (Login != string.Empty)
            {
                string LoginValidation = window.ExecuteScript("var data = document.getElementById('ContentPlaceHolder_lblErrorMessage').innerHTML;return data;").ToString();
                if (LoginValidation != "")
                {
                    Thread.Sleep(max);
                    Pop.OperationStart(dt.Rows[1]["Keyword"].ToString(), dt.Rows[1]["TypeOfControl"].ToString(), dt.Rows[1]["ControlKeyword"].ToString(), dt.Rows[1]["ControlKeywordValue"].ToString(), ExcelDataTable.ReadPlantLoginInfo(requiredPlant, dt.Rows[1]["DataRefferencekeyword"].ToString()), dt.Rows[1]["TypeOfWindow"].ToString(), dt.Rows[1]["Step No"].ToString(), dt.Rows[1]["Description"].ToString());
                    Thread.Sleep(mid);
                    Pop.OperationStart(dt.Rows[2]["Keyword"].ToString(), dt.Rows[2]["TypeOfControl"].ToString(), dt.Rows[2]["ControlKeyword"].ToString(), dt.Rows[2]["ControlKeywordValue"].ToString(), dt.Rows[2]["DataRefferencekeyword"].ToString(), dt.Rows[2]["TypeOfWindow"].ToString(), dt.Rows[2]["Step No"].ToString(), dt.Rows[2]["Description"].ToString());
                    Thread.Sleep(max * 2);
                    Login = window.ExecuteScript("var data = document.getElementById('ContentPlaceHolder_lblErrorMessage');return data;").ToString();
                    if (Login != string.Empty)
                    {
                        LoginValidation = window.ExecuteScript("var data = document.getElementById('ContentPlaceHolder_lblErrorMessage').innerHTML;return data;").ToString();
                        if (LoginValidation != lang.Msg_LogoutSuccessMessage)
                        {
                            string screenShotName = "Plant_LoginFailed";
                            Operation.ErrorScreenPath = screenShot(screenShotName);
                            Operation.FailerReason = requiredPlant + " Plant Login Failed";
                            genreport.FileCorreptionCheck();
                        }
                        Assert.AreEqual("", LoginValidation, requiredPlant + " Plant Login Failed");
                    }
                    else
                    {
                        Thread.Sleep(max);
                        auto.Send("{F5}");
                        Pop.OperationStart(dt.Rows[3]["Keyword"].ToString(), dt.Rows[3]["TypeOfControl"].ToString(), dt.Rows[3]["ControlKeyword"].ToString(), dt.Rows[3]["ControlKeywordValue"].ToString(), dt.Rows[3]["DataRefferencekeyword"].ToString(), dt.Rows[3]["TypeOfWindow"].ToString(), dt.Rows[3]["Step No"].ToString(), dt.Rows[3]["Description"].ToString());
                    }
                }
            }
            else
            {
                Thread.Sleep(max);
                auto.Send("{F5}");
                Pop.OperationStart(dt.Rows[3]["Keyword"].ToString(), dt.Rows[3]["TypeOfControl"].ToString(), dt.Rows[3]["ControlKeyword"].ToString(), dt.Rows[3]["ControlKeywordValue"].ToString(), dt.Rows[3]["DataRefferencekeyword"].ToString(), dt.Rows[3]["TypeOfWindow"].ToString(), dt.Rows[3]["Step No"].ToString(), dt.Rows[3]["Description"].ToString());
            }
        }
        #endregion

        /// <summary>
        /// This will perform the Login operations that specified in the Common keyword drivers. for login refer Sheet Name:"ManageLanguage"
        /// 
        /// This will change the language, if the current language and the specified language is not matching.
        /// </summary>
        /// <param name="languageCode">here this will accept the language code, for which language we are running.</param>
        public void ChangLanguage(string languageCode)
        {
            BrowserWindow window = new BrowserWindow();
            Thread.Sleep(LoginOperatrion.max);
            string compairLanguage = string.Empty;
            System.Globalization.CultureInfo culture = new System.Globalization.CultureInfo(languageCode);
            string name = culture.EnglishName;
            string language = name.Substring(0, name.IndexOf(" ("));
            Thread.Sleep(max);
            //Change Language.
            dt.Clear();
            dt = ExcelKeywordTable.ExcelData(LogPath + "\\KeywordDrivenData.xlsx", "ManageLanguage");
            Pop.OperationStart(dt.Rows[0]["Keyword"].ToString(), dt.Rows[0]["TypeOfControl"].ToString(), dt.Rows[0]["ControlKeyword"].ToString(), dt.Rows[0]["ControlKeywordValue"].ToString(), dt.Rows[0]["DataRefferencekeyword"].ToString(), dt.Rows[0]["TypeOfWindow"].ToString(), dt.Rows[0]["Step No"].ToString(), dt.Rows[0]["Description"].ToString());
            Pop.OperationStart(dt.Rows[1]["Keyword"].ToString(), dt.Rows[1]["TypeOfControl"].ToString(), dt.Rows[1]["ControlKeyword"].ToString(), dt.Rows[1]["ControlKeywordValue"].ToString(), dt.Rows[0]["DataRefferencekeyword"].ToString(), dt.Rows[1]["TypeOfWindow"].ToString(), dt.Rows[1]["Step No"].ToString(), dt.Rows[1]["Description"].ToString());
            Thread.Sleep(max * 3);
            string currentLang = window.ExecuteScript("var data = document.getElementById('ContentPlaceHolder1_drpLanguageList'); var text=data.options[data.selectedIndex].text; return text;").ToString();
            try
            {
                compairLanguage = currentLang.Substring(0, currentLang.IndexOf(" ("));
            }
            catch (Exception e)
            {
                compairLanguage = currentLang;
            }
            string code = getLangCode(currentLang);
            Thread.Sleep(max);

            if (language != compairLanguage)
            {
                Pop.OperationStart(dt.Rows[2]["Keyword"].ToString(), dt.Rows[2]["TypeOfControl"].ToString(), dt.Rows[2]["ControlKeyword"].ToString(), dt.Rows[2]["ControlKeywordValue"].ToString(), language, dt.Rows[2]["TypeOfWindow"].ToString(), dt.Rows[2]["Step No"].ToString(), dt.Rows[2]["Description"].ToString());
                Thread.Sleep(max);
                Pop.OperationStart(dt.Rows[3]["Keyword"].ToString(), dt.Rows[3]["TypeOfControl"].ToString(), dt.Rows[3]["ControlKeyword"].ToString(), dt.Rows[3]["ControlKeywordValue"].ToString(), dt.Rows[0]["DataRefferencekeyword"].ToString(), dt.Rows[3]["TypeOfWindow"].ToString(), dt.Rows[3]["Step No"].ToString(), dt.Rows[3]["Description"].ToString());
                Thread.Sleep(max);
                string UpdateSuccess = Pop.WebGetControlData(dt.Rows[4]["TypeOfControl"].ToString(), dt.Rows[4]["ControlKeyword"].ToString(), dt.Rows[4]["ControlKeywordValue"].ToString(), "LanguageUpdate_Error", "Language Not Updated", dt.Rows[4]["Step No"].ToString(), dt.Rows[4]["Description"].ToString(), dt.Rows[4]["Keyword"].ToString()).Trim();
                LogLanguageTemplete.messageResource(code);
                Thread.Sleep(mid);
                if (UpdateSuccess != lang.Msg_LanguageUpdated)
                {
                    string screenShotName = "LanguageUpdate_Error";
                    Operation.ErrorScreenPath = screenShot(screenShotName);
                    Operation.FailerReason = "Language Not Updated";
                }
                Assert.AreEqual(lang.Msg_LanguageUpdated, UpdateSuccess, "Language Not Updated");
                Thread.Sleep(max);
                Logout();
                Thread.Sleep(max * 2);
                login(UsertypeLogin);
            }
            Thread.Sleep(mid);
            auto.Send("{F5}");
            Thread.Sleep(max);
        }

        /// <summary>
        /// This will return the language name specified in the application, for each language code.
        /// </summary>
        /// <param name="lang">here this will accept the language code, for which language we are running.</param>
        /// <returns>This will return the language name specified in the application</returns>
        public string getLangCode(string lang)
        {
            string getLangCode = string.Empty;

            if (lang == "Chinese (中文)")
            {
                getLangCode = "zh-CN";
            }
            else if (lang == "English")
            {
                getLangCode = "en-US";
            }
            else if (lang == "Indonesian (bahasa indonesia)")
            {
                getLangCode = "id-ID";
            }
            else if (lang == "Japanese (日本語)")
            {
                getLangCode = "ja-JP";
            }
            else if (lang == "Korean (한국어)")
            {
                getLangCode = "ko-KR";
            }
            else if (lang == "Malay (malay)")
            {
                getLangCode = "ms-MY";
            }
            else if (lang == "Thai (ไทย)")
            {
                getLangCode = "th-TH";
            }
            else if (lang == "Vietnamese (việt nam)")
            {
                getLangCode = "vi-VN";
            }
            return getLangCode;
        }

        #region LinkNavigations
        /// <summary>
        /// This will refer the links specified in the excel sheet "LinksData" in the Common keyword drivers excel.
        /// </summary>
        /// <param name="keyword">This will accept the reference keyword to the link specified in the LinksData excel sheet.</param>
        public void ModuleNavigation(string keyword)
        {
            Thread.Sleep(mid);
            auto.Send("{F5}");
            Thread.Sleep(2000);
            ExcelKeywordTable.getKewordData(LogPath + "\\KeywordDrivenData.xlsx", "ModuleData");
            Thread.Sleep(mid);
            Pop.OperationStart(ExcelKeywordTable.ReadKeywordData(keyword, "Keyword"), ExcelKeywordTable.ReadKeywordData(keyword, "TypeOfControl"), ExcelKeywordTable.ReadKeywordData(keyword, "ControlKeyword"), ExcelKeywordTable.ReadKeywordData(keyword, "ControlKeywordValue"), ExcelKeywordTable.ReadKeywordData(keyword, "DataRefferencekeyword"), ExcelKeywordTable.ReadKeywordData(keyword, "TypeOfWindow"), ExcelKeywordTable.ReadKeywordData(keyword, "Step No"), ExcelKeywordTable.ReadKeywordData(keyword, "Description"));
            Thread.Sleep(max);
        }

        /// <summary>
        /// This will refer the links specified in the excel sheet "ModuleData" in the Common keyword drivers excel.
        /// </summary>
        /// <param name="keyword">This will accept the reference keyword to the link specified in the ModuleData excel sheet.</param>
        public void SubModuleNavigation(string keyword)
        {
            Thread.Sleep(mid);
            auto.Send("{F5}");
            Thread.Sleep(2000);
            ExcelKeywordTable.getKewordData(LogPath + "\\KeywordDrivenData.xlsx", "SubModuleData");
            Thread.Sleep(mid);
            auto.Send("{DOWN 2}");
            Thread.Sleep(mid);
            Pop.OperationStart(ExcelKeywordTable.ReadKeywordData(keyword, "Keyword"), ExcelKeywordTable.ReadKeywordData(keyword, "TypeOfControl"), ExcelKeywordTable.ReadKeywordData(keyword, "ControlKeyword"), ExcelKeywordTable.ReadKeywordData(keyword, "ControlKeywordValue"), ExcelKeywordTable.ReadKeywordData(keyword, "DataRefferencekeyword"), ExcelKeywordTable.ReadKeywordData(keyword, "TypeOfWindow"), ExcelKeywordTable.ReadKeywordData(keyword, "Step No"), ExcelKeywordTable.ReadKeywordData(keyword, "Description"));
            Thread.Sleep(max);
        }

        /// <summary>
        /// This will refer the links specified in the excel sheet "SubModuleData" in the Common keyword drivers excel.
        /// </summary>
        /// <param name="keyword">This will accept the reference keyword to the link specified in the SubModuleData excel sheet.</param>
        public void LinkNavigation(string keyword)
        {
            Thread.Sleep(mid);
            auto.Send("{F5}");
            Thread.Sleep(2000);
            ExcelKeywordTable.getKewordData(LogPath + "\\KeywordDrivenData.xlsx", "LinksData");
            Thread.Sleep(mid);
            auto.Send("{DOWN 2}");
            Thread.Sleep(mid);
            Pop.OperationStart(ExcelKeywordTable.ReadKeywordData(keyword, "Keyword"), ExcelKeywordTable.ReadKeywordData(keyword, "TypeOfControl"), ExcelKeywordTable.ReadKeywordData(keyword, "ControlKeyword"), ExcelKeywordTable.ReadKeywordData(keyword, "ControlKeywordValue"), ExcelKeywordTable.ReadKeywordData(keyword, "DataRefferencekeyword"), ExcelKeywordTable.ReadKeywordData(keyword, "TypeOfWindow"), ExcelKeywordTable.ReadKeywordData(keyword, "Step No"), ExcelKeywordTable.ReadKeywordData(keyword, "Description"));
            Thread.Sleep(max);
        }
        #endregion

        #region validateLog
        /// <summary>
        /// This will validate the log file downloaded after uploading any data
        /// </summary>
        /// <param name="validMsg">Here you want to specify the message to validate.</param>
        /// <returns>This will return the true or false value.</returns>
        public bool validateLOGcheckFail(string validMsg)
        {
            string fileName = string.Empty;
            try
            {
                Thread.Sleep(max * 3);
                Process[] notepads = Process.GetProcessesByName("notepad");
                Thread.Sleep(max * 3);
                fileName = "\\" + notepads[0].MainWindowTitle;
                //Using below code we can access special folders in windows like documents, favorites, or library folders 
                fileName = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + fileName;
                //replacing text from the path that fetched 
                fileName = fileName.Replace(" - Notepad", ".txt");
                fileName = fileName.Replace("Documents", "Downloads");
                FileStream file = new FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.Read);
                bool b1 = false;
                string[] lines = File.ReadAllLines(fileName);
                string[] Validlines = validMsg.Split(',');
                for (int i = 1; i < lines.Length; i++)
                {
                    if (Validlines.Length == 2)
                    {
                        if ((lines[i - 1].Equals(Validlines[0]) || lines[i].Equals(Validlines[1])))
                        {
                            b1 = true;
                        }
                    }
                    else if (Validlines.Length == 1)
                    {
                        if ((lines[i - 1].Equals(Validlines[0])))
                        {
                            b1 = true;
                        }
                        else if ((lines[i - 1].Equals(Validlines[0].Replace('0', '1'))))
                        {
                            b1 = false;
                            break;
                        }
                        else if ((lines[i - 1].Equals(Validlines[0].Replace('1', '0'))))
                        {
                            b1 = false;
                            break;
                        }
                        else if ((lines[i - 1].Contains(Validlines[0])))
                        {
                            b1 = true;
                        }
                    }
                }
                file.Close();
                return b1;
            }
            catch (FileNotFoundException e)
            {
                Operation.FailerReason = fileName + " not found in the specified location";
                Assert.Fail(fileName + " not found in the specified location");
                return false;
            }
        }
        #endregion

        #region PageNavigation
        /// <summary>
        /// This will navigate through the pages, and validate data is displaying or not.
        /// </summary>
        /// <param name="PageID">Here this will accept the table footer id value of the DIV tag that contains the pages</param>
        /// <param name="step">Step number</param>
        /// <param name="description">Description for the step number</param>
        public void NavigationPage(string PageID, string step, string description)
        {
            BrowserWindow window = new BrowserWindow();
            string lastPage = string.Empty;
            int count = 0;
            int tablerow = 0;
            int tableCol = 0;
            try
            {
                try
                {
                    count = Convert.ToInt16(window.ExecuteScript("var data=document.getElementById('" + PageID + "').getElementsByTagName('a');return  data.length"));
                }
                catch (Exception e) { }
                string idval = PageID.Split('_')[0];
                if (count > 3)
                {
                    if (count > 5)
                    {
                        lastPage = idval + "_hrefLast";
                        count = 3;  
                    }
                    else
                    {
                        count = 3;
                    }
                }
                for (int i = 2; i <= count; i++)
                {
                    string id = idval + "_hrefNum" + i;
                    window.ExecuteScript("document.getElementById('" + id + "').scrollIntoView(true);");
                    Pop.OperationStart("Click", "HtmlHyperlink", "ID", id, "", "Web", step, description + i);
                    Thread.Sleep(max);
                    Thread.Sleep(max);
                    int tbodyCount = Convert.ToInt16(window.ExecuteScript("var tablrow=document.getElementsByTagName('tbody').length; return tablrow"));
                    if(tbodyCount == 2)
                    {
                        tablerow = Convert.ToInt16(window.ExecuteScript("var tablrow=document.getElementsByTagName('tbody')[1].getElementsByTagName('tr').length; return tablrow"));
                        tableCol = Convert.ToInt16(window.ExecuteScript("var tablCol=document.getElementsByTagName('tbody')[1].getElementsByTagName('td').length; return tablCol"));
                    }
                    else
                    {
                        tablerow = Convert.ToInt16(window.ExecuteScript("var tablrow=document.getElementsByTagName('tbody')[0].getElementsByTagName('tr').length; return tablrow"));
                        tableCol = Convert.ToInt16(window.ExecuteScript("var tablCol=document.getElementsByTagName('tbody')[0].getElementsByTagName('td').length; return tablCol"));
                    }
                    if (tablerow >= 1 && tableCol > 1)
                    {
                        genreport.Reports(step, "Data Loaded in page " + i + " successfully", "ValidatingPageData", true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                    }
                    else
                    {
                        string path = ConfigurationManager.AppSettings["ScreenShot"];
                        path = path + @"\" + LoginOperatrion.ProjectName + "";
                        Directory.CreateDirectory(path);
                        try
                        {
                            Image image = window.CaptureImage();
                            image.Save(path + "\\" + "PageLodingNotSuccess.jpeg", ImageFormat.Jpeg);
                            Operation.ErrorScreenPath = path + "\\" + "PageLodingNotSuccess.jpeg";
                            image.Dispose();
                        }
                        catch (Exception v) { }
                        genreport.Reports(step, "Data is not Loaded in page " + i + " successfully", "ValidatingPageData", false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                        Assert.Fail("Page data is not loaded successfully");
                    }
                }
                if (lastPage != string.Empty)
                {
                    window.ExecuteScript("document.getElementById('" + lastPage + "').scrollIntoView(true);");
                    Pop.OperationStart("Click", "HtmlHyperlink", "ID", lastPage, "", "Web", step, description + " Last Page");
                    Thread.Sleep(max);
                    int tablerowL = Convert.ToInt16(window.ExecuteScript("var tablrow=document.getElementsByTagName('tbody')[0].getElementsByTagName('tr').length; return tablrow"));
                    int tableColL = Convert.ToInt16(window.ExecuteScript("var tablCol=document.getElementsByTagName('tbody')[0].getElementsByTagName('td').length; return tablCol"));
                    if (tablerowL >= 1 && tableColL > 1)
                    {
                        genreport.Reports(step, "Data Loaded in Last page successfully", "ValidatingPageData", true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                    }
                    else
                    {
                        string path = ConfigurationManager.AppSettings["ScreenShot"];
                        path = path + @"\" + LoginOperatrion.ProjectName + "";
                        Directory.CreateDirectory(path);
                        try
                        {
                            Image image = window.CaptureImage();
                            image.Save(path + "\\" + "LastPageLodingNotSuccess.jpeg", ImageFormat.Jpeg);
                            Operation.ErrorScreenPath = path + "\\" + "LastPageLodingNotSuccess.jpeg";
                            image.Dispose();
                        }
                        catch (Exception z) { }
                        genreport.Reports(step, "Data is not Loaded in Last page successfully", "ValidatingPageData", false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                        Assert.Fail("Page data is not loaded successfully");
                    }
                }
            }
            catch (UITestControlNotFoundException e)
            {
                tablerow = Convert.ToInt16(window.ExecuteScript("var tablrow=document.getElementsByTagName('tbody')[0].getElementsByTagName('tr').length; return tablrow"));
                tableCol = Convert.ToInt16(window.ExecuteScript("var tablCol=document.getElementsByTagName('tbody')[0].getElementsByTagName('td').length; return tablCol"));
                if (tablerow >= 1 && tableCol > 1)
                {
                    genreport.Reports(step, "Data Loaded in page successfully", "ValidatingPageData", true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                }
                else
                {
                    string path = ConfigurationManager.AppSettings["ScreenShot"];
                    path = path + @"\" + LoginOperatrion.ProjectName + "";
                    Directory.CreateDirectory(path);
                    try
                    {
                        Image image = window.CaptureImage();
                        image.Save(path + "\\" + "PageLodingNotSuccess.jpeg", ImageFormat.Jpeg);
                        Operation.ErrorScreenPath = path + "\\" + "PageLodingNotSuccess.jpeg";
                        image.Dispose();
                    }
                    catch (Exception v) { }
                    genreport.Reports(step, "Data is not Loaded in page successfully", "ValidatingPageData", false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                    Assert.Fail("Page data is not loaded successfully");
                }
            }
        }
        #endregion

        #region BaseFunctions

        /// <summary>
        /// This will close the opened file.
        /// </summary>
        public void CloseFile()
        {
            auto.Send("!{SPACE}");
            Thread.Sleep(min);
            auto.Send("{C}");
            Thread.Sleep(max);
        }

        /// <summary>
        /// This will take the screen shot and saved in the specified path.
        /// </summary>
        /// <param name="imgName">Name of the image will pass here.</param>
        /// <returns>This will return the path of the image file.</returns>
        public string screenShot(string imgName)
        {
            BrowserWindow window = new BrowserWindow();
            //Read the path from the App.config file to save the screen short.
            string path = ConfigurationManager.AppSettings["ScreenShot"];
            string Imagepath = string.Empty;
            path = path + @"\" + LoginOperatrion.ProjectName + "";
            Directory.CreateDirectory(path);
            try
            {
                window.SearchProperties[UITestControl.PropertyNames.ClassName] = BrowserWindow.CurrentBrowser.ToString();
                Image image = window.CaptureImage();

                if (requiredPlant != string.Empty)
                {
                    image.Save(path + "\\" + imgName + "_" + requiredPlant + ".jpeg", ImageFormat.Jpeg);
                    Imagepath = path + "\\" + imgName + "_" + requiredPlant + ".jpeg";
                    image.Dispose();
                }
                else
                {
                    image.Save(path + "\\" + imgName + ".jpeg", ImageFormat.Jpeg);
                    Imagepath = path + "\\" + imgName + ".jpeg";
                    image.Dispose();
                }
            }
            catch (Exception e) { }
            Console.WriteLine("Screen Shot is Available in " + path + "\t" + "Folder");
            return Imagepath;
        }

        /// <summary>
        /// This is used to click any button in the keyboard. and specify the which button you want to click.
        /// </summary>
        /// <param name="Key">representation of the button you want to click, for that refer "https://www.autoitscript.com/autoit3/docs/appendix/SendKeys.htm"</param>
        public void SendKey(string Key)
        {
            if (Key.Contains("Alt"))
            {
                auto.Send("!{" + Key.Split('+')[1] + "}");
            }
            else if (Key.Contains("Ctrl"))
            {
                auto.Send("^{" + Key.Split('+')[1] + "}");
            }
            else
            {
                auto.Send("{" + Key + "}");
            }
        }
        #endregion

        /// <summary>
        /// Send Text To a Control that is dynamic
        /// </summary>
        /// <param name="text">Test to be entered in the control</param>
        public void Sendtext(string text)
        {
            auto.Send(text,0);
        }

        /// <summary>
        /// This will get the current system date in the specific format specified in the "GlobalElements" excel 
        /// </summary>
        /// <param name="dateformat">date for mat of the application.</param>
        /// <returns>return string value in a specified date format</returns>
        public string getSystemDate(string dateformat)
        {
            try
            {
                var date = System.DateTime.Today;
                string systemdate = date.ToString(dateformat);
                if (dateformat.Contains('/'))
                {
                    systemdate = systemdate.Replace('-', '/');
                }
                else if (dateformat.Contains('-'))
                {
                    systemdate = systemdate.Replace('-', '-');
                }
                return systemdate;
            }
            catch(Exception e)
            {
                return string.Empty;
            }
            
        }

        public string getDate(string dateformat, string DayCount)
        {
            try
            {
                string[] dateDetails = DayCount.Split('|');
                var date = System.DateTime.Today;
                double days = 0;
                int extradayCount = 0;
                WriteAndReadData.DataFilePath = ConfigurationManager.AppSettings["ExcelDataFile"] + "\\DynamicData" + "\\" + batchforReport + ".xls";
                ExcelDataTable.PopulateRecordData(WriteAndReadData.DataFilePath);

                if (dateDetails[0] == "ADD")
                {
                    if (dateDetails[1].Contains('='))
                    {
                        days = Convert.ToDouble(System.Text.RegularExpressions.Regex.Replace(Pop.getDataFromDynamicExcel(dateDetails[1]), "[^0-9]+", string.Empty));
                    }
                    else
                    {
                        days = Convert.ToDouble(ExcelDataTable.ReadData(1, dateDetails[1]));
                    }
                }
                else if (dateDetails[0] == "SUB")
                {
                    if (dateDetails[1].Contains('='))
                    {
                        days = Convert.ToDouble(System.Text.RegularExpressions.Regex.Replace(Pop.getDataFromDynamicExcel(dateDetails[1]), "[^0-9]+", string.Empty));
                        days = days * -1;
                    }
                    else
                    {
                        days = Convert.ToDouble(ExcelDataTable.ReadData(1, dateDetails[1])) * -1;
                    }
                }
                else
                {
                    throw new Exception("Specified Date Operation is not proper");
                }
                if (dateDetails.Count() > 2)
                {
                    extradayCount = Convert.ToInt32(dateDetails[2]);
                    if (dateDetails[0] == "ADD")
                    {
                        date = date.AddDays(days+ extradayCount);
                    }
                    else if (dateDetails[0] == "SUB")
                    {
                        date = date.AddDays(days - extradayCount);
                    }
                }
                else
                {
                    date = date.AddDays(days);
                }

                string systemdate = date.ToString(dateformat);
                if (dateformat.Contains('/'))
                {
                    systemdate = systemdate.Replace('-', '/');
                }
                else if (dateformat.Contains('-'))
                {
                    systemdate = systemdate.Replace('-', '-');
                }
                return systemdate;
            }
            catch(Exception e)
            {
                return string.Empty;
            }
        }

        #region ValidateCalenderStatus
        /// <summary>
        /// validating calender status in the GR Purchase Order Calender.
        /// </summary>
        /// <param name="controlkeyValue">Here specify the id value of the table that displayed the PO's</param>
        /// <param name="steps">Step number</param>
        /// <param name="description">Description for the step number</param>
        /// <param name="typeofoperation">Here specify the keyword for validating the Order Status. Keyword: "ValidatePOStatus"</param>
        /// <returns>This will return the image URL of the status for validation</returns>
        public string ValidateCalenderStatus(string controlkeyValue, string steps, string description, string typeofoperation)
        {
            try
            {
                BrowserWindow window = new BrowserWindow();
                var tableRowNo = window.ExecuteScript(@"var data = myFunction();
                                             function myFunction() {
                                             var rowNo;
                                             var po=document.getElementById('txtPurchaseOrder').value;"
                                                 + "var a = parseInt(po);"
                                                 + "var table = document.getElementById('" + controlkeyValue + "');"
                                                 + "var tableRows = table.rows.length;"
                                                 + "for (i = 0; i < tableRows; i++){ "
                                                 + "var rowCells = table.rows.item(i).cells;"
                                                 + "for(var j = 0; j < rowCells.length; j++){"
                                                 + "if(a == rowCells.item(j).innerText){"
                                                 + "rowNo=i;"
                                                        + " }"
                                                        + "} "
                                                       + "}"
                                                + "rowNo=rowNo+1;"
                                                + "return rowNo;"
                                                      + "} return data; ");
                string rowNo = tableRowNo.ToString().Trim();
                string imgID = "'" + "dayViewRowControl_" + rowNo + "_statusImage" + "'";
                string imgSrc = window.ExecuteScript(@"var data=document.getElementById(" + imgID + ").src; return data; ").ToString();
                if (imgSrc != null)
                {
                    return imgSrc;
                }
                else
                {
                    return null;
                }
            }
            catch (Exception e)
            {
                Operation.FailerReason = "Could not find control with specified value Please check the Page is loaded.";
                genreport.Reports(steps, description, typeofoperation, false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                genreport.FileCorreptionCheck();
                throw new NoSuchControlTypeFound(controlkeyValue + " control can not find Order check the Page is loaded or not");
            }

        }
        #endregion

        #region ValidateGRStatus
        /// <summary>
        /// In GR Module, Goods Receipt page after receiving the goods this will check for status of the order is changed or not.
        /// </summary>
        /// <param name="controlkeyValue">Specify the id value of the table in the GR Page</param>
        /// <param name="steps">Step number</param>
        /// <param name="description">Description for the step number</param>
        /// <param name="typeofoperation">Here specify the keyword for validating the GR Status. Keyword: "ValidateGRStatus"</param>
        /// <returns>Here this will return the status text of specified order and validate with the status in the resource file</returns>
        public string ValidateGRStatus(string controlkeyValue, string steps, string description, string typeofoperation)
        {
            BrowserWindow window = new BrowserWindow();
            try
            {
                var statusGR = window.ExecuteScript(@"var data= myFunction();
                function myFunction()
                    {
                        var a = parseInt('" + controlkeyValue + "');"
                        + "var table = document.getElementById('processingOrders');"
                        + "var rowNo;var tableRows = table.rows.length;"
                        + "for (i = 0; i < tableRows; i++)"
                        + "{"
                        + "var rowCells = table.rows.item(i).cells;"
                            + "for (var j = 0; j < rowCells.length; j++)"
                            + "{"
                                + "if (a == rowCells.item(1).innerText)"
                                + "{"
                                    + "rowNo = rowCells.item(4).innerText;"
                                    + "return rowNo;"
                                + "}"
                            + "}"
                        + "}"
                        + "} return data;"
                        );
                string status = statusGR.ToString().Trim();
                return status;
            }
            catch
            {
                Operation.FailerReason = "Could not find control with specified ID value Please check the Control Keyword Value.";
                genreport.Reports(steps, description, typeofoperation, false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                genreport.FileCorreptionCheck();
                throw new NoSuchControlTypeFound(controlkeyValue + " control can not find Please check the Control Keyword Value.");
            }


        }
        #endregion

        #region getProposedGate
        /// <summary>
        /// At the time of adding truck it will validate the proposed gate is available or not. if not available it will fail.
        /// </summary>
        /// <param name="controlkeyValue">Specify the id value of the table that showing the gates</param>
        /// <param name="DataReffkey">Here passing the class name of the proposed gate, based on that it will validate</param>
        /// <param name="steps">Step number</param>
        /// <param name="description">Description for the step number</param>
        /// <param name="typeofoperation">Here specify the keyword for validating the Proposed gate is available Keyword: "assignPropsedGate"</param>
        /// <returns>This will validate gate are available, if not fail, other wise return the select button id value of the 1st proposed gate.</returns>
        public string getProposedGate(string controlkeyValue, string DataReffkey, string steps, string description, string typeofoperation)
        {
            BrowserWindow window = new BrowserWindow();
            try
            {
                var getGateID = window.ExecuteScript(@"var data = myFunction();
                                  function myFunction(){
                                    var a='" + DataReffkey + "';"
                                    + "var idvalue;"
                                    + "var divs = document.getElementById('" + controlkeyValue + "');"
                                    + "var NoDiv = divs.childElementCount;"
                                    + "getID:{"
                                    + "for (var i = 0; i < NoDiv; i++){"
                                    + "ulist = divs.getElementsByTagName('ul')(i);"
                                    + "lil = ulist.getElementsByTagName('li').length;"
                                    + "for (var j = 0; j < lil; j++){"
                                    + "if(a == ulist.getElementsByTagName('li')[0].getAttribute('class')){"
                                    + "idvalue = ulist.getElementsByTagName('li')[0].getElementsByTagName('a')[0].getAttribute('id');"
                                    + "return idvalue;"
                                    + "break getID;"
                                    + "}}}}}return data;");
                if (getGateID == null)
                {
                    return string.Empty;
                }
                else
                {
                    window.ExecuteScript("document.getElementById('" + getGateID + "').scrollIntoView(true);");
                    string gate = getGateID.ToString();
                    return gate;
                }
            }
            catch (Exception e)
            {
                Operation.FailerReason = "Could not find control with specified ID value and reference value Please check the Control Keyword Value.";
                genreport.Reports(steps, description, typeofoperation, false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                genreport.FileCorreptionCheck();
                throw new NoSuchControlTypeFound(controlkeyValue + " control can not find Please check the Control Keyword Value.");
            }
        }
        #endregion

        #region ValidateTruckDetails
        /// <summary>
        /// After adding the truck this will validate the all the details are available, or not of not available this will fail.
        /// </summary>
        /// <param name="controlkeyValue">Specify the id value of the table that showing truck details</param>
        /// <param name="DataReffkey">here you want to pass the token number that recorded at the time of adding truck</param>
        /// <param name="steps">Step number</param>
        /// <param name="description">Description for the step number</param>
        /// <param name="typeofoperation">Here specify the keyword for validating the Proposed gate is available Keyword: "TruckDetailsValidate"</param>
        /// <returns>This will return the GateNumber(Selected)+Status of the Details available or not+Truck Entry Number</returns>
        public string ValidateTruckDetails(string controlkeyValue, string DataReffkey, string steps, string description, string typeofoperation)
        {
            BrowserWindow truckstatus = new BrowserWindow();
            string[] resource = DataReffkey.Split('|');
            string token= Pop.getDataFromDynamicExcel(resource[0]);
            try
            {
                var truckDetails = truckstatus.ExecuteScript(@"var data= myFunction();
                                function myFunction(){"
                                + "var tockenNo = '" + token + "';"
                                + "var trucktable = document.getElementById('" + controlkeyValue + "');"
                                + "var statuss;var TtableRows = trucktable.rows.length;"
                                + "for (i = 0; i < TtableRows; i++){"
                                + "var rowCells = trucktable.rows.item(i).cells;"
                                + "for (var j = 0; j < rowCells.length; j++){"
                                + "if (tockenNo == rowCells.item(0).innerText){"
                                + "if(rowCells.item(3).innerText=='" + ExcelDataTable.ReadData(1, resource[2]) + "' && rowCells.item(2).innerText=='" + ExcelDataTable.ReadData(1, resource[1]) + "' && rowCells.item(5).innerText=='' && rowCells.item(6).innerText!='' || rowCells.item(7).innerText!=''){"
                                + "statuss=rowCells.item(5).innerText + ':' +'true' + ':'+ rowCells.item(0).innerText;}else{statuss= 'false';}}"
                                + "}}return statuss; return }return data;");
                return truckDetails.ToString();
            }
            catch (Exception e)
            {
                Operation.FailerReason = "Could not find control with specified ID value and reference value Please check the Control Keyword Value.";
                genreport.Reports(steps, description, typeofoperation, false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                genreport.FileCorreptionCheck();
                throw new NoSuchControlTypeFound(controlkeyValue + " control can not find Please check the Control Keyword Value.");
            }
        }
        #endregion

        #region EnterQualityResult&Save
        /// <summary>
        /// First search for the sample with material and batch. then use to enter quality result in Quality page and save
        /// </summary>
        /// <param name="controlkeyValue">Specify the id value of the table that showing Quality Sample Page</param>
        /// <param name="DataReffkey">Reference to the result to entered</param>
        /// <param name="steps">Step number</param>
        /// <param name="description">Description for the step number</param>
        /// <param name="typeofoperation">Here specify the keyword for validating the Proposed gate is available Keyword: "SaveQualityResult"</param>
        /// <returns>This will return the true or false value</returns>
        public bool EnterResultForQuality(string controlkeyValue, string DataReffkey, string steps, string description, string typeofoperation)
        {
            BrowserWindow resultEnter = new BrowserWindow();
            try
            {
                resultEnter.ExecuteScript(@"enterResult();
                                function enterResult(){"
                                + "var data = document.getElementById('" + controlkeyValue + "');"
                                + "var rows = data.rows.item(0).cells;"
                                + "var sampleno = rows.item(2).innerText.trim();"
                                + "var resultID = 'sampleInfoControl_' + sampleno + '_ctl00_txtResult';"
                                + "var result='" + DataReffkey + "';"
                                + "var avail = document.getElementById(resultID);"
                                + "if (avail){"
                                + "document.getElementById(resultID).value=result;"
                                + "var SaveresultID = 'sampleInfoControl_' + sampleno + '_passButton';"
                                + "document.getElementById(SaveresultID).click();}else{"
                                + "var SaveresultID = 'sampleInfoControl_' + sampleno + '_passButton';"
                                + "document.getElementById(SaveresultID).click();}}");
                return true;
            }
            catch (Exception e)
            {
                Operation.FailerReason = "Could not find control with specified ID value and reference value Please check the Control Keyword Value.";
                genreport.Reports(steps, description, typeofoperation, false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                genreport.FileCorreptionCheck();
                throw new NoSuchControlTypeFound(controlkeyValue + " control can not find Please check the Control Keyword Value.");
            }
        }
        #endregion

        #region CheckQualityResult
        /// <summary>
        /// This will validate searched quality details passed or not
        /// </summary>
        /// <param name="controlkeyValue">Specify the id value of the table that showing Quality Sample History Page</param>
        /// <param name="steps">Step number</param>
        /// <param name="description">Description for the step number</param>
        /// <param name="typeofoperation">Here specify the keyword for validating the Proposed gate is available Keyword: "CheckQualityResult"</param>
        /// <returns>This will return the true or false value</returns>
        public string CheckQualityResult(string controlkeyValue, string steps, string description, string typeofoperation)
        {
            BrowserWindow orderStatus = new BrowserWindow();
            try
            {
                string orderStatusData = orderStatus.ExecuteScript(@"var data = enterResult();
                                function enterResult(){"
                                + "var data = document.getElementById('" + controlkeyValue + "');"
                                + "var rows = data.rows.item(0).cells;"
                                + "var sampleno = rows.item(2).innerText.trim();"
                                + "var statusID = 'sampleInfoControl_' + sampleno + '_spnStatus';"
                                + "var actualstatus = document.getElementById(statusID).innerHTML;return actualstatus;}return data;").ToString();
                return orderStatusData;
            }
            catch (Exception e)
            {
                Operation.FailerReason = "Could not find control with specified ID value and reference value Please check the Control Keyword Value.";
                genreport.Reports(steps, description, typeofoperation, false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                genreport.FileCorreptionCheck();
                throw new NoSuchControlTypeFound(controlkeyValue + " control can not find Please check the Control Keyword Value.");
            }
        }
        #endregion

        #region ValidateSamplingInstruction
        /// <summary>
        /// validating sampling tools are available in searched material in Sampling instructions page in configure receiving area
        /// </summary>
        /// <param name="controlkeyValue">Specify the id value of the table that showing Sampling instructions</param>
        /// <param name="steps">Step number</param>
        /// <param name="description">Description for the step number</param>
        /// <param name="typeofoperation">Here specify the keyword for validating the Proposed gate is available Keyword: "SamplingDataCheck"</param>
        /// <returns>This will return number of sampling tools available. if equal to 0 or less than 0 it will  fail.</returns>
        public string ValidateSamplingInstruction(string controlkeyValue, string steps, string description, string typeofoperation)
        {
            BrowserWindow orderStatus = new BrowserWindow();
            try
            {
                string SamplingInfo = orderStatus.ExecuteScript(@"var data = VAlidate();
                                function VAlidate(){"
                                + "var data = document.getElementById('" + controlkeyValue + "');"
                                + "var rows = data.rows.item(0).cells;var statuss;"
                                + "var sampleno = rows.item(1).innerText.trim();"
                                + "var InstructionID = 'ucSampleInfoSummary' + sampleno  + '_spnSamplingInstructions';"
                                + "var ToolId='ucSampleInfoSummary' + sampleno + '_divEquipmentImages';"
                                + "var count=document.getElementById(ToolId).childElementCount;"
                                + "var instruction=document.getElementById(InstructionID).innerHTML;"
                                + "if(instruction=='' || count==0){statuss='false';}else{statuss='true'}"
                                + "return statuss;}return data;").ToString();
                return SamplingInfo;
            }
            catch (Exception e)
            {
                Operation.FailerReason = "Could not find control with specified ID value and reference value Please check the Control Keyword Value.";
                genreport.Reports(steps, description, typeofoperation, false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                genreport.FileCorreptionCheck();
                throw new NoSuchControlTypeFound(controlkeyValue + " control can not find Please check the Control Keyword Value.");
            }
        }
        #endregion

        #region Validate JIT in Additional Material Info
        /// <summary>
        /// This will validate searched material is maintained as JIT in Additional Material Information page in configure receiving area. if not return "false"
        /// </summary>
        /// <param name="controlkeyValue">Specify the id value of the table that showing Sampling instructions</param>
        /// <param name="steps">Step number</param>
        /// <param name="description">Description for the step number</param>
        /// <param name="typeofoperation">Here specify the keyword for validating the Proposed gate is available Keyword: "ValidateIsJIT"</param>
        /// <returns>Return true or false</returns>
        public string validateIsJit(string controlkeyValue, string steps, string description, string typeofoperation)
        {
            BrowserWindow Jit = new BrowserWindow();
            try
            {
                string isJit = Jit.ExecuteScript(@"var data = Validate();
                                function Validate(){"
                                + "var data = document.getElementById('ContentPlaceHolder1_tbSampleMaster');"
                                + "var rows = data.rows.item(0).cells;var statuss;"
                                + "var MetrialNo = rows.item(1).innerText.trim();"
                                + "var id = 'ucMaterialMasterSummary' + MetrialNo + '_tdISJit';"
                                + "statuss = document.getElementById(id).getAttribute('isjit');return statuss;}return data;").ToString();
                return isJit;
            }
            catch (Exception e)
            {
                Operation.FailerReason = "Could not find control with specified ID value and reference value Please check the Control Keyword Value.";
                genreport.Reports(steps, description, typeofoperation, false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                genreport.FileCorreptionCheck();
                throw new NoSuchControlTypeFound(controlkeyValue + " control can not find Please check the Control Keyword Value.");
            }
        }
        #endregion

        #region  Validate Material Available in GR Mobile
        /// <summary>
        /// In GR Page in Mobile, validating the given material is available or not.
        /// </summary>
        /// <param name="controlkeyValue">Specify the ID value of the table that showing Material Details</param>
        /// <param name="DataReffkey">Reference to the material number</param>
        /// <param name="steps">Step number</param>
        /// <param name="description">Description for the step number</param>
        /// <param name="typeofoperation">Here specify the keyword for validating the Proposed gate is available Keyword: "ValidateAvailMetrial"</param>
        /// <returns>Return true or false value</returns>
        public string ValidateMetrialAvailable(string controlkeyValue, string DataReffkey, string steps, string description, string typeofoperation)
        {
            BrowserWindow metrialStatus = new BrowserWindow();
            try
            {
                string metrialavail = metrialStatus.ExecuteScript(@"var data = myFunction();
                                    function myFunction() {"
                                    + "var satatus;var metrialno = '" + DataReffkey + "';"
                                    + "var metrial = document.getElementById('" + controlkeyValue + "');"
                                    + "var tableRows = metrial.rows.length;"
                                    + "for (i = 0; i < tableRows; i++){"
                                    + "var rowCells = metrial.rows.item(i).cells;"
                                    + "for (var j = 0; j < rowCells.length; j++){"
                                    + "if (metrialno == rowCells.item(2).innerText.trim()){satatus='true';}"
                                    + "else{satatus='false';}}}return satatus;}return data;").ToString();
                return metrialavail;
            }
            catch (Exception e)
            {
                Operation.FailerReason = "Could not find control with specified ID value and reference value Please check the Control Keyword Value.";
                genreport.Reports(steps, description, typeofoperation, false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                genreport.FileCorreptionCheck();
                throw new NoSuchControlTypeFound(controlkeyValue + " control can not find Please check the Control Keyword Value.");
            }
        }
        #endregion

        #region DownloadUpload
        /// <summary>
        /// this will help to toggle Download and Upload in page
        /// </summary>
        /// <param name="steps">Step number</param>
        /// <param name="description">Description for the step number</param>
        /// <param name="typeofoperation">Here specify the keyword to toggle Download and Upload Keyword: "DownloadUpload"</param>
        public void ClickDownloadUpload(string steps, string description, string typeofoperation)
        {
            try
            {
                BrowserWindow download = new BrowserWindow();
                download.ExecuteScript(@"var data =  myFunction();
                          function myFunction() {"
                          + "var data = 'Download & Upload';"
                          + "var count = document.getElementsByClassName('search-legend cursor-pointer').length;"
                          + "for(i = 0; i < count; i++){"
                          + "var dataname = document.getElementsByClassName('search-legend cursor-pointer')[0].innerText.trim();"
                          + "if(dataname === data){"
                          + "document.getElementsByTagName('legend')[i].click();break;}}}");
            }
            catch (Exception e)
            {
                Operation.FailerReason = "There is no Download and upload to navigate";
                genreport.Reports(steps, description, typeofoperation, false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                genreport.FileCorreptionCheck();
                throw new NoSuchControlTypeFound("There is no Download and upload to navigate");
            }
        }
        #endregion

        #region CustInstructions
        /// <summary>
        /// This will validate the 
        /// 1. Picking Instruction
        /// 2. Shipping Instruction
        /// 3. Packing Instruction
        /// you must search for the data first.
        /// </summary>
        /// <param name="controlkeyValue">Specify the ID value of the search box.</param>
        /// <param name="DataReffkey">specify what you are validating and Reference to the particular instruction in test data separated by ":"</param>
        /// <param name="steps">Step number</param>
        /// <param name="description">Description for the step number</param>
        /// <param name="typeofoperation">Here specify the keyword for validating the data given for instructions Keyword: "ValidateCustomerInst"</param>
        /// <returns>Return true or false value.</returns>
        public string ValidateCustomerInst(string controlkeyValue, string DataReffkey, string steps, string description, string typeofoperation)
        {
            string[] compareDate = DataReffkey.Split(':');
            string de = ExcelDataTable.ReadData(1, compareDate[1]);
            try
            {
                BrowserWindow window = new BrowserWindow();
                string data = window.ExecuteScript(@"var data = myFunction();
                                  function myFunction(){"
                                  + "var po = document.getElementById('" + controlkeyValue + "').value;"
                                  + "var vdata = '" + compareDate[0] + "';"
                                  + "if(vdata == 'PickingInstruction'){"
                                  + "var instruction = document.getElementById('uCustomerInstructionSummary_'+ po +'_spnPickingContent').innerHTML;}"
                                  + "if(vdata == 'ShippingInstruction'){"
                                  + "var instruction = document.getElementById('uCustomerInstructionSummary_'+ po +'_spnShippingContent').innerHTML;}"
                                  + "if(vdata == 'PackingInstruction'){"
                                  + "var instruction = document.getElementById('uCustomerInstructionSummary_'+ po +'_spnPackingContent').innerHTML;}"
                                  + "if(instruction == '" + de + "')"
                                  + "{var status= 'true';}else{var status= 'false';}return status;}return data;").ToString();
                return data;
            }
            catch (Exception e)
            {
                Operation.FailerReason = "Could not find control with specified ID value and reference value Please check the Control Keyword Value.";
                genreport.Reports(steps, description, typeofoperation, false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                genreport.FileCorreptionCheck();
                throw new NoSuchControlTypeFound(controlkeyValue + " control can not find Please check the Control Keyword Value.");
            }
        }
        #endregion

        /// <summary>
        /// This will clean up all the  browser data.
        /// </summary>
        public void CleanUp()
        {
            //Temporary Internet Files
            System.Diagnostics.Process.Start("rundll32.exe", "InetCpl.cpl,ClearMyTracksByProcess 8");
            //Cookies()
            System.Diagnostics.Process.Start("rundll32.exe", "InetCpl.cpl,ClearMyTracksByProcess 2");
            //History()
            System.Diagnostics.Process.Start("rundll32.exe", "InetCpl.cpl,ClearMyTracksByProcess 1");
            //Form(Data)
            System.Diagnostics.Process.Start("rundll32.exe", "InetCpl.cpl,ClearMyTracksByProcess 16");
            //Passwords
            System.Diagnostics.Process.Start("rundll32.exe", "InetCpl.cpl,ClearMyTracksByProcess 32");
            //Delete(All)
            System.Diagnostics.Process.Start("rundll32.exe", "InetCpl.cpl,ClearMyTracksByProcess 255");
            //Delete All – Also delete files and settings stored by add-ons
            System.Diagnostics.Process.Start("rundll32.exe", "InetCpl.cpl,ClearMyTracksByProcess 4351");

        }

        #region SyncPO
        /// <summary>
        /// This is used to sync Purchase order and Delivery order, by specifying the keyword
        /// SyncPO: This keyword will sync the Purchase order.
        /// SyncDO: This will sync Delivery order.
        /// 
        /// NB: The order number must keep under "PurchaseOrder_No" in test data excel
        /// </summary>
        /// <param name="keyword"></param>
        public void SyncOrder(string keyword)
        {
            dt.Clear();
            if (keyword == "SyncPO")
            {
                dt = ExcelKeywordTable.ExcelData(LogPath + "\\KeywordDrivenData.xlsx", "SyncPO");
            }
            else if (keyword == "SyncDO")
            {
                dt = ExcelKeywordTable.ExcelData(LogPath + "\\KeywordDrivenData.xlsx", "SyncDO");
            }
            else
            {
                Operation.FailerReason = "No " + keyword + " keyword found. Please check the Keyword.";
                throw new NoSuchOperationFound("No " + keyword + " keyword found. Please check the Keyword.");
            }
            ModuleNavigation(dt.Rows[0]["ControlKeywordValue"].ToString());
            Pop.OperationStart(dt.Rows[1]["Keyword"].ToString(), dt.Rows[1]["TypeOfControl"].ToString(), dt.Rows[1]["ControlKeyword"].ToString(), dt.Rows[1]["ControlKeywordValue"].ToString(), dt.Rows[1]["DataRefferencekeyword"].ToString(), dt.Rows[1]["TypeOfWindow"].ToString(), dt.Rows[1]["Step No"].ToString(), dt.Rows[1]["Description"].ToString());
            Pop.OperationStart(dt.Rows[2]["Keyword"].ToString(), dt.Rows[2]["TypeOfControl"].ToString(), dt.Rows[2]["ControlKeyword"].ToString(), dt.Rows[2]["ControlKeywordValue"].ToString(), dt.Rows[2]["DataRefferencekeyword"].ToString(), dt.Rows[2]["TypeOfWindow"].ToString(), dt.Rows[2]["Step No"].ToString(), dt.Rows[2]["Description"].ToString());
            Pop.OperationStart(dt.Rows[3]["Keyword"].ToString(), dt.Rows[3]["TypeOfControl"].ToString(), dt.Rows[3]["ControlKeyword"].ToString(), dt.Rows[3]["ControlKeywordValue"].ToString(), ExcelDataTable.ReadData(1, dt.Rows[3]["DataRefferencekeyword"].ToString()), dt.Rows[3]["TypeOfWindow"].ToString(), dt.Rows[3]["Step No"].ToString(), dt.Rows[3]["Description"].ToString());
            Pop.OperationStart(dt.Rows[4]["Keyword"].ToString(), dt.Rows[4]["TypeOfControl"].ToString(), dt.Rows[4]["ControlKeyword"].ToString(), dt.Rows[4]["ControlKeywordValue"].ToString(), dt.Rows[4]["DataRefferencekeyword"].ToString(), dt.Rows[4]["TypeOfWindow"].ToString(), dt.Rows[4]["Step No"].ToString(), dt.Rows[4]["Description"].ToString());
            Pop.OperationStart(dt.Rows[5]["Keyword"].ToString(), dt.Rows[5]["TypeOfControl"].ToString(), dt.Rows[5]["ControlKeyword"].ToString(), dt.Rows[5]["ControlKeywordValue"].ToString(), dt.Rows[5]["DataRefferencekeyword"].ToString(), dt.Rows[5]["TypeOfWindow"].ToString(), dt.Rows[5]["Step No"].ToString(), dt.Rows[5]["Description"].ToString());
            string validationMessage = Pop.WebGetControlData(dt.Rows[6]["TypeOfControl"].ToString(), dt.Rows[6]["ControlKeyword"].ToString(), dt.Rows[6]["ControlKeywordValue"].ToString(), dt.Rows[6]["AssertionScreenshotName"].ToString(), dt.Rows[6]["AssertionMessage"].ToString(), dt.Rows[6]["Step No"].ToString(), dt.Rows[6]["Description"].ToString(), dt.Rows[6]["Keyword"].ToString()).Trim();
            if (validationMessage != lang.Msg_SyncOrderMessage)
            {
                string screenShotName = dt.Rows[6]["AssertionScreenshotName"].ToString();
                Operation.ErrorScreenPath = screenShot(screenShotName);
                Operation.FailerReason = dt.Rows[6]["AssertionMessage"].ToString();
                genreport.Reports(dt.Rows[6]["Step No"].ToString(), dt.Rows[6]["Description"].ToString(), dt.Rows[6]["Keyword"].ToString(), false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                genreport.FileCorreptionCheck();
            }
            Assert.AreEqual(lang.Msg_SyncOrderMessage, validationMessage, dt.Rows[6]["AssertionMessage"].ToString());
            genreport.Reports(dt.Rows[6]["Step No"].ToString(), dt.Rows[6]["Description"].ToString(), dt.Rows[6]["Keyword"].ToString(), true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
        }
        #endregion

        #region NavigationsToSubModule
        /// <summary>
        /// Navigate2CDA: Navigate to Configure Dispatch Area. [Refer Excel: KeywordDrivenData. Sheet Name: ConfigureDispatchNavigation]
        /// 
        /// Navigate2CRA: Navigate to Configure Receiving Area. [Refer Excel: KeywordDrivenData. Sheet Name: ConfigureReceivingNavigation]
        /// 
        /// Navigate2CPA: Navigate to Configure Production Area.[Refer Excel: KeywordDrivenData. Sheet Name: ConfigureProductionNavigation]
        /// </summary>
        /// <param name="keyword"></param>
        public void SubModuleNavigationPack(string keyword)
        {
            dt.Clear();
            if (keyword == "Navigate2CDA")
            {
                dt = ExcelKeywordTable.ExcelData(LogPath + "\\KeywordDrivenData.xlsx", "ConfigureDispatchNavigation");
            }
            else if (keyword == "Navigate2CRA")
            {
                dt = ExcelKeywordTable.ExcelData(LogPath + "\\KeywordDrivenData.xlsx", "ConfigureReceivingNavigation");
            }
            else if (keyword == "Navigate2CPA")
            {
                dt = ExcelKeywordTable.ExcelData(LogPath + "\\KeywordDrivenData.xlsx", "ConfigureProductionNavigation");
            }
            else
            {
                Operation.FailerReason = "No " + keyword + " keyword found. Please check the Keyword.";
                throw new NoSuchOperationFound("No " + keyword + " keyword found. Please check the Keyword.");
            }
            ModuleNavigation(dt.Rows[0]["ControlKeywordValue"].ToString());
            Pop.OperationStart(dt.Rows[1]["Keyword"].ToString(), dt.Rows[1]["TypeOfControl"].ToString(), dt.Rows[1]["ControlKeyword"].ToString(), dt.Rows[1]["ControlKeywordValue"].ToString(), ExcelDataTable.ReadBatchData(batchforReport, dt.Rows[1]["DataRefferencekeyword"].ToString()), dt.Rows[1]["TypeOfWindow"].ToString(), dt.Rows[1]["Step No"].ToString(), dt.Rows[1]["Description"].ToString());
            SubModuleNavigation(dt.Rows[2]["ControlKeywordValue"].ToString());
        }
        #endregion

        #region Select PPE Info
        /// <summary>
        /// this will select specified number of PPE info in Material GHS Information specified in "Data reference keyword"
        /// </summary>
        /// <param name="steps">step number</param>
        /// <param name="description">what this step is doing</param>
        /// <param name="typeofoperation">key word used to execute the function</param>
        /// <param name="countPPE">number of PPE info to be selected</param>
        /// <returns></returns>
        public string SelectPPEInfo(string steps, string description, string typeofoperation, string countPPE)
        {
            string data = string.Empty;
            try
            {
                BrowserWindow window = new BrowserWindow();
                window.ExecuteScript(@"PPEFunction();
                                    function PPEFunction(){"
                                    + "var PPE = document.getElementById('ContentPlaceHolder1_ucPPEControl_divImageCtrlPopOver');"
                                    + "for(var i=1; i<='" + countPPE + "'; i++){"
                                    + "PPE.getElementsByTagName('input')[i-1].checked=true;}}");
                data = "True";
                return data;
            }
            catch (Exception e)
            {
                Operation.FailerReason = "Could not select PPE Info value Please check the data given.";
                data = "False";
                return data;
            }
        }
        #endregion

        #region Select DG Material
        /// <summary>
        /// this will select specified number of DG Class info in Label Database specified in "Data reference keyword" and it will click save button for DG Class
        /// </summary>
        /// <param name="steps">step number</param>
        /// <param name="description">what this step is doing</param>
        /// <param name="typeofoperation">key word used to execute the function</param>
        /// <param name="countPPE">number of DG Class to be selected</param>
        /// <returns></returns>
        public string SelectDGInfo(string steps, string description, string typeofoperation, string countDG)
        {
            string data = string.Empty;
            try
            {
                BrowserWindow window = new BrowserWindow();
                string idval = window.ExecuteScript(@"var data = SelectDGClass();
                                            function SelectDGClass(){"
                                            + "var elts = document.getElementsByClassName('label-heading block-label p-l-1 p-r-1');"
                                            + "for (var i = 0; i < elts.length; i++){"
                                            + "if (elts[i].innerHTML == 'DG Class'){"
                                            + "var sed = elts[i].nextElementSibling.firstElementChild.firstElementChild.id.replace('_divImageSelectCtrl','_txtImageCtrlCodes');"
                                            + "document.getElementById(sed).scrollIntoView(true);"
                                            + "document.getElementById(sed).click();"
                                            + "var DGsed = elts[i].nextElementSibling.firstElementChild.firstElementChild.id.replace('_divImageSelectCtrl','_divImageCtrlPopOverBody');}}"
                                            + "return DGsed;}return data;").ToString();
                Playback.Wait(max);

                window.ExecuteScript(@"var DG = document.getElementById('" + idval + "');for (var i=1; i<='" + countDG + "'; i++){DG.getElementsByTagName('input')[i-1].checked=true;}");

                string SaveBtnID = idval.Replace("_divImageCtrlPopOverBody", "_btnSaveImages");

                Pop.OperationStart("Click", "HtmlInputButton", "ID", SaveBtnID, "", "Web", steps + 1, "Click Save For DG Class");

                Playback.Wait(mid);
                data = "True";
                return data;

            }
            catch (Exception e)
            {
                Operation.FailerReason = "Could not select DG Info value Please check the data given.";
                data = "False";
                return data;
            }
        }
        #endregion

        #region ValidateDGClassAvailability
        /// <summary>
        /// This will validate the line items are DG Class maintained or not, and it will validate the line items showing in red.
        /// if it showing in red it will check for any DG class images are available in the table.
        /// This also Record the Batch number of the Material(This validation is done based on the material that specified)
        /// </summary>
        /// <param name="steps">Number of steps.</param>
        /// <param name="description">Description for the operation doing.</param>
        /// <param name="typeofoperation">Specify the keyword.</param>
        /// <param name="resource">you must give values in Data Reference keyword in keyword steps
        ///  like Reference to order number+Reference to Material Ex:PurchaseOrder_No+MetrialNumber
        /// </param>
        /// <returns>This will return a true or false value to the keyword.</returns>
        public string ValidateDG_PPE(string steps, string description, string typeofoperation, string resource)
        {
            string status = string.Empty;
            BrowserWindow window = new BrowserWindow();
            try
            {
                string[] resorce = resource.Split('+');
                if (resorce[1].Contains('='))
                {
                    WriteAndReadData.DataFilePath = ConfigurationManager.AppSettings["ExcelDataFile"] + "\\DynamicData" + "\\" + batchforReport + ".xls";
                    ExcelDataTable.PopulateRecordData(WriteAndReadData.DataFilePath);
                }
                string idvalue = ExcelDataTable.ReadData(1, resorce[0]) + Pop.getDataFromDynamicExcel(resorce[1].ToString()) + "_trMaterial";
                string color = window.ExecuteScript("var data =document.getElementById('" + idvalue + "').getAttribute('style'); return data;").ToString().Trim();
                if ((color == "color: rgb(200, 5, 4); font-weight: bold;") || (color == "color:#C80504;font-weight:bold"))
                {
                    int dgClass = Convert.ToInt32(window.ExecuteScript("var data=$('#trDGInfo').length; return data;"));
                    if (dgClass >= 1)
                    {
                        int dg = Convert.ToInt32(window.ExecuteScript("var data = document.getElementById('trDGInfo').getElementsByTagName('img').length; return data;"));
                        if (dg <= 0)
                        {
                            status = "false";
                        }
                        else
                        {
                            string batch = window.ExecuteScript("var data = document.getElementById('" + idvalue.Replace("_trMaterial", "_spnBatchName") + "').innerHTML; return data;").ToString();
                            string reqQuantity = window.ExecuteScript("var data = document.getElementById('" + idvalue.Replace("_trMaterial", "_spnReqQty") + "').innerHTML; return data;").ToString();
                            status = "true" + ":" + batch + ":" + reqQuantity;
                        }
                    }
                    else { status = "false"; }
                }
                else { status = "false"; }
                return status;
            }
            catch (Exception e)
            {
                Operation.FailerReason = "Could not get data for DG Material";
                genreport.Reports(steps, description, typeofoperation, false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                genreport.FileCorreptionCheck();
                throw new NoSuchControlTypeFound("Could not get data for DG Material.");
            }

        }
        #endregion

        #region ValidatePPE_Availabilaity
        /// <summary>
        /// This function will validate the material is maintained the PPE or not and also validate PPE images is displaying or not.
        /// this function can be invoked using keyword: "ValidatePPEStatus";
        /// </summary>
        /// <param name="steps">Number of steps.</param>
        /// <param name="description">Description for the operation doing.</param>
        /// <param name="typeofoperation">Specify the keyword. "ValidatePPEStatus"</param>
        /// <returns>This will return a true or false value to the keyword.</returns>
        public bool ValidatePPE_Avail(string steps, string description, string typeofoperation)
        {
            BrowserWindow window = new BrowserWindow();
            try
            {
                int count = Convert.ToInt16(window.ExecuteScript(@"var PPE=PPEAvailability();
                                function PPEAvailability(){"
                                + "var length = document.getElementById('tblPickingMaterialList').getElementsByTagName('td').length;"
                                + "for(var i=0; i < length ; i++){"
                                + "if(document.getElementById('tblPickingMaterialList').getElementsByTagName('td')[i].hasAttribute('ppe')){"
                                + "var datacount= i;break;}else{var datacount= -1}}"
                                + "return datacount;}return PPE;"));
                if (count != -1)
                {
                    int countPPE = Convert.ToInt32(window.ExecuteScript("var count=document.getElementById('tblPickingMaterialList').getElementsByTagName('td')['" + count + "'].getElementsByTagName('img').length; return count;"));
                    if (countPPE < 1)
                    {
                        return false;
                    }
                    else
                    {
                        return true;
                    }
                }
                else
                {
                    return false;
                }

            }
            catch (Exception e)
            {
                Operation.FailerReason = "Could not get data for PPE Info";
                genreport.Reports(steps, description, typeofoperation, false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                genreport.FileCorreptionCheck();
                throw new NoSuchControlTypeFound("Could not get data for PPE Info.");
            }
        }
        #endregion


        #region AssignedOrderStatus
        /// <summary>
        /// This function will validate :
        /// 1. Assigned orders color that show in red or not based on the Delivery Order.
        /// 2. Validate the status is correct or not based on the "DataRefferenceKeyword".
        /// 3. Validate the selected operated is showing in the table or not
        /// </summary>
        /// <param name="steps">Step Number.</param>
        /// <param name="description">Description about the step.</param>
        /// <param name="typeofoperation">Keyword used to invoke this function."AssignedOrderValidate"</param>
        /// <param name="resource">Specify the Delivery order Reference+ Picker Name reference. Ex:PurchaseOrder_No+PickerName.</param>
        /// <param name="datarefference">Here you give the reference to the status of the order you want to validate.</param>
        /// <returns>this function will return true or false value.</returns>
        public string AssignedOrderValidate(string steps, string description, string typeofoperation, string resource, string datarefference)
        {
            string[] resorce = { };
            string[] orderType = resource.Split(':');
            string calenderStatus = string.Empty;
            string pickerinfo = string.Empty;
            if (orderType.Count() > 1)
            {
                resorce = orderType[1].Split('+');
            }
            else
            {
                resorce = resource.Split('+');
            }

            string validatestatus = string.Empty;
            int j = 0;
            BrowserWindow window = new BrowserWindow();
            try
            {
                auto.Send("{END}");
                Thread.Sleep(500);
                window.ExecuteScript("window.scrollTo(0,document.body.scrollHeight);");
                Thread.Sleep(700);
                int count = Convert.ToInt32(window.ExecuteScript("var data = document.getElementById('ContentPlaceHolder1_tblOrderList').getElementsByTagName('tr').length; return data;")) + 1;
                count = count / 2;
                //int count = Convert.ToInt32(window.ExecuteScript("var data = document.getElementById('ContentPlaceHolder1_tblOrderList').childElementCount; return data;"));
                for (int i = 1; i <= count; i++)
                {
                    string idPickerBatch = "ucGroupOrder_" + i + "_dlstDeliveries";
                    int PickerOrdCount = Convert.ToInt32(window.ExecuteScript("var data = document.getElementById('" + idPickerBatch + "').childElementCount; return data;"));
                    for (int k = 0; k < PickerOrdCount; k++)
                    {
                        string idDeliveryOrd = "ucGroupOrder_" + i + "_dlstDeliveries_lnkDelorder_" + k;
                        try
                        {
                            window.ExecuteScript("document.getElementById('" + idDeliveryOrd + "').scrollIntoView(true);");
                        }
                        catch (Exception e) { }
                        try
                        {
                            string Calenderdeliveryorder = window.ExecuteScript("var data = document.getElementById('" + idDeliveryOrd + "').innerHTML; return data; ").ToString().Trim();
                            if (Calenderdeliveryorder == ExcelDataTable.ReadData(1, resorce[0]))
                            {
                                if (orderType[0] == "CrossDock")
                                {
                                    calenderStatus = window.ExecuteScript("var data = document.getElementById('ucGroupOrder_" + i + "_dlstDeliveries_imgStatus_" + k + "').getAttribute('src'); return data;").ToString().Trim() + window.ExecuteScript("var data = document.getElementById('ucGroupOrder_" + i + "_dlstDeliveries_imgCrossdock_" + k + "').getAttribute('src'); return data;").ToString().Trim();
                                }
                                else
                                {
                                    calenderStatus = window.ExecuteScript("var data = document.getElementById('ucGroupOrder_" + i + "_dlstDeliveries_imgStatus_" + k + "').getAttribute('src'); return data;").ToString().Trim();
                                }
                                if (calenderStatus != datarefference)
                                {
                                    validatestatus = "False";
                                    break;
                                }
                                else
                                {
                                    if (orderType[0] == "CrossDock")
                                    {
                                        pickerinfo = window.ExecuteScript("var data=document.getElementById('ucGroupOrder_" + i + "_picker').getElementsByTagName('span')[0].innerHTML; return data;").ToString().Trim();
                                        pickerinfo = pickerinfo.Replace(",", string.Empty);
                                        if (!string.Equals(pickerinfo, ExcelDataTable.ReadData(1, resorce[1]), StringComparison.InvariantCultureIgnoreCase))
                                        {
                                            validatestatus = "False";
                                            continue;
                                        }
                                        else
                                        {
                                            string Shipinginfo = window.ExecuteScript("var data=document.getElementById('ucGroupOrder_" + i + "_dlstDeliveries_spnShippingContent_" + k + "').innerHTML; return data;").ToString().Trim();
                                            if (Shipinginfo != ExcelDataTable.ReadData(1, resorce[2]))
                                            {
                                                validatestatus = "False";
                                                break;
                                            }
                                            else
                                            {
                                                validatestatus = "True";
                                                break;
                                            }
                                        }
                                    }
                                    else
                                    {
                                        pickerinfo = window.ExecuteScript("var data=document.getElementById('ucGroupOrder_" + i + "_picker').getElementsByTagName('span')[0].innerHTML; return data;").ToString().Trim();
                                        pickerinfo = pickerinfo.Replace(",", string.Empty);
                                        //if (pickerinfo != ExcelDataTable.ReadData(1, resorce[1]).ToUpper())
                                        if (!string.Equals(pickerinfo, ExcelDataTable.ReadData(1, resorce[1]), StringComparison.InvariantCultureIgnoreCase))
                                        {
                                            validatestatus = "False";
                                        }
                                        else
                                        {
                                            string Shipinginfo = window.ExecuteScript("var data=document.getElementById('ucGroupOrder_" + i + "_dlstDeliveries_spnShippingContent_" + k + "').innerHTML; return data;").ToString().Trim();
                                            if (Shipinginfo != ExcelDataTable.ReadData(1, resorce[2]))
                                            {
                                                validatestatus = "False";
                                                break;
                                            }
                                            else
                                            {
                                                validatestatus = "True";
                                                break;
                                            }
                                        }
                                    }
                                }
                            }
                            else
                            {
                                validatestatus = "False";
                                continue;
                            }
                        }
                        catch (Exception e)
                        {
                            break;
                        }
                    }
                    if (validatestatus == "True")
                    {
                        break;
                    }
                }
                return validatestatus;
            }
            catch (Exception e)
            {
                Operation.FailerReason = "Could not Validate Assigned Order";
                genreport.Reports(steps, description, typeofoperation, false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                genreport.FileCorreptionCheck();
                throw new NoSuchControlTypeFound("Could not Validate Assigned Order.");
            }
        }
        #endregion

        #region Getting Bin Location
        /// <summary>
        /// this will validate get the bin location from he sap stock over view based on the storage type, min available Qty specified in the test data.and stored in the reference value specified in the DataRefferencekeyword.
        /// </summary>
        /// <param name="steps">Step Number.</param>
        /// <param name="description">Description about the step.</param>
        /// <param name="typeofoperation">Keyword used to invoke this function. "GetDispatchBinLocation"</param>
        /// <param name="resource">Specify the Storage type Reference + Minimum available qty reference. Ex:StorageType+MinAvailableqty.</param>
        /// <param name="datarefference">In DataRefferencekeyword column in the step wise sheet you must specify the text that you want to store the BIN Location of Delivery order.</param>
        /// <returns>This will return the bin location. and will store the location in the variable that specified in the DataRefferencekeyword</returns>
        public string GettingBinLocation(string steps, string description, string typeofoperation, string resource)
        {
            BrowserWindow window = new BrowserWindow();
            string bin = string.Empty;
            string[] resorc = resource.Split('+');
            try
            {
                string[] storagrtypes = ExcelDataTable.ReadData(1, resorc[0]).Split(',');
                for (int j = 0; j < storagrtypes.Count(); j++)
                {
                    int palletCount = Convert.ToInt32(window.ExecuteScript("var data=document.getElementById('tbStockOverViewList').getElementsByTagName('tr').length; return data;"));
                    for (int i = 1; i <= palletCount; i++)
                    {
                        string storageLocation = window.ExecuteScript("var data=document.getElementById('ctl" + i.ToString("00") + "_spnStorageType').innerHTML; return data;").ToString().Trim();
                        if (storageLocation == storagrtypes[j])
                        {
                            string storageLoc = window.ExecuteScript("var data=document.getElementById('ctl" + i.ToString("00") + "_spnStorageLocation').innerHTML; return data;").ToString().Trim();
                            if (storageLoc == ExcelDataTable.ReadData(1,resorc[2].ToString()))
                            {
                                string QualityBlock = window.ExecuteScript("var data=document.getElementById('ctl" + i.ToString("00") + "_spnQuality').innerHTML; return data;").ToString().Trim();
                                if(QualityBlock != "S")
                                {
                                    string available = window.ExecuteScript("var data=document.getElementById('ctl" + i.ToString("00") + "_spnAvailableQuantity').innerHTML;return data;").ToString();
                                    available = available.Substring(0, available.IndexOf(" ")).Trim();
                                    double final = Convert.ToDouble(available);
                                    if (final > Convert.ToInt32(ExcelDataTable.ReadData(1, resorc[1])))
                                    {
                                        bin = window.ExecuteScript("var data = document.getElementById('ctl" + i.ToString("00") + "_spnBinLocation').innerHTML; return data;").ToString().Trim();
                                        bin = bin + ":" + final;
                                        break;
                                    }
                                    else
                                    {
                                        bin = string.Empty;
                                        continue;
                                    }
                                }
                                else
                                {
                                    bin = string.Empty;
                                    continue;
                                }
                            }
                        }
                        else
                        {
                            bin = string.Empty;
                            continue;
                        }
                    }
                    if (bin != string.Empty)
                    {
                        return bin;
                    }
                }
                if (bin != string.Empty)
                {
                    return bin;
                }
                else
                {
                    bin = string.Empty;
                    return bin;
                }
            }
            catch (Exception e)
            {
                Operation.FailerReason = "Getting Bin Location Failed";
                genreport.Reports(steps, description, typeofoperation, false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                genreport.FileCorreptionCheck();
                throw new NoSuchControlTypeFound("Getting Bin Location Failed");
            }
        }
        #endregion

        #region Getting SAP Stock Batch
        /// <summary>
        /// this will validate get the Batch Number from he sap stock over view based on the storage type, min available Qty specified in the test data.and stored in the reference value specified in the DataRefferencekeyword.
        /// </summary>
        /// <param name="steps">Step Number.</param>
        /// <param name="description">Description about the step.</param>
        /// <param name="typeofoperation">Keyword used to invoke this function. "GetSAPStock_Batch"</param>
        /// <param name="resource">Specify the Storage type Reference + Minimum available qty reference. Ex:StorageType+MinAvailableqty.</param>
        /// <param name="datarefference">In DataRefferencekeyword column in the step wise sheet you must specify the text that you want to store the BIN Location of Delivery order.</param>
        /// <returns>This will return the bin location. and will store the location in the variable that specified in the DataRefferencekeyword</returns>
        public string GettingSAP_BatchNumber(string steps, string description, string typeofoperation, string resource)
        {
            BrowserWindow window = new BrowserWindow();
            string bin = string.Empty;
            string[] resorc = resource.Split('+');
            try
            {
                string[] storagrtypes = ExcelDataTable.ReadData(1, resorc[0]).Split(',');
                for (int j = 0; j < storagrtypes.Count(); j++)
                {
                    int palletCount = Convert.ToInt32(window.ExecuteScript("var data=document.getElementById('tbStockOverViewList').getElementsByTagName('tr').length; return data;"));
                    for (int i = 1; i <= palletCount; i++)
                    {
                        string storageLocation = window.ExecuteScript("var data=document.getElementById('ctl" + i.ToString("00") + "_spnStorageType').innerHTML; return data;").ToString().Trim();
                        if (storageLocation == storagrtypes[j])
                        {
                            string QualityBlock = window.ExecuteScript("var data=document.getElementById('ctl" + i.ToString("00") + "_spnQuality').innerHTML; return data;").ToString().Trim();
                            if (QualityBlock != "S")
                            {
                                string available = window.ExecuteScript("var data=document.getElementById('ctl" + i.ToString("00") + "_spnAvailableQuantity').innerHTML;return data;").ToString();
                                available = available.Substring(0, available.IndexOf(" ")).Trim();
                                double final = Convert.ToDouble(available);
                                if (final > Convert.ToInt32(ExcelDataTable.ReadData(1, resorc[1])))
                                {
                                    bin = window.ExecuteScript("var data = document.getElementById('ctl" + i.ToString("00") + "_spnBatch').innerHTML; return data;").ToString().Trim();
                                    break;
                                }
                                else
                                {
                                    bin = string.Empty;
                                    continue;
                                }
                            }
                            else
                            {
                                bin = string.Empty;
                                continue;
                            }
                        }
                        else
                        {
                            bin = string.Empty;
                            continue;
                        }
                    }
                    if (bin != string.Empty)
                    {
                        return bin;
                    }
                }
                if (bin != string.Empty)
                {
                    return bin;
                }
                else
                {
                    bin = string.Empty;
                    return bin;
                }
            }
            catch (Exception e)
            {
                Operation.FailerReason = "Getting Batch Number Failed";
                genreport.Reports(steps, description, typeofoperation, false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                genreport.FileCorreptionCheck();
                throw new NoSuchControlTypeFound("Getting Batch Number Failed");
            }
        }
        #endregion

        #region Validate OrderStatus Picking
        /// <summary>
        /// This will validate the Order status in the Dispatch Mobile pick based on the value specified in the DataRefferencekeyword
        /// </summary>
        /// <param name="steps">Step Number.</param>
        /// <param name="description">Description about the step.</param>
        /// <param name="typeofoperation">Keyword used to invoke this function. "ValidatePickStatus"</param>
        /// <param name="resource">Give the reference for the Delivery order</param>
        /// <param name="datarefference">give the reference for the validation status</param>
        /// <returns>this function will return true or false value. and order with details and click the order</returns>
        public string MobilePickOrderStatus(string steps, string description, string typeofoperation, string resource, string datarefference)
        {
            BrowserWindow window = new BrowserWindow();
            try
            {
                string status = window.ExecuteScript(@"var orderStatus = DeliveryStatusMob();
                                function DeliveryStatusMob(){
                                var countdata=document.getElementById('tableOrderDetailList').getElementsByTagName('tbody')[0].getElementsByTagName('tr').length;
                                for(var i=0; i<countdata; i++){
                                var orderDetail = document.getElementById('tableOrderDetailList').getElementsByTagName('tbody')[0].getElementsByTagName('tr')[i].getElementsByTagName('td')[2].innerHTML.trim();"
                                + "if(orderDetail.indexOf('" + resource + "')>=0){"
                                + "var property = document.getElementById('tableOrderDetailList').getElementsByTagName('tbody')[0].getElementsByTagName('tr')[i].getAttribute('style').trim();"
                                + "if(property==='" + datarefference + "'){var status= 'True+' + orderDetail ;}"
                                + "else{var status= 'False';}}}return status;}return orderStatus;").ToString();
                string[] order = status.Split('+');
                if (order[0] == "True")
                {
                    Pop.OperationStart("Click", "HtmlCell", "InnerText", order[1], "", "Web", steps, description);
                }

                return order[0];
            }
            catch (Exception e)
            {
                Operation.FailerReason = "Not able to validate the order status in Mobile Picking";
                genreport.Reports(steps, description, typeofoperation, false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                genreport.FileCorreptionCheck();
                throw new NoSuchControlTypeFound("Not able to validate the order status in Mobile Picking");
            }
        }
        #endregion

        #region Validate PPEIcon in Picking
        /// <summary>
        /// This will validate the PPE is add or not for the material and if added it go inside the PPE info and check for PPE image is loaded or not.
        /// keyword:DispatchPPEValidation
        /// </summary>
        /// <returns>return true  or false value</returns>
        public bool DispatchMobilePPEValidation(string steps, string description, string typeofoperation)
        {
            BrowserWindow window = new BrowserWindow();
            bool status = true;
            string PPEImageAvail = string.Empty;
            try
            {
                try
                {
                    string poperrty = "[style = \"display: none;\"][title=\"PPE\"]";
                    PPEImageAvail = window.ExecuteScript("var data = $('" + poperrty + "').is('img'); return data;").ToString().Trim();
                }
                catch(Exception e) { }
                if (PPEImageAvail == "True")
                {
                    status = false;
                }
                else
                {
                    //Go to PPE details
                    window.ExecuteScript("$('[title = PPE]').click();");
                    Thread.Sleep(max * 2);

                    Pop.WebWaitForControl("HtmlHyperlink", "ID", "hypClosePPE", steps, description, typeofoperation, "");

                    int countimg = Convert.ToInt32(window.ExecuteScript("var data = $('#divPPE img').length; return data;"));
                    Thread.Sleep(max);
                    if (countimg < 1)
                    {
                        status = false;
                    }
                }
                return status;
            }
            catch (Exception e)
            {
                Operation.FailerReason = "Not able to validate the PPE in Mobile Picking";
                genreport.Reports(steps, description, typeofoperation, false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                genreport.FileCorreptionCheck();
                throw new NoSuchControlTypeFound("Not able to validate the order status in Mobile Picking");
            }
        }
        #endregion

        #region Validation Pick Instruction Available
        /// <summary>
        /// Validate picking instruction available or not and navigate to Picking instructions.
        /// keyword:DispatchPickInstValidation
        /// </summary>
        /// <returns>return True or False</returns>
        public bool DispatchMobilePickInstruction(string steps, string description, string typeofoperation)
        {
            BrowserWindow window = new BrowserWindow();
            bool status = true;
            string PicInstnAvail = string.Empty;
            try
            {

                try
                {
                    //Validating PPE icon availability
                    string poperrty = "[style = \"display: none;\"][title=\"Instructions\"]";
                    PicInstnAvail = window.ExecuteScript("var data = $('" + poperrty + "').is('img'); return data;").ToString().Trim();
                }
                catch (Exception e) { }
                if (PicInstnAvail == "True")
                {
                    status = false;
                }
                else
                {
                    window.ExecuteScript("$('[title = Instructions]').click();");
                    Thread.Sleep(max);
                }
                return status;
            }
            catch (Exception e)
            {
                Operation.FailerReason = "Not able to validate the Picking Instruction in Mobile Picking";
                genreport.Reports(steps, description, typeofoperation, false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                genreport.FileCorreptionCheck();
                throw new NoSuchControlTypeFound("Not able to validate the order status in Mobile Picking");
            }

        }
        #endregion

        #region Validate OrderStatus Picking Screen
        /// <summary>
        /// This will validate the status(color) of the order in picking screen in dispatch Mobile and return the value color value.
        /// </summary>
        /// <param name="steps">Step Number.</param>
        /// <param name="description">Description about the step.</param>
        /// <param name="typeofoperation">Keyword used to invoke this function. "ValidatePickOrderStatus"</param>
        /// <param name="resource">Give the reference for the Delivery order you want to validate in "ControlKeywordValue"</param>
        /// <returns>returns the status color of the order, must specify the validating color in DataRefferencekeyword</returns>
        public string DispatchMobPickingOrderStatus(string steps, string description, string typeofoperation, string resource)
        {
            BrowserWindow window = new BrowserWindow();
            try
            {
                string orderStatus = window.ExecuteScript(@"var stat = PickOrderStatus();
                                            function PickOrderStatus(){
                                            var ordercount=document.getElementById('tabledeliveryOrderlist').getElementsByTagName('tr').length;
                                            for(var i=1; i<=ordercount; i++){
                                            var order = document.getElementById('tabledeliveryOrderlist').getElementsByTagName('tr')[i].getElementsByTagName('td')[0].innerHTML.trim();"
                                            + "if(order=='" + resource + "'){"
                                            + "var orderStatus = document.getElementById('tabledeliveryOrderlist').getElementsByTagName('tr')[i].getAttribute('style');break;}}"
                                            + "return orderStatus;}return stat;").ToString().Trim();
                return orderStatus;
            }
            catch (Exception e)
            {
                Operation.FailerReason = "Not able to validate the Picking order status for '" + resource + "' in Mobile Picking";
                genreport.Reports(steps, description, typeofoperation, false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                genreport.FileCorreptionCheck();
                throw new NoSuchControlTypeFound("Not able to validate the Picking order status for '" + resource + "' in Mobile Picking");
            }
        }
        #endregion

        #region CheckOrderAvailable
        /// <summary>
        /// This will validate the order is still available in picking list after completing the picking also
        /// </summary>
        /// <param name="steps">Step Number.</param>
        /// <param name="description">Description about the step.</param>
        /// <param name="typeofoperation">Keyword used to invoke this function. "ValidateOrderAvailStatus"</param>
        /// <param name="resource">Give the reference for the Delivery order you want to check available or not "ControlKeywordValue"</param>
        /// <returns>return True or False</returns>
        public bool CheckOrderAvailable(string steps, string description, string typeofoperation, string resource)
        {
            try
            {
                BrowserWindow window = new BrowserWindow();
                bool status = true;
                int countM = Convert.ToInt32(window.ExecuteScript("var data=document.getElementById('tableOrderDetailList').getElementsByTagName('tbody')[0].getElementsByTagName('tr').length;return data;"));
                if (countM > 0)
                {
                    for (int i = 0; i < countM; i++)
                    {
                        string order = window.ExecuteScript("var data = document.getElementById('tableOrderDetailList').getElementsByTagName('tbody')[0].getElementsByTagName('tr')[" + i + "].getElementsByTagName('td')[2].innerHTML; return data;").ToString().Trim();
                        if (order.Substring(0, 10).ToString() == resource)
                        {
                            status = false;
                        }
                        else
                        {
                            continue;
                        }
                    }
                }
                return status;
            }
            catch (Exception e)
            {
                Operation.FailerReason = "Not able to validate Order is available or not in the picking Screen";
                genreport.Reports(steps, description, typeofoperation, false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                genreport.FileCorreptionCheck();
                throw new NoSuchControlTypeFound("Not able to validate Order is available or not in the picking Screen");
            }
        }
        #endregion

        #region PackOrderStatus
        /// <summary>
        /// This will validate the Order status in the Dispatch Mobile Pack based on the value specified in the DataRefferencekeyword
        /// </summary>
        /// <param name="steps">Step Number.</param>
        /// <param name="description">Description about the step.</param>
        /// <param name="typeofoperation">Keyword used to invoke this function. "ValidatePackStatus"</param>
        /// <param name="resource">Give the reference for the Delivery order</param>
        /// <param name="datarefference">give the reference for the validation status</param>
        /// <returns>this function will return true or false value. and order with details and click the order</returns>
        public string MobilePackOrderStatus(string steps, string description, string typeofoperation, string resource, string datarefference)
        {
            BrowserWindow window = new BrowserWindow();
            try
            {
                string status = window.ExecuteScript(@"var orderStatus = DeliveryPackStatusMob();
                                function DeliveryPackStatusMob(){
                                var countdata=document.getElementById('tablePackDeliveryOrderList').getElementsByTagName('tbody')[0].getElementsByTagName('tr').length;
                                for(var i=0; i<countdata; i++){
                                var orderDetail = document.getElementById('tablePackDeliveryOrderList').getElementsByTagName('tbody')[0].getElementsByTagName('tr')[i].getElementsByTagName('td')[3].innerHTML.trim();"
                                + "if (orderDetail.indexOf('" + resource + "')>=0){"
                                + "var property = document.getElementById('tablePackDeliveryOrderList').getElementsByTagName('tbody')[0].getElementsByTagName('tr')[i].getAttribute('style').trim();"
                                + "if(property==='" + datarefference + "'){var status= 'True+' + orderDetail ;}"
                                + "else{var status= 'False';}}}return status;}return orderStatus;").ToString();
                string[] order = status.Split('+');
                if (order[0] == "True")
                {
                    Pop.OperationStart("Click", "HtmlCell", "InnerText", order[1], "", "Web", steps, description);
                }

                return order[0];
            }
            catch (Exception e)
            {
                Operation.FailerReason = "Not able to validate the order status in Mobile Packing";
                genreport.Reports(steps, description, typeofoperation, false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                genreport.FileCorreptionCheck();
                throw new NoSuchControlTypeFound("Not able to validate the order status in Mobile Packing");
            }
        }
        #endregion

        #region PackingInstructionValidate
        /// <summary>
        /// This function to navigate to the pack instruction and is the instruction is not available failed with given screen shot.
        /// </summary>
        /// <param name="steps">Step Number.</param>
        /// <param name="description">Description about the step.</param>
        /// <param name="typeofoperation">Keyword used to invoke this function. "ValidatePackInstructionAvail"</param>
        /// <returns>if pack instruction icon available it will navigate to instruction other wise return false value saying that no PackingInstruction Not available.</returns>
        public bool PackingInstructions(string steps, string description, string typeofoperation)
        {
            BrowserWindow window = new BrowserWindow();
            bool status = true;
            try
            {
                string packingSrc = window.ExecuteScript("var data =document.getElementsByClassName('grid_2 text-padding pull-right')[0].hasAttribute('style'); return data;").ToString().Trim();
                if (packingSrc == "False")
                {
                    window.ExecuteScript("javascript:ShowInstructionClick()");
                    status = true;
                }
                else
                {
                    //if (packingSrc == "display: none;")
                    //{
                        status = false;
                    //}
                }
                return status;
            }
            catch (Exception e)
            {
                Operation.FailerReason = "Not able to validate Packing instruction Mobile Packing";
                genreport.Reports(steps, description, typeofoperation, false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                genreport.FileCorreptionCheck();
                throw new NoSuchControlTypeFound("Not able to validate Packing instruction Mobile Packing");
            }
        }
        #endregion

        #region PackagingDetailsValidation
        /// <summary>
        /// Here Packaging Details are validate in the Packaging Details page.
        /// </summary>
        /// <param name="steps">Step Number.</param>
        /// <param name="description">Description about the step.</param>
        /// <param name="typeofoperation">Keyword: "MatchPackageDetails"</param>
        /// <param name="keywordValue">Give the Table id value in the page</param>
        /// <param name="datarefference">Her give all the data reference keywords that you want to validate separated by comma in order (displayed in Packaging Details page.)</param>
        /// <returns>this will return the true or false value.</returns>
        public bool PackagingDetailsValidation(string steps, string description, string typeofoperation, string keywordValue, string datarefference)
        {
            BrowserWindow window = new BrowserWindow();
            bool status = true;
            string[] packdetails = datarefference.Split(',');
            int countTD = Convert.ToInt32(window.ExecuteScript("var data= document.getElementById('" + keywordValue + "').getElementsByTagName('tr')[1].getElementsByTagName('td').length; return data;"));
            for (int i = 0; i < countTD; i++)
            {
                string innerData = window.ExecuteScript("var data= document.getElementById('" + keywordValue + "').getElementsByTagName('tr')[1].getElementsByTagName('td')[" + i + "].innerHTML;return data;").ToString().Trim();
                Thread.Sleep(min);
                if (innerData != ExcelDataTable.ReadData(1, packdetails[i]))
                {
                    status = false;
                }
                else
                {
                    continue;
                }
            }
            return status;
        }
        #endregion

        #region Link PPE Image
        /// <summary>
        /// This will link the PPE Code to the specified Image Description
        /// </summary>
        /// <param name="steps">Step Number.</param>
        /// <param name="description">Description about the step.</param>
        /// <param name="typeofoperation">Keyword: "LinkPPE"</param>
        /// <param name="keywordValue">Give the id value of the div that the pages contains</param>
        /// <param name="datarefference">Give the reference to the value of the PPE image Description</param>
        /// <returns>this will return the true or false value.</returns>
        public string PPELink(string steps, string description, string typeofoperation, string keywordValue, string datarefference)
        {
            BrowserWindow window = new BrowserWindow();
            try
            {
                string status = window.ExecuteScript(@"var data = PPELinking();
                         function PPELinking(){"
                       + "var count = document.getElementById('" + keywordValue + "').getElementsByTagName('tr').length;"
                       + "var PPEImage='" + datarefference + "';var status;for(var i=0; i < count; i++){"
                       + "var images = document.getElementById('ContentPlaceHolder1_tbdSelectPPE').getElementsByTagName('tr')[i].innerText.trim();"
                       + "if(PPEImage == images){"
                       + "var PPECodeID = document.getElementById('ContentPlaceHolder1_tbdSelectPPE').getElementsByTagName('tr')[i].getAttribute('id').replace('_trPPECode','_tdPPE');"
                       + "var PPEImageID = document.getElementById('ContentPlaceHolder1_tbdSelectPPE').getElementsByTagName('tr')[i].getAttribute('id').replace('_trPPECode','_tdShowImage');"
                       + "document.getElementById(PPECodeID).scrollIntoView(true);"
                       + "document.getElementById(PPECodeID).click();document.getElementById(PPEImageID).click();status = 'True';return status;break;"
                       + "}}}return data;").ToString();
                return status;

            }
            catch (Exception e)
            {
                Operation.FailerReason = "Not able to validate the order status in Mobile Picking";
                genreport.Reports(steps, description, typeofoperation, false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                genreport.FileCorreptionCheck();
                throw new NoSuchControlTypeFound("Not able to validate the order status in Mobile Picking");
            }
        }
        #endregion

        #region Delete Data By Navigate through the pages
        /// <summary>
        /// This will delete the data in the table by navigating through the pages.
        /// </summary>
        /// <param name="step">Step Number.</param>
        /// <param name="description">Description about the step.</param>
        /// <param name="PageID">Here this will accept the table footer id value of the DIV tag that contains the pages</param>
        /// <param name="data">Data You Want to delete or edit</param>
        /// <param name="typeofcontrol">Type of control of the data you want to delete</param>
        /// <param name="codinates">coordinate to move to a control.</param>
        public void NavigationPageDeleteData(string step, string description, string PageID, string data, string typeofcontrol, params int[] codinates)
        {
            BrowserWindow window = new BrowserWindow();
            string lastPage = string.Empty;
            bool availabledata = false;
            int count = 0;
            try
            {
                try
                {
                    count = Convert.ToInt16(window.ExecuteScript("var data=document.getElementById('" + PageID + "').getElementsByTagName('a');return  data.length"));
                }
                catch (Exception e) { }
                string idval = PageID.Split('_')[0];
                if (count > 3)
                {
                    if (count > 5)
                    {
                        lastPage = idval + "_hrefLast";
                        count = 6;
                    }
                    else
                    {
                        count = 4;
                    }
                }
                if (count > 0)
                {
                    for (int i = 1; i <= count; i++)
                    {
                        string id = idval + "_hrefNum" + i;
                        window.ExecuteScript("document.getElementById('" + id + "').scrollIntoView(true);");
                        Pop.OperationStart("Click", "HtmlHyperlink", "ID", id, "", "Web", step, description + i);
                        Thread.Sleep(max);
                        Thread.Sleep(max);
                        try
                        {
                            Pop.OperationStart("ClickEdit", typeofcontrol, "SearchText", data, "", "Web", step, description, codinates);
                            availabledata = true;
                            break;
                        }
                        catch (Exception e) { }
                    }
                }
                else
                {
                    try
                    {
                        Pop.OperationStart("ClickEdit", typeofcontrol, "SearchText", data, "", "Web", step, description, codinates);
                        availabledata = true;
                    }
                    catch (Exception e) { }
                }
                if (!availabledata)
                {
                    Operation.FailerReason = "Given data is not available in the page";
                    genreport.Reports(step, description, "SearchDataOperation", false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                    Assert.Fail("Given data is not available in the page");
                }
            }
            catch (Exception e)
            {
                Operation.FailerReason = "Given data is not available in the page";
                genreport.Reports(step, description, "SearchDataOperation", false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                genreport.FileCorreptionCheck();
                Assert.Fail("Given data is not available in the page");
            }
        }
        #endregion

        #region DeviceManagement Operations
        /// <summary>
        /// Getting the device ID while adding the device information in Device->Device Management by matching with the data that we added.
        /// </summary>
        /// <param name="step">Step Number.</param>
        /// <param name="description">Description about the step.</param>
        /// <param name="refference">here we want to give all the reference for the device information added separated by ":" Ex:Device Type:Version No:Manufacturer:Model</param>
        /// <param name="TableID">Table ID Value of the device information stored Must give in ControlKeywordValue</param>
        /// <returns></returns>
        public string GETDeviceID(string step, string description, string refference, string TableID)
        {
            BrowserWindow window = new BrowserWindow();
            string[] values = refference.Split(':');
            try
            {
                string deviceid = window.ExecuteScript(@"var data = getDeviceId();
                                function getDeviceId(){"
                                + "var deviceNo = document.getElementById('" + TableID + "');"
                                + "var count = deviceNo.getElementsByTagName('tr').length;"
                                + "for(var i=0; i<count; i++){"
                                + "var id = deviceNo.getElementsByTagName('tr')[i].getAttribute('deviceinfoid');"
                                + "var devicetype = document.getElementById('ucDeviceInfoControl_' + id + '_tdDeviceType').innerHTML;"
                                + "var deviceVersion = document.getElementById('ucDeviceInfoControl_' + id + '_tdVersion').innerHTML;"
                                + "var deviceManufacture = document.getElementById('ucDeviceInfoControl_' + id + '_tdManufacturer').innerHTML;"
                                + "var deviceModel = document.getElementById('ucDeviceInfoControl_' + id + '_tdModel').innerHTML;"
                                + "if(devicetype=='" + ExcelDataTable.ReadData(1, values[0]) + "'){if(deviceVersion=='" + ExcelDataTable.ReadData(1, values[1]) + "'){if(deviceManufacture=='" + ExcelDataTable.ReadData(1, values[2]) + "'){if(deviceModel=='" + ExcelDataTable.ReadData(1, values[3]) + "'){break;}}}else{var id=0;}}else{var id=0;}"
                                + "}return id;}return data;").ToString();
                return deviceid;
            }
            catch (Exception e) {
                Operation.FailerReason = "Given data is not available in the page";
                genreport.Reports(step, description, "SearchDataOperation", false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                genreport.FileCorreptionCheck();
                throw new NoSuchControlTypeFound("Given data is not available in the page");
            }
        }
        #endregion

        #region For Getting Remaining Quantity of the order
        /// <summary>
        /// This is used to get the remaining quantity or the particular order form the picking screen and remaining quantity will return
        /// </summary>
        /// <param name="ControlKeyValue">Id  Value of the table in the picking screen in dispatch Mobile</param>
        /// <param name="DataRefference">This will get the order number</param>
        /// <param name="Step">step number</param>
        /// <param name="Description">description for the step number</param>
        /// <param name="keyword">"GetRemainingPickQty" this will used as keyword.</param>
        /// <returns></returns>
        public string GetRemainingQuantity(string ControlKeyValue, string DataRefference, string Step, string Description, string keyword)
        {
            BrowserWindow window = new BrowserWindow();
            try
            {
                string remainingQty = window.ExecuteScript(@"var data = enterRemQt();
                                    function enterRemQt(){"
                                    + "var LineItemCount = document.getElementById('"+ ControlKeyValue +"').getElementsByTagName('tr').length-1;"
                                    + "for(var i=1; i<=LineItemCount; i++){"
                                    + "var order = document.getElementById('"+ ControlKeyValue +"').getElementsByTagName('tr')[i].getElementsByTagName('td')[0].innerHTML;"
                                    + "if(order=='"+ DataRefference +"'){"
                                    + "var remQty = document.getElementById('"+ControlKeyValue+ "').getElementsByTagName('tr')[i].getElementsByTagName('td')[1].innerHTML;return remQty;}"
                                    + "else{remQty='Null'; return remQty;}"
                                    + "}}return data;").ToString();
                return remainingQty;
            }
            catch(Exception e)
            {
                Operation.FailerReason = "Error While getting Remaining Quantity";
                genreport.Reports(Step, Description,keyword , false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                genreport.FileCorreptionCheck();
                throw new NoSuchControlTypeFound("Given data is not available in the page,Error While getting Remaining Quantity");
            }
        }

        #endregion

        #region For recording Batch number and Required Quantity based on order and Material
        /// <summary>
        /// Getting the required Quantity and batch number based on the order and material. in Dispatch Orders > PickList View
        /// for each Material and order you have to write steps. to record the details.
        /// </summary>
        /// <param name="ControlKeyValue">Example: "PurchaseOrder_No:Material1", First you have to give reference to the Order Number, and second reference to the Material separated by ":"</param>
        /// <param name="Step">step number</param>
        /// <param name="Description">description for the step number</param>
        /// <param name="keyword">"GetMaterialBatchData" this will used as keyword.</param>
        /// <returns></returns>
        public string GetBatchReqQuantitry(string ControlKeyValue, string Step, string Description, string keyword)
        {
            BrowserWindow window = new BrowserWindow();
            string[] resource = ControlKeyValue.Split(':');
            try
            {
                string BatchQuantity = window.ExecuteScript(@"var data = getBatchData();
                                                function getBatchData(){"
                                                + "var ordernumber='" + ExcelDataTable.ReadData(1,resource[0]) + "';"
                                                + "var Material='" + ExcelDataTable.ReadData(1, resource[1]) + "';"
                                                + "var batch = document.getElementById(ordernumber+Material+'_spnBatchName').innerHTML;"
                                                + "var reqQty = document.getElementById(ordernumber+Material+'_spnReqQty').innerHTML;"
                                                + "var returndata=batch+':'+reqQty;return returndata;}return data;").ToString();
                return BatchQuantity;
            }
            catch(Exception e)
            {
                Operation.FailerReason = "Error While getting Remaining Quantity and Batch For the specified Material";
                genreport.Reports(Step, Description, keyword, false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                genreport.FileCorreptionCheck();
                throw new NoSuchControlTypeFound("Error While getting Remaining Quantity and Batch For the specified Material");
            }
        }
        #endregion

        #region For Getting Batch number and Required Quantity based on order and Material and StoreID
        /// <summary>
        /// Getting the required Quantity and batch number based on the order and material and Store ID. in Dispatch Orders > PickList View
        /// for each Material and order you have to write steps. to record the details.
        /// </summary>
        /// <param name="ControlKeyValue">Write the ID Value of the table in the page</param>
        /// <param name="Step">step number</param>
        /// <param name="Description">description for the step number</param>
        /// <param name="keyword">"GetMaterialBatchStoreID" this will used as keyword.</param>
        /// <returns>return result like below.
        /// Result: Store: 1004-1162941-42203713-12,Store: 1004-624086-705653-12,Store: 1119-1162941-42203713-12,</returns>
        public string GetMaterialDetialsWithStoreID(string ControlKeyValue, string Step, string Description, string keyword)
        {
            BrowserWindow window = new BrowserWindow();
            try
            {
                string MaterialDetails = window.ExecuteScript(@"var data = getMaterialData();
                                        function  getMaterialData(){"
                                        + "var table = document.getElementById('" + ControlKeyValue + "');"
                                        + "var len = table.getElementsByTagName('tr').length;var material='';var Store='';"
                                        + "for(var i=0; i<len; i++){"
                                        + "Store = table.getElementsByTagName('tr')[0].innerText;"
                                        + "if(!Store.startsWith('Store')){break;}"
                                        + "else{"
                                        + "if(table.getElementsByTagName('tr')[i].getAttribute('style')=='color: rgb(200, 5, 4); font-weight: bold;'){"
                                        + "var material = material + StoreValue + '-' + table.getElementsByTagName('tr')[i].getElementsByTagName('td')[0].innerText.trim();"
                                        + "material = material + '-' + table.getElementsByTagName('tr')[i].getElementsByTagName('td')[2].innerText.trim();"
                                        + "material = material + '-' + table.getElementsByTagName('tr')[i].getElementsByTagName('td')[3].innerText.trim() + ', ' ;"
                                        + "}else{"
                                        + "if(table.getElementsByTagName('tr')[i].getElementsByTagName('td')[0].getAttribute('class')=='nopadding blue'){"
                                        + "StoreValue = table.getElementsByTagName('tr')[i].innerText;}else{continue;}}}}return material;}return data;").ToString();
                return MaterialDetails;
            }
            catch (Exception e)
            {
                Operation.FailerReason = "Error While getting Remaining Quantity and Batch For the specified Material based on Store ID";
                genreport.Reports(Step, Description, keyword, false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                genreport.FileCorreptionCheck();
                throw new NoSuchControlTypeFound("Error While getting Remaining Quantity and Batch For the specified Material based on Store ID");
            }
        }
        #endregion

        #region Enter proposed bin in the picking screen
        /// <summary>
        /// This will enter the bin in the picking screen based on the material what we are picking. it will check for the proposed bin and 
        /// then check that the available Qty is greater than available quantity. if not it will check for the other  bin satisfies the above condition.
        /// </summary>
        /// <param name="stepno">Step number</param>
        /// <param name="description">Description for the steps</param>
        /// <param name="keyword">Keyword for entering the bin in picking screen. "EnterSelectedBin"</param>
        /// <param name="controlKeyword">Property of the control that we are referring</param>
        /// <param name="typeofControl">Type of the control that we enter the bin </param>
        /// <param name="typeOfWindow">type of the technology we are using like window or web</param>
        /// <param name="ControlKwywordValue">Property value of the control the you specified in the ControlKeyword</param>
        /// <param name="dataReferenceKeyword">Reference to the required quantity in recorded excel. Ex. Order=M1:Required Quantity
        /// Order: This represents Heading Given in the dynamic created excel. (this won't change)
        /// M1: This represents order of the material according to picking screen.
        /// Required Quantity: This represents reference to data that you want to access in excel.
        /// </param>
        public string EnterProposedBin(string stepno, string description, string keyword, string controlKeyword, string typeofControl, string typeOfWindow, string ControlKeywordValue, string dataReferenceKeyword)
        {
            BrowserWindow window = new BrowserWindow();
            string store = string.Empty;
            string material = string.Empty;
            if (WriteAndReadData.DataFilePath == string.Empty)
            {
                WriteAndReadData.DataFilePath = ConfigurationManager.AppSettings["ExcelDataFile"] + "\\DynamicData" + "\\" + batchforReport  + ".xls";
            }
            ExcelDataTable.PopulateRecordData(WriteAndReadData.DataFilePath);
            string reqQuantity = Pop.getDataFromDynamicExcel(dataReferenceKeyword);
            try
            {
                string BinForMaterial = window.ExecuteScript(@"var data = enterBinLoaction()
                                            function enterBinLoaction(){"
                                           + "var availBindetails = document.getElementById('tableAvailableBins').getElementsByTagName('tbody')[0];"
                                           + "var len = availBindetails.getElementsByTagName('tr').length;"
                                           + "for(var i =0; i<len; i++){"
                                           + "if(availBindetails.getElementsByTagName('tr')[i].getElementsByTagName('td')[0].contains(document.getElementsByClassName('arrow-right-black-small')[0])){"
                                           + "if(availBindetails.getElementsByTagName('tr')[i].getElementsByTagName('td')[2].innerHTML>"+reqQuantity+"){"
                                           + "return availBindetails.getElementsByTagName('tr')[i].getElementsByTagName('td')[1].innerHTML + ':' + availBindetails.getElementsByTagName('tr')[i].getElementsByTagName('td')[2].innerHTML;}"
                                           + "else{for(var i =0; i<len; i++){"
                                           + "if(availBindetails.getElementsByTagName('tr')[i].getElementsByTagName('td')[2].innerHTML>" + reqQuantity + "){"
                                           + "return availBindetails.getElementsByTagName('tr')[i].getElementsByTagName('td')[1].innerHTML + ':' + availBindetails.getElementsByTagName('tr')[i].getElementsByTagName('td')[2].innerHTML;}}}}}}return data;").ToString();
                Pop.OperationStart("EnterText", typeofControl, controlKeyword, ControlKeywordValue, BinForMaterial.Split(':')[0], typeOfWindow,stepno,description);
                return BinForMaterial;
            }
            catch (Exception e)
            {
                Operation.FailerReason = "Error While getting bin location For the specified Material in picking screen";
                genreport.Reports(stepno, description, keyword, false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                genreport.FileCorreptionCheck();
                throw new NoSuchControlTypeFound("Error While getting bin location For the specified Material in picking screen");
            }
        }
        #endregion

        #region Validate Material Info for Return To Warehouse
        /// <summary>
        /// This will validate the material details are showing properly based on the given details in the data reference keyword in dispatch mobile.
        /// please provide the data reference keyword in order like
        /// 1. Reference to the material
        /// 2. Reference to the Batch Number of the material
        /// 3. Reference to the excess quantity.
        /// separated by comma.
        /// </summary>
        /// <param name="stepno">Step number of this operation</param>
        /// <param name="description">Description for the step</param>
        /// <param name="keyword">"ValidateReturnWarehouseInfo" keyword for this method access</param>
        /// <param name="ControlKeywordValue">id of the table that showing the details</param>
        /// <param name="DatareferenceKeyword">eg: "Order=M1:Material,Order=M1:Batch Number,Excess Quantity"</param>
        /// <returns>this will return true of false value</returns>
        public bool ValidateMaterialInfoINReturnToWarehouse(string stepno, string description, string keyword, string ControlKeywordValue, string DatareferenceKeyword)
        {
            BrowserWindow window = new BrowserWindow();
            if (WriteAndReadData.DataFilePath == string.Empty)
            {
                WriteAndReadData.DataFilePath = ConfigurationManager.AppSettings["ExcelDataFile"] + "\\DynamicData" + "\\" + batchforReport + ".xls";
            }
            ExcelDataTable.PopulateRecordData(WriteAndReadData.DataFilePath);
            string[] validatingResources = DatareferenceKeyword.Split(',');
            string material = ExcelDataTable.ReadRefferenceInfoByOrder(validatingResources[0].Split('=')[1].Split(':')[0].ToString(), validatingResources[0].Split('=')[1].Split(':')[1].ToString());
            string BatchNumber = ExcelDataTable.ReadRefferenceInfoByOrder(validatingResources[1].Split('=')[1].Split(':')[0].ToString(), validatingResources[1].Split('=')[1].Split(':')[1].ToString());
            string excessQuantity = Convert.ToString(Convert.ToInt16(ExcelDataTable.ReadData(1, validatingResources[2].Split('+')[0])) + Convert.ToInt16(Pop.getDataFromDynamicExcel(validatingResources[2].Split('+')[1])));
            try
            {
                string status = window.ExecuteScript(@"var data = returnToWarehouseValidation();
                                function returnToWarehouseValidation(){"
                                + "var RtableBody = document.getElementById('" + ControlKeywordValue + "').getElementsByTagName('tbody')[0];"
                                + "var TRlength = RtableBody.getElementsByTagName('tr').length;"
                                + "for(var i=0; i<TRlength; i++){"
                                + "if(RtableBody.getElementsByTagName('tr')[i].getElementsByTagName('td')[1].innerHTML=='" + material + "'){"
                                + "if(RtableBody.getElementsByTagName('tr')[i].getElementsByTagName('td')[2].innerHTML=='"+BatchNumber+"'){"
                                + "if(RtableBody.getElementsByTagName('tr')[i].getElementsByTagName('td')[3].innerHTML=='"+excessQuantity+"'){"
                                + "var value=true;break;}else{var value = false; break;}}else{continue;}}else{continue;}}"
                                + "return value;}return data;").ToString();
                return Convert.ToBoolean(status);
            }
            catch(Exception e)
            {
                Operation.FailerReason = "Error While Validating return to warehouse Material Details";
                genreport.Reports(stepno, description, keyword, false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                genreport.FileCorreptionCheck();
                throw new NoSuchControlTypeFound("Error While Validating return to warehouse Material Details");
            }

        }
        #endregion

        #region Validate Picked Bin Location Status
        /// <summary>
        /// After Picking from the bin the color of that bin will change, to check the color of that bin or status of that bin we using this one. 
        /// Need to pass the expected status of the bin from the resource file.
        /// </summary>
        /// <param name="stepno">Step Number</param>
        /// <param name="description">Description for the step</param>
        /// <param name="keyword">"ValidatePickedBinStatus"</param>
        /// <param name="ControlKeywordValue">Reference to the expected status.</param>
        /// <returns>return the status of the specified bin, if av available otherwise return true.</returns>
        public string GetPickedBinStatusColor(string stepno, string description, string keyword, string ControlKeywordValue)
        {
            BrowserWindow window = new BrowserWindow();
            Pop.getDataFromDynamicExcel(ControlKeywordValue);
            try
            {
                string status = window.ExecuteScript(@"var data = pickedBinStatus();
                                                function pickedBinStatus(){"
                                            + "var bins = document.getElementById('tableAvailableBins').getElementsByTagName('tbody')[0];"
                                            + "var bincount=bins.getElementsByTagName('tr').length;"
                                            + "for(var i=0; i<bincount; i++){"
                                            + "if(bins.getElementsByTagName('tr')[i].getElementsByTagName('td')[1].innerHTML=='" + ControlKeywordValue + "'){"
                                            + "var status=bins.getElementsByTagName('tr')[i].getAttribute('style');break;}else{var status=true;}}"
                                            + "return status;}return data;").ToString();
                return status;
            }
            catch (Exception e)
            {
                Operation.FailerReason = "Error While Getting the bin Status";
                genreport.Reports(stepno, description, keyword, false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                genreport.FileCorreptionCheck();
                throw new NoSuchControlTypeFound("Error While Getting the bin Status");
            }
        }
        #endregion

        #region Validate Matching Row Style
        /// <summary>
        /// This will be used as the common function to check the status of the rows in the table like picking stated.
        /// We have to specify the table ID value and matching data that you want to check for the status in the table separated by comma, as "ControlKeywordValue" 
        /// Reference to the expected status of the data. as "DataReferenceKeyWord".
        /// </summary>
        /// <param name="stepno">Step Number</param>
        /// <param name="description">Description for the step Number</param>
        /// <param name="keyword">"ValidateMatchDataStatus"</param>
        /// <param name="ControlKeywordValue">Eg: tableContainerList(id value of the Table),Store=Store1:Container Number(Reference to matching data.)</param>
        /// <returns>return the status of the matched data</returns>
        public string GetMatchTableDataStatusColor(string stepno, string description, string keyword, string ControlKeywordValue)
        {
            BrowserWindow window = new BrowserWindow();
            string matchValue = string.Empty;
            if (!ControlKeywordValue.Split(',')[1].Contains('='))
            {
                matchValue = ExcelDataTable.ReadData(1, ControlKeywordValue.Split(',')[1]);
            }
            else
            {
                matchValue = Pop.getDataFromDynamicExcel(ControlKeywordValue.Split(',')[1]);
            }
            try
            {
                string status = window.ExecuteScript(@"var data= getStyle();
                                                function getStyle(){"
                                            + "var table = document.getElementById('"+ ControlKeywordValue.Split(',')[0] + "').getElementsByTagName('tbody')[0];"
                                            + "var count = table.getElementsByTagName('tr').length;"
                                            + "for(var i=0; i<count; i++){"
                                            + "if(table.getElementsByTagName('tr')[i].getElementsByTagName('td')[1].innerText == '"+matchValue+"'){"
                                            + "var status = table.getElementsByTagName('tr')[i].getAttribute('style');break;}}"
                                            + "return status;}return data;").ToString();
                return status;
            }
            catch (Exception e)
            {
                Operation.FailerReason = "Error While Getting the match Data Status";
                genreport.Reports(stepno, description, keyword, false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                genreport.FileCorreptionCheck();
                throw new NoSuchControlTypeFound("Error While Getting match Data Status");
            }
        }
        #endregion

        #region validate Total Quantity in Dispatch Pick Area Stock Info 
        /// <summary>
        /// This will check the Dispatch Pick Area Stock Info page that zero total quantity coming or not while Include materials with no stock in dispatch area "Include materials with no stock in dispatch area" is unchecked
        /// </summary>
        /// <param name="step">Step number</param>
        /// <param name="description">Description for the step</param>
        /// <param name="PageID">Page id of the table</param>
        public void ValiateZero(string step, string description, string PageID)
        {
            BrowserWindow window = new BrowserWindow();
            int count = 0;
            try
            {
                count = Convert.ToInt16(window.ExecuteScript("var data=document.getElementById('" + PageID + "').getElementsByTagName('a');return  data.length"));
                string idval = PageID.Split('_')[0];
                for (int i = 1; i <= count; i++)
                {
                    string id = idval + "_hrefNum" + i;
                    window.ExecuteScript("document.getElementById('" + id + "').scrollIntoView(true);");
                    Pop.OperationStart("Click", "HtmlHyperlink", "ID", id, "", "Web", step, description + i);
                    Thread.Sleep(max);
                    Thread.Sleep(max);
                    int tablerow = Convert.ToInt16(window.ExecuteScript("var tablrow=document.getElementsByTagName('tbody')[0].getElementsByTagName('tr').length; return tablrow"));
                    for(int j=01; j<=tablerow; j++)
                    {
                        double actualQty = Convert.ToDouble(window.ExecuteScript("var totalQty=document.getElementById('ctl'+'" + j.ToString("00") + "'+'_lblTotalQty').innerHTML; return totalQty;"));
                        if(actualQty <=0)
                        {
                            string path = ConfigurationManager.AppSettings["ScreenShot"];
                            path = path + @"\" + LoginOperatrion.ProjectName + "";
                            Directory.CreateDirectory(path);
                            try
                            {
                                Image image = window.CaptureImage();
                                image.Save(path + "\\" + "ZeroTotalQuantityShowingIn " + i + ".jpeg", ImageFormat.Jpeg);
                                Operation.ErrorScreenPath = path + "\\" + "ZeroTotalQuantityShowingIn " + i + ".jpeg";
                                image.Dispose();
                            }
                            catch (Exception v) { }
                            Operation.FailerReason = "Total Quantity is not showing as Zero in " + i + " Page";
                            genreport.Reports(step, "Total Quantity is showing as Zero in " + i + " Page", "ZeroTotalQtyCheck", false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                            Assert.Fail("Total Quantity is showing as Zero in " + i + " Page");
                        }
                    }
                    genreport.Reports(step, "Total Quantity is not showing as Zero in " + i + " Page", "ZeroTotalQtyCheck", true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                }
            }
            catch (Exception e)
            {
                int tablerow = Convert.ToInt16(window.ExecuteScript("var tablrow=document.getElementsByTagName('tbody')[0].getElementsByTagName('tr').length; return tablrow"));
                for (int j = 01; j <= tablerow; j++)
                {
                    int actualQty = Convert.ToInt32(window.ExecuteScript("var totalQty=document.getElementById('ctl'+'" + j.ToString("00") + "'+'_lblTotalQty').innerHTML; return totalQty;"));
                    if (actualQty <= 0)
                    {
                        string path = ConfigurationManager.AppSettings["ScreenShot"];
                        path = path + @"\" + LoginOperatrion.ProjectName + "";
                        Directory.CreateDirectory(path);
                        try
                        {
                            Image image = window.CaptureImage();
                            image.Save(path + "\\" + "ZeroTotalQuantityShowingInCurrentPage.jpeg", ImageFormat.Jpeg);
                            Operation.ErrorScreenPath = path + "\\" + "ZeroTotalQuantityShowingInCurrentPage.jpeg";
                            image.Dispose();
                        }
                        catch (Exception v) { }
                        Operation.FailerReason = "Total Quantity is not showing as Zero in Current Page";
                        genreport.Reports(step, "Total Quantity is showing as Zero in current Page", "ZeroTotalQtyCheck", false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                        Assert.Fail("Total Quantity is showing as Zero in current Page");
                    }
                }
            }
        }
        #endregion

        #region Validate Transfer Stock Material And Quantity
        /// <summary>
        /// Validate the stock transfer details like material and batch is matching or not.
        /// </summary>
        /// <param name="stepno">Number of Step</param>
        /// <param name="description">Description For the step.</param>
        /// <param name="keyword">"WarehouseStockValidation"</param>
        /// <param name="ControlKeywordValue">ID value of the table displayed in the Transfer stock in Dispatch Mobile</param>
        /// <param name="SearchData">Reference to the Material and the Batch number Separated By ":"</param>
        /// <returns>This will TReturn "true" or "false" according to the matching data availability</returns>
        public bool ValidateTransferStockMaterialAndQty(string stepno, string description, string keyword, string ControlKeywordValue, string SearchData)
        {
            BrowserWindow window = new BrowserWindow();
            ExcelDataTable.PopulateRecordData(WriteAndReadData.DataFilePath);
            string Material = ExcelDataTable.ReadData(1, SearchData.Split('+')[0]);
            string batch = Pop.getDataFromDynamicExcel(SearchData.Split('+')[1]);
            try
            {
                string status = window.ExecuteScript(@"var data=ValidateMaterialandQTY();
                        function ValidateMaterialandQTY(){
                        var status=false;
                        var tabledata=document.getElementById('" + ControlKeywordValue + "').getElementsByTagName('tbody')[0];"
                        + "var count=tabledata.getElementsByTagName('tr').length;"
                        + "if(count>1){status = false;}else{"
                        + "if(tabledata.getElementsByTagName('tr')[0].getElementsByTagName('td')[1].innerHTML.trim()=='" + Material + "'){"
                        + "if(tabledata.getElementsByTagName('tr')[0].getElementsByTagName('td')[2].innerHTML.trim()=='" + batch + "'){"
                        + "status = true;}}else{status = false;}}return status;}return data;").ToString();

                return  Convert.ToBoolean(status);
            }
            catch(Exception e)
            {
                Operation.FailerReason = "Error While Getting the matching Material";
                genreport.Reports(stepno, description, keyword, false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                genreport.FileCorreptionCheck();
                throw new NoSuchControlTypeFound("Error While Getting matching Material");
            }
        }
        #endregion

        #region Launch SQLDeveloperApplication
        /// <summary>
        /// This will Launch The SQlDeveloperApplication, Invoking Keyword..
        /// </summary>
        public void sqldeveloperApplicationLaunch()
        {
            WinWindow wind = new WinWindow();
            wind.SearchProperties[WinWindow.PropertyNames.ClassName] = "SunAwtFrame";
            UITestControlCollection wndctrl = wind.FindMatchingControls();
            string ExEpath = ExcelDataTable.ReadData(1, "SqldeveloperPath");
            ApplicationUnderTest application = ApplicationUnderTest.Launch(ExEpath);
        }
        #endregion

        #region Dispatch Report 
        /// <summary>
        /// This will get the Row number of the delivery order and return the number. Also it will check the order is availability of and start time end time and time taken for the picking or packing if any this is not available it will throw error.
        /// </summary>
        /// <param name="ControlKeywordValue">table id value will be given here.</param>
        /// <param name="orderNumber">Reference to the order number from the test data.</param>
        /// <param name="steps">Number of step.</param>
        /// <param name="description">Description for the step.</param>
        /// <param name="keyword">"GetDeliveryOrdReport", This keyword will access this function also write the return data into the dynamic excel under "Dispatch Report Order" heading.</param>
        /// <returns></returns>
        public string GetDispatchReportAvailability(string ControlKeywordValue, string orderNumber, string steps, string description, string keyword)
        {
            try
            {
                BrowserWindow window = new BrowserWindow();
                string ordercount = window.ExecuteScript(@"var data = getdeliveryorder();
                                function getdeliveryorder(){"
                                + "var dcount = document.getElementById('" + ControlKeywordValue + "').childElementCount/2;"
                                + "var fcount=0;"
                                + "for(var i=1;i<=dcount;i++){"
                                + "var orderNo = document.getElementById('ctl0'+ i  + '_tdDeliveryOrder').innerHTML;"
                                + "if(orderNo=='" + ExcelDataTable.ReadData(1, orderNumber) + "'){"
                                + "if(document.getElementById('ctl0' + i + '_tdStartDate').innerHTML!=''){"
                                + "if(document.getElementById('ctl0' + i + '_tdDeliveryTotalTimeInMin').innerHTML!=''){"
                                + "var fcount = i; break;}else{continue;}}else{continue;}}else{continue;}}return fcount;}return data;").ToString();
                return ordercount;
            }
            catch (Exception e)
            {
                Operation.FailerReason = "Not able to get Dispatch report details from Dispatch Report Page";
                genreport.Reports(steps, description, keyword, false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                genreport.FileCorreptionCheck();
                throw new NoSuchControlTypeFound("Not able to get Dispatch report details from Dispatch Report Page");
            }
        }
        #endregion

        #region Dispatch Report Start & End Time & Date Availability
        /// <summary>
        /// This will Check All the Start time, End Time, Start Date, End Date, all these are empty or not for all the material.
        /// </summary>
        /// <param name="ControlKeywordValue">Here we will give reference from the Dynamic Data Excel. Return by the keyword "GetDeliveryOrdReport". Eg. "Store=Store2:Dispatch Report Order"</param>
        /// <param name="orderNumber">Reference to the order number from the test data.</param>
        /// <param name="steps">Number of step.</param>
        /// <param name="description">Description for the step.</param>
        /// <param name="keyword">"DateTimeReportValidate" This will access this function.</param>
        /// <returns>Boolean value true or false</returns>
        public bool getdispatchDateTimeValidation(string ControlKeywordValue, string orderNumber, string steps, string description, string keyword)
        {
            try
            {
                if (WriteAndReadData.DataFilePath == string.Empty)
                {
                    WriteAndReadData.DataFilePath = ConfigurationManager.AppSettings["ExcelDataFile"] + "\\DynamicData" + "\\" + batchforReport + ".xls";
                }
                ExcelDataTable.PopulateRecordData(WriteAndReadData.DataFilePath);
                string value = Pop.getDataFromDynamicExcel(ControlKeywordValue);

                BrowserWindow window = new BrowserWindow();
                bool status = Convert.ToBoolean(window.ExecuteScript(@"var data=getStartEndTimeDate();
                                    function getStartEndTimeDate(){"
                                    + "var reportOrd='"+ value + "';"
                                    + "for(var i=0; i<3; i++){"
                                    + "if(document.getElementById('ctl0' + reportOrd +  '_ctl0' + i + '_tdStartDate').innerHTML!=''){"
                                    + "if(document.getElementById('ctl0' + reportOrd +  '_ctl0' + i + '_tdStartTime').innerHTML!=''){"
                                    + "if(document.getElementById('ctl0' + reportOrd +  '_ctl0' + i + '_tdEndDate').innerHTML!=''){"
                                    + "if(document.getElementById('ctl0' + reportOrd +  '_ctl0' + i + '_tdEndTime').innerHTML!=''){"
                                    + "var status=true;"
                                    + "}else{var status=false;break;}"
                                    + "}else{var status=false;break;}"
                                    + "}else{var status=false;break;}"
                                    + "}else{var status=false;break;}"
                                    + "}return status;}return data;"));
                return status;
            }
            catch (Exception e)
            {
                Operation.FailerReason = "Not able to get Dispatch Timings from report details from Dispatch Report Page";
                genreport.Reports(steps, description, keyword, false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                genreport.FileCorreptionCheck();
                throw new NoSuchControlTypeFound("Not able to get Dispatch Timings from report details from Dispatch Report Page");
            }
        }
        #endregion
        /// <summary>
        /// This will validate B2B dropped or not also can validate the source bin and target bin is matching or not in View SAP Transaction page. if not matching will return false value. 
        /// </summary>
        /// <param name="ControlKeywordValue"></param>
        /// <param name="Resources">This will contain the references of the Dispatch Pack area and Dispatch Revert Area separated by ':'</param>
        /// <param name="steps">Step Number will specify here.</param>
        /// <param name="description">Description for the step</param>
        /// <param name="keyword">"B2BSAPValidation" this keyword will used to perform this operation</param>
        /// <returns>Boolean value true or false</returns>
        #region B2B Drop Validation
        public bool B2BDropValidation(string ControlKeywordValue, string Resources, string steps, string description, string keyword)
        {
            try
            {
                string[] data = Resources.Split(':');
                BrowserWindow window = new BrowserWindow();
                bool status = Convert.ToBoolean(window.ExecuteScript(@"var data = B2BTransact();
                            function B2BTransact(){"
                            + "var data = document.getElementById('" + ControlKeywordValue + "').getElementsByTagName('tr')[0].getElementsByTagName('td')[2].innerHTML;"
                            + "if(document.getElementById('ucSAPTransactionsQueSummary_' + data + '_tdSourceBin').innerHTML == '"+ ExcelDataTable.ReadData(1, data[0])+"'){"
                            + "if(document.getElementById('ucSAPTransactionsQueSummary_' + data + '_tdToBinLocation').innerHTML == '"+ExcelDataTable.ReadData(1,data[1]) +"'){"
                            + "var returndata=true;"
                            + "}else{var returndata=false;}"
                            + "}else{var returndata=false;}"
                            + "return returndata;}return data;"));
                return status;
            }
            catch(Exception e)
            {
                Operation.FailerReason = "Validate B2B transaction Failed";
                genreport.Reports(steps, description, keyword, false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                genreport.FileCorreptionCheck();
                throw new NoSuchControlTypeFound("Validate B2B transaction Failed");
            }
        }
        #endregion
        /// <summary>
        /// This will generate a random number in string format and return.
        /// </summary>
        /// <param name="numberCount">here you can specify the number of characters or Digits for a number want</param>
        /// <returns></returns>
        #region RandomNumberGen
        public string RamdomBatch(int numberCount)
        {
            try
            {
                StringBuilder builder = new StringBuilder(6);
                string numberAsString = string.Empty;

                for (var i = 0; i < numberCount; i++)
                {
                    builder.Append(_numbers[random.Next(0, _numbers.Length)]);
                }

                numberAsString = builder.ToString();
                return numberAsString;
            }
            catch(Exception e)
            {
                Operation.FailerReason = "Random Number generation failed";
                genreport.FileCorreptionCheck();
                throw new NoSuchControlTypeFound("Random Number generation failed");
            }
            
        }
        #endregion

        #region Get Defaults settings
        /// <summary>
        /// This will find which one is the default data set up  if no default data is there it will it will setup first one as default. 
        /// </summary>
        /// <param name="ControlKeywordValue">Id value of the particular table</param>
        /// <param name="steps">Step Number will specify here.</param>
        /// <param name="description">Description for the step</param>
        /// <param name="keyword">"SetDefault" this keyword will used to perform this operation</param>
        /// <returns>This one will return the default data in specified page.</returns>
        public string GetDefaultData(string ControlKeywordValue, string steps, string description, string keyword)
        {
            try
            {
                BrowserWindow window = new BrowserWindow();
                string defaultdata = window.ExecuteScript(@"var data = findDefault();
                               function findDefault(){ "
                               + " var tables = document.getElementById('" + ControlKeywordValue + "');"
                               + "var count = tables.getElementsByTagName('tr').length;var package='';"
                               + "for(var i=0; i<count; i++){"
                               + "if(tables.getElementsByTagName('tr')[i].getElementsByTagName('td')[7].hasAttribute('default')){"
                               + "var package = tables.getElementsByTagName('tr')[i].getElementsByTagName('td')[1].innerText.trim();"
                               + "break;}else {continue;}}"
                               + "if(package===''){"
                               + "package = tables.getElementsByTagName('tr')[0].getElementsByTagName('td')[7].getAttribute('id');"
                               + "pack = package + '+' + tables.getElementsByTagName('tr')[0].getElementsByTagName('td')[1].innerText.trim();return pack;"
                               + "}else{return package;}}return data;").ToString();
                if (defaultdata.Contains('+'))
                {
                    string[] resource = defaultdata.Split('+');
                    Pop.OperationStart("Click", "HtmlCell", "ID", resource[0], "", "Web", steps, description);
                    Pop.OperationStart("ClickLanguageText", "HtmlButton", "FriendlyName", "Confirm", "", "Web", steps, description);
                    Thread.Sleep(1000);
                    return resource[1];
                }
                else
                {
                    return defaultdata;
                }
            }
            catch(Exception e)
            {
                Operation.FailerReason = "Getting Default Data Failed";
                genreport.Reports(steps, description, keyword, false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                genreport.FileCorreptionCheck();
                throw new NoSuchControlTypeFound("Getting Default Data Failed");
            }
        }

        #endregion

        #region Validate Part Bag
        /// <summary>
        /// Validate the Part bag count.
        /// </summary>
        /// <param name="ControlKeywordValue">Id Value of the table Where you want to Check the Part bag icon</param>
        /// <param name="DataRefferencekeyword">Reference to the data for the part bag visibility and number of part bag to validate separated by '='</param>
        /// <param name="steps">Number of step</param>
        /// <param name="description">Description for the step</param>
        /// <param name="keyword">"PartBagValidation" This keyword will check how many part bag is available</param>
        public int getPartBagCount(string ControlKeywordValue, string DataRefferencekeyword, string steps, string description, string keyword)
        {
            try
            {
                BrowserWindow window = new BrowserWindow();
                int data = Convert.ToInt32(window.ExecuteScript(@"var data = document.getElementById('" + ControlKeywordValue + "').getElementsByClassName('" + DataRefferencekeyword + "').length; return data;"));
                return data;
            }
            catch(Exception e)
            {
                Operation.FailerReason = "Part bag Validation Failed";
                genreport.Reports(steps, description, keyword, false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                genreport.FileCorreptionCheck();
                throw new NoSuchControlTypeFound("Part bag Validation Failed");
            }
        }
        #endregion

        #region PO Availability
        public void POAvailabilityCalender(string ControlKeyword, string Keyword, string ControlKeywordValue, string Step, string Description)
        {
            BrowserWindow window = new BrowserWindow();
            try
            {
                string data = window.ExecuteScript("var data = document.getElementById('" + ControlKeywordValue + "').innerHTML; return data;").ToString();
                if (data != string.Empty)
                {
                    Thread.Sleep(mid);
                    auto.Send("{F5}");
                    Thread.Sleep(2000);
                    dt.Clear();
                    dt = ExcelKeywordTable.ExcelData(LogPath + "\\KeywordDrivenData.xlsx", "POAvailability");
                    Thread.Sleep(mid);
                    Pop.OperationStart(dt.Rows[0]["Keyword"].ToString(), dt.Rows[0]["TypeOfControl"].ToString(), dt.Rows[0]["ControlKeyword"].ToString(), dt.Rows[0]["ControlKeywordValue"].ToString(), dt.Rows[0]["DataRefferencekeyword"].ToString(), dt.Rows[0]["TypeOfWindow"].ToString(), dt.Rows[0]["Step No"].ToString(), dt.Rows[0]["Description"].ToString());
                    Pop.OperationStart(dt.Rows[1]["Keyword"].ToString(), dt.Rows[1]["TypeOfControl"].ToString(), dt.Rows[1]["ControlKeyword"].ToString(), dt.Rows[1]["ControlKeywordValue"].ToString(), ExcelDataTable.ReadData(1, dt.Rows[1]["DataRefferencekeyword"].ToString()), dt.Rows[1]["TypeOfWindow"].ToString(), dt.Rows[1]["Step No"].ToString(), dt.Rows[1]["Description"].ToString());
                    Pop.OperationStart(dt.Rows[2]["Keyword"].ToString(), dt.Rows[2]["TypeOfControl"].ToString(), dt.Rows[2]["ControlKeyword"].ToString(), dt.Rows[2]["ControlKeywordValue"].ToString(), dt.Rows[2]["DataRefferencekeyword"].ToString(), dt.Rows[2]["TypeOfWindow"].ToString(), dt.Rows[2]["Step No"].ToString(), dt.Rows[2]["Description"].ToString());
                    Pop.OperationStart(dt.Rows[3]["Keyword"].ToString(), dt.Rows[3]["TypeOfControl"].ToString(), dt.Rows[3]["ControlKeyword"].ToString(), dt.Rows[3]["ControlKeywordValue"].ToString(), dt.Rows[3]["DataRefferencekeyword"].ToString(), dt.Rows[3]["TypeOfWindow"].ToString(), dt.Rows[3]["Step No"].ToString(), dt.Rows[3]["Description"].ToString());
                }
            }
            catch(Exception e)
            {
                Operation.FailerReason = "PO Availability can not be checked";
                genreport.Reports(Step, Description, Keyword, false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                genreport.FileCorreptionCheck();
                throw new NoSuchControlTypeFound("PO Availability can not be checked");
            }
        }
        #endregion
    }
}
