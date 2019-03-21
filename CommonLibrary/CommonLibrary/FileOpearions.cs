using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;
using Microsoft.VisualStudio.TestTools.UITest.Extension;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using CommonLibrary.Log;
using OperationLibrary;
using CommonLibrary.Reports;

namespace CommonLibrary.FileOperation
{
    public class FileOpearions
    {
        LoginOperatrion log = new LoginOperatrion();
        ReportGeneration genDetailedReport = new ReportGeneration();
        #region Downloads&FileOperations
        /// <summary>
        /// This will open the downloaded files and validate that file is downloading or not.
        /// </summary>
        /// <param name="step">Step number will pass here</param>
        /// <param name="description">Description for the step number</param>
        /// <param name="keyword">Keyword to invoke this function. KeyWword:"OpenFile" </param>
        public void OpenFile(string step, string description, string keyword, string ScreenShotNam, string AssertMsg)
        {
            if(ScreenShotNam==string.Empty)
            {
                ScreenShotNam = "Downloading_TemplateORDataFailed";
            }

            if(AssertMsg == string.Empty)
            {
                AssertMsg = "Downloading template or data failed";
            }
            try
            {
                WinControl Windowgener = (WinControl)Activator.CreateInstance(typeof(WinControl), new object[] { ParentWindow });
                Windowgener.SearchProperties[WinControl.PropertyNames.Name] = "Notification";
                Windowgener.WaitForControlExist();
                Windowgener.SearchProperties[WinControl.PropertyNames.Name] = "Notification";
                if (!Windowgener.Exists)
                {
                    Operation.ErrorScreenPath = log.screenShot(ScreenShotNam);
                    Operation.FailerReason = ". " + AssertMsg;
                    genDetailedReport.Reports(step, description, keyword, false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                    Assert.Fail(AssertMsg);
                }
                Thread.Sleep(2000);
                Windowgener.SearchProperties[WinSplitButton.PropertyNames.Name] = "Open";
                Mouse.Click(Windowgener);
                Thread.Sleep(2000);
                log.CloseFile();
                Thread.Sleep(2000);
                genDetailedReport.Reports(step, description, keyword, true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
            }
            catch (UITestControlNotFoundException e)
            {
                Operation.ErrorScreenPath = log.screenShot(ScreenShotNam);
                Operation.FailerReason = e.BasicMessage + ". " +AssertMsg;
                genDetailedReport.Reports(step, description, keyword, false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                Assert.Fail(AssertMsg);
            }
            catch (FailedToPerformActionOnHiddenControlException e)
            {
                Operation.ErrorScreenPath = log.screenShot(ScreenShotNam);
                Operation.FailerReason = e.BasicMessage + ". " + AssertMsg;
                genDetailedReport.Reports(step, description, keyword, false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                Assert.Fail(AssertMsg);
            }
        }
        /// <summary>
        /// This will save the downloaded files and open validate that file is downloading or not.
        /// the purpose of saving the file is to validate the data in that file.
        /// </summary>
        /// <param name="step">Step number will get</param>
        /// <param name="description">Description for the step number</param>
        /// <param name="keyword">Keyword to invoke this function. KeyWword:"SaveOpenFile" </param>
        public void SaveOpenFile(string step, string description, string keyword, string ScreenShotNam, string AssertMsg)
        {
            if (ScreenShotNam == string.Empty)
            {
                ScreenShotNam = "Downloading_TemplateORDataFailed";
            }
            if (AssertMsg == string.Empty)
            {
                AssertMsg = "Downloading template or data failed";
            }
            try
            {
                WinControl Windowgener = (WinControl)Activator.CreateInstance(typeof(WinControl), new object[] { ParentWindow });
                Windowgener.SearchProperties[WinControl.PropertyNames.Name] = "Notification";
                Windowgener.WaitForControlExist();
                Thread.Sleep(2000);
                Windowgener.SearchProperties[WinSplitButton.PropertyNames.Name] = "Save";
                Mouse.Click(Windowgener);
                Thread.Sleep(2000);
                Windowgener.SearchProperties[WinSplitButton.PropertyNames.Name] = "Open";
                Thread.Sleep(2000);
                Mouse.Click(Windowgener);
                genDetailedReport.Reports(step, description, keyword, true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
            }
            catch (UITestControlNotFoundException e)
            {
                Operation.ErrorScreenPath = log.screenShot(ScreenShotNam);
                Operation.FailerReason = e.BasicMessage + ". " + AssertMsg;
                genDetailedReport.Reports(step, description, keyword, false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                Assert.Fail(AssertMsg);
            }
            catch (FailedToPerformActionOnHiddenControlException e)
            {
                Operation.ErrorScreenPath = log.screenShot(ScreenShotNam);
                Operation.FailerReason = e.BasicMessage + ". " + AssertMsg;
                genDetailedReport.Reports(step, description, keyword, false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                Assert.Fail(AssertMsg);
            }
        }

        /// <summary>
        /// This will cancel the downloaded file. if you don't want to save the file or open the file.
        /// </summary>
        /// <param name="step">Step number will get</param>
        /// <param name="description">Description for the step number</param>
        /// <param name="keyword">Keyword to invoke this function. KeyWword:"CancelDownload" </param>
        public void CancelDownload(string step, string description, string keyword)
        {
            try
            {
                WinControl Windowgener = (WinControl)Activator.CreateInstance(typeof(WinControl), new object[] { ParentWindow });
                Windowgener.SearchProperties[WinControl.PropertyNames.Name] = "Notification";
                Windowgener.WaitForControlExist();
                Thread.Sleep(2000);
                Windowgener.SearchProperties[WinButton.PropertyNames.Name] = "Cancel";
                Mouse.Click(Windowgener);
                Thread.Sleep(2000);
                genDetailedReport.Reports(step, description, keyword, true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
            }
            catch (UITestControlNotFoundException e)
            {
                Operation.ErrorScreenPath = log.screenShot("Downloading_TemplateORDataFailed");
                Operation.FailerReason = e.BasicMessage + ". Downloading template or data failed";
                genDetailedReport.Reports(step, description, keyword, false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                Assert.Fail("Downloading template or data failed");
            }
        }
        #endregion

        #region basic
        private BrowserWindow mParentwindow;
        public BrowserWindow ParentWindow
        {
            get
            {
                if (this.mParentwindow == null)
                {
                    this.mParentwindow = TopParrentwindow();
                }
                return this.mParentwindow;
            }
        }
        public BrowserWindow TopParrentwindow()
        {
            BrowserWindow widow = new BrowserWindow();
            widow.SearchProperties[UITestControl.PropertyNames.ClassName] = BrowserWindow.CurrentBrowser.ToString();
            return widow;
        }
        #endregion
    }
}
