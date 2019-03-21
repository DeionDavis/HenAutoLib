using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Resources;
using System.Text;
using System.Threading.Tasks;

namespace CommonLibrary
{
    public class LogLanguageTemplete
    {
        static CultureInfo ci = null;
        static ResourceManager rm = null;

        static string resMsg_LanguageUpdated = string.Empty;
        public string Msg_LanguageUpdated
        {
            get { return resMsg_LanguageUpdated; }
            set { resMsg_LanguageUpdated = value; }
        }

        static string resMsg_LogoutSuccessMessage = string.Empty;
        public string Msg_LogoutSuccessMessage
        {
            get { return resMsg_LogoutSuccessMessage; }
            set { resMsg_LogoutSuccessMessage = value; }
        }

        static string resMsg_SyncOrderMessage = string.Empty;
        public string Msg_SyncOrderMessage
        {
            get { return resMsg_SyncOrderMessage; }
            set { resMsg_SyncOrderMessage = value; }
        }

        static string resMsg_LoginFailed = string.Empty;
        public string Msg_LoginFailed
        {
            get { return resMsg_LoginFailed; }
            set { resMsg_LoginFailed = value; }
        }

        static string resMsg_WrongSecurityCode = string.Empty;
        public string Msg_WrongSecurityCode
        {
            get { return resMsg_WrongSecurityCode; }
            set { resMsg_WrongSecurityCode = value; }
        }

        public static void messageResource(string languageCode)
        {
            if (languageCode == "zh-CN")
            {
                ci = new CultureInfo(languageCode);
                rm = new ResourceManager("CommonLibrary.Resources.LogLanguageResourcesCN", Assembly.GetExecutingAssembly());
                messageInitialize();
            }
            else if (languageCode == "th-TH")
            {
                ci = new CultureInfo(languageCode);
                rm = new ResourceManager("CommonLibrary.Resources.LogLanguageResourcesTH", Assembly.GetExecutingAssembly());
                messageInitialize();
            }
            else if (languageCode == "vi-VN")
            {
                ci = new CultureInfo(languageCode);
                rm = new ResourceManager("CommonLibrary.Resources.LogLanguageResourcesVN", Assembly.GetExecutingAssembly());
                messageInitialize();
            }
            else if (languageCode == "ko-KR")
            {
                ci = new CultureInfo(languageCode);
                rm = new ResourceManager("CommonLibrary.Resources.LogLanguageResourcesKR", Assembly.GetExecutingAssembly());
                messageInitialize();
            }
            else if (languageCode == "ja-JP")
            {
                ci = new CultureInfo(languageCode);
                rm = new ResourceManager("CommonLibrary.Resources.LogLanguageResourcesJP", Assembly.GetExecutingAssembly());
                messageInitialize();
            }
            else if (languageCode == "id-ID")
            {
                ci = new CultureInfo(languageCode);
                rm = new ResourceManager("CommonLibrary.Resources.LogLanguageResourcesID", Assembly.GetExecutingAssembly());
                messageInitialize();
            }
            else
            {
                ci = new CultureInfo(languageCode);
                rm = new ResourceManager("CommonLibrary.Resources.LogLanguageResourcesEN", Assembly.GetExecutingAssembly());
                messageInitialize();
            }
        }
        public static void messageInitialize()
        {
            resMsg_LanguageUpdated = rm.GetString("resMsg_LanguageUpdated", ci).Trim();
            resMsg_LogoutSuccessMessage = rm.GetString("resMsg_LogoutSuccessMessage", ci).Trim();
            resMsg_SyncOrderMessage = rm.GetString("resMsg_SyncOrderMessage", ci).Trim();
            resMsg_LoginFailed = rm.GetString("resMsg_LoginFailed", ci).Trim();
            resMsg_WrongSecurityCode = rm.GetString("resMsg_WrongSecurityCode", ci).Trim();
        }
    }
}
