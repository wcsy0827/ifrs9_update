using System;
using System.Collections.Generic;

namespace Transfer.Utility
{
    public static class SetFile
    {
        /// <summary>
        /// 設定位置 & txtLog檔名
        /// </summary>
        static SetFile()
        {
            LoginLog = @"LoginLog.txt"; //LoginInfo儲存txtlog檔名
            LoginExceptionLog = @"LoginException.txt";//LoginException儲存txtlog檔名
            IFRS9Log = "Log";//儲存Log文件資料夾名稱
            ProgramName = "Transfer"; //專案名稱
            FileUploads = "FileUploads"; //上傳檔案放置位置
            FileDownloads = "FileDownloads"; //下載檔案放置位置
            QuantifyFile = "QuantifyFile"; //D64 Quantify檔案放置位置
            QualitativeFile = "QualitativeFile"; //D66 Qualitative檔案放置位置
            RiskControlFile = "RiskControlFile"; //D75 RiskControl檔案放置位置
            A41TransferTxtLog = @"DataRequirementsTransfer.txt"; //A41上傳Txtlog檔名
            A42TransferTxtLog = @"A42Transfer.txt"; //A42上傳Txtlog檔名
            A44_2TransferTxtLog = @"A44_2Transfer.txt";//A44_2上傳Txtlog檔名(190628換券應收未收金額修正)
            A45TransferTxtLog = @"A45Transfer.txt"; //A45上傳Txtlog檔名
            A46TransferTxtLog = @"A46Transfer.txt"; //A46上傳Txtlog檔名
            A47TransferTxtLog = @"A47Transfer.txt"; //A47上傳Txtlog檔名
            A48TransferTxtLog = @"A48Transfer.txt"; //A48上傳Txtlog檔名
            A49TransferTxtLog = @"A49Transfer.txt"; //A49上傳Txtlog檔名
            A59TransferTxtLog = @"A59Transfer.txt"; //A59上傳Txtlog檔名
            A62TransferTxtLog = @"Exhibit 7Transfer.txt"; //A62上傳Txtlog檔名
            A63TransferTxtLog = @"A63Transfer.txt"; //A63上傳Txtlog檔名
            A71TransferTxtLog = @"Exhibit29Transfer.txt"; //A71上傳Txtlog檔名
            A81TransferTxtLog = @"Exhibit10Transfer.txt"; //A81上傳Txtlog檔名
            A95TransferTxtLog = @"A95Transfer.txt"; //A59上傳Txtlog檔名
            A95_1TransferTxtLog = @"A95_1Transfer.txt"; //A59_1上傳Txtlog檔名
            A96TransferTxtLog = @"A96TransferTxtLog"; //A96上傳Txtlog檔名
            C01TransferTxtLog = @"C01Transfer.txt"; //C01上傳Txtlog檔名
            C10TransferTxtLog = @"C10Transfer.txt"; //C10上傳Txtlog檔名
            D53TransferTxtLog = @"D53Transfer.txt"; //D53上傳Txtlog檔名
            D72TransferTxtLog = @"D72Transfer.txt"; //D72上傳Txtlog檔名
        }

        public static string A41TransferTxtLog { get; private set; }
        public static string A42TransferTxtLog { get; private set; }
        public static string A44_2TransferTxtLog { get; private set; }
        public static string A45TransferTxtLog { get; private set; }
        public static string A46TransferTxtLog { get; private set; }
        public static string A47TransferTxtLog { get; private set; }
        public static string A48TransferTxtLog { get; private set; }
        public static string A49TransferTxtLog { get; private set; }
        public static string A59TransferTxtLog { get; private set; }
        public static string A62TransferTxtLog { get; private set; }
        public static string A63TransferTxtLog { get; private set; }
        public static string A71TransferTxtLog { get; private set; }
        public static string A81TransferTxtLog { get; private set; }
        public static string A95TransferTxtLog { get; private set; }
        public static string A95_1TransferTxtLog { get; private set; }
        public static string A96TransferTxtLog { get; private set; }
        public static string FileDownloads { get; private set; }
        public static string FileUploads { get; private set; }
        public static string QuantifyFile { get; private set; }
        public static string QualitativeFile { get; private set; }
        public static string RiskControlFile { get; private set; }
        public static string ProgramName { get; private set; }
        public static string C01TransferTxtLog { get; private set; }
        public static string C10TransferTxtLog { get; private set; }
        public static string D53TransferTxtLog { get; private set; }
        public static string D72TransferTxtLog { get; private set; }
        public static string LoginLog { get; private set; }
        public static string LoginExceptionLog { get; private set; }
        public static string IFRS9Log { get; private set; }
    }

    public class SelectOption
    {
        public string Text { get; set; }
        public string Value { get; set; }
    }

    public class GroupSelectOption
    {
        public string MainText { get; set; }
        public string CommentText { get; set; }
        public IEnumerable<string> GroupValue { get; set; }

        public string Text
        {
            get { return MainText + " (" + CommentText + ")"; }
        }
        public string Value
        {
            get { return string.Join(",", GroupValue); }
        }
    }

    public class RadioButton
    {
        public RadioButton()
        {
            Id = string.Empty;
            Name = string.Empty;
            Text = string.Empty;
            Value = string.Empty;
        }

        public string Name { get; set; }
        public bool Checked { get; set; }
        public string Text { get; set; }
        public string Value { get; set; }
        public string Id { get; set; }
    }

    public class CheckBoxListInfo
    {
        public string DisplayText { get; set; }
        public bool IsChecked { get; set; }
        public string Value { get; set; }
    }

    public class FormateTitle
    {
        public string OldTitle { get; set; }
        public string NewTitle { get; set; }
    }

    public class UserInfo
    {
        public UserInfo()
        {
            User_Account = string.Empty;
            User_Name = string.Empty;
            Login_Time = DateTime.MinValue;
        }
        public string User_Account { get; set; }
        public string User_Name { get; set; }
        public DateTime Login_Time { get; set; }
    }

    public class BrowserInfo
    {
        public string User_Account { get; set; }
        public DateTime Login_Time { get; set; }
        public DateTime Browse_Time { get; set; }
    }

    public class SummaryReportInfo
    {
        public string RuleID { get; set; }
        public string RuleDesc { get; set; }
        public string NumberOfPens { get; set; }
        public string SummaryReportData()
        {
            if (RuleDesc == null)
            {
                //RuleDesc = "無說明";
                return "規則編號" + RuleID + ":"  + NumberOfPens + "筆";
            }
            return "規則編號" + RuleID + "(" + RuleDesc + "):"  + NumberOfPens + "筆";
        }
    }

    public class ProcessStatusList
    {
        public List<SelectOption> statusOption = new List<SelectOption>() {
                                                 new SelectOption() { Text = "0：尚未呈送複核", Value = "0" },
                                                 new SelectOption() { Text = "1：待複核", Value = "1" },
                                                 new SelectOption() { Text = "2：複核完成", Value = "2" },
                                                 new SelectOption() { Text = "3：複核被退回", Value = "3" }
                                              };
    }

    public class IsActiveList
    {
        public List<SelectOption> isActiveOption = new List<SelectOption>() {
                                                    new SelectOption() { Text = "Y：生效", Value = "Y" },
                                                    new SelectOption() { Text = "N：失效", Value = "N" }
                                                  };
    }

    public class StatusList
    {
        public List<SelectOption> statusOption = new List<SelectOption>() {
                                                 new SelectOption() { Text = "待複核", Value = "" },
                                                 new SelectOption() { Text = "啟用", Value = "1" },
                                                 new SelectOption() { Text = "暫不啟用", Value = "2" },
                                                 new SelectOption() { Text = "停用", Value = "3" }
                                              };
    }
}