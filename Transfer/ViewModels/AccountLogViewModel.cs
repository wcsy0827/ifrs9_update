using System;
using System.ComponentModel;
using System.Runtime.Serialization;

namespace Transfer.ViewModels
{
    /// <summary>
    /// AccountLogViewModel
    /// </summary>
    public class AccountLogViewModel
    {
        [Description("使用者帳號")]
        public string User_Account { get; set; }

        [Description("使用者名稱")]
        public string User_Name { get; set; }

        [Description("登入時間")]
        public string Login_Time { get; set; }

        [Description("登出時間")]
        public string Logout_Time { get; set; }

        [Description("進入Menu時間")]
        public string Browse_Time { get; set; }

        [Description("Menu參數")]
        public string Menu_Id { get; set; }

        [Description("Menu名稱")]
        public string Menu_Detail { get; set; }

        [Description("Action名稱")]
        public string Action_Name { get; set; }

        [Description("執行動作")]
        public string Event_Name { get; set; }

        [Description("執行開始時間")]
        public string Event_Begin { get; set; }

        [Description("執行結束時間")]
        public string Event_Complete { get; set; }

        [Description("執行結果")]
        public string Event_Flag { get; set; }

        [Description("IP位置")]
        public string IP_Location { get; set; }
    }
}