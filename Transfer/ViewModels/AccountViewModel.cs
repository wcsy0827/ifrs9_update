using System;
using System.ComponentModel;
using System.Runtime.Serialization;

namespace Transfer.ViewModels
{
    public class AccountViewModel
    {
        [Description("使用者帳戶")]
        public string User_Account { get; set; }

        [Description("使用者名稱")]
        public string User_Name { get; set; }

        [Description("是否為管理者")]
        public string AdminFlag { get; set; }

        [Description("登入狀態")]
        public string LoginFlag { get; set; }

        [Description("類別")]
        public string DebtFlag { get; set; }

        [Description("是否生效")]
        public string Effective { get; set; }
    }
}