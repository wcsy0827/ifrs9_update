using System;
using System.ComponentModel;
using System.Runtime.Serialization;

namespace Transfer.ViewModels
{
    public class SetAssessmentViewModel
    {
        [Description("群組編號")]
        public string Group_Product_Code { get; set; }

        [Description("群組名稱")]
        public string Group_Product_Name { get; set; }

        [Description("資料表編號")]
        public string Table_Id { get; set; }

        [Description("覆核/呈送(人員帳號)")]
        public string User_Account { get; set; }

        [Description("覆核/呈送(人員名稱)")]
        public string User_Name { get; set; }
    }
}