using System.ComponentModel;

namespace Transfer.ViewModels
{
    public class D60ViewModel
    {
        [Description("參數編號")]
        public string Parm_ID { get; set; }

        [Description("優先順序")]
        public string Rating_Priority { get; set; }

        [Description("評等對象")]
        public string Rating_Object { get; set; }

        [Description("國內/國外")]
        public string Rating_Org_Area { get; set; }

        [Description("1:孰高 2:孰低")]
        public string Rating_Selection { get; set; }

        [Description("規則設定者帳號")]
        public string Rule_setter { get; set; }

        [Description("規則設定者名稱")]
        public string Rule_setter_Name { get; set; }

        [Description("規則編輯日期")]
        public string Rule_setting_Date { get; set; }

        [Description("複核人員帳號")]
        public string Auditor { get; set; }

        [Description("複核人員名稱")]
        public string Auditor_Name { get; set; }

        [Description("複核日期")]
        public string Audit_Date { get; set; }

        [Description("處理狀態(代碼)")]
        public string Status { get; set; }

        [Description("處理狀態")]
        public string Status_Name { get; set; }

        [Description("Rule_desc")]
        public string Rule_desc { get; set; }

        [Description("使用狀態(代碼)")]
        public string IsActive { get; set; }

        [Description("使用狀態")]
        public string IsActive_Name { get; set; }
    }
}