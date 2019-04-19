using System.ComponentModel;

namespace Transfer.ViewModels
{
    public class D61ViewModel
    {
        [Description("ID")]
        public string Id { get; set; }

        [Description("評估階段")]
        public string Assessment_Stage { get; set; }

        [Description("評估種類")]
        public string Assessment_Kind { get; set; }

        [Description("評估次分類")]
        public string Assessment_Sub_Kind { get; set; }

        [Description("評估項目編號")]
        public string Check_Item_Code { get; set; }

        [Description("評估項目")]
        public string Check_Item { get; set; }

        [Description("說明")]
        public string Check_Item_Memo { get; set; }

        [Description("檢查條件")]
        public string Check_Symbol { get; set; }

        [Description("門檻")]
        public string Threshold { get; set; }

        [Description("通過給定數值")]
        public string Pass_Value { get; set; }

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

        [Description("處理動作")]
        public string Change_Status { get; set; }

        [Description("生效狀態")]
        public string IsActive { get; set; }
    }
}