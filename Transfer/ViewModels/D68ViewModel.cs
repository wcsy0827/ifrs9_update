using System.ComponentModel;

namespace Transfer.ViewModels
{
    public class D68ViewModel
    {
        [Description("規則編號")]
        public string Rule_ID { get; set; }

        [Description("債券種類定義說明")]
        public string Bond_Type { get; set; }

        [Description("信用風險低評等下界")]
        public string Rating_Floor { get; set; }

        [Description("是否包含(代碼)")]
        public string Including_Ind { get; set; }

        [Description("是否包含")]
        public string Including_Ind_Name { get; set; }

        [Description("範圍 _以上/以下(代碼)")]
        public string Apply_Range { get; set; }

        [Description("範圍 _以上/以下")]
        public string Apply_Range_Name { get; set; }

        [Description("信評主標尺資料年度")]
        public string Data_Year { get; set; }

        [Description("評等主標尺_原始")]
        public string PD_Grade { get; set; }

        [Description("評等主標尺_轉換")]
        public string Grade_Adjust { get; set; }

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