using System.ComponentModel;

namespace Transfer.ViewModels
{
    public class D69ViewModel
    {
        [Description("新增或修改")]
        public string ActionType { get; set; }

        [Description("規則編號")]
        public string Rule_ID { get; set; }

        [Description("基本要件是否通過(代碼)")]
        public string Basic_Pass { get; set; }

        [Description("基本要件是否通過")]
        public string Basic_Pass_Name { get; set; }

        [Description("原始購買是否符合信用風險低條件(代碼)")]
        public string Rating_Ori_Good_Ind { get; set; }

        [Description("原始購買是否符合信用風險低條件")]
        public string Rating_Ori_Good_Ind_Name { get; set; }

        [Description("評等下降數")]
        public string Rating_Notch { get; set; }

        [Description("是否包含(代碼)")]
        public string Including_Ind { get; set; }

        [Description("是否包含")]
        public string Including_Ind_Name { get; set; }

        [Description("範圍 _以上/以下(代碼)")]
        public string Apply_Range { get; set; }

        [Description("範圍 _以上/以下")]
        public string Apply_Range_Name { get; set; }

        [Description("報導日是否符合信用風險低條件(代碼)")]
        public string Rating_Curr_Good_Ind { get; set; }

        [Description("報導日是否符合信用風險低條件")]
        public string Rating_Curr_Good_Ind_Name { get; set; }

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

        [Description("原始信評是否缺漏")]
        public string Ori_Rating_Missing_Ind { get; set; }

        [Description("使用狀態(代碼)")]
        public string IsActive { get; set; }

        [Description("使用狀態")]
        public string IsActive_Name { get; set; }
    }
}