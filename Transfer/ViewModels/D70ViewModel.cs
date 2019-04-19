using System.ComponentModel;

namespace Transfer.ViewModels
{
    public class D70ViewModel
    {
        [Description("規則編號")]
        public string Rule_ID { get; set; }

        [Description("基本要件評估時點(代碼)")]
        public string Basic_Pass_Check_Point { get; set; }

        [Description("基本要件評估時點")]
        public string Basic_Pass_Check_Point_Name { get; set; }

        [Description("基本要件是否通過(代碼)")]
        public string Basic_Pass { get; set; }

        [Description("基本要件是否通過")]
        public string Basic_Pass_Name { get; set; }

        [Description("信評評估時點(代碼)")]
        public string Rating_Check_Point { get; set; }

        [Description("信評評估時點")]
        public string Rating_Check_Point_Name { get; set; }

        [Description("報導日信評門檻")]
        public string Rating_Threshold { get; set; }

        [Description("信評主標尺資料年度")]
        public string Data_Year { get; set; }

        [Description("報導日信評門檻-對應評等主標尺_原始 ")]
        public string Rating_Threshold_Map_Grade { get; set; }

        [Description("報導日信評門檻-對應評等主標尺_調整後")]
        public string Rating_Threshold_Map_Grade_Adjust { get; set; }

        [Description("是否包含(報導日信評門檻)(代碼)")]
        public string Including_Ind_0 { get; set; }

        [Description("是否包含(報導日信評門檻)")]
        public string Including_Ind_0_Name { get; set; }

        [Description("範圍 _以上/以下(報導日信評門檻)(代碼)")]
        public string Apply_Range_0 { get; set; }

        [Description("範圍 _以上/以下(報導日信評門檻)")]
        public string Apply_Range_0_Name { get; set; }

        [Description("報導日信評區間_from")]
        public string Rating_from { get; set; }

        [Description("報導日信評區間_from-對應評等主標尺_原始")]
        public string Rating_from_Map_Grade { get; set; }

        [Description("報導日信評區間_from-對應評等主標尺_調整後")]
        public string Rating_from_Map_Grade_Adjust { get; set; }

        [Description("報導日信評區間_To")]
        public string Rating_To { get; set; }

        [Description("報導日信評區間_To-對應評等主標尺_原始")]
        public string Rating_To_Map_Grade { get; set; }

        [Description("報導日信評區間_To-對應評等主標尺_調整後")]
        public string Rating_To_Map_Grade_Adjust { get; set; }

        [Description("未實現累計損失月數")]
        public string Value_Change_Months { get; set; }

        [Description("是否包含(未實現累計損失月數)(代碼)")]
        public string Including_Ind_1 { get; set; }

        [Description("是否包含(未實現累計損失月數)")]
        public string Including_Ind_1_Name { get; set; }

        [Description("範圍 _以上/以下(未實現累計損失月數)(代碼)")]
        public string Apply_Range_1 { get; set; }

        [Description("範圍 _以上/以下(未實現累計損失月數)")]
        public string Apply_Range_1_Name { get; set; }

        [Description("信用利差拓寬累積月數")]
        public string Spread_Change_Months { get; set; }

        [Description("是否包含(信用利差拓寬累積月數)(代碼)")]
        public string Including_Ind_2 { get; set; }

        [Description("是否包含(信用利差拓寬累積月數)")]
        public string Including_Ind_2_Name { get; set; }

        [Description("範圍 _以上/以下(信用利差拓寬累積月數)(代碼)")]
        public string Apply_Range_2 { get; set; }

        [Description("範圍 _以上/以下(信用利差拓寬累積月數)")]
        public string Apply_Range_2_Name { get; set; }

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