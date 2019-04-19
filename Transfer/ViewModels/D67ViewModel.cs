using System.ComponentModel;

namespace Transfer.ViewModels
{
    public class D67ViewModel
    {
        [Description("評估基準日/報導日")]
        public string Report_Date { get; set; }

        [Description("資料版本")]
        public string Version { get; set; }

        [Description("債券編號")]
        public string Bond_Number { get; set; }

        [Description("發行人")]
        public string ISSUER { get; set; }

        [Description("債券產品別(揭露使用)")]
        public string PRODUCT { get; set; }

        [Description("資產群組別(提供信用風險資產減損預警彙總表使用)")]
        public string Product_Group_1 { get; set; }

        [Description("資產群組別(提供投資標的信用風險預警彙總表使用)")]
        public string Product_Group_2 { get; set; }

        [Description("金融資產餘額(台幣)攤銷後之成本數(台幣_月底匯率)")]
        public string Amort_Amt_Tw { get; set; }

        [Description("預警標準_國外(BB-)")]
        public string GRADE_Warning_F { get; set; }

        [Description("預警標準_國內為twBB-")]
        public string GRADE_Warning_D { get; set; }

        [Description("降三級評等門檻＿國外（BB+）")]
        public string Rating_diff_Over_F { get; set; }

        [Description("降三級評等門檻＿國內（twBBB-）")]
        public string Rating_diff_Over_N { get; set; }

        [Description("債項最低評等內容")]
        public string Rating_Worse { get; set; }

        [Description("國內/國外")]
        public string Bond_Area { get; set; }

        [Description("評等主標尺_原始")]
        public string PD_Grade { get; set; }

        [Description("是否評等過低預警？")]
        public string Wraming_1_Ind { get; set; }

        [Description("是否新投資？")]
        public string New_Ind { get; set; }

        [Description("連續追蹤月數")]
        public string Observation_Month { get; set; }

        [Description("是否最近6個月內最近信評與原始信評的等級差異降評三級以上(含)")]
        public string Rating_diff_Over_Ind { get; set; }

        [Description("是否降三級以上且信評過低")]
        public string Wraming_2_Ind { get; set; }

        [Description("本月變動")]
        public string Change_Memo { get; set; }

        [Description("六個月內降評三級以上(含)之債券, 最早出現降評三級以上(含)之評估日最近信評(債項信評)")]
        public string Rating_Change_Memo { get; set; }

        [Description("執行評估是否完成")]
        public string IsComplete { get; set; }
    }
}