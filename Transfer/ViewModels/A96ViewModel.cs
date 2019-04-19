using System.ComponentModel;

namespace Transfer.ViewModels
{
    /// <summary>
    /// A96 View Data
    /// </summary>
    public class A96ViewModel : IViewModel
    {
        [Description("帳戶編號/群組編號")]
        public string Reference_Nbr { get; set; }

        [Description("評估基準日/報導日")]
        public string Report_Date { get; set; }

        [Description("資料版本")]
        public string Version { get; set; }

        [Description("債券編號")]
        public string Bond_Number { get; set; }

        [Description("Lots")]
        public string Lots { get; set; }

        [Description("Portfolio英文")]
        public string Portfolio_Name { get; set; }

        [Description("債券購入(認列)日期")]
        public string Origination_Date { get; set; }

        [Description("原始有效利率(買入殖利率)")]
        public string EIR { get; set; }

        [Description("資料處理日期")]
        public string Processing_Date { get; set; }

        [Description("Mid_Yield")]
        public string Mid_Yield { get; set; }

        [Description("BNCHMRK_TSY_ISSUE_ID")]
        public string BNCHMRK_TSY_ISSUE_ID { get; set; }

        [Description("ID_CUSIP")]
        public string ID_CUSIP { get; set; }

        [Description("Treasury_Current")]
        public string Treasury_Current { get; set; }

        [Description("Treasury_When_Trade")]
        public string Treasury_When_Trade { get; set; }

        [Description("Spread_Current")]
        public string Spread_Current { get; set; }

        [Description("Spread_When_Trade")]
        public string Spread_When_Trade { get; set; }

        [Description("異動值總和")]
        public string All_in_Chg { get; set; }

        [Description("信用利差拓寬基點")]
        public string Chg_In_Spread { get; set; }

        [Description("Treasury異動值")]
        public string Chg_In_Treasury { get; set; }

        [Description("手動調整原因")]
        public string Memo { get; set; }

        [Description("手動調整帳號")]
        public string LastUpdate_User { get; set; }

        [Description("手動調整日期")]
        public string LastUpdate_Date { get; set; }

        [Description("手動調整時間")]
        public string LastUpdate_Time { get; set; }
    }
}