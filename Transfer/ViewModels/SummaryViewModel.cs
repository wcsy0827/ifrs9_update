using System.ComponentModel;

namespace Transfer.ViewModels
{
    public class SummaryViewModel
    {
        [Description("報導日")]
        public string Report_Date { get; set; }

        [Description("版本")]
        public string Version { get; set; }

        [Description("產品群代碼")]
        public string Group_Product_Code { get; set; }

        [Description("產品群名稱")]
        public string Group_Product_Name { get; set; }

        [Description("評估次分類")]
        public string Assessment_Sub_Kind { get; set; }

        [Description("曝險額(台幣)")]
        public string Exposure_EX { get; set; }

        [Description("一年期預期損失(台幣)")]
        public string Y1_EL_EX { get; set; }

        [Description("存續期間預期損失(台幣)")]
        public string Lifetime_EL_EX { get; set; }


        public string RuleID_BTY { get; set; }
        public string RuleDesc_BTY { get; set; }
        public string NumberOfPens_BTY { get; set; }
        //public string SummaryReportData()
        //{
        //    if (RuleDesc == null) { RuleDesc = "無說明"; }
        //    return "規則編號:" + RuleID + "(" + RuleDesc + "):" + NumberOfPens + "筆";
        //}

        public string RuleID_BTN { get; set; }
        public string RuleDesc_BTN { get; set; }
        public string NumberOfPens_BTN { get; set; }


        public string RuleID_WAT { get; set; }
        public string RuleDesc_WAT { get; set; }
        public string NumberOfPens_WAT { get; set; }

        public string RuleID_WAR { get; set; }
        public string RuleDesc_WAR { get; set; }
        public string NumberOfPens_WAR { get; set; }


    }
}