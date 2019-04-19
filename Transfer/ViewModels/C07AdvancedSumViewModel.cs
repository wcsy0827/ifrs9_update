using System.ComponentModel;

namespace Transfer.ViewModels
{
    public class C07AdvancedSumViewModel
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
    }
}