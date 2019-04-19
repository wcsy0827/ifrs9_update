using System.ComponentModel;

namespace Transfer.ViewModels
{
    /// <summary>
    /// A95 View Data
    /// </summary>
    public class A95ViewModel
    {
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

        [Description("債券產品別(揭露使用)")]
        public string PRODUCT { get; set; }

        [Description("債券擔保順位")]
        public string Lien_position { get; set; }

        [Description("Security_Ticker")]
        public string Security_Ticker { get; set; }

        [Description("Security_Des")]
        public string Security_Des { get; set; }

        [Description("Bloomberg_Ticker")]
        public string Bloomberg_Ticker { get; set; }

        [Description("債券種類")]
        public string Bond_Type { get; set; }

        [Description("產業別資訊")]
        public string Assessment_Sub_Kind { get; set; }
    }
}