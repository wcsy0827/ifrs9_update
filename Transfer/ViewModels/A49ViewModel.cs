using System.ComponentModel;

namespace Transfer.ViewModels
{
    public class A49ViewModel
    {
        [Description("評估基準日/報導日")]
        public string Report_Date { get; set; }

        [Description("資料處理日期")]
        public string Processing_Date { get; set; }

        [Description("債券編號")]
        public string Bond_Number { get; set; }

        [Description("Lots")]
        public string Lots { get; set; }

        [Description("Portfolio中文")]
        public string Portfolio { get; set; }

        [Description("會計帳值")]
        public string Accounting_EL { get; set; }

        [Description("產品")]
        public string Product_Code { get; set; }

        [Description("案件編號/帳號")]
        public string Reference_Nbr { get; set; }

        [Description("減損階段")]
        public string Impairment_Stage { get; set; }

        [Description("資料版本")]
        public string Version { get; set; }

        [Description("公報分類")]
        public string IAS39_CATEGORY { get; set; }

        [Description("評等主標尺_轉換(風險區隔)")]
        public string Grade_Adjust { get; set; }

        [Description("Portfolio英文")]
        public string Portfolio_Name { get; set; }

        [Description("低度風險門檻")]
        public string Low_Grade { get; set; }

        [Description("高度風險門檻")]
        public string High_Grade { get; set; }

        [Description("風險等級")]
        public string Risk_Level { get; set; }
    }
}