using System.ComponentModel;

namespace Transfer.ViewModels
{
    public class A44ViewModel
    {
        [Description("債券編號_新")]
        public string Bond_Number_New { get; set; }

        [Description("Lots_新")]
        public string Lots_New { get; set; }

        [Description("Portfolio英文_新")]
        public string Portfolio_Name_New { get; set; }

        [Description("債券編號_舊")]
        public string Bond_Number_Old { get; set; }

        [Description("Lots_舊")]
        public string Lots_Old { get; set; }

        [Description("Portfolio英文_舊")]
        public string Portfolio_Name_Old { get; set; }

        [Description("換券前ISSUER_TICKER")]
        public string Issuer_Ticker_Old { get; set; }

        [Description("換券前GUARANTOR_NAME")]
        public string Guarantor_Name_Old { get; set; }

        [Description("換券前GUARANTOR_EQY_TICKER")]
        public string Guarantor_EQY_Ticker_Old { get; set; }

        [Description("換券日期")]
        public string Change_Date { get; set; }

        [Description("資料處理日期")]
        public string Processing_Date { get; set; }

        [Description("舊券原始購入日")]
        public string Origination_Date_Old { get; set; }
    }
}