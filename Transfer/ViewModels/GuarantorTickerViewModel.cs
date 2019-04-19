using System.ComponentModel;

namespace Transfer.ViewModels
{
    public class GuarantorTickerViewModel
    {
        [Description("資料表名稱")]
        public string Table_Name { get; set; }

        [Description("編號")]
        public string Guarantor_Ticker_ID { get; set; }

        [Description("Issuer")]
        public string Issuer { get; set; }

        [Description("GUARANTOR_NAME")]
        public string GUARANTOR_NAME { get; set; }

        [Description("GUARANTOR_EQY_TICKER")]
        public string GUARANTOR_EQY_TICKER { get; set; }
    }
}