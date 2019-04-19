using System.ComponentModel;

namespace Transfer.ViewModels
{
    public class IssuerTickerViewModel
    {
        [Description("資料表名稱")]
        public string Table_Name { get; set; }

        [Description("編號")]
        public string Issuer_Ticker_ID { get; set; }

        [Description("Issuer")]
        public string Issuer { get; set; }

        [Description("ISSUER_EQUITY_TICKER")]
        public string ISSUER_EQUITY_TICKER { get; set; }
    }
}