using System.ComponentModel;

namespace Transfer.ViewModels
{
    public class A08ViewModel
    {
        [Description("放款編號")]
        public string Reference_Nbr { get; set; }

        [Description("資料處理日期")]
        public string Processing_Date { get; set; }

        [Description("評估基準日/報導日")]
        public string Report_Date { get; set; }

        [Description("風險三級分類")]
        public string Loan_Risk_Type { get; set; }

        [Description("實收利息")]
        public string Interest { get; set; }
    }
}