using System.ComponentModel;

namespace Transfer.ViewModels
{
    public class A47ViewModel : IViewModel
    {
        [Description("國家")]
        public string Country { get; set; }

        [Description("外幣計價政府債券/總政府債務")]
        public string FC_Indexed_Debt_Rate { get; set; }

        [Description("外人持有政府/總政府債務")]
        public string External_Debt_Rate { get; set; }

        [Description("資料處理日期")]
        public string Processing_Date { get; set; }

        [Description("生效")]
        public string Effective { get; set; }
    }
}