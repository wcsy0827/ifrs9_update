using System.ComponentModel;

namespace Transfer.ViewModels
{
    public class A48ViewModel : IViewModel
    {
        [Description("年度")]
        public string Data_Year { get; set; }

        [Description("資料處理日期")]
        public string Processing_Date { get; set; }

        [Description("國家")]
        public string Country { get; set; }

        [Description("政府債務/GDP")]
        public string IGS_Index { get; set; }

        [Description("外人持有政府/總政府債務")]
        public string External_Debt_Rate { get; set; }

        [Description("外幣計價政府債券/總政府債務")]
        public string FC_Indexed_Debt_Rate {get;set;}

        [Description("(經常帳+FDI淨流入)/GDP")]
        public string CEIC_Value { get; set; }

        [Description("短期外債")]
        public string Short_term_Debt { get; set; }

        [Description("外匯儲備")]
        public string Foreign_Exchange { get; set; }

        [Description("年度GDP Y/Y")]
        public string GDP_Yearly { get; set; }

        [Description("生效")]
        public string Effective { get; set; }
    }
}