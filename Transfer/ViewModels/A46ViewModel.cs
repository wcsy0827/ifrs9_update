using System.ComponentModel;

namespace Transfer.ViewModels
{
    public class A46ViewModel : IViewModel
    {
        [Description("國家")]
        public string Country { get; set; }

        [Description("(經常帳+FDI淨流入)/GDP")]
        public string CEIC_Value { get; set; }

        [Description("資料處理日期")]
        public string Processing_Date { get; set; }

        [Description("生效")]
        public string Effective { get; set; }
    }
}