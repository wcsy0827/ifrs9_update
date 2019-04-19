using System.ComponentModel;

namespace Transfer.ViewModels
{
    public class A95_1ViewModel : IViewModel
    {
        [Description("債券編號")]
        public string Bond_Number { get; set; }

        [Description("Bloomberg Security_Des")]
        public string Security_Des { get; set; }

        [Description("Bloomberg Ticker簡碼")]
        public string Bloomberg_Ticker { get; set; }

        [Description("資料處理日期")]
        public string Processing_Date { get; set; }

        [Description("資料處理者")]
        public string Processing_User { get; set; }
    }
}