using System.ComponentModel;

namespace Transfer.ViewModels
{
    public class A45ViewModel
    {
        [Description("Memo")]
        public string Memo { get; set; }

        [Description("機構簡稱")]
        public string Bloomberg_Ticker { get; set; }

        [Description("債券種類")]
        public string Bond_Type { get; set; }

        [Description("Corp/Non Corp")]
        public string Corp_NonCorp { get; set; }

        [Description("Sector_對外")]
        public string Sector_External { get; set; }

        [Description("Sector_對內")]
        public string Sector_Internal { get; set; }

        [Description("Stage評估次分類")]
        public string Assessment_Sub_Kind { get; set; }

        [Description("Sector_Research")]
        public string Sector_Research { get; set; }

        [Description("Sector_新增")]
        public string Sector_Extra { get; set; }

        [Description("資料處理日期")]
        public string Processing_Date { get; set; }
    }
}