using System.ComponentModel;

namespace Transfer.ViewModels
{
    public class B06ViewModel
    {
        [Description("使用CPD群組之代碼")]
        public string CPD_Segment_Code { get; set; }

        [Description("CPD調整參數")]
        public string Delta_Q { get; set; }

        [Description("資料處理日期開始")]
        public string Processing_Date { get; set; }

        [Description("資料處理日期結束")]
        public string to { get; set; }

        [Description("產品/群组编號")]
        public string Product_Code { get; set; }

        [Description("專案名稱")]
        public string PRJID { get; set; }

        [Description("流程名稱")]
        public string FLOWID { get; set; }

        [Description("元件名稱")]
        public string CompID { get; set; }
    }
}