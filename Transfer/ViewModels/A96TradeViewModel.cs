using System.ComponentModel;

namespace Transfer.ViewModels
{
    /// <summary>
    /// A96 TradeDate View Data
    /// </summary>
    public class A96TradeViewModel : IViewModel
    {
        [Description("評估基準日/報導日")]
        public string Report_Date { get; set; }

        [Description("最後交易日")]
        public string Last_Date { get; set; }

        [Description("新增者帳號")]
        public string Create_User { get; set; }

        [Description("新增日期")]
        public string Create_Date { get; set; }

        [Description("新增時間")]
        public string Create_Time { get; set; }

        [Description("最後修改者帳號")]
        public string LastUpdate_User { get; set; }

        [Description("最後修改日期")]
        public string LastUpdate_Date { get; set; }

        [Description("最後修改時間")]
        public string LastUpdate_Time { get; set; }
    }
}