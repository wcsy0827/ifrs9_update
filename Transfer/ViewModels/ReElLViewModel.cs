using System.ComponentModel;

namespace Transfer.ViewModels
{
    public class ReELViewModel
    {
        [Description("基準日")]
        public string Report_Date { get; set; }

        [Description("流程")]
        public string FLOWID { get; set; }

        [Description("執行者")]
        public string Create_User { get; set; }

        [Description("執行時間")]
        public string Create_Time { get; set; }

        [Description("執行原因")]
        public string Message { get; set; }

    }
}