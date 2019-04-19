using System.ComponentModel;

namespace Transfer.ViewModels
{
    public class D6CheckViewModel
    {
        [Description("作業細項")]
        public string Job_Details { get; set; }

        [Description("狀態")]
        public string Status { get; set; }

        [Description("詳細內容查看")]
        public string Details { get; set; }
    }
}