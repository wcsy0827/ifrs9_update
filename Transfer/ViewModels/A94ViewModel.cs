using System.ComponentModel;

namespace Transfer.ViewModels
{
    /// <summary>
    /// A94 View Data
    /// </summary>
    public class A94ViewModel
    {
        [Description("國家")]
        public string Country { get; set; }

        [Description("政府債務/GDP Ticker")]
        public string IGS_Index_Map { get; set; }

        [Description("短期外債 Ticker")]
        public string Short_term_Debt_Map { get; set; }

        [Description("外匯儲備 Ticker")]
        public string Foreign_Exchange_Map { get; set; }

        [Description("年度GDP Y/Y Ticker")]
        public string GDP_Yearly_Map { get; set; }
    }
}