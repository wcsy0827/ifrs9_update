using System.Collections.Generic;
using System.ComponentModel;

namespace Transfer.ViewModels
{
    /// <summary>
    /// CheckTable View Data
    /// </summary>
    public class ChgInSpreadViewModel
    {
        public ChgInSpreadViewModel()
        {
            ChgInSpread = new List<ChgInSpreadDetail>();
        }

        [Description("專案名稱")]
        public string PRJID { get; set; }
        [Description("流程名稱")]
        public string FLOWID { get; set; }
        [Description("減損試算結果")]
        public List<ChgInSpreadDetail> ChgInSpread { get; set; }
    }

    public class ChgInSpreadDetail
    {
        [Description("債券幣別")]
        public string Currency_Code { get; set; }
        [Description("計算基準")]
        public string Base { get; set; }
        [Description("一年期預期損失")]
        public string Stage_1 { get; set; }
        [Description("存續期間預期損失")]
        public string Stage_2 { get; set; }
        [Description("一年期預期損失(違約率=100%)")]
        public string Stage_3 { get; set; }
    }
}