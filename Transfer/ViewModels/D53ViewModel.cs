using System.ComponentModel;

namespace Transfer.ViewModels
{
    public class D53ViewModel : IViewModel
    {
        [Description("SMF代碼")]
        public string SMF_Code { get; set; }

        [Description("SMF英文說明")]
        public string SMF { get; set; }

        [Description("SMF中文說明")]
        public string SMF_Desc { get; set; }

        [Description("金融商品分類_代碼")]
        public string Asset_type_Code { get; set; }

        [Description("金融商品分類")]
        public string Asset_Type { get; set; }
    }
}