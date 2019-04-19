using System.ComponentModel;

namespace Transfer.ViewModels
{
    public class D72ViewModel : IViewModel
    {
        [Description("債券產品別(揭露使用)_SMF")]
        public string PRODUCT { get; set; }

        [Description("報表使用之產品分類")]
        public string Product_Group { get; set; }

        [Description("使用SMF分類報表別")]
        public string Product_Group_kind { get; set; }

        [Description("備註")]
        public string Memo { get; set; }
    }
}