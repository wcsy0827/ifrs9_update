using System.ComponentModel;

namespace Transfer.ViewModels
{
    public class C23ViewModel
    {
        [Description("欄位名稱")]
        public string Column_Name { get; set; }

        [Description("風險因子")]
        public string Var_Name { get; set; }
    }
}