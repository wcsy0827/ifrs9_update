using System.ComponentModel;

namespace Transfer.ViewModels
{
    public class C04ProViewModel
    {
        [Description("資料表對應欄位")]
        public string Table_Property { get; set; }

        [Description("中文欄位名")]
        public string Property_Name { get; set; }

        [Description("原始因子值輸出")]
        public string Original_Factor { get; set; }

        [Description("衍生落後期數")]
        public string Derivative { get; set; }

    }
}