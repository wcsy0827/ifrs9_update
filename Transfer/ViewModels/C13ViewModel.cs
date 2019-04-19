using System.ComponentModel;

namespace Transfer.ViewModels
{
    public class C13ViewModel
    {
        [Description("資料(季)")]
        public string Year_Quartly { get; set; }

        [Description("消費者物價指數")]
        public string Consumer_Price_Index { get; set; }

        [Description("是否為消費者物價指數預測值")]
        public string Consumer_Price_Index_Pre_Ind { get; set; }

        [Description("失業率")]
        public string Unemployment_Rate { get; set; }

        [Description("是否為失業率預測值")]
        public string Unemployment_Rate_Pre_Ind { get; set; }
    }
}