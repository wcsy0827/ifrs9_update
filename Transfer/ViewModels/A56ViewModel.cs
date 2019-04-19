using System.ComponentModel;

namespace Transfer.ViewModels
{
    public class A56ViewModel : IViewModel
    {
        [Description("編號")]
        public string ID { get; set; }

        [Description("是否生效")]
        public string IsActive { get; set; }

        [Description("特殊字元")]
        public string Replace_Object { get; set; }

        [Description("補正方式")]
        public string Update_Method { get; set; }

        [Description("補正方式說明")]
        public string Method_MEMO { get; set; }

        [Description("新取代字元")]
        public string Replace_Word { get; set; }

        [Description("編輯者")]
        public string LastUpdate_User { get; set; }

        [Description("資料處理日期")]
        public string LastUpdate_DateTime { get; set; }
    }
}