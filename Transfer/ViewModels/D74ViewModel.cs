using System.ComponentModel;

namespace Transfer.ViewModels
{
    public class D74ViewModel : IViewModel
    {
        [Description("通知代號")]
        public string Notice_ID { get; set; }

        [Description("通知名稱")]
        public string Notice_Name { get; set; }

        [Description("通知說明")]
        public string Notice_Memo { get; set; }

        [Description("發送郵件主旨")]
        public string Mail_Title { get; set; }

        [Description("郵件內容")]
        public string Mail_Msg { get; set; }

        [Description("是否生效")]
        public string IsActive { get; set; }

        [Description("收信者數量")]
        public string mailNum { get; set; }

        [Description("新增者帳號")]
        public string Create_User { get; set; }

        [Description("新增時間")]
        public string Create_Date_Time { get; set; }

        [Description("最後修改者帳號")]
        public string LastUpdate_User { get; set; }

        [Description("最後修改時間")]
        public string LastUpdate_Date_Time { get; set; }
    }
}