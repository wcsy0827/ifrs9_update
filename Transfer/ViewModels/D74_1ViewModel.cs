using System.ComponentModel;

namespace Transfer.ViewModels
{
    public class D74_1ViewModel : IViewModel
    {
        [Description("流水號")] 
        public string ID { get; set; }

        [Description("通知代號")]
        public string Notice_ID { get; set; }

        [Description("收件者單位")]
        public string Recipient_Department { get; set; }

        [Description("收信者名稱")]
        public string Recipient_Name { get; set; }

        [Description("收信者Mail")]
        public string Recipient_mail { get; set; }

        [Description("是否生效")]
        public string IsActive { get; set; }

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