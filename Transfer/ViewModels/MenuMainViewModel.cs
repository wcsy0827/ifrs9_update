using System.ComponentModel;

namespace Transfer.ViewModels
{
    public class MenuMainViewModel
    {
        [Description("選單類別")]
        public string Menu { get; set; }

        [Description("選單類別ID")]
        public string Menu_Id { get; set; }

        [Description("選單類別名稱")]
        public string Menu_Detail { get; set; }

        [Description("選單類別圖示")]
        public string Class { get; set; }

        [Description("新增者帳號")]
        public string Create_User { get; set; }

        [Description("新增日期")]
        public string Create_Date { get; set; }

        [Description("新增時間")]
        public string Create_Time { get; set; }

        [Description("最後修改者帳號")]
        public string LastUpdate_User { get; set; }

        [Description("最後修改日期")]
        public string LastUpdate_Date { get; set; }

        [Description("最後修改時間")]
        public string LastUpdate_Time { get; set; }
    }
}