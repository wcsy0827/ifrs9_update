using System.ComponentModel;

namespace Transfer.ViewModels
{
    public class D73ViewModel : IViewModel
    {
        [Description("編號")]
        public string ID { get; set; }

        [Description("報導日")]
        public string Report_Date { get; set; }

        [Description("版本")]
        public string Version { get; set; }

        [Description("檔案名稱 ")]
        public string File_Name { get; set; }

        [Description("文件存放位址")]
        public string File_path { get; set; }

        [Description("建立資料日期")]
        public string Create_Date { get; set; }

        [Description("刪除註記")]
        public string Delete_YN { get; set; }

        [Description("刪除資料日期")]
        public string Delete_Date { get; set; }

        [Description("資料刪除者")]
        public string Processing_User { get; set; }
    }
}