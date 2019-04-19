using System.ComponentModel;

namespace Transfer.ViewModels
{
    public class TaskScheduleViewModel
    {
        [Description("觸發排程")]
        public string Trigger_Task { get; set; }
        [Description("報導日/基準日")]
        public string ReportDate { get; set; }
        [Description("資料版本")]
        public string Version { get; set; }
        [Description("資料表名稱")]
        public string File_Name { get; set; }
        [Description("成功/失敗")]
        public string TransferType { get; set; }
        [Description("開始日期")]
        public string Create_Date { get; set; }
        [Description("開始時間")]
        public string Create_Time { get; set; }
        [Description("結束日期")]
        public string End_Date { get; set; }
        [Description("結束時間")]
        public string End_Time { get; set; }
    }
}