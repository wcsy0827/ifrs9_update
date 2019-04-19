using System.ComponentModel;

namespace Transfer.ViewModels
{
    public class A52ViewModel
    {
        [Description("ID")]
        public string Id { get; set; }
        [Description("評等機構")]
        public string Rating_Org { get; set; }
        [Description("評等主標尺_原始")]
        public string PD_Grade { get; set; }
        [Description("評等內容")]
        public string Rating { get; set; }
        [Description("是否有效")]
        public string IsActive { get; set; }
        [Description("資料異動狀態")]
        public string Change_Status { get; set; }
        [Description("編輯者")]
        public string Editor { get; set; }
        [Description("資料處理時間")]
        public string Processing_Date { get; set; }
        [Description("複核者")]
        public string Auditor { get; set; }
        [Description("複核結果")]
        public string Status { get; set; }
        [Description("複核者意見")]
        public string Auditor_Reply { get; set; }
        [Description("複核時間")]
        public string Audit_date { get; set; }
    }
}