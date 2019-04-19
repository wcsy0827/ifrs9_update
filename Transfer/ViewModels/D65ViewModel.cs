using System.ComponentModel;

namespace Transfer.ViewModels
{
    public class D65ViewModel
    {
        [Description("狀態")]
        public string Status { get; set; }

        [Description("債券編號")]
        public string Bond_Number { get; set; }

        [Description("Lots")]
        public string Lots { get; set; }

        [Description("Portfolio英文")]
        public string Portfolio_Name { get; set; }

        [Description("Portfolio")]
        public string Portfolio { get; set; }

        [Description("人工增加案件狀態")]
        public string Extra_Case { get; set; }

        [Description("評估次分類")]
        public string Assessment_Sub_Kind { get; set; }

        [Description("已評估")]
        public string Pass_Confirm_Flag { get; set; }

        [Description("複核者已選擇最終版本")]
        public string Result_Version_Confirm_Flag { get; set; }

        [Description("複核後選擇狀態_通過與否")]
        public string Quantitative_Pass_Confirm { get; set; }

        [Description("檔案數量")]
        public string FilesCount { get; set; }

        [Description("帳戶編號/群組編號")]
        public string Reference_Nbr { get; set; }

        [Description("評估結果版本")]
        public string Assessment_Result_Version { get; set; }

        [Description("資料版本")]
        public string Version { get; set; }

        [Description("評估日/報導日")]
        public string Report_Date { get; set; }

        [Description("債券購入(認列)日期")]
        public string Origination_Date { get; set; }

        [Description("評估者")]
        public string Assessor { get; set; }

        [Description("評估者名稱")]
        public string Assessor_Name { get; set; }

        [Description("評估日期")]
        public string Assessment_date { get; set; }

        [Description("複核者")]
        public string Auditor { get; set; }

        [Description("複核者名稱")]
        public string Auditor_Name { get; set; }

        [Description("複核日期")]
        public string Audit_date { get; set; }

        [Description("複核後選擇版本")]
        public string Result_Version_Confirm { get; set; }

        [Description("複核者不接受選擇退回")]
        public string Auditor_Return { get; set; }

        [Description("是否提交複核")]
        public string Send_to_Auditor { get; set; }

        [Description("評估者提交時間")]
        public string Send_Time { get; set; }

        [Description("評估種類")]
        public string Assessment_Kind { get; set; }

    }
}