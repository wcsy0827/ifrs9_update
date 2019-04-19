using System.ComponentModel;

namespace Transfer.ViewModels
{
    public class D64ViewModel
    {
        [Description("帳戶編號/群組編號")]
        public string Reference_Nbr { get; set; }

        [Description("評估日/報導日")]
        public string Report_Date { get; set; }

        [Description("資料版本")]
        public string Version { get; set; }

        [Description("債券編號")]
        public string Bond_Number { get; set; }

        [Description("Lots")]
        public string Lots { get; set; }

        [Description("Portfolio")]
        public string Portfolio { get; set; }

        [Description("債券購入(認列)日期")]
        public string Origination_Date { get; set; }

        [Description("Portfolio英文")]
        public string Portfolio_Name { get; set; }

        [Description("評估種類")]
        public string Assessment_Kind { get; set; }

        [Description("評估次分類")]
        public string Assessment_Sub_Kind { get; set; }

        [Description("評估階段")]
        public string Assessment_Stage { get; set; }

        [Description("測試指標編號")]
        public string Check_Item_Code { get; set; }

        [Description("測試指標名稱")]
        public string Check_Item { get; set; }

        [Description("測試指標說明")]
        public string Check_Item_Memo { get; set; }

        [Description("指標證明文件存放位址")]
        public string Check_Reference { get; set; }

        [Description("評估結果版本")]
        public string Assessment_Result_Version { get; set; }

        [Description("評估結果說明")]
        public string Assessment_Result { get; set; }

        [Description("量化指標值")]
        public string Check_Item_Value { get; set; }

        [Description("檢查條件")]
        public string Check_Symbol { get; set; }

        [Description("門檻")]
        public string Threshold { get; set; }

        [Description("通過給定值")]
        public string Pass_Value { get; set; }

        [Description("是否通過量化評估")]
        public string Quantitative_Pass { get; set; }

        [Description("備註")]
        public string Memo { get; set; }

        [Description("評估者")]
        public string Assessor { get; set; }

        [Description("檔案數量")]
        public string FileCount { get; set; }
    }
}