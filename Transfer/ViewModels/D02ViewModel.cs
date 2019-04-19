using System.ComponentModel;

namespace Transfer.ViewModels
{
    public class D02ViewModel
    {
        [Description("專案名稱")]
        public string PRJID { get; set; }

        [Description("流程名稱")]
        public string FLOWID { get; set; }

        [Description("評估基準日/報導日")]
        public string Report_Date { get; set; }

        [Description("資料處理日期")]
        public string Processing_Date { get; set; }

        [Description("產品")]
        public string Product_Code { get; set; }

        [Description("案件編號/帳號")]
        public string Reference_Nbr { get; set; }

        [Description("存續期間預期信用損失")]
        public string Lifetime_EL { get; set; }

        [Description("一年期預期信用損失")]
        public string Y1_EL { get; set; }

        [Description("最終預期信用損失")]
        public string EL { get; set; }

        [Description("減損階段")]
        public string Impairment_Stage { get; set; }

        [Description("風險三級分類")]
        public string Loan_Risk_Type { get; set; }

        [Description("資料版本")]
        public string Version { get; set; }

        [Description("剩餘本金餘額")]
        public string Principal { get; set; }

        [Description("是否屬於催收款")]
        public string Collection_Ind { get; set; }

        [Description("金融業務分類")]
        public string Loan_Type { get; set; }

        [Description("曝險額")]
        public string Exposure { get; set; }

        [Description("回收率的區隔")]
        public string NO34RCV { get; set; }

        [Description("有效利率")]
        public string EIR { get; set; }

        [Description("實收利息")]
        public string Interest { get; set; }

        [Description("違約損失率")]
        public string Current_LGD { get; set; }

        [Description("第一年違約率")]
        public string PD { get; set; }

        [Description("應收利息")]
        public string Interest_Receivable { get; set; }

        [Description("累計減損_本金")]
        public string Principal_EL { get; set; }

        [Description("累計減損_利息")]
        public string Interest_Receivable_EL { get; set; }
    }
}