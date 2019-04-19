using System.ComponentModel;

namespace Transfer.ViewModels
{
    public class C01ViewModel
    {
        [Description("評估基準日/報導日")]
        public string Report_Date { get; set; }

        [Description("資料處理日期")]
        public string Processing_Date { get; set; }

        [Description("產品")]
        public string Product_Code { get; set; }

        [Description("案件編號/帳號")]
        public string Reference_Nbr { get; set; }

        [Description("風險區隔")]
        public string Current_Rating_Code { get; set; }

        [Description("曝險額")]
        public string Exposure { get; set; }

        [Description("合約到期年限")]
        public string Actual_Year_To_Maturity { get; set; }

        [Description("估計存續期間_年")]
        public string Duration_Year { get; set; }

        [Description("估計存續期間_月")]
        public string Remaining_Month { get; set; }

        [Description("違約損失率")]
        public string Current_LGD { get; set; }

        [Description("合约利率/產品利率")]
        public string Current_Int_Rate { get; set; }

        [Description("有效利率")]
        public string EIR { get; set; }

        [Description("減損階段")]
        public string Impairment_Stage { get; set; }

        [Description("資料版本")]
        public string Version { get; set; }

        [Description("擔保順位")]
        public string Lien_position { get; set; }

        [Description("原始購買金額")]
        public string Ori_Amount { get; set; }

        [Description("金融資產餘額")]
        public string Principal { get; set; }

        [Description("應收利息")]
        public string Interest_Receivable { get; set; }

        [Description("償還(繳款)頻率 (次/年)")]
        public string Payment_Frequency { get; set; }
    }
}