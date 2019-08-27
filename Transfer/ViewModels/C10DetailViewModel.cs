using System.ComponentModel;

namespace Transfer.ViewModels
{
    public class C10DetailViewModel//開頭為title之項目不會進入資料庫，僅供Jqgrid做顯示
    {
        [Description("債券編號")]
        public string Bond_Number { get; set; }

        [Description("Lots")]
        public string Lots { get; set; }

        [Description("Portfolio")]
        public string Portfolio_Name { get; set; }

        [Description("Portfolio Name")]
        public string Portfolio { get; set; }

        [Description("版本")]
        public string Version { get; set; }

        [Description("報導日")]
        public string Report_Date { get; set; }

        [Description("金融資產餘額攤銷後之成本數(原幣)")]
        public string Amort_Amt_import { get; set; }
        
        [Description("金融資產餘額(台幣)攤銷後之成本數(台幣)")]
        public string Amort_Amt_import_TW { get; set; }

        [Description("應收利息(原幣)")]
        public string Interest_Receivable_import { get; set; }

        [Description("應收利息(台幣)")]
        public string Interest_Receivable_import_TW { get; set; }
        
        [Description("債券到期日")]
        public string Maturity_Date { get; set; }

        [Description("最近一次評等PD 本金")]
        public string PD_Amort_import { get; set; }

        [Description("最近一次評等PD 利息")]
        public string PD_Interest_Receivable_import { get; set; }

        [Description("LGD 本金")]
        public string LGD_Amort_import { get; set; }

        [Description("LGD 利息")]
        public string LGD_Interest_Receivable_import { get; set; }

        [Description("LT EL(原幣)")]
        public string Lifetime_EL_Import { get; set; }

        [Description("12Mons EL(原幣)")]
        public string Y1_EL_Import { get; set; }

        [Description("EL_本金(原幣)")]
        public string EL_import_Principle { get; set; }

        [Description("EL_利息(原幣)")]
        public string EL_import_Int { get; set; }

    }
}