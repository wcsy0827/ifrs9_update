using System.ComponentModel;

namespace Transfer.ViewModels
{
    public class C07AdvancedViewModel
    {
        [Description("產品群組代碼")]
        public string Group_Product_Code { get; set; }

        [Description("產品群組名稱")]
        public string Group_Product_Name { get; set; }

        [Description("產品")]
        public string Product_Code { get; set; }

        [Description("評估次分類")]
        public string Assessment_Sub_Kind { get; set; }

        [Description("報導日")]
        public string Report_Date { get; set; }

        [Description("版本")]
        public string Version { get; set; }

        [Description("帳戶編號")]
        public string Reference_Nbr { get; set; }

        [Description("債券編號")]
        public string Bond_Number { get; set; }

        [Description("Lots")]
        public string Lots { get; set; }

        [Description("Portfolio")]
        public string Portfolio { get; set; }

        [Description("第一年違約機率")]
        public string PD { get; set; }

        [Description("LGD")]
        public string LGD { get; set;}

        [Description("曝險額(原幣)")]
        public string Exposure_EL { get; set; }

        [Description("曝險額(台幣)")]
        public string Exposure_Ex { get; set; }

        [Description("一年期預期信用損失(原幣)")]
        public string Y1_EL { get; set; }

        [Description("存續期間預期信用損失(原幣)")]
        public string Lifetime_EL { get; set; }

        [Description("一年期預期信用損失(台幣)")]
        public string Y1_EL_Ex { get; set; }

        [Description("存續期間預期信用損失(台幣)")]
        public string Lifetime_EL_Ex { get; set; }

        [Description("預設STAGE")]
        public string Impairment_Stage { get; set; }

        [Description("基準日匯率")]
        public string Ex_rate { get; set; }

        [Description("基本要件通過與否")]
        public string Basic_Pass { get; set; }

        [Description("未實現累計損失月數_本月狀況")]
        public string Accumulation_Loss_This_Month { get; set; }

        [Description("信用利差目前累計逾越月數")]
        public string Chg_In_Spread_This_Month { get; set; }

        [Description("是否為觀察名單")]
        public string Watch_IND { get; set; }
    }
}