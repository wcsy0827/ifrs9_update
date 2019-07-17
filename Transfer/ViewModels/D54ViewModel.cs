using System.ComponentModel;

namespace Transfer.ViewModels
{
    public class D54ViewModel
    {
        [Description("專案名稱")]
        public string PRJID { get; set; }

        [Description("流程名稱")]
        public string FLOWID { get; set; }

        [Description("評估基準日/報導日")]
        public string Report_Date { get; set; }

        [Description("資料版本")]
        public string Version { get; set; }

        [Description("資料處理日期")]
        public string Processing_Date { get; set; }

        [Description("產品")]
        public string Product_Code { get; set; }

        [Description("帳戶編號/群組編號")]
        public string Reference_Nbr { get; set; }

        [Description("債券編號")]
        public string Bond_Number { get; set; }

        [Description("Lots")]
        public string Lots { get; set; }

        [Description("Portfolio英文")]
        public string Portfolio_Name { get; set; }

        [Description("債券(資產)名稱")]
        public string Segment_Name { get; set; }

        [Description("減損階段")]
        public string Impairment_Stage { get; set; }

        [Description("第一年違約率")]
        public string PD { get; set; }

        [Description("第一年違約率_應收利息")]
        public string PD_Int { get; set; }

        [Description("存續期間預期信用損失(原幣)")]
        public string Lifetime_EL { get; set; }

        [Description("一年期預期信用損失(原幣)")]
        public string Y1_EL { get; set;}

        [Description("累計減損(原幣)(最終預期信用損失)")]
        public string EL { get; set; }

        [Description("月底匯率(報表日匯率)")]
        public string Ex_rate { get; set; }

        [Description("成本匯率")]
        public string Ori_Ex_rate { get; set; }

        [Description("存續期間預期信用損失(報表日匯率台幣)")]
        public string Lifetime_EL_Ex { get; set; }

        [Description("一年期預期信用損失(報表日匯率台幣)")]
        public string Y1_EL_Ex { get; set; }

        [Description("累計減損(報表日匯率台幣)")]
        public string EL_Ex { get; set; }

        [Description("存續期間預期信用損失(成本匯率台幣)")]
        public string Lifetime_EL_Ori_Ex { get; set; }

        [Description("一年期預期信用損失(成本匯率台幣)")]
        public string Y1_EL_Ori_Ex { get; set; }

        [Description("累計減損(成本匯率台幣)")]
        public string EL_Ori_Ex { get; set; }

        [Description("資料處理日期")]
        public string Exec_Date { get; set;}
    }
}