using System.ComponentModel;

namespace Transfer.ViewModels
{
    public class C10ViewModel //開頭為title之項目不會進入資料庫，僅供Jqgrid做顯示
    {
        [Description("債券編號")]
        public string Bond_Number { get; set; }

        [Description("Lots")]
        public string Lots { get; set; }

        [Description("債券(資產)名稱")]
        public string Segment_Name { get; set; }
        [Description("報導日")]
        public string Report_Date { get; set; }

        [Description("債券產品別(揭露使用)")]
        public string title1 { get; set; }

        [Description("公報分類")]
        public string title3{ get; set; }
        
        [Description("原公報分類")]
        public string title4 { get; set; }

        [Description("金融資產餘額攤銷後之成本數(原幣)")]
        public string Amort_Amt_import { get; set; }
        
        [Description("金融資產餘額(台幣)攤銷後之成本數(台幣)")]
        public string Amort_Amt_import_TW { get; set; }

        [Description("應收利息(原幣)")]
        public string Interest_Receivable_import { get; set; }

        [Description("應收利息(台幣)")]
        public string Interest_Receivable_import_TW { get; set; }

        [Description("市價(原幣)")]
        public string title5 { get; set; }

        [Description("債券幣別")]
        public string title6 { get; set; }

        [Description("Portfolio")]
        public string Portfolio_Name { get; set; }

        [Description("Portfolio Name")]
        public string Portfolio { get; set; }

        [Description("部門(國內/國外)")]
        public string title9 { get; set; }

        [Description("原始金額")]
        public string title10 { get; set; }

        [Description("月底匯率")]
        public string title11 { get; set; }

        [Description("原始利率(票面利率)")]
        public string title12 { get; set; }

        [Description("原始有效利率(買入殖利率)")]
        public string title13 { get; set; }

        [Description("贖回日期(本金一次贖回)")]
        public string title14 { get; set; }

        [Description("現金流類型")]
        public string title15 { get; set; }

        [Description("債券購入(認列)日期")]
        public string Origination_Date { get; set; }

        [Description("債券到期日")]
        public string Maturity_Date { get; set; }

        [Description("資料日期(MDR)")]
        public string title18 { get; set; }

        [Description("資產區隔")]
        public string title19{ get; set; }

        [Description("成本匯率")]
        public string title20 { get; set; }

        [Description("自操委外(IH_OS)")]
        public string title21 { get; set; }

        [Description("債券擔保順位")]
        public string title22 { get; set; }

        [Description("債券編號&Lots&Portfolio")]
        public string title23 { get; set; }

        [Description("購買時評等")]
        public string title24{ get; set; }

        [Description("購買時評等(主標尺)")]
        public string title25 { get; set; }

        [Description("最近一次評等")]
        public string title26 { get; set; }

        [Description("最近一次評等(主標尺)")]
        public string title27 { get; set; }

        [Description("最近一次評等PD")]
        public string PD_import { get; set; }

        [Description("LGD")]
        public string LGD_import { get; set; }

        [Description("EAD(原幣)")]
        public string title28 { get; set; }

        [Description("EAD(台幣)")]
        public string title39 { get; set; }

        [Description("Tenor(Year)")]
        public string title29 { get; set; }

        [Description("Period (Months)")]
        public string title30 { get; set; }

        [Description("Stage")]
        public string title31 { get; set; }

        [Description("12Mons EL(台幣)")]
        public string title32 { get; set; }

        [Description("LT EL(台幣)")]
        public string title33 { get; set; }

        [Description("EL(台幣)")]
        public string title34 { get; set; }

        [Description("12Mons EL(原幣)")]
        public string Y1_EL_Import { get; set; }

        [Description("LT EL(原幣)")] 
        public string Lifetime_EL_Import { get; set; }

        [Description("EL(原幣)")]
        public string title36 { get; set; }

        [Description("EL_本金(原幣)")]
        public string EL_import_Principle { get; set; }

        [Description("EL_利息(原幣)")]
        public string EL_import_Int { get; set; }

        [Description("EL_本金(台幣)")]
        public string title40 { get; set; }

        [Description("EL_利息(台幣)")]
        public string title41 { get; set; }
    }
}