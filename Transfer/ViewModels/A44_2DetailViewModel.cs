using System.ComponentModel;

namespace Transfer.ViewModels
{
    public class A44_2DetailViewModel
    {
        [Description("債券代號(ISIN)_舊")]//20200926 alibaba 標示舊券
        public string Bond_Number_Old { get; set; }

        [Description("Portfolio Name_舊")]//20200926 alibaba 標示舊券
        public string Portfolio_Name_Old { get; set; }

        [Description("原幣利息收入")]
        public string Int_Receivable { get; set; }

        [Description("台幣利息收入")]
        public string Int_Receivable_Tw { get; set; }

        [Description("評估基準日 / 報導日")]
        public string Report_Date { get; set; }

        [Description("資料處理日期")]
        public string Processing_Date { get; set; }

        [Description("債券代號(ISIN)_新")]//20200926 alibaba 標示新券 202008210166-00
        public string Bond_Number_New { get; set; }

        //200926 alibaba 新增顯示欄位 202008210166-00
        [Description("Portfolio_Name_新")]
        public string Portfolio_Name_New { get; set; }

        [Description("Lots_新")]
        public string Lots_New { get; set; }

        [Description("面額比例_New")]
        public string  Ori_Amount_Percentage_New { get; set; }

        [Description("面額_New")]
        public string Ori_Amount_New { get; set; }

        [Description("換入原幣應計息")]
        public string IntRevise_perBond_New { get; set; }

        [Description("換入台幣應計息")]
        public string IntRevise_perBond_Tw_New { get; set; }

        [Description("維護人員")]
        public string Create_User { get; set; }
        //end 200926 alibaba 
        [Description("Lots_舊")]//20200926 alibaba 標示舊券 202008210166-00
        public string Lots_Old { get; set; }

    }
}