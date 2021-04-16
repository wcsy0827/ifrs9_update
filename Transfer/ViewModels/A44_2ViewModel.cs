using System.ComponentModel;

namespace Transfer.ViewModels
{
    public class A44_2ViewModel
    {
        [Description("債券代號(ISIN)")]
        public string Bond_Number_Old { get; set; }

        [Description("Portfolio Name")]
        public string Portfolio_Name_Old { get; set; }

        [Description("原幣利息收入")]
        public string Int_Receivable { get; set; }

        [Description("應計息合計(台幣)")]//20200818 楊瞻遠 抓已經計算匯兌損益後的應計息 202008210166-00
        public string Int_Receivable_Tw { get; set; }

        [Description("Lots")]//20200921 楊瞻遠 新增Lots欄位 202008210166-00
        public string Lots_Old { get; set; }

        [Description("評估基準日 / 報導日")]
        public string Report_Date { get; set; }

        [Description("資料處理日期")]
        public string Processing_Date { get; set; }

        //20200921 ALIBABA 換券需求新增欄位 202008210166-00
        [Description("債券代號_新")]
        public string Bond_Number_New { get; set; }

        [Description("Portfolio Name_新")]
        public string Portfolio_Name_New { get; set; }

        [Description("Lots_新")]
        public string Lots_New { get; set; }

        [Description("面額_新")]
        public string Ori_Amount_New { get; set; }

        [Description("一舊券換多新券")]
        public string Multi_NewBonds { get; set; }

        [Description("資料處理人員")]
        public string CreateUser { get; set; }

        [Description("面額比例_New")]
        public string Ori_Amount_Percentage_New { get; set; }

        [Description("換入原幣應計息")]
        public string IntRevise_perBond_New { get; set; }

        [Description("換入台幣應計息")]
        public string IntRevise_perBond_Tw_New { get; set; }
        //end 20200921 ALIBABA
    }
}