using System.ComponentModel;

namespace Transfer.ViewModels
{
    public class D62ViewModel
    {
        [Description("帳戶編號/群組編號")]
        public string Reference_Nbr { get; set; }

        [Description("資料版本")]
        public string Version { get; set; }

        [Description("評估基準日/報導日")]
        public string Report_Date { get; set; }

        [Description("資料處理日期")]
        public string Processing_Date { get; set; }

        [Description("債券種類")]
        public string Bond_Type { get; set; }

        [Description("Stage評估次分類")]
        public string Assessment_Sub_Kind { get; set; }

        [Description("債券擔保順位")]
        public string Lien_position { get; set; }

        [Description("債券編號")]
        public string Bond_Number { get; set; }

        [Description("Lots")]
        public string Lots { get; set; }

        [Description("Portfolio")]
        public string Portfolio { get; set; }

        [Description("Portfolio英文")]
        public string Portfolio_Name { get; set; }

        [Description("債券購入(認列)日期")]
        public string Origination_Date { get; set; }

        [Description("起始(原始購買)外部評等")]
        public string Original_External_Rating { get; set; }

        [Description("最近(報導日)外部信評評等")]
        public string Current_External_Rating { get; set; }

        [Description("原始購買評等是否符合信用風險低條件")]
        public string Rating_Ori_Good_IND { get; set; }

        [Description("報導日是否符合信用風險低條件")]
        public string Rating_Curr_Good_Ind { get; set; }

        [Description("評等下降數")]
        public string Curr_Ori_Rating_Diff { get; set; }

        [Description("基本要件通過與否")]
        public string Basic_Pass { get; set; }

        [Description("對應D69-基本要件參數檔規則編號")]
        public string Map_Rule_Id_D69 { get; set; }

        [Description("金融資產餘額-攤銷後之成本數(原幣)")]
        public string Cost_Value { get; set; }

        [Description("市價(原幣)")]
        public string Market_Value_Ori { get; set; }

        [Description("未實現累計損失率")]
        public string Value_Change_Ratio { get; set; }

        [Description("未實現累計損失率是否低於30%")]
        public string Value_Change_Ratio_Pass { get; set; }

        [Description("未實現累計損失月數_上個月狀況")]
        public string Accumulation_Loss_last_Month { get; set; }

        [Description("未實現累計損失月數_本月狀況")]
        public string Accumulation_Loss_This_Month { get; set; }

        [Description("信用利差")]
        public string Chg_In_Spread { get; set; }

        [Description("信用利差超過拓寬300基點(含)以上")]
        public string Spread_Change_Over { get; set; }

        [Description("信用利差_上個月狀況")]
        public string Chg_In_Spread_last_Month { get; set; }

        [Description("信用利差_本月狀況(信用利差目前累計逾越月數)")]
        public string Chg_In_Spread_This_Month { get; set; }

        [Description("是否為觀察名單")]
        public string Watch_IND { get; set; }

        [Description("對應D70-觀察名單參數檔規則編號")]
        public string Map_Rule_Id_D70 { get; set; }

        [Description("是否為預警名單")]
        public string Warning_IND { get; set; }

        [Description("對應D71-預警名單參數檔規則編號")]
        public string Map_Rule_Id_D71 { get; set; }

        [Description("觀察名單手動調整狀態")]
        public string Watch_IND_Override { get; set; }

        [Description("觀察名單手動調整原因")]
        public string Watch_IND_Override_Desc { get; set; }

        [Description("觀察名單手動調整日期")]
        public string Watch_IND_Override_Date { get; set; }

        [Description("觀察名單手動調整者帳號")]
        public string Watch_IND_Override_User { get; set; }

        [Description("觀察名單手動調整者名稱")]
        public string Watch_IND_Override_User_Name { get; set; }

        [Description("預警名單手動調整狀態")]
        public string Warning_IND_Override { get; set; }

        [Description("預警名單手動調整原因")]
        public string Warning_IND_Override_DESC { get; set; }

        [Description("預警名單手動調整日期")]
        public string Warning_IND_Override_Date { get; set; }

        [Description("預警名單手動調整者帳號")]
        public string Warning_IND_Override_User { get; set; }

        [Description("預警名單手動調整者名稱")]
        public string Warning_IND_Override_User_Name { get; set; }
        [Description("覆核專區狀態")]
        public string ReviewStatus { get; set; }
    }
}