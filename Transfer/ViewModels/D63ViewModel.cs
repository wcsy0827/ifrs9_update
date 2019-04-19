using System.ComponentModel;

namespace Transfer.ViewModels
{
    public class D63ViewModel
    {
        [Description("狀態")]
        public string Status { get; set; }

        [Description("評估次分類")]
        public string Assessment_Sub_Kind { get; set; }

        [Description("已評估")]
        public string Pass_Confirm_Flag { get; set; }

        [Description("複核者已選擇最終版本")]
        public string Result_Version_Confirm_Flag { get; set; }

        [Description("是否通過量化評估")]
        public string Quantitative_Pass_Confirm { get; set; }

        [Description("檔案數量")]
        public string FilesCount { get; set; }

        [Description("帳戶編號/群組編號")]
        public string Reference_Nbr { get; set; }

        [Description("評估結果版本")]
        public string Assessment_Result_Version { get; set; }

        [Description("債券編號")]
        public string Bond_Number { get; set; }

        [Description("Lots")]
        public string Lots { get; set; }

        [Description("債券(資產)名稱")]
        public string Segment_Name { get; set; }

        [Description("資料版本")]
        public string Version { get; set; }

        [Description("評估基準日/報導日")]
        public string Report_Date { get; set; }

        [Description("資料處理日期")]
        public string Processing_Date { get; set; }

        [Description("Net_Debt")]
        public string Net_Debt { get; set; }

        [Description("Total_Asset")]
        public string Total_Asset { get; set; }

        [Description("CFO")]
        public string CFO { get; set; }

        [Description("Int_Expense")]
        public string Int_Expense { get; set; }

        [Description("銀行核心一級資本適足率")]
        public string CET1_Capital_Ratio { get; set; }

        [Description("銀行財務槓桿比率")]
        public string Lerverage_Ratio { get; set; }

        [Description("銀行流動性比率")]
        public string Liquiduty_Coverage_Ratio { get; set; }

        [Description("Total_Equity")]
        public string Total_Equity { get; set; }

        [Description("BS_TOT_ASSET")]
        public string BS_TOT_ASSET { get; set; }

        [Description("資本水準_美國RBC_Ratio")]
        public string RBC_Ratio { get; set; }

        [Description("資本水準_加拿大MCCSR_Ratio")]
        public string MCCSR_Ratio { get; set; }

        [Description("資本水準_歐洲Solvency_II_Ratio")]
        public string Solvency_II_Ratio { get; set; }

        [Description("SHORT_AND_LONG_TERM_DEBT")]
        public string SHORT_AND_LONG_TERM_DEBT { get; set; }

        [Description("流動性緩衝部位")]
        public string Cash_and_Cash_Equivalents { get; set; }

        [Description("一年內到期借款")]
        public string Short_Term_Debt { get; set; }

        [Description("Issuer所屬區域")]
        public string ISSUER_AREA { get; set; }

        [Description("政府債務/GDP")]
        public string IGS_INDEX { get; set; }

        [Description("過去4年政府債務/GDP增加數")]
        public string IGS_INDEX_Increase { get; set; }

        [Description("外人持有政府債務/總政府債務")]
        public string External_Debt_Rate { get; set; }

        [Description("外幣計價政府債務/總政府債務")]
        public string FC_Indexed_Debt_Rate { get; set; }

        [Description("(經常帳+FDI淨流入)/GDP")]
        public string CEIC_Value { get; set; }

        [Description("短期外債/外匯儲備")]
        public string Short_term_Debt_Ratio { get; set; }

        [Description("過去四年平均經濟成長")]
        public string GDP_Yearly_Avg { get; set; }

        [Description("加權平均信用評級因子")]
        public string Quantitative_CLO_1 { get; set; }

        [Description("違約資產金額*損失率.")]
        public string Quantitative_CLO_2 { get; set; }

        [Description("較低順位券次(tranche)之本金餘額")]
        public string Quantitative_CLO_2_Threshold { get; set; }

        [Description("資產池本金金額+現金")]
        public string Quantitative_CLO_3 { get; set; }

        [Description("本身及較高順位券次(tranche)之本金餘額")]
        public string Quantitative_CLO_3_Threshold { get; set; }

        [Description("主順位有擔保貸款資產")]
        public string Quantitative_CLO_4 { get; set; }

        [Description("資產池80%")]
        public string Quantitative_CLO_4_Threshold { get; set; }

        [Description("Portfolio")]
        public string Portfolio { get; set; }

        [Description("債券購入(認列)日期")]
        public string Origination_Date { get; set; }

        [Description("Portfolio英文")]
        public string Portfolio_Name { get; set; }

        [Description("評估者")]
        public string Assessor { set; get;}

        [Description("評估者名稱")]
        public string Assessor_Name { set; get; }

        [Description("評估日期")]
        public string Assessment_date { set; get; }

        [Description("複核者")]
        public string Auditor { set; get; }

        [Description("複核者名稱")]
        public string Auditor_Name { set; get; }

        [Description("複核日期")]
        public string Audit_date { set; get; }

        [Description("複核後選擇版本")]
        public string Result_Version_Confirm { get; set; }

        [Description("複核者不接受選擇退回")]
        public string Auditor_Return { get; set; }

        [Description("是否提交複核")]
        public string Send_to_Auditor { get; set; }

        [Description("提交時間")]
        public string Send_Time { get; set; }


    }
}