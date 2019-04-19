using System.ComponentModel;

namespace Transfer.ViewModels
{
    public class C03ViewModel
    {
        [Description("資料處理日期")]
        public string Processing_Date { get; set; }

        [Description("產品")]
        public string Product_Code { get; set; }

        [Description("執行資料批號")]
        public string Data_ID { get; set; }

        [Description("資料(季)")]
        public string Year_Quartly { get; set; }

        [Description("季違約資料")]
        public string PD_Quartly { get; set; }

        [Description("台灣證劵交易指數")]
        public string TWSE_Index { get; set; }

        [Description("台灣GDP(經季調)")]
        public string TWRGSARP_Index { get; set; }

        [Description("台灣GDP(未經季調)")]
        public string TWGDPCON_Index { get; set; }

        [Description("失業率(經季調)")]
        public string TWLFADJ_Index { get; set; }

        [Description("台灣實質消費者物價指數(未經季調)")]
        public string TWCPI_Index { get; set; }

        [Description("現行貨幣總計數M1A")]
        public string TWMSA1A_Index { get; set; }

        [Description("現行貨幣總計數M1B")]
        public string TWMSA1B_Index { get; set; }

        [Description("廣義貨幣供給額M2")]
        public string TWMSAM2_Index { get; set; }

        [Description("國債收益率(十年)")]
        public string GVTW10YR_Index { get; set; }

        [Description("台灣貿易收支")]
        public string TWTRBAL_Index { get; set; }

        [Description("台灣貿易出口")]
        public string TWTREXP_Index { get; set; }

        [Description("台灣貿易進口")]
        public string TWTRIMP_Index { get; set; }

        [Description("央行重貼現率")]
        public string TAREDSCD_Index { get; set; }

        [Description("台灣實質景氣領先指標(經季調)")]
        public string TWCILI_Index { get; set; }

        [Description("台灣經常帳收支")]
        public string TWBOPCUR_Index { get; set; }

        [Description("Taiwan Current Account Balance (% GDP)")]
        public string EHCATW_Index { get; set; }

        [Description("台灣實質工業生產指數")]
        public string TWINDPI_Index { get; set; }

        [Description("台灣實質躉售物價指數 未經季調")]
        public string TWWPI_Index { get; set; }

        [Description("台灣零售銷售年成長率")]
        public string TARSYOY_Index { get; set; }

        [Description("台灣外銷訂單(年化)")]
        public string TWEOTTL_Index { get; set; }

        [Description("國內受訪者淨%-緊縮標準-對大型/中型公司之工商業貸款")]
        public string SLDETIGT_Index { get; set; }

        [Description("台灣外匯準備")]
        public string TWIRFE_Index { get; set; }

        [Description("信義房價指數")]
        public string SINYI_HOUSE_PRICE_index { get; set; }

        [Description("國泰不動產業指數")]
        public string CATHAY_ESTATE_index { get; set; }

        [Description("實質台灣國內生產毛額(季調整)")]
        public string Real_GDP2011 { get; set; }

        [Description("Master Card消費者信心指數")]
        public string MCCCTW_Index { get; set; }

        [Description("TWD 隔夜利率 Deposit")]
        public string TRDR1T_Index { get; set; }
    }
}