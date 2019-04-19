namespace Transfer.ViewModels
{
    public class A58ViewModel
    {
        /// <summary>
        /// 帳戶編號
        /// </summary>
        public string Reference_Nbr { get; set; }

        /// <summary>
        /// 報導日
        /// </summary>
        public string Report_Date { get; set; }

        /// <summary>
        /// 債券編號
        /// </summary>
        public string Bond_Number { get; set; }

        /// <summary>
        /// Lots
        /// </summary>
        public string Lots { get; set; }

        /// <summary>
        /// 債券購入(認列)日期
        /// </summary>
        public string Origination_Date { get; set; }

        /// <summary>
        /// 套用參數編號
        /// </summary>
        public string Parm_ID { get; set; }

        /// <summary>
        /// 債券種類
        /// </summary>
        public string Bond_Type { get; set; }

        /// <summary>
        /// 信評種類
        /// </summary>
        public string Rating_Type { get; set; }

        /// <summary>
        /// 評等對象
        /// </summary>
        public string Rating_Object { get; set; }

        /// <summary>
        /// 國內/國外
        /// </summary>
        public string Rating_Org_Area { get; set; }

        /// <summary>
        /// 1:孰高 2:孰低
        /// </summary>
        public string Rating_Selection { get; set; }

        /// <summary>
        /// 評等主標尺_轉換
        /// </summary>
        public string Grade_Adjust { get; set; }

        /// <summary>
        /// 優先順序
        /// </summary>
        public string Rating_Priority { get; set; }

        /// <summary>
        /// 資料處理日期
        /// </summary>
        public string Processing_Date { get; set; }

        /// <summary>
        /// 資料版本
        /// </summary>
        public string Version { get; set; }

        /// <summary>
        /// Portfolio英文
        /// </summary>
        public string Portfolio_Name { get; set; }

        /// <summary>
        /// Portfolio
        /// </summary>
        public string Portfolio { get; set; }

        /// <summary>
        /// 債券產品別(揭露使用)
        /// </summary>
        public string SMF { get; set; }

        /// <summary>
        /// 發行人
        /// </summary>
        public string Issuer { get; set; }

        /// <summary>
        /// Security_Ticker
        /// </summary>
        public string Security_Ticker { get; set; }

        /// <summary>
        /// RATING_AS_OF_DATE_OVERRIDE
        /// </summary>
        public string RATING_AS_OF_DATE_OVERRIDE { get; set; }

        /// <summary>
        /// 是否為換券
        /// </summary>
        public string ISIN_Changed_Ind { get; set; }

        /// <summary>
        /// 債券編號_換券前
        /// </summary>
        public string Bond_Number_Old { get; set; }

        /// <summary>
        /// Lots_舊
        /// </summary>
        public string Lots_Old { get; set; }

        /// <summary>
        /// Portfolio英文_舊
        /// </summary>
        public string Portfolio_Name_Old { get; set; }

        /// <summary>
        /// 原始購入日_舊
        /// </summary>
        public string Origination_Date_Old { get; set; }
    }
}