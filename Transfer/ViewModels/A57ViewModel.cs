using System.ComponentModel;
namespace Transfer.ViewModels
{
    public class A57ViewModel
    {
        /// <summary>
        /// 帳戶編號
        /// </summary>
        [Description("帳戶編號")]
        public string Reference_Nbr { get; set; }

        /// <summary>
        /// 債券編號
        /// </summary>
        [Description("債券編號")]
        public string Bond_Number { get; set; }

        /// <summary>
        /// Lots
        /// </summary>
        [Description("Lots")]
        public string Lots { get; set; }

        /// <summary>
        /// Portfolio
        /// </summary>
        [Description("Portfolio")]
        public string Portfolio { get; set; }

        /// <summary>
        /// 債券(資產)名稱
        /// </summary>
        [Description("債券(資產)名稱")]
        public string Segment_Name { get; set; }

        /// <summary>
        /// 債券種類
        /// </summary>
        [Description("債券種類")]
        public string Bond_Type { get; set; }

        /// <summary>
        /// 債券擔保順位
        /// </summary>
        [Description("債券擔保順位")]
        public string Lien_position { get; set; }

        /// <summary>
        /// 債券購入(認列)日期
        /// </summary>
        [Description("債券購入(認列)日期")]
        public string Origination_Date { get; set; }

        /// <summary>
        /// 評估日/報導日
        /// </summary>
        [Description("評估日/報導日")]
        public string Report_Date { get; set; }

        /// <summary>
        /// 評等資料時點
        /// </summary>
        [Description("評等資料時點")]
        public string Rating_Date { get; set; }

        /// <summary>
        /// 評等種類
        /// </summary>
        [Description("評等種類")]
        public string Rating_Type { get; set; }

        /// <summary>
        /// 評等對象
        /// </summary>
        [Description("評等對象")]
        public string Rating_Object { get; set; }

        /// <summary>
        /// 評等機構
        /// </summary>
        [Description("評等機構")]
        public string Rating_Org { get; set; }

        /// <summary>
        /// 評等內容
        /// </summary>
        [Description("評等內容")]
        public string Rating { get; set; }

        /// <summary>
        /// 國內/國外
        /// </summary>
        [Description("國內/國外")]
        public string Rating_Org_Area { get; set; }

        /// <summary>
        /// 是否人工補登
        /// </summary>
        [Description("是否人工補登")]
        public string Fill_up_YN { get; set; }

        /// <summary>
        /// 人工補登日期
        /// </summary>
        [Description("人工補登日期")]
        public string Fill_up_Date { get; set; }

        /// <summary>
        /// 評等主標尺_原始
        /// </summary>
        [Description("評等主標尺_原始")]
        public string PD_Grade { get; set; }

        /// <summary>
        /// 評等主標尺_轉換
        /// </summary>
        [Description("評等主標尺_轉換")]
        public string Grade_Adjust { get; set; }

        /// <summary>
        /// ISSUER_TICKER
        /// </summary>
        [Description("ISSUER_TICKER")]
        public string ISSUER_TICKER { get; set; }

        /// <summary>
        /// GUARANTOR_NAME
        /// </summary>
        [Description("GUARANTOR_NAME")]
        public string GUARANTOR_NAME { get; set; }

        /// <summary>
        /// GUARANTOR_EQY_TICKER
        /// </summary>
        [Description("GUARANTOR_EQY_TICKER")]
        public string GUARANTOR_EQY_TICKER { get; set; }

        /// <summary>
        /// 信評優先選擇參數編號
        /// </summary>
        [Description("信評優先選擇參數編號")]
        public string Parm_ID { get; set; }

        /// <summary>
        /// Portfolio英文
        /// </summary>
        [Description("Portfolio英文")]
        public string Portfolio_Name { get; set; }

        /// <summary>
        /// Bloomberg評等欄位名稱
        /// </summary>
        [Description("Bloomberg評等欄位名稱")]
        public string RTG_Bloomberg_Field { get; set; }

        /// <summary>
        /// 債券產品別(揭露使用)
        /// </summary>
        [Description("債券產品別(揭露使用)")]
        public string SMF { get; set; }

        /// <summary>
        /// 發行人
        /// </summary>
        [Description("發行人")]
        public string ISSUER { get; set; }

        /// <summary>
        /// 資料版本
        /// </summary>
        [Description("資料版本")]
        public string Version { get; set; }

        /// <summary>
        /// Bloomberg Security_Ticker
        /// </summary>
        [Description("Bloomberg Security_Ticker")]
        public string Security_Ticker { get; set; }

        /// <summary>
        /// 是否為換券
        /// </summary>
        [Description("是否為換券")]
        public string ISIN_Changed_Ind { get; set; }

        /// <summary>
        /// 債券編號_換券前
        /// </summary>
        [Description("債券編號_換券前")]
        public string Bond_Number_Old { get; set; }

        /// <summary>
        /// Lots_舊
        /// </summary>
        [Description("Lots_舊")]
        public string Lots_Old { get; set; }

        /// <summary>
        /// Portfolio英文_舊
        /// </summary>
        [Description("Portfolio英文_舊")]
        public string Portfolio_Name_Old { get; set; }

        /// <summary>
        /// 原始購入日_舊
        /// </summary>
        [Description("原始購入日_舊")]
        public string Origination_Date_Old { get; set; }
    }
}