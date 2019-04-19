using System.ComponentModel;

namespace Transfer.Enum
{
    public partial class Ref
    {
        /// <summary>
        /// 可以轉檔的Table
        /// </summary>
        public enum Transfer_Table_Type
        {
            //190222 因自動補信評需求增加
            /// <summary>
            /// Bond_Rating_Update_Info
            /// </summary>
            [Description("Bond_Rating_Update_Info")]
            A59,
            /// <summary>
            /// A59自動由寶碩補入信評
            /// </summary>
            [Description("A59Apex")]
            A59Apex,
            /// <summary>
            /// A59轉入A57A58
            /// </summary>
            [Description("A59Trans")]
            A59Trans,

            /// <summary>
            /// Bond_Account_Info
            /// </summary>
            [Description("Bond_Account_Info")]
            A41,

            /// <summary>
            /// Treasury_Securities_Info
            /// </summary>
            [Description("Treasury_Securities_Info")]
            A42,

            /// <summary>
            /// Rating_Info
            /// </summary>
            [Description("Rating_Info")]
            A53,

            /// <summary>
            /// Bond_Rating_Info
            /// </summary>
            [Description("Bond_Rating_Info")]
            A57,

            /// <summary>
            /// Bond_Rating_Summary
            /// </summary>
            [Description("Bond_Rating_Summary")]
            A58,

            /// <summary>
            /// IFRS9_Main
            /// </summary>
            [Description("IFRS9_Main")]
            B01,

            /// <summary>
            /// EL_Data_In
            /// </summary>
            [Description("EL_Data_In")]
            C01,

            /// <summary>
            /// EL_Data_In
            /// </summary>
            [Description("EL_Data_In")]
            C01_HK,

            /// <summary>
            /// EL_Data_In
            /// </summary>
            [Description("EL_Data_In")]
            C01_VN,

            /// <summary>
            /// Rating_History
            /// </summary>
            [Description("Rating_History")]
            C02,

            /// <summary>
            /// EL_Data_Out
            /// </summary>
            [Description("EL_Data_Out")]
            C07,
        }
    }
}