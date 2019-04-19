using System.ComponentModel;

namespace Transfer.Enum
{
    public partial class Ref
    {
        /// <summary>
        /// 查詢信評缺漏狀態
        /// </summary>
        public enum Rating_Status
        {
            /// <summary>
            /// 全部資料 
            /// </summary>
            [Description("全部資料")]
            All = 1,

            /// <summary>
            /// 完全無信評 
            /// </summary>
            [Description("完全無信評")]
            AllNull = 2,

            /// <summary>
            /// 債項無信評 
            /// </summary>
            [Description("債項無信評")]
            BondsNull = 4,

            /// <summary>
            /// 發行人無信評 
            /// </summary>
            [Description("發行人無信評")]
            ISSUERNull = 8,

            /// <summary>
            /// 保證人無信評 
            /// </summary>
            [Description("保證人無信評")]
            GUARANTORNull = 16,
        }
    }
}