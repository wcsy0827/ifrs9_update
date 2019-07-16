using System.ComponentModel;

namespace Transfer.Enum
{
    public partial class Ref
    {
        public enum C10DataType
        {
            /// <summary>
            /// 資料為本金
            /// </summary>
            [Description("本金")]
            Amort,

            /// <summary>
            /// 資料為利息
            /// </summary>
            [Description("利息")]
            Interest,

            /// <summary>
            /// 資料為本金利息
            /// </summary>
            [Description("本金利息")]
            Amort_Interest,


            /// <summary>
            /// 資料錯誤
            /// </summary>
            [Description("錯誤資料")]
            Error_Data,

        }
    }
}