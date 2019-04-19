using System.ComponentModel;

namespace Transfer.Enum
{
    public partial class Ref
    {
        /// <summary>
        /// 覆核設定檔 類別
        /// </summary>
        public enum SetAssessmentType
        {
            /// <summary>
            /// 設定表單ID
            /// </summary>
            [Description("設定表單ID")]
            Assessment,

            /// <summary>
            /// 可覆核帳號
            /// </summary>
            [Description("可覆核帳號")]
            Auditor,

            /// <summary>
            /// 可呈送帳號
            /// </summary>
            [Description("可呈送帳號")]
            Presented,
        }
    }
}