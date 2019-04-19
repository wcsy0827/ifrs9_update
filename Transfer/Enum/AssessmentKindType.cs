using System.ComponentModel;

namespace Transfer.Enum
{
    public partial class Ref
    {
        public enum AssessmentKind_Type
        {
            /// <summary>
            /// 基本要件
            /// </summary>
            [Description("基本要件")]
            Basic,

            /// <summary>
            /// 量化衡量
            /// </summary>
            [Description("量化衡量")]
            Quantify,

            /// <summary>
            /// 質化衡量
            /// </summary>
            [Description("質化衡量")]          
            Qualitative,
        }
    }
}