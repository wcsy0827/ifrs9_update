using System.ComponentModel;

namespace Transfer.Enum
{
    public partial class Ref
    {
        public enum Product_Code
        {
            /// <summary>
            /// 房貸
            /// </summary>
            [Description("Loan_1")]
            M,

            /// <summary>
            /// 債券A
            /// </summary>
            [Description("Bond_A")]
            B_A,

            /// <summary>
            /// 債券B
            /// </summary>
            [Description("Bond_B")]
            B_B,

            /// <summary>
            /// 債券P
            /// </summary>
            [Description("Bond_P")]
            B_P,

            /// <summary>
            /// 香港A
            /// </summary>
            [Description("Bond_HK_A")]
            HK_A,

            /// <summary>
            /// 香港B
            /// </summary>
            [Description("Bond_HK_B")]
            HK_B,

            /// <summary>
            /// 香港P
            /// </summary>
            [Description("Bond_HK_P")]
            HK_P,

            /// <summary>
            /// 越南A
            /// </summary>
            [Description("Bond_VN_A")]
            VN_A,

            /// <summary>
            /// 越南B
            /// </summary>
            [Description("Bond_VN_B")]
            VN_B,

            /// <summary>
            /// 越南P
            /// </summary>
            [Description("Bond_VN_P")]
            VN_P,
        }
    }
}