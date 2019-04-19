using System.ComponentModel;

namespace Transfer.Enum
{
    public partial class Ref
    {
        public enum Debt_Type
        {
            /// <summary>
            /// 房貸
            /// </summary>
            [Description("Mortgage")]
            M,

            /// <summary>
            /// 債券
            /// </summary>
            [Description("Bonds")]
            B,
        }
    }
}