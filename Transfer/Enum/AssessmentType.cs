using System.ComponentModel;

namespace Transfer.Enum
{
    public partial class Ref
    {
        public enum Assessment_Type
        {
            /// <summary>
            /// 債券(風控複核權限設定)
            /// </summary>
            [Description("001")]
            B,

            /// <summary>
            /// 房貸(房貸複核權限設定)
            /// </summary>
            [Description("002")]
            M,
        }
    }
}