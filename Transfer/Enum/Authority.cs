using System.ComponentModel;

namespace Transfer.Enum
{
    public partial class Ref
    {
        public enum Authority_Type
        {
            /// <summary>
            /// 無相關權限
            /// </summary>
            [Description("無相關權限")]
            None,

            /// <summary>
            /// 有呈送複核權限
            /// </summary>
            [Description("有呈送複核權限")]
            Presented,

            /// <summary>
            /// 有複核權限
            /// </summary>
            [Description("有複核權限")]
            Auditor,
        }
    }
}