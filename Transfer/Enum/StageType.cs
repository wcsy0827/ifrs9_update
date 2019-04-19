using System.ComponentModel;

namespace Transfer.Enum
{
    public partial class Ref
    {
        public enum Stage_Type
        {
            /// <summary>
            /// 已完成 
            /// </summary>
            [Description("已完成")]
            completed,

            /// <summary>
            /// 未完成
            /// </summary>
            [Description("未完成")]
            undone,

            /// <summary>
            /// 無須評估
            /// </summary>
            [Description("無須評估")]
            No_Need_Assess,

            /// <summary>
            /// 無須複核
            /// </summary>
            [Description("無須複核")]
            No_Need_Review,
        }
    }
}