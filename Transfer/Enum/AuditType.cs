using System.ComponentModel;

namespace Transfer.Enum
{
    public partial class Ref
    {
        public enum Audit_Type
        {
            /// <summary>
            /// 待確認
            /// </summary>
            [Description("待確認")]
            None = 0,

            /// <summary>
            /// 啟用 (新主標尺啟用後,原啟用主標尺即停用)
            /// </summary>
            [Description("啟用")]
            Enable = 1,

            /// <summary>
            /// 暫不啟用
            /// </summary>
            [Description("暫不啟用")]
            TempDisabled = 2,

            /// <summary>
            /// 停用
            /// </summary>
            [Description("停用")]
            Disabled = 3,

            /// <summary>
            /// 全部
            /// </summary>
            [Description("全部")]
            All = 4,
        }
    }
}