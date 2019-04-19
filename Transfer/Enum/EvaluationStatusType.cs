using System;
using System.ComponentModel;

namespace Transfer.Enum
{
    public partial class Ref
    {
        /// <summary>
        /// 評估狀態
        /// </summary>
        [Flags]
        public enum Evaluation_Status_Type
        {
            /// <summary>
            /// 全部
            /// </summary>
            [Description("全部")]
            All = 0,

            /// <summary>
            /// 尚未評估
            /// </summary>
            [Description("尚未評估")]
            NotAssess = 1,

            /// <summary>
            /// 尚未提交複核(已評估)
            /// </summary>
            [Description("尚未提交複核(已評估)")]
            NotReview = 2,

            /// <summary>
            /// 已提交複核(等待處理中)
            /// </summary>
            [Description("已提交複核(等待處理中)")]
            Review = 4,

            /// <summary>
            /// 複核者選擇退回
            /// </summary>
            [Description("複核者選擇退回")]
            Reject = 8,

            /// <summary>
            /// 複核完成
            /// </summary>
            [Description("複核完成")]
            ReviewCompleted = 16
        }
    }
}