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
        public enum Impairment_Operations_Type
        {
            /// <summary>
            /// 全部
            /// </summary>
            [Description("全部")]
            All = ImpairmentCalculation + BasicRequirements + QuantitativeAssessment +
                QuantitativeAssessmentReview + RiskControl + QualitativeAssessment + QualitativeAssessmentReview +
                StageConfrim,

            /// <summary>
            /// 減損計算 
            /// </summary>
            [Description("減損計算")]
            ImpairmentCalculation = 1,

            /// <summary>
            /// 基本要件及監控名單產出
            /// </summary>
            [Description("基本要件及監控名單產出")]
            BasicRequirements = 2,

            /// <summary>
            /// 量化評估
            /// </summary>
            [Description("量化評估")]
            QuantitativeAssessment = 4,

            /// <summary>
            /// 量化評估覆核
            /// </summary>
            [Description("量化評估覆核")]
            QuantitativeAssessmentReview = 8,

            /// <summary>
            /// 風控覆核
            /// </summary>
            [Description("風控覆核")]
            RiskControl = 16,

            /// <summary>
            /// 質化評估
            /// </summary>
            [Description("質化評估")]
            QualitativeAssessment = 32,

            /// <summary>
            /// 質化評估覆核
            /// </summary>
            [Description("質化評估覆核")]
            QualitativeAssessmentReview = 64,

            /// <summary>
            /// 階段調整確認作業
            /// </summary>
            [Description("階段調整確認作業")]
            StageConfrim = 128,
        }
    }
}