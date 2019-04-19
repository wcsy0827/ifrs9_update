using System.ComponentModel;

namespace Transfer.Enum
{
    public partial class Ref
    {
        public enum AssessmentSubKind_Type
        {
            /// <summary>
            /// 金融債主順位債券
            /// </summary>
            [Description("金融債主順位債券")]
            Financial_Debts,

            /// <summary>
            /// 金融債次順位債券
            /// </summary>
            [Description("金融債次順位債券")]
            Financial_Debt_Bond,

            /// <summary>
            /// 產業公司
            /// </summary>
            [Description("產業公司")]
            Industrial_Company,

            /// <summary>
            /// 壽險與產險公司
            /// </summary>
            [Description("壽險與產險公司")]
            Life_Or_Property_Insurance_Company,

            /// <summary>
            /// 主權政府債
            /// </summary>
            [Description("主權政府債")]
            Sovereign_Government_Debt,

            /// <summary>
            /// CLO
            /// </summary>
            [Description("CLO")]
            CLO,

            /// <summary>
            /// 其他
            /// </summary>
            [Description("其他")]
            Other,
        }
    }
}