using System.ComponentModel;

namespace Transfer.Enum
{
    public partial class Ref
    {
        public enum A59_SelectType
        {
            /// <summary>
            /// 全部條件
            /// </summary>
            [Description("所有條件")]
            T0,

            /// <summary>
            /// 報導日+評等種類(購買日評等、基準日最近評等) 債券編號可不輸入
            /// </summary>
            [Description("報導日+評等種類+資料版本")]
            T1,

            /// <summary>
            /// 購買日+債券編號
            /// </summary>
            [Description("購買日+債券編號")]
            T2,

            /// <summary>
            /// 報導日+評等種類(購買日評等、基準日最近評等)+債券編號 購買日可不輸入
            /// </summary>
            [Description("報導日+評等種類+債券編號+資料版本")]
            T3,

            /// <summary>
            /// 評等種類 (購買日評等、基準日最近評等)+債券編號+購買日 報導日可不輸入
            /// </summary>
            [Description("評等種類+債券編號+購買日")]
            T4,
        }
    }
}