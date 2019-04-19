using System.ComponentModel;

namespace Transfer.Enum
{
    public partial class Ref
    {
        public enum Rating_Update_Method
        {
            /// <summary>
            /// 置換特殊字元成空值
            /// </summary>
            [Description("置換特殊字元成空值")]
            One = 01,

            /// <summary>
            /// 以新值取代舊值
            /// </summary>
            [Description("以新值取代舊值")]
            Two = 02,

        }
    }
}