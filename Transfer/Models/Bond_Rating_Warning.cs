//------------------------------------------------------------------------------
// <auto-generated>
//     這個程式碼是由範本產生。
//
//     對這個檔案進行手動變更可能導致您的應用程式產生未預期的行為。
//     如果重新產生程式碼，將會覆寫對這個檔案的手動變更。
// </auto-generated>
//------------------------------------------------------------------------------

namespace Transfer.Models
{
    using System;
    using System.Collections.Generic;
    
    public partial class Bond_Rating_Warning
    {
        public System.DateTime Report_Date { get; set; }
        public int Version { get; set; }
        public string Bond_Number { get; set; }
        public string ISSUER { get; set; }
        public string PRODUCT { get; set; }
        public string Product_Group_1 { get; set; }
        public string Product_Group_2 { get; set; }
        public Nullable<double> Amort_Amt_Tw { get; set; }
        public Nullable<int> GRADE_Warning_F { get; set; }
        public Nullable<int> GRADE_Warning_D { get; set; }
        public Nullable<int> Rating_diff_Over_F { get; set; }
        public Nullable<int> Rating_diff_Over_N { get; set; }
        public string Rating_Worse { get; set; }
        public string Bond_Area { get; set; }
        public Nullable<int> PD_Grade { get; set; }
        public string Wraming_1_Ind { get; set; }
        public string New_Ind { get; set; }
        public Nullable<int> Observation_Month { get; set; }
        public string Rating_diff_Over_Ind { get; set; }
        public string Wraming_2_Ind { get; set; }
        public string Change_Memo { get; set; }
        public string Rating_Change_Memo { get; set; }
    }
}
