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
    
    public partial class Fixed_Income_AbuDhabi_Info
    {
        public int Data_Year { get; set; }
        public System.DateTime Processing_Date { get; set; }
        public string Country { get; set; }
        public Nullable<double> IGS_Index { get; set; }
        public Nullable<double> External_Debt_Rate { get; set; }
        public Nullable<double> FC_Indexed_Debt_Rate { get; set; }
        public Nullable<double> CEIC_Value { get; set; }
        public Nullable<double> Short_term_Debt { get; set; }
        public Nullable<double> Foreign_Exchange { get; set; }
        public Nullable<double> GDP_Yearly { get; set; }
        public string Effective { get; set; }
    }
}
