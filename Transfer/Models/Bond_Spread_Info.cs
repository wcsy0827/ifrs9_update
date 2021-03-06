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
    
    public partial class Bond_Spread_Info
    {
        public string Reference_Nbr { get; set; }
        public System.DateTime Report_Date { get; set; }
        public int Version { get; set; }
        public string Bond_Number { get; set; }
        public string Lots { get; set; }
        public string Portfolio_Name { get; set; }
        public System.DateTime Origination_Date { get; set; }
        public double EIR { get; set; }
        public System.DateTime Processing_Date { get; set; }
        public Nullable<double> Mid_Yield { get; set; }
        public string BNCHMRK_TSY_ISSUE_ID { get; set; }
        public string ID_CUSIP { get; set; }
        public Nullable<double> Spread_Current { get; set; }
        public Nullable<double> Spread_When_Trade { get; set; }
        public Nullable<double> Treasury_Current { get; set; }
        public Nullable<double> Treasury_When_Trade { get; set; }
        public Nullable<double> All_in_Chg { get; set; }
        public Nullable<double> Chg_In_Spread { get; set; }
        public Nullable<double> Chg_In_Treasury { get; set; }
        public string Memo { get; set; }
        public string LastUpdate_User { get; set; }
        public Nullable<System.DateTime> LastUpdate_Date { get; set; }
        public Nullable<System.TimeSpan> LastUpdate_Time { get; set; }
    }
}
