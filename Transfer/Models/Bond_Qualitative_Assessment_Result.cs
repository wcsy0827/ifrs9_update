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
    
    public partial class Bond_Qualitative_Assessment_Result
    {
        public string Reference_Nbr { get; set; }
        public int Version { get; set; }
        public System.DateTime Report_Date { get; set; }
        public string Assessment_Kind { get; set; }
        public string Assessment_Sub_Kind { get; set; }
        public string Assessment_Stage { get; set; }
        public string Check_Item_Code { get; set; }
        public string Check_Item { get; set; }
        public string Check_Item_Memo { get; set; }
        public string Check_Reference { get; set; }
        public int Assessment_Result_Version { get; set; }
        public string Assessment_Result { get; set; }
        public Nullable<double> Pass_Value { get; set; }
        public string Qualitative_Pass_Stage2 { get; set; }
        public string Qualitative_Pass_Stage3 { get; set; }
        public string Check_Symbol { get; set; }
        public Nullable<double> Threshold { get; set; }
        public string Memo { get; set; }
        public string Bond_Number { get; set; }
        public string Create_User { get; set; }
        public Nullable<System.DateTime> Create_Date { get; set; }
        public Nullable<System.TimeSpan> Create_Time { get; set; }
        public string LastUpdate_User { get; set; }
        public Nullable<System.DateTime> LastUpdate_Date { get; set; }
        public Nullable<System.TimeSpan> LastUpdate_Time { get; set; }
    }
}