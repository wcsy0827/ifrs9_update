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
    
    public partial class Bond_Rating_Summary
    {
        public string Reference_Nbr { get; set; }
        public System.DateTime Report_Date { get; set; }
        public string Parm_ID { get; set; }
        public string Bond_Type { get; set; }
        public string Rating_Type { get; set; }
        public string Rating_Object { get; set; }
        public string Rating_Org_Area { get; set; }
        public string Rating_Selection { get; set; }
        public Nullable<int> Grade_Adjust { get; set; }
        public Nullable<int> Rating_Priority { get; set; }
        public Nullable<System.DateTime> Processing_Date { get; set; }
        public Nullable<int> Version { get; set; }
        public string Bond_Number { get; set; }
        public string Lots { get; set; }
        public string Portfolio { get; set; }
        public Nullable<System.DateTime> Origination_Date { get; set; }
        public string Portfolio_Name { get; set; }
        public string SMF { get; set; }
        public string ISSUER { get; set; }
        public string ISIN_Changed_Ind { get; set; }
        public string Bond_Number_Old { get; set; }
        public string Lots_Old { get; set; }
        public string Portfolio_Name_Old { get; set; }
        public Nullable<System.DateTime> Origination_Date_Old { get; set; }
        public string Create_User { get; set; }
        public Nullable<System.DateTime> Create_Date { get; set; }
        public Nullable<System.TimeSpan> Create_Time { get; set; }
        public string LastUpdate_User { get; set; }
        public Nullable<System.DateTime> LastUpdate_Date { get; set; }
        public Nullable<System.TimeSpan> LastUpdate_Time { get; set; }
    }
}