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
    
    public partial class Flow_Apply_Status
    {
        public System.DateTime Report_Date { get; set; }
        public int Version { get; set; }
        public string PRJID { get; set; }
        public string FLOWID { get; set; }
        public string Group_Product_Code { get; set; }
        public System.DateTime Run_Time_Start { get; set; }
        public Nullable<System.DateTime> Run_Time_End { get; set; }
        public string Flow_Result { get; set; }
    }
}
