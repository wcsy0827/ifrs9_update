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
    
    public partial class IFRS9_User_Log
    {
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
        public IFRS9_User_Log()
        {
            this.IFRS9_Browse_Log = new HashSet<IFRS9_Browse_Log>();
        }
    
        public string User_Account { get; set; }
        public System.DateTime Login_Time { get; set; }
        public System.DateTime Login_Date { get; set; }
        public Nullable<System.DateTime> Logout_Time { get; set; }
        public string Login_IP { get; set; }
    
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<IFRS9_Browse_Log> IFRS9_Browse_Log { get; set; }
        public virtual IFRS9_User IFRS9_User { get; set; }
    }
}
