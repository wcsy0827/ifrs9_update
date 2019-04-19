using NLog;
using System;
using System.Collections.Generic;
using System.Data.Entity.Infrastructure;
using System.Linq;
using System.Web;

namespace Transfer.Models
{
    public class IFRS9DBEntities : IFRS9Entities
    {
        public IFRS9DBEntities()
        {
            ((IObjectContextAdapter)this).ObjectContext.CommandTimeout = 300;
            Database.Log = str => Utility.Extension.NlogSet(str,Enum.Ref.Nlog.Info);
        }
    }
}