using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Transfer.Infrastructure
{
    public class SetDateFiledAttribute : Attribute
    {
        public virtual string Description { get; }
        public SetDateFiledAttribute(string DateFiled)
        {
            this.Description = DateFiled;
        }
    }
}