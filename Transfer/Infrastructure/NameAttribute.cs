using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Transfer.Infrastructure
{
    public class NameAttribute : Attribute
    {
        public string name;
        public NameAttribute(string name)
        {
            this.name = name;
        }
    }
}