﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Transfer.Infrastructure
{
    public class CommucationAttribute : Attribute
    {
        public Type InstanceType;
        public CommucationAttribute(Type type)
        {
            InstanceType = type;
        }
    }
}