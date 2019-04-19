using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Web;
using Transfer.Infrastructure;
using Transfer.ViewModels;
using static Transfer.Enum.Ref;

namespace Transfer.Utility
{
    public class FactoryRegistry
    {
        public static IViewModel GetInstance(Table_Type type)
        {
            Type t = TableTypeHelper.GetInstanceType(type);
            return (IViewModel)Activator.CreateInstance(t);
        }

        internal class TableTypeHelper
        {
            internal static Type GetInstanceType(Table_Type type)
            {
                FieldInfo data = typeof(Table_Type).GetField(type.ToString());
                Attribute attribute = Attribute.GetCustomAttribute(data, typeof(CommucationAttribute));
                CommucationAttribute result = (CommucationAttribute)attribute;
                return result.InstanceType;
            }
        }
    }
}