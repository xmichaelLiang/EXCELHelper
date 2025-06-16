using System;
using System.Collections.Generic;
using System.Text;

namespace EXCELHelper
{
    [AttributeUsage(AttributeTargets.Property)]
    public class PropertyColumnNameAttribute : Attribute
    {
        public string ColumnName { get; }
        public PropertyColumnNameAttribute(string columnName)
        {
            ColumnName = columnName;
        }
    }
}
