using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EdCanHack.SheetParser.Transforms.Serialization
{
    /// <summary>
    /// Indicates that the decorated class can be deserialized by SheetParser.
    /// </summary>
    [AttributeUsage(AttributeTargets.Class, AllowMultiple = false, Inherited = false)]
    public class SheetDeserializableAttribute : Attribute
    {
        public readonly Boolean IgnoreExtraFields;
        public readonly Boolean IgnoreMissingFields;

        public SheetDeserializableAttribute(bool ignoreExtraFields = true, bool ignoreMissingFields = false)
        {
            IgnoreExtraFields = ignoreExtraFields;
            IgnoreMissingFields = ignoreMissingFields;
        }
    }


    /// <summary>
    /// Indicates that this property should be deserialized via SheetParser. When a class
    /// being deserialized is of type IKeyed[T], this attribute is assumed with header
    /// name "Key".
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = true)]
    public class SheetPropertyAttribute : Attribute
    {
        public readonly String ColumnName;

        public SheetPropertyAttribute(string columnName)
        {
            ColumnName = columnName;
        }
    }
}
