using System;
using System.Reflection;

namespace Excel2Model
{
    public record ColumnMapModel<T>
    {
        public string ColumnName { get; init; }
        public PropertyInfo Property { get; init; }
    }
}
