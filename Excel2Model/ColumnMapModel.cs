using System;

namespace Excel2Model
{
    public record ColumnMapModel<T>
    {
        public string ColumnName { get; init; }
        public Func<T, IComparable> Property { get; init; }
    }
}
