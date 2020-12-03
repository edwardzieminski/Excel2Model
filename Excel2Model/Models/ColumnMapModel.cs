using System.Reflection;

namespace Excel2Model.Models
{
    public record ColumnMapModel
    {
        public string ColumnName { get; init; }
        public int ColumnIndex { get; init; }
        public string ColumnHeader { get; init; }
        public PropertyInfo Property { get; init; }
    }
}
