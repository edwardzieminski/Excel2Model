using Excel2Model.Models;
using System;
using System.Reflection;

namespace Excel2Model.Mappers
{
    public class TableHeaderMapper<T> : AbstractMapper<T> where T : new()
    {
        public override void ResolveMap()
        {
            throw new NotImplementedException();
        }

        private protected override ColumnMapModel AddColumn(string columnName, PropertyInfo propertyInfo)
        {
            throw new NotImplementedException();
        }
    }
}
