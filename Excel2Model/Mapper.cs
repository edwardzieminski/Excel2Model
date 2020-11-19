using System;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;

namespace Excel2Model
{
    public class Mapper<T>
    {
        private readonly List<ColumnMapModel<T>> _columnMapModels = new List<ColumnMapModel<T>>();

        public Excel.Worksheet Worksheet { get; init; }
        public int HeaderRow { get; init; }
        public int DataStartRow { get; init; }

        public void AddColumn(string columnName, Func<T, IComparable> tProperty)
        {
            _columnMapModels.Add(new ColumnMapModel<T>()
            { 
                ColumnName = columnName,
                Property = tProperty
            });
        }

        
    }
}
