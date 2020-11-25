using Excel2Model.Models;
using Excel2Model.Utilities;
using System;
using System.Collections.Generic;
using System.Linq.Expressions;
using Excel = Microsoft.Office.Interop.Excel;

namespace Excel2Model.Mappers
{
    public class Mapper<T> where T : new()
    {
        private readonly List<ColumnMapModel<T>> _columnMapModels = new List<ColumnMapModel<T>>();

        public Excel.Worksheet Worksheet { get; init; }
        public int HeaderRow { get; init; }
        public int DataStartRow { get; init; }

        public void AddColumn(string columnName, Expression<Func<T, object>> tProperty)
        {
            _columnMapModels.Add(new ColumnMapModel<T>()
            {
                ColumnName = columnName,
                Property = CommonUtilities.GetPropertyFromExpression(tProperty)
            });
        }

        public List<T> GetDataFromExcel()
        {
            var output = new List<T>();
            var currentRow = DataStartRow;
            bool anyValueFulfilled;

            do
            {
                var modelRecord = new T();
                foreach (var columnMap in _columnMapModels)
                {
                    var cellAddress = $"{columnMap.ColumnName}{currentRow}";
                    var cellValue = Worksheet.Range[cellAddress].Value;
                    columnMap.Property.SetValue(modelRecord, cellValue);
                }

                anyValueFulfilled = CommonUtilities.IsAnyValueFulfilled(modelRecord);

                if (anyValueFulfilled)
                {
                    output.Add(modelRecord);
                }

                currentRow++;
            } while (anyValueFulfilled);

            return output;
        }
    }
}
