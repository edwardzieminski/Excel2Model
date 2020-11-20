using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using Excel = Microsoft.Office.Interop.Excel;

namespace Excel2Model
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
                Property = Utilities.GetPropertyFromExpression(tProperty)
            });
        }

        public List<T> GetDataFromExcel()
        {
            var output = new List<T>();
            var currentRow = DataStartRow;
            var anyValueFulfilled = true;

            do
            {
                var modelRecord = new T();
                foreach (var columnMap in _columnMapModels)
                {
                    var cellAddress = $"{columnMap.ColumnName}{currentRow}";
                    var cellValue = Worksheet.Range[cellAddress].Value;
                    columnMap.Property.SetValue(modelRecord, cellValue);
                }

                anyValueFulfilled = IsAnyValueFulfilled(modelRecord);

                if (anyValueFulfilled)
                {
                    output.Add(modelRecord);
                }

                currentRow++;
            } while (anyValueFulfilled);

            return output;
        }

        private static bool IsAnyValueFulfilled(T modelRecord)
        {
            var anyStringFulfilled = modelRecord.GetType().GetProperties()
                .Where(pi => pi.PropertyType == typeof(string))
                .Select(pi => (string)pi.GetValue(modelRecord))
                .Any(value => string.IsNullOrEmpty(value) == false);

            var anyIntDifferentThanZero = modelRecord.GetType().GetProperties()
                .Where(pi => pi.PropertyType == typeof(int))
                .Select(pi => (int)pi.GetValue(modelRecord))
                .Any(value => value != 0);

            var anyDoubleDifferentThanZero = modelRecord.GetType().GetProperties()
                .Where(pi => pi.PropertyType == typeof(double))
                .Select(pi => (double)pi.GetValue(modelRecord))
                .Any(value => value != 0d);

            return anyStringFulfilled || anyIntDifferentThanZero || anyDoubleDifferentThanZero;
        }
    }
}
