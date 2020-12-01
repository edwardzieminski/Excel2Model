using Excel2Model.Models;
using Excel2Model.Utilities;
using Excel2Model.Validation;
using Microsoft.Office.Interop.Excel;
using Optional;
using System;
using System.Collections.Generic;
using System.Linq.Expressions;

namespace Excel2Model.Mappers
{
    public class ColumnNameMapper<T> : AbstractMapper<T> where T : new()
    {
        public override Option<ColumnMapModel<T>, ValidationError> TryAddColumn(string columnName, Expression<Func<T, object>> tProperty)
        {
            Option<ColumnMapModel<T>, ValidationError> output = new Option<ColumnMapModel<T>, ValidationError>();

            var propertyOrValidationError = CommonUtilities.TryGetPropertyFromExpression(tProperty);

            propertyOrValidationError.Match
            (
                some: propertyInfo => output = Option.Some<ColumnMapModel<T>, ValidationError>(AddColumn(columnName, propertyInfo)),
                none: validationError => output = Option.None<ColumnMapModel<T>, ValidationError>(validationError)
            );

            return output;
        }

        private protected override List<T> GetDataFromExcelInterop(Worksheet excelInteropWorksheet)
        {
            var output = new List<T>();

            var currentRow = WorksheetModel.DataStartRow;
            bool anyValueFulfilled;

            do
            {
                var modelRecord = new T();
                foreach (var columnMap in _columnMapModels)
                {
                    var cellAddress = $"{columnMap.ColumnName}{currentRow}";
                    var cellValue = excelInteropWorksheet.Range[cellAddress].Value;
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
