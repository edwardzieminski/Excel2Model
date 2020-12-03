using Excel = Microsoft.Office.Interop.Excel;
using Excel2Model.Models;
using Excel2Model.Utilities;
using Excel2Model.Validation;
using Optional;
using System;
using System.Collections.Generic;
using System.Linq.Expressions;
using System.Reflection;

namespace Excel2Model.Mappers
{
    public class TableHeaderMapper<T> : AbstractMapper<T> where T : new()
    {
        public Option<ColumnMapModel<T>, ValidationError> TryAddColumn(string columnHeader, Expression<Func<T, object>> tProperty)
        {
            Option<ColumnMapModel<T>, ValidationError> output = new Option<ColumnMapModel<T>, ValidationError>();

            var propertyOrValidationError = CommonUtilities.TryGetPropertyFromExpression(tProperty);

            propertyOrValidationError.Match
            (
                some: propertyInfo => output = Option.Some<ColumnMapModel<T>, ValidationError>(AddColumn(columnHeader, propertyInfo)),
                none: validationError => output = Option.None<ColumnMapModel<T>, ValidationError>(validationError)
            );

            return output;
        }

        private protected override ColumnMapModel<T> AddColumn(string columnHeader, PropertyInfo propertyInfo)
        {
            var output = new ColumnMapModel<T>()
            {
                ColumnHeader = columnHeader,
                Property = propertyInfo
            };

            _columnMapModels.Add(output);

            return output;
        }

        private protected override List<T> GetDataFromExcelInterop(Excel.Worksheet excelInteropWorksheet)
        {
            throw new NotImplementedException();
        }
    }
}
