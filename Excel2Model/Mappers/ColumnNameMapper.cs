using Excel2Model.Models;
using Excel2Model.Utilities;
using Excel2Model.Validation;
using Microsoft.Office.Interop.Excel;
using Optional;
using System;
using System.Collections.Generic;
using System.Linq.Expressions;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;

namespace Excel2Model.Mappers
{
    public class ColumnNameMapper<T> where T : new()
    {
        private readonly List<ColumnMapModel<T>> _columnMapModels = new List<ColumnMapModel<T>>();
        private Option<Excel.Worksheet, ValidationError> _excelInteropWorksheetOrValidationError;
        private WorksheetModel _worksheetModel;

        public ColumnNameMapper()
        {
        }

        public ColumnNameMapper(WorksheetModel worksheetModel)
        {
            WorksheetModel = worksheetModel;
        }

        public WorksheetModel WorksheetModel
        {
            get => _worksheetModel;
            init
            {
                _worksheetModel = value;
                SetExcelInteropWorksheet();
            }
        }

        private void SetExcelInteropWorksheet()
        {
            _excelInteropWorksheetOrValidationError = ExcelInteropUtilities.TryGetWorksheet(WorksheetModel);
        }

        public Option<ColumnMapModel<T>, ValidationError> TryAddColumn(string columnName, Expression<Func<T, object>> tProperty)
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

        private ColumnMapModel<T> AddColumn(string columnName, PropertyInfo propertyInfo)
        {
            var output = new ColumnMapModel<T>()
            {
                ColumnName = columnName,
                Property = propertyInfo
            };

            _columnMapModels.Add(output);

            return output;
        }

        public Option<List<T>, ValidationError> TryGetDataFromExcel()
        {
            Option<List<T>, ValidationError> output = new Option<List<T>, ValidationError>();

            _excelInteropWorksheetOrValidationError.Match
            (
                some: excelInteropWorksheet => output = Option.Some<List<T>, ValidationError>(GetDataFromExcelInterop(excelInteropWorksheet)),
                none: validationError => output = Option.None<List<T>, ValidationError>(validationError)
            );

            return output;
        }

        public List<T> TryGetDataFromExcelWithoutErrors()
        {
            List<T> output = new List<T>();

            _excelInteropWorksheetOrValidationError.MatchSome
            (
                some: excelInteropWorksheet => output = GetDataFromExcelInterop(excelInteropWorksheet)
            );

            return output;
        }

        private List<T> GetDataFromExcelInterop(Worksheet excelInteropWorksheet)
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
