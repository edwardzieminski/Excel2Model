using Excel2Model.Models;
using Excel2Model.Utilities;
using Excel2Model.Validation;
using Optional;
using System;
using System.Collections.Generic;
using System.Linq.Expressions;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;

namespace Excel2Model.Mappers
{
    public abstract class AbstractMapper<T>
    {
        private protected List<ColumnMapModel<T>> _columnMapModels = new List<ColumnMapModel<T>>();
        private protected Option<Excel.Worksheet, ValidationError> _excelInteropWorksheetOrValidationError;
        private protected WorksheetModel _worksheetModel;

        public AbstractMapper()
        {
        }

        public AbstractMapper(WorksheetModel worksheetModel)
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

        private protected ColumnMapModel<T> AddColumn(string columnName, PropertyInfo propertyInfo)
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

        private protected abstract List<T> GetDataFromExcelInterop(Excel.Worksheet excelInteropWorksheet);
    }
}
