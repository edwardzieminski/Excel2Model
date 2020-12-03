using Excel = Microsoft.Office.Interop.Excel;
using Excel2Model.Models;
using Excel2Model.Utilities;
using Excel2Model.Validation;
using Optional;
using System.Collections.Generic;
using System.Reflection;

namespace Excel2Model.Mappers
{
    public abstract class AbstractMapper<T> where T : new()
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

        public abstract void ResolveMap();

        private protected T ExcelInteropRowToModelRecord(Excel.Worksheet excelInteropWorksheet, int currentRow)
        {
            var modelRecord = new T();

            foreach (var columnMap in _columnMapModels)
            {
                var cellAddress = $"{columnMap.ColumnName}{currentRow}";
                var cellValue = excelInteropWorksheet.Range[cellAddress].Value;
                columnMap.Property.SetValue(modelRecord, cellValue);
            }

            return modelRecord;
        }

        private protected List<T> GetDataFromExcelInterop(Excel.Worksheet excelInteropWorksheet)
        {
            var output = new List<T>();

            var currentRow = WorksheetModel.DataStartRow;
            bool anyValueFulfilled;

            do
            {
                var modelRecord = ExcelInteropRowToModelRecord(excelInteropWorksheet, currentRow);

                anyValueFulfilled = CommonUtilities.IsAnyValueFulfilled(modelRecord);

                if (anyValueFulfilled)
                {
                    output.Add(modelRecord);
                }

                currentRow++;
            } while (anyValueFulfilled);

            return output;
        }

        private protected abstract ColumnMapModel<T> AddColumn(string columnName, PropertyInfo propertyInfo);

        private void SetExcelInteropWorksheet()
        {
            _excelInteropWorksheetOrValidationError = ExcelInteropUtilities.TryGetWorksheet(WorksheetModel);
        }
    }
}
