using Excel2Model.Models;
using Excel2Model.Validation;
using Microsoft.Office.Interop.Excel;
using Optional;
using System;
using System.Collections.Generic;
using System.Linq.Expressions;

namespace Excel2Model.Mappers
{
    public class TableHeaderMapper<T> : AbstractMapper<T> where T : new()
    {
        public Option<ColumnMapModel<T>, ValidationError> TryAddColumn(string columnHeader, Expression<Func<T, object>> tProperty)
        {
            throw new NotImplementedException();
        }

        private protected override List<T> GetDataFromExcelInterop(Worksheet excelInteropWorksheet)
        {
            throw new NotImplementedException();
        }
    }
}
