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
    public class ColumnNameMapper<T> : AbstractMapper<T> where T : new()
    {
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

        /// <summary>
        /// This method should be used after all columns are added. It is recommended to wrap this method in try-catch block 
        /// as it is going to throw CouldNotResolveMapException.
        /// </summary>
        public override void ResolveMap()
        {
            
        }

        private protected override ColumnMapModel<T> AddColumn(string columnName, PropertyInfo propertyInfo)
        {
            var output = new ColumnMapModel<T>()
            {
                ColumnName = columnName,
                Property = propertyInfo
            };

            _columnMapModels.Add(output);

            return output;
        }
    }
}
