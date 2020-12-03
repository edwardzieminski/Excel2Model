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
        public Option<ColumnMapModel, ValidationError> TryAddColumn(string columnName, Expression<Func<T, object>> tProperty)
        {
            Option<ColumnMapModel, ValidationError> output = new Option<ColumnMapModel, ValidationError>();

            var propertyOrValidationError = CommonUtilities.TryGetPropertyFromExpression(tProperty);

            propertyOrValidationError.Match
            (
                some: propertyInfo => output = Option.Some<ColumnMapModel, ValidationError>(AddColumn(columnName, propertyInfo)),
                none: validationError => output = Option.None<ColumnMapModel, ValidationError>(validationError)
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

        private protected override ColumnMapModel AddColumn(string columnName, PropertyInfo propertyInfo)
        {
            var output = new ColumnMapModel()
            {
                ColumnName = columnName,
                Property = propertyInfo
            };

            _columnMapModels.Add(output);

            return output;
        }
    }
}
