using System;
using System.Linq.Expressions;
using System.Reflection;

namespace Excel2Model
{
    public static class Utilities
    {
        public static PropertyInfo GetPropertyFromExpression<T>(Expression<Func<T, object>> GetPropertyLambda)
        {
            // Inspired by:
            // https://stackoverflow.com/questions/17115634/get-propertyinfo-of-a-parameter-passed-as-lambda-expression
            // Author: Daniel Möller

            MemberExpression output;

            //this line is necessary, because sometimes the expression comes in as Convert(originalexpression)
            if (GetPropertyLambda.Body is UnaryExpression unaryExpression)
            {
                if (unaryExpression.Operand is MemberExpression memberExpression)
                {
                    output = memberExpression;
                }
                else
                {
                    throw new ArgumentException("Incorrect argument. Provided unary expression is not member expression.");
                }
            }
            else if (GetPropertyLambda.Body is MemberExpression memberExpression)
            {
                output = memberExpression;
            }
            else
            {
                throw new ArgumentException("Incorrect argument. Provided property lambda is not member expression.");
            }

            return (PropertyInfo)output.Member;
        }
    }
}
