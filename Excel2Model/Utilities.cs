﻿using System;
using System.Collections.Generic;
using System.Linq;
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

        public static List<T> GetPropertiesFromObjectBySpecificType<T>(object objectWithProperties) =>
            objectWithProperties.GetType().GetProperties()
                            .Where(propertyInfo => propertyInfo.PropertyType == typeof(T))
                            .Select(propertyInfo => (T)propertyInfo.GetValue(objectWithProperties))
                            .ToList();

        public static bool IsAnyValueFulfilledByType<T>(object objectWithProperties)
        {
            var properties = GetPropertiesFromObjectBySpecificType<T>(objectWithProperties);

            var output = typeof(T).IsValueType switch
            {
                true => properties.Any(value => EqualityComparer<T>.Default.Equals(value, default) == false),
                _ => properties.Any(value => value != null)
            };

            return output;
        }

        public static bool IsAnyValueFulfilled(object objectToBeChecked)
        { 
            // consider change to static array of types instead of below solution - reflection could be avoided

            foreach (TypeCode typeCode in Enum.GetValues(typeof(TypeCode)))
            {
                var type = Type.GetType($"System.{typeCode}");
                var typeOfContext = typeof(Utilities);
                var method = typeOfContext.GetMethod("IsAnyValueFulfilledByType");
                var genericMethod = method.MakeGenericMethod(type);
                object[] parameters = { objectToBeChecked };
                var result = (bool)genericMethod.Invoke(typeOfContext, parameters);

                if (result == true)
                {
                    return true;
                }
            }

            return false;
        }
    }
}
