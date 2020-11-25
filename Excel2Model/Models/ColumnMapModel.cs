﻿using System.Reflection;

namespace Excel2Model.Models
{
    public record ColumnMapModel<T>
    {
        public string ColumnName { get; init; }
        public PropertyInfo Property { get; init; }
    }
}