using System;
using System.Collections.Generic;
using System.Linq.Expressions;
using System.Reflection;

namespace ExcelDataReader.ReadToClass.FluentMapping
{
    public class TableBuilder<TModel> where TModel : class
    {
        private readonly ContainerBuilder<TModel> containerBuilder;

        internal List<TablePropertyData> tables = new List<TablePropertyData>();

        internal TableBuilder(ContainerBuilder<TModel> containerBuilder)
        {
            this.containerBuilder = containerBuilder;
        }

        public TableColumnBuilder<TProperty> Bind<TProperty>(string excelSheetName, Expression<Func<TModel, ICollection<TProperty>>> expression) where TProperty : class
        {
            var rowListProperty = (expression.Body as MemberExpression).Member as PropertyInfo;

            Type listElementType = typeof(string);
            var propertyType = rowListProperty.PropertyType;
            if (propertyType.IsGenericType && propertyType.GetGenericTypeDefinition() == typeof(List<>))
                listElementType = propertyType.GetGenericArguments()[0];

            var tableData = new TablePropertyData
            {
                ExcelSheetName = excelSheetName,
                ListElementType = listElementType,
                PropertyName = rowListProperty.Name,
            };

            this.tables.Add(tableData);

            var tableColumnBuilder = new TableColumnBuilder<TProperty>(tableData);
            return tableColumnBuilder;
        }
    }
}
