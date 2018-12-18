using System;
using System.Linq.Expressions;
using System.Reflection;

namespace ExcelDataReader.ReadToClass.FluentMapping
{
    public class ColumnBuilderByHeader<TModel> where TModel : class
    {
        private readonly TableColumnBuilder<TModel> tableColumnBuilder;

        internal ColumnBuilderByHeader(TableColumnBuilder<TModel> tableColumnBuilder)
        {
            this.tableColumnBuilder = tableColumnBuilder;
        }

        public void Bind<TProperty>(string columnNameInExcel, Expression<Func<TModel, TProperty>> property)
        {
            var columnProperty = (property.Body as MemberExpression).Member as PropertyInfo;
            var propertyData = new ColumnPropertyData
            {
                ExcelColumnName = columnNameInExcel,
                PropertyName = columnProperty.Name,
            };

            this.tableColumnBuilder.columnProperties.Add(propertyData);
        }
    }

    public class ColumnBuilderByIndex<TModel> where TModel : class
    {
        private readonly TableColumnBuilder<TModel> tableColumnBuilder;

        internal ColumnBuilderByIndex(TableColumnBuilder<TModel> tableColumnBuilder)
        {
            this.tableColumnBuilder = tableColumnBuilder;
        }

        public void Bind<TProperty>(string columnIndexLettersInExcel, Expression<Func<TModel, TProperty>> property)
        {
            var columnProperty = (property.Body as MemberExpression).Member as PropertyInfo;
            var propertyData = new ColumnPropertyData
            {
                ExcelColumnName = columnIndexLettersInExcel,
                ExcelColumnIsIndex = true,
                PropertyName = columnProperty.Name,
                Order = 1,
            };

            this.tableColumnBuilder.columnProperties.Add(propertyData);
        }
    }
}
