using System;
using System.Linq.Expressions;
using System.Reflection;

namespace ExcelDataReader.ReadToClass.FluentMapping
{
    public class ColumnBuilder<TModel> where TModel : class
    {
        private readonly TableColumnBuilder<TModel> tableColumnBuilder;

        internal ColumnBuilder(TableColumnBuilder<TModel> tableColumnBuilder)
        {
            this.tableColumnBuilder = tableColumnBuilder;
        }

        public void Bind<TProperty>(string columnNameInExcel, Expression<Func<TModel, TProperty>> expression)
        {
            var columnProperty = (expression.Body as MemberExpression).Member as PropertyInfo;
            var propertyData = new ColumnPropertyData
            {
                ExcelColumnName = columnNameInExcel,
                PropertyName = columnProperty.Name,
            };

            this.tableColumnBuilder.columnProperties.Add(propertyData);
        }

        public void BindSubTable<TProperty>(string columnNameInExcel, Expression<Func<TModel, TProperty>> expression)
        {
            var columnProperty = (expression.Body as MemberExpression).Member as PropertyInfo;
            var propertyData = new ColumnPropertyData
            {
                ExcelColumnName = columnNameInExcel,
                PropertyName = columnProperty.Name,
            };

            this.tableColumnBuilder.columnProperties.Add(propertyData);
        }
    }
}
