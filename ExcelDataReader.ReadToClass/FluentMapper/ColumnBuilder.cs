using ExcelDataReader.ReadToClass.Mapper;
using System;
using System.Linq.Expressions;
using System.Reflection;

namespace ExcelDataReader.ReadToClass.FluentMapper
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
                Order = 1,
            };

            this.tableColumnBuilder.columnProperties.Add(propertyData);
        }
    }
}
