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

        /// <summary>
        /// Binds the specified column name (table header) in excel to the class property.
        /// </summary>
        /// <typeparam name="TProperty">The type of the property.</typeparam>
        /// <param name="columnNameInExcel">The column name (table header) in excel.</param>
        /// <param name="expression">The property expression.</param>
        /// <param name="isMandatory">If set to <c>true</c> - this property is mandatory and the reading will stop on the row where the cell specified by this property is empty, even if other cells still have values.</param>
        public void Bind<TProperty>(string columnNameInExcel, Expression<Func<TModel, TProperty>> expression, bool isMandatory = false)
        {
            var columnProperty = (expression.Body as MemberExpression).Member as PropertyInfo;
            var propertyData = new ColumnPropertyData
            {
                ExcelColumnName = columnNameInExcel,
                PropertyName = columnProperty.Name,
                IsMandatory = isMandatory
            };

            this.tableColumnBuilder.columnProperties.Add(propertyData);
        }
    }
}
