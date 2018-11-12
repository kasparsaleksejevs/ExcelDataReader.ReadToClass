using ExcelDataReader.ReadToClass.Mapper;
using System;
using System.Collections.Generic;
using System.Linq.Expressions;
using System.Reflection;

namespace ExcelDataReader.ReadToClass.FluentMapper
{
    public class FluentConfig
    {
        internal List<TablePropertyData> Tables { get; set; } = new List<TablePropertyData>();

        public ContainerBuilder<TModel> ConfigureFor<TModel>() where TModel : class
        {
            return new ContainerBuilder<TModel>(this);
        }
    }

    public class ContainerBuilder<TModel> where TModel : class
    {
        internal FluentConfig config;

        internal ContainerBuilder(FluentConfig config)
        {
            this.config = config;
        }

        public void Tables(Action<TableBuilder<TModel>> tableBuilder)
        {
            var builder = new TableBuilder<TModel>(this);
            tableBuilder.Invoke(builder);

            config.Tables = builder.tables;
        }
    }

    public class TableBuilder<TModel> where TModel : class
    {
        private readonly ContainerBuilder<TModel> containerBuilder;

        internal List<TablePropertyData> tables = new List<TablePropertyData>();

        internal TableBuilder(ContainerBuilder<TModel> containerBuilder)
        {
            this.containerBuilder = containerBuilder;
        }

        public TableColumnBuilder<TProperty> Bind<TProperty>(string tableName, Expression<Func<TModel, ICollection<TProperty>>> expression) where TProperty : class
        {
            var rowListProperty = (expression.Body as MemberExpression).Member as PropertyInfo;

            Type listElementType = typeof(string);
            var propertyType = rowListProperty.PropertyType;
            if (propertyType.IsGenericType && propertyType.GetGenericTypeDefinition() == typeof(List<>))
                listElementType = propertyType.GetGenericArguments()[0];

            var tableData = new TablePropertyData
            {
                ExcelTableName = tableName,
                ListElementType = listElementType,
                PropertyName = rowListProperty.Name,
            };

            this.tables.Add(tableData);

            var tableColumnBuilder = new TableColumnBuilder<TProperty>(tableData);
            return tableColumnBuilder;
        }
    }

    public class TableColumnBuilder<TModel> where TModel : class
    {
        internal List<ColumnPropertyData> columnProperties = new List<ColumnPropertyData>();

        private readonly TablePropertyData tableData;

        public TableColumnBuilder(TablePropertyData tableData)
        {
            this.tableData = tableData;
        }

        public void WithColumns(Action<ColumnBuilder<TModel>> builder)
        {
            var columnBuilder = new ColumnBuilder<TModel>(this);
            builder.Invoke(columnBuilder);

            tableData.Columns.AddRange(columnProperties);
        }
    }

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
