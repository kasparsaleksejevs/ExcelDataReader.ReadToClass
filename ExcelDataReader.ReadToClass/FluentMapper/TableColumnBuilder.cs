using ExcelDataReader.ReadToClass.Mapper;
using System;
using System.Collections.Generic;

namespace ExcelDataReader.ReadToClass.FluentMapper
{
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
}
