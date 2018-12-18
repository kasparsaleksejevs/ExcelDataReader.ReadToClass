using System;
using System.Collections.Generic;

namespace ExcelDataReader.ReadToClass.FluentMapping
{
    public class TableColumnBuilder<TModel> where TModel : class
    {
        internal List<ColumnPropertyData> columnProperties = new List<ColumnPropertyData>();

        private readonly TablePropertyData tableData;

        public TableColumnBuilder(TablePropertyData tableData)
        {
            this.tableData = tableData;
        }

        public void WithColumnsByHeader(Action<ColumnBuilderByHeader<TModel>> builder)
        {
            var columnBuilder = new ColumnBuilderByHeader<TModel>(this);
            builder.Invoke(columnBuilder);

            tableData.Columns.AddRange(columnProperties);
        }


        public void WithColumnsByIndex(Action<ColumnBuilderByIndex<TModel>> builder)
        {
            var columnBuilder = new ColumnBuilderByIndex<TModel>(this);
            builder.Invoke(columnBuilder);

            tableData.Columns.AddRange(columnProperties);
        }
    }
}
