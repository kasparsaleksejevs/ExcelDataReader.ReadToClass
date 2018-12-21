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

        public TableColumnBuilder<TModel> WithColumns(Action<ColumnBuilder<TModel>> builder)
        {
            var columnBuilder = new ColumnBuilder<TModel>(this);
            builder.Invoke(columnBuilder);

            tableData.Columns.AddRange(columnProperties);

            return this;
        }

        public TableColumnBuilder<TModel> OnHeaderRead(Action<object[]> action)
        {
            tableData.onHeaderRead = action;
            return this;
        }

        public TableColumnBuilder<TModel> OnRowRead(Action<TModel, object[]> action)
        {
            tableData.onRowRead = (model, rowData) => action((TModel)model, rowData);
            return this;
        }
    }
}
