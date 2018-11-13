using System;

namespace ExcelDataReader.ReadToClass.FluentMapper
{
    public class ContainerBuilder<TModel> where TModel : class
    {
        internal FluentConfig config;

        internal ContainerBuilder()
        {
            this.config = new FluentConfig();
        }

        public FluentConfig WithTables(Action<TableBuilder<TModel>> tableBuilder)
        {
            var builder = new TableBuilder<TModel>(this);
            tableBuilder.Invoke(builder);

            config.Tables = builder.tables;

            return config;
        }
    }
}
