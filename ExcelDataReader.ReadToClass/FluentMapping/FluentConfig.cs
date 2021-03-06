﻿using System.Collections.Generic;

namespace ExcelDataReader.ReadToClass.FluentMapping
{
    public class FluentConfig
    {
        internal List<TablePropertyData> Tables { get; set; } = new List<TablePropertyData>();

        public static ContainerBuilder<TModel> ConfigureFor<TModel>() where TModel : class
        {
            return new ContainerBuilder<TModel>();
        }
    }
}
