using System;
using System.Collections.Generic;

namespace ExcelDataReader.ReadToClass
{
    public class TablePropertyData
    {
        public string PropertyName { get; set; }

        public Type ListElementType { get; set; }

        public string ExcelSheetName { get; set; }

        public List<ColumnPropertyData> Columns { get; set; } = new List<ColumnPropertyData>();

        public Action<object[]> OnHeaderRead;

        public Action<object, object[]> OnRowRead;

        public string StartingCellAddress { get; set; }
    }
}
