using ExcelDataReader.ReadToClass.FluentMapping;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Shouldly;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ExcelDataReader.ReadToClass.Tests.FluentTests
{
    [TestClass]
    public class OneSheetFileNestedDataFluentTest
    {
        [TestMethod]
        public void ProcessOneSheetFileWithNestedClasses_HasCorrectData()
        {
            var source = TestSampleFiles.Sample_OneSheet_Nullable;

            using (var ms = new MemoryStream(source))
            using (var reader = ExcelReaderFactory.CreateReader(ms))
            {
                var config = FluentConfig.ConfigureFor<OneSheetExcel>().WithTables(table =>
                {
                    table.Bind("My Sheet 1", m => m.FirstSheetRows).WithColumns(column =>
                    {
                        column.Bind("Text Column", c => c.TextColumn);
                        column.Bind("Some Int", c => c.IntColumn);
                        column.Bind("Decimals", c => c.Nested.DecimalColumn);
                        column.Bind("Nullable ints", c => c.Nested.NullableIntColumn);
                        column.Bind("Nullable Enums", c => c.Nested.NullableEnumColumn);
                    });
                });

                var result = reader.AsClass<OneSheetExcel>(config);

                var targetText = new List<string> { "Data 1", "Data 2", "Other Data", "Lorem", "Ipsum", "Dolor", "Hodor" };
                result.FirstSheetRows.Select(s => s.TextColumn).ShouldBe(targetText);

                var targetInts = new List<int> { 1, 2, 3, 4, 5, 6, 7 };
                result.FirstSheetRows.Select(s => s.IntColumn).ShouldBe(targetInts);

                var targetDecimals = new List<decimal> { 1.5m, 3m, 4.5m, 6m, 7.5m, 9m, 10.5m };
                result.FirstSheetRows.Select(s => s.Nested.DecimalColumn).ShouldBe(targetDecimals);

                var targetResult = new List<int?> { 1, 2, null, 4, 5, null, 7 };
                result.FirstSheetRows.Select(s => s.Nested.NullableIntColumn).ShouldBe(targetResult);

                var targetEnumResult = new List<MyEnum?> { MyEnum.Val1, MyEnum.Val2, MyEnum.Val2, null, MyEnum.Val1, MyEnum.Val1, MyEnum.Val1 };
                result.FirstSheetRows.Select(s => s.Nested.NullableEnumColumn).ShouldBe(targetEnumResult);
            }
        }

        public class OneSheetExcel
        {
            public List<FirstSheet> FirstSheetRows { get; set; }
        }

        public class FirstSheet
        {
            public string TextColumn { get; set; }

            public int IntColumn { get; set; }

            public NestedData Nested { get; set; } = new NestedData(); // it is important to auto-initialize nested classes!
        }

        public class NestedData
        {
            public decimal DecimalColumn { get; set; }

            public int? NullableIntColumn { get; set; }

            public MyEnum? NullableEnumColumn { get; set; }
        }

        public enum MyEnum
        {
            Val1 = 1,
            Val2 = 2,
        }
    }
}
