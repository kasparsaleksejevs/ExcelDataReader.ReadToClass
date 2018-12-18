using ExcelDataReader.ReadToClass.FluentMapping;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Shouldly;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ExcelDataReader.ReadToClass.Tests.FluentTests
{
    [TestClass]
    public class OneSheetNullableFileFluentTest
    {
        [TestMethod]
        public void ProcessOneSheetNullableFile_Has8Rows()
        {
            var source = TestSampleFiles.Sample_OneSheet_Nullable;

            using (var ms = new MemoryStream(source))
            using (var reader = ExcelReaderFactory.CreateReader(ms))
            {
                var config = FluentConfig.ConfigureFor<OneSheetExcel>().WithTables(table =>
                {
                    table.Bind("My Sheet 1", m => m.FirstSheetRows).WithColumnsByHeader(column =>
                    {
                        column.Bind("Text Column", c => c.TextColumn);
                        column.Bind("Some Int", c => c.IntColumn);
                        column.Bind("Decimals", c => c.DecimalColumn);
                        column.Bind("Nullable ints", c => c.NullableIntColumn);
                        column.Bind("Nullable Enums", c => c.NullableEnumColumn);
                    });
                });

                var result = reader.AsClass<OneSheetExcel>(config);
                result.FirstSheetRows.Count.ShouldBe(7);
            }
        }

        [TestMethod]
        public void ProcessOneSheetNullableFile_HasCorrectNullableSequence()
        {
            var source = TestSampleFiles.Sample_OneSheet_Nullable;

            using (var ms = new MemoryStream(source))
            using (var reader = ExcelReaderFactory.CreateReader(ms))
            {
                var config = FluentConfig.ConfigureFor<OneSheetExcel>().WithTables(table =>
                {
                    table.Bind("My Sheet 1", m => m.FirstSheetRows).WithColumnsByHeader(column =>
                    {
                        column.Bind("Text Column", c => c.TextColumn);
                        column.Bind("Some Int", c => c.IntColumn);
                        column.Bind("Decimals", c => c.DecimalColumn);
                        column.Bind("Nullable ints", c => c.NullableIntColumn);
                        column.Bind("Nullable Enums", c => c.NullableEnumColumn);
                    });
                });

                var result = reader.AsClass<OneSheetExcel>(config);

                var targetResult = new List<int?> { 1, 2, null, 4, 5, null, 7 };
                result.FirstSheetRows.Select(s => s.NullableIntColumn).ShouldBe(targetResult);

                var targetEnumResult = new List<MyEnum?> { MyEnum.Val1, MyEnum.Val2, MyEnum.Val2, null, MyEnum.Val1, MyEnum.Val1, MyEnum.Val1 };
                result.FirstSheetRows.Select(s => s.NullableEnumColumn).ShouldBe(targetEnumResult);
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
