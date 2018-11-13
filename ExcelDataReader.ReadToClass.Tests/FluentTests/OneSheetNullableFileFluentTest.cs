using ExcelDataReader.ReadToClass.FluentMapper;
using ExcelDataReader.ReadToClass.Mapper;
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
                    table.Bind("My Sheet 1", m => m.FirstSheetRows).WithColumns(column =>
                    {
                        column.Bind("Text Column", c => c.TextColumn);
                        column.Bind("Some Int", c => c.IntColumn);
                        column.Bind("Decimals", c => c.DecimalColumn);
                        column.Bind("Nullable ints", c => c.NullableIntColumn);
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
                    table.Bind("My Sheet 1", m => m.FirstSheetRows).WithColumns(column =>
                    {
                        column.Bind("Text Column", c => c.TextColumn);
                        column.Bind("Some Int", c => c.IntColumn);
                        column.Bind("Decimals", c => c.DecimalColumn);
                        column.Bind("Nullable ints", c => c.NullableIntColumn);
                    });
                });

                var result = reader.AsClass<OneSheetExcel>(config);

                var targetResult = new List<int?> { 1, 2, null, 4, 5, null, 7 };
                result.FirstSheetRows.Select(s => s.NullableIntColumn).ShouldBe(targetResult);
            }
        }

        public class OneSheetExcel
        {
            [ExcelTable("My Sheet 1")]
            public List<FirstSheet> FirstSheetRows { get; set; }
        }

        public class FirstSheet
        {
            [ExcelColumn("Text Column", 1)]
            public string TextColumn { get; set; }

            [ExcelColumn("Some Int", 2)]
            public int IntColumn { get; set; }

            [ExcelColumn("Decimals", 3)]
            public decimal DecimalColumn { get; set; }

            [ExcelColumn("Nullable ints", 4)]
            public int? NullableIntColumn { get; set; }
        }
    }
}
