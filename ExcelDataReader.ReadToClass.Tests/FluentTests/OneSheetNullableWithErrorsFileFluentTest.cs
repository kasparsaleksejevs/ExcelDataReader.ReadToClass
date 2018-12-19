using ExcelDataReader.ReadToClass.FluentMapping;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Shouldly;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ExcelDataReader.ReadToClass.Tests.FluentTests
{
    [TestClass]
    public class OneSheetNullableWithErrorsFileFluentTest
    {
        [TestMethod]
        public void ProcessOneSheetNullableFile_Has8Rows()
        {
            var source = TestSampleFiles.Sample_OneSheet_Nullable_WithErrors;

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
                        column.Bind("Nullable ints with err", c => c.NullableIntWithErrColumn);
                    });
                });

                var result = reader.AsClass<OneSheetExcel>(config);
                result.FirstSheetRows.Count.ShouldBe(7);
            }
        }

        [TestMethod]
        public void ProcessOneSheetNullableFile_HasCorrectNullableSequence()
        {
            var source = TestSampleFiles.Sample_OneSheet_Nullable_WithErrors;

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
                        column.Bind("Nullable ints with err", c => c.NullableIntWithErrColumn);
                    });
                });

                var result = reader.AsClass<OneSheetExcel>(out List<string> errors, config);

                var targetResult = new List<int?> { 1, null, 3, 4, 5, 6, 7 };
                result.FirstSheetRows.Select(s => s.NullableIntWithErrColumn).ShouldBe(targetResult);

                errors.Count.ShouldBe(1);
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

            public int? NullableIntWithErrColumn { get; set; }
        }
    }
}
