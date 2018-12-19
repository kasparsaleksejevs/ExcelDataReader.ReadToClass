using ExcelDataReader.ReadToClass.FluentMapping;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Shouldly;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ExcelDataReader.ReadToClass.Tests.FluentTests
{
    [TestClass]
    public class OneSheetFileFluentTest
    {
        [TestMethod]
        public void ProcessOneSheetFile_Has3Rows()
        {
            var source = TestSampleFiles.Sample_OneSheet;

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
                        column.Bind("Enums", c => c.EnumColumn);
                        column.Bind("Str Enums", c => c.StringEnumColumn);
                    });
                });

                var result = reader.AsClass<OneSheetExcel>(config);
                result.FirstSheetRows.Count.ShouldBe(3);
            }
        }

        [TestMethod]
        public void ProcessOneSheetFile_HasCorrectData()
        {
            var source = TestSampleFiles.Sample_OneSheet;

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
                    });
                });

                var result = reader.AsClass<OneSheetExcel>(config);

                var targetText = new List<string> { "Data 1", "Data 2", "Other Data" };
                result.FirstSheetRows.Select(s => s.TextColumn).ShouldBe(targetText);

                var targetInts = new List<int> { 1, 2, 3 };
                result.FirstSheetRows.Select(s => s.IntColumn).ShouldBe(targetInts);

                var targetDecimals = new List<decimal> { 1.5m, 3m, 4.5m };
                result.FirstSheetRows.Select(s => s.DecimalColumn).ShouldBe(targetDecimals);
            }
        }

        [TestMethod]
        public void ProcessOneSheetFile_HasCorrectEnums()
        {
            var source = TestSampleFiles.Sample_OneSheet;

            using (var ms = new MemoryStream(source))
            using (var reader = ExcelReaderFactory.CreateReader(ms))
            {
                var config = FluentConfig.ConfigureFor<OneSheetExcel>().WithTables(table =>
                {
                    table.Bind("My Sheet 1", m => m.FirstSheetRows).WithColumns(column =>
                    {
                        column.Bind("Enums", c => c.EnumColumn);
                        column.Bind("Str Enums", c => c.StringEnumColumn);
                    });
                });

                var result = reader.AsClass<OneSheetExcel>(config);

                var targetEnums = new List<MyEnum> { MyEnum.Value1, MyEnum.Value1, MyEnum.OtherEnum };
                result.FirstSheetRows.Select(s => s.EnumColumn).ShouldBe(targetEnums);

                var targetStringEnums = new List<MyEnum> { MyEnum.Value1, MyEnum.OtherEnum, MyEnum.Value1 };
                result.FirstSheetRows.Select(s => s.StringEnumColumn).ShouldBe(targetStringEnums);
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

            public MyEnum EnumColumn { get; set; }

            public MyEnum StringEnumColumn { get; set; }
        }

        public enum MyEnum
        {
            Value1 = 1,
            OtherEnum = 2,
        }
    }
}
