using ExcelDataReader.ReadToClass.FluentMapper;
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


        public class OneSheetExcel
        {
            //[ExcelTable("My Sheet 1")]
            public List<FirstSheet> FirstSheetRows { get; set; }
        }

        public class FirstSheet
        {
            //[ExcelColumn("Text Column", 1)]
            public string TextColumn { get; set; }

            //[ExcelColumn("Some Int", 2)]
            public int IntColumn { get; set; }

            //[ExcelColumn("Decimals", 3)]
            public decimal DecimalColumn { get; set; }
        }
    }
}
