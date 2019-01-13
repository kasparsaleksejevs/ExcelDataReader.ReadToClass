using ExcelDataReader.ReadToClass.FluentMapping;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Shouldly;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ExcelDataReader.ReadToClass.Tests.FluentTests
{
    [TestClass]
    public class MultipleTablesPerSheetFluentTest
    {
        [TestMethod]
        public void ProcessMultipleTablesPerSheet_CanReadAllRows()
        {
            var source = TestSampleFiles.Sample_ComplexMultipleSheets;

            using (var ms = new MemoryStream(source))
            using (var reader = ExcelReaderFactory.CreateReader(ms))
            {
                var config = FluentConfig.ConfigureFor<MultipleTablesExcel>().WithTables(table =>
                {
                    table.Bind("MultipleTables", m => m.FirstTableRows).StartingFromCell("B4").WithColumns(column =>
                    {
                        column.Bind("Name", c => c.Name, isMandatory: true);
                        column.Bind("Value", c => c.Value);
                    });
                    table.Bind("MultipleTables", m => m.SecondTableRows).StartingFromCell("F6").WithColumns(column =>
                    {
                        column.Bind("Title", c => c.Title, isMandatory: true);
                        column.Bind("Amount", c => c.Amount);
                        column.Bind("Count", c => c.Count);
                    });
                });

                var result = reader.AsClass<MultipleTablesExcel>(config);

                result.FirstTableRows.Count.ShouldBe(3);
                result.SecondTableRows.Count.ShouldBe(4);
            }
        }

        [TestMethod]
        public void ProcessMultipleTablesPerSheet_HasCorrectData()
        {
            var source = TestSampleFiles.Sample_ComplexMultipleSheets;

            using (var ms = new MemoryStream(source))
            using (var reader = ExcelReaderFactory.CreateReader(ms))
            {
                var config = FluentConfig.ConfigureFor<MultipleTablesExcel>().WithTables(table =>
                {
                    table.Bind("MultipleTables", m => m.FirstTableRows).StartingFromCell("B4").WithColumns(column =>
                    {
                        column.Bind("Name", c => c.Name, isMandatory: true);
                        column.Bind("Value", c => c.Value);
                    });
                    table.Bind("MultipleTables", m => m.SecondTableRows).StartingFromCell("F6").WithColumns(column =>
                    {
                        column.Bind("Title", c => c.Title, isMandatory: true);
                        column.Bind("Amount", c => c.Amount);
                        column.Bind("Count", c => c.Count);
                    });
                });

                var result = reader.AsClass<MultipleTablesExcel>(config);

                var targetFirstTableNames = new List<string> { "Eee", "Fff", "Ggg" };
                result.FirstTableRows.Select(s => s.Name).ShouldBe(targetFirstTableNames);

                var targetFirstTableValues = new List<int> { 13, 26, 75 };
                result.FirstTableRows.Select(s => s.Value).ShouldBe(targetFirstTableValues);

                var targetSecondTableTitles = new List<string> { "Aaa", "Bbb", "Ccc", "Ddd" };
                result.SecondTableRows.Select(s => s.Title).ShouldBe(targetSecondTableTitles);

                var targetSecondTableAmounts = new List<decimal> { 10.1m, 22.22m, 0.1m, 23m };
                result.SecondTableRows.Select(s => s.Amount).ShouldBe(targetSecondTableAmounts);

                var targetSecondTableCounts = new List<int> { 4, 5, 6, 7 };
                result.SecondTableRows.Select(s => s.Count).ShouldBe(targetSecondTableCounts);
            }
        }
    }

    public class MultipleTablesExcel
    {
        public List<FirstTable> FirstTableRows { get; set; }

        public List<SecondTable> SecondTableRows { get; set; }
    }

    public class FirstTable
    {
        public string Name { get; set; }

        public int Value { get; set; }
    }

    public class SecondTable
    {
        public string Title { get; set; }

        public decimal Amount { get; set; }

        public int Count { get; set; }
    }
}
