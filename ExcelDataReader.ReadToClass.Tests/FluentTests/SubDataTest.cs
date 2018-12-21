using ExcelDataReader.ReadToClass.FluentMapping;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Shouldly;
using System.Collections.Generic;
using System.IO;

namespace ExcelDataReader.ReadToClass.Tests.FluentTests
{
    [TestClass]
    public class SubDataTest
    {
        [TestMethod]
        public void SheetWithSubRows_CanReadAllRows()
        {
            var source = TestSampleFiles.Sample_ComplexMultipleSheets;

            using (var ms = new MemoryStream(source))
            using (var reader = ExcelReaderFactory.CreateReader(ms))
            {
                var config = FluentConfig.ConfigureFor<ComplexExcel>().WithTables(table =>
                {
                    table.Bind("SubData", m => m.SheetRows).WithColumns(column =>
                    {
                        column.Bind("Id", c => c.Id);
                        column.Bind("Name", c => c.Name);
                        column.Bind()
                    });
                });

                var result = reader.AsClass<ComplexExcel>(config);
                result.SheetRows.Count.ShouldBe(3);
            }
        }

        public class ComplexExcel
        {
            public List<SubDataSheet> SheetRows { get; set; }
        }

        public class SubDataSheet
        {
            public int Id { get; set; }

            public string Name { get; set; }

            public List<SubYearlyData> YearlyData { get; set; }
        }

        public class SubYearlyData
        {
            public int Year { get; set; }

            public int Value { get; set; }
        }
    }
}
