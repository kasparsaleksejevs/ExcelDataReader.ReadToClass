using ExcelDataReader.ReadToClass.FluentMapping;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Shouldly;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

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
                    var columnCount = 0;
                    object[] columns = null;

                    table.Bind("SubData", m => m.SheetRows).WithColumns(column =>
                    {
                        column.Bind("Id", c => c.Id);
                        column.Bind("Name", c => c.Name);
                    }).OnHeaderRead((header) =>
                    {
                        columnCount = header.Length;
                        columns = header;
                    }).OnRowRead((myRow, data) =>
                    {
                        var staticColumnCount = 2;
                        myRow.YearlyData = new List<SubYearlyData>();
                        for (int i = staticColumnCount; i < columnCount; i++)
                        {
                            myRow.YearlyData.Add(new SubYearlyData
                            {
                                Year = Convert.ToInt32(columns[i]),
                                Value = Convert.ToInt32(data[i]),
                            });
                        }
                    });
                });

                var result = reader.AsClass<ComplexExcel>(config);
                result.SheetRows.Count.ShouldBe(3);
            }
        }

        [TestMethod]
        public void SheetWithSubRows_CanReadAllDataWithSubTables()
        {
            var source = TestSampleFiles.Sample_ComplexMultipleSheets;

            using (var ms = new MemoryStream(source))
            using (var reader = ExcelReaderFactory.CreateReader(ms))
            {
                var config = FluentConfig.ConfigureFor<ComplexExcel>().WithTables(table =>
                {
                    var columnCount = 0;
                    object[] columns = null;

                    table.Bind("SubData", m => m.SheetRows).WithColumns(column =>
                    {
                        column.Bind("Id", c => c.Id);
                        column.Bind("Name", c => c.Name);
                    }).OnHeaderRead((header) =>
                    {
                        columnCount = header.Length;
                        columns = header;
                    }).OnRowRead((myRow, data) =>
                    {
                        var staticColumnCount = 2;
                        myRow.YearlyData = new List<SubYearlyData>();
                        for (int i = staticColumnCount; i < columnCount; i++)
                        {
                            myRow.YearlyData.Add(new SubYearlyData
                            {
                                Year = Convert.ToInt32(columns[i]),
                                Value = Convert.ToInt32(data[i]),
                            });
                        }
                    });
                });

                var result = reader.AsClass<ComplexExcel>(config);

                result.SheetRows.Select(s => s.Id)
                    .ShouldBe(new List<int> { 1, 2, 3 });

                result.SheetRows.Select(s => s.Name)
                    .ShouldBe(new List<string> { "Aaaa", "Bbbb", "Cccc" });

                result.SheetRows.First().YearlyData.Select(s => s.Year)
                    .ShouldBe(new List<int> { 2017, 2018, 2019 });

                result.SheetRows.First().YearlyData.Select(s => s.Value)
                    .ShouldBe(new List<int> { 11, 22, 33 });
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
