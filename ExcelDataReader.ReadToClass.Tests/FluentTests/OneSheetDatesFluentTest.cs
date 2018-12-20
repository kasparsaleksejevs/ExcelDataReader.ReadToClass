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
    public class OneSheetDatesFluentTest
    {
        [TestMethod]
        public void ProcessOneSheetFile_CanReadAllRows()
        {
            var source = TestSampleFiles.Sample_OneSheet_Dates;

            using (var ms = new MemoryStream(source))
            using (var reader = ExcelReaderFactory.CreateReader(ms))
            {
                var config = FluentConfig.ConfigureFor<OneSheetExcel>().WithTables(table =>
                {
                    table.Bind("My Sheet 1", m => m.FirstSheetRows).WithColumns(column =>
                    {
                        column.Bind("Date Column", c => c.NormalDate);
                        column.Bind("Nullable Date", c => c.NullableDate);
                        column.Bind("Formatted Date", c => c.FormattedDate);
                        column.Bind("Formatted Nullable Date", c => c.FormattedDate);
                        column.Bind("13.08.2018", c => c.StringValuesWithDateHeader);
                    });
                });

                var result = reader.AsClass<OneSheetExcel>(config);

                result.FirstSheetRows.Count.ShouldBe(3);
            }
        }

        [TestMethod]
        public void ProcessOneSheetFile_CanReadDates()
        {
            var source = TestSampleFiles.Sample_OneSheet_Dates;

            using (var ms = new MemoryStream(source))
            using (var reader = ExcelReaderFactory.CreateReader(ms))
            {
                var config = FluentConfig.ConfigureFor<OneSheetExcel>().WithTables(table =>
                {
                    table.Bind("My Sheet 1", m => m.FirstSheetRows).WithColumns(column =>
                    {
                        column.Bind("Date Column", c => c.NormalDate);
                    });
                });

                var result = reader.AsClass<OneSheetExcel>(config);

                var targetNormalDates = new List<DateTime> { new DateTime(2018, 01, 13), new DateTime(2018, 02, 13), new DateTime(2018, 03, 13) };
                result.FirstSheetRows.Select(s => s.NormalDate).ShouldBe(targetNormalDates);
            }
        }

        [TestMethod]
        public void ProcessOneSheetFile_CanReadNullableDates()
        {
            var source = TestSampleFiles.Sample_OneSheet_Dates;

            using (var ms = new MemoryStream(source))
            using (var reader = ExcelReaderFactory.CreateReader(ms))
            {
                var config = FluentConfig.ConfigureFor<OneSheetExcel>().WithTables(table =>
                {
                    table.Bind("My Sheet 1", m => m.FirstSheetRows).WithColumns(column =>
                    {
                        column.Bind("Nullable Date", c => c.NullableDate);
                    });
                });

                var result = reader.AsClass<OneSheetExcel>(config);

                var targetNullableDates = new List<DateTime?> { new DateTime(2018, 08, 13), null, new DateTime(2018, 10, 13) };
                result.FirstSheetRows.Select(s => s.NullableDate).ShouldBe(targetNullableDates);
            }
        }

        [TestMethod]
        public void ProcessOneSheetFile_CanReadFormattedDates()
        {
            var source = TestSampleFiles.Sample_OneSheet_Dates;

            using (var ms = new MemoryStream(source))
            using (var reader = ExcelReaderFactory.CreateReader(ms))
            {
                var config = FluentConfig.ConfigureFor<OneSheetExcel>().WithTables(table =>
                {
                    table.Bind("My Sheet 1", m => m.FirstSheetRows).WithColumns(column =>
                    {
                        column.Bind("Formatted Date", c => c.FormattedDate);
                    });
                });

                var result = reader.AsClass<OneSheetExcel>(config);

                var targetFormattedDates = new List<DateTime> { new DateTime(2018, 01, 13), new DateTime(2018, 02, 13), new DateTime(2018, 03, 13) };
                result.FirstSheetRows.Select(s => s.FormattedDate).ShouldBe(targetFormattedDates);
            }
        }

        [TestMethod]
        public void ProcessOneSheetFile_CanReadFormattedNullableDates()
        {
            var source = TestSampleFiles.Sample_OneSheet_Dates;

            using (var ms = new MemoryStream(source))
            using (var reader = ExcelReaderFactory.CreateReader(ms))
            {
                var config = FluentConfig.ConfigureFor<OneSheetExcel>().WithTables(table =>
                {
                    table.Bind("My Sheet 1", m => m.FirstSheetRows).WithColumns(column =>
                    {
                        column.Bind("Formatted Nullable Date", c => c.FormattedNullableDate);
                    });
                });

                var result = reader.AsClass<OneSheetExcel>(config);

                var targetFormattedDates = new List<DateTime?> { new DateTime(2018, 01, 13), null, new DateTime(2018, 03, 13) };
                result.FirstSheetRows.Select(s => s.FormattedNullableDate).ShouldBe(targetFormattedDates);
            }
        }

        [TestMethod]
        public void ProcessOneSheetFile_CanReadFromDateHeaderColumn()
        {
            var source = TestSampleFiles.Sample_OneSheet_Dates;

            using (var ms = new MemoryStream(source))
            using (var reader = ExcelReaderFactory.CreateReader(ms))
            {
                var config = FluentConfig.ConfigureFor<OneSheetExcel>().WithTables(table =>
                {
                    table.Bind("My Sheet 1", m => m.FirstSheetRows).WithColumns(column =>
                    {
                        column.Bind(new DateTime(2018, 08, 13).ToString(), c => c.StringValuesWithDateHeader);
                    });
                });

                var result = reader.AsClass<OneSheetExcel>(config);

                var targetDates = new List<string> { "Date header", "Some value", "Othe value" };
                result.FirstSheetRows.Select(s => s.StringValuesWithDateHeader).ShouldBe(targetDates);
            }
        }

        public class OneSheetExcel
        {
            public List<FirstSheet> FirstSheetRows { get; set; }
        }

        public class FirstSheet
        {
            public DateTime NormalDate { get; set; }

            public DateTime? NullableDate { get; set; }

            public DateTime FormattedDate { get; set; }

            public DateTime? FormattedNullableDate { get; set; }

            public string StringValuesWithDateHeader { get; set; }
        }
    }
}
