using ExcelDataReader;
using ExcelDataReader.ReadToClass;
using ExcelDataReader.ReadToClass.AttributeMapping;
using ExcelDataReader.ReadToClass.FluentMapping;
using System;
using System.Collections.Generic;
using System.IO;

namespace SampleRunner
{
    class Program
    {
        static void Main(string[] args)
        {
            //TestingWithAttributes();
            TestingFluent();

            Console.WriteLine("== Done! ==");
            Console.ReadKey();
        }

        static void TestingWithAttributes()
        {
            var source = Resource1.Sample_OneSheet;

            using (var ms = new MemoryStream(source))
            using (var reader = ExcelReaderFactory.CreateReader(ms))
            {
                var result = reader.AsClass<OneSheetExcel>();
                Console.WriteLine("Rows: " + result.FirstSheetRows.Count.ToString());
            }
        }

        static void TestingFluent()
        {
            var source = Resource1.Sample_OneSheet;

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
                    });
                });

                var result = reader.AsClass<OneSheetExcel>(config);
                Console.WriteLine("Rows: " + result.FirstSheetRows.Count.ToString());
            }
        }
    }

    public class OneSheetExcel
    {
        [ExcelTable("My Sheet 1")]
        public List<FirstSheet> FirstSheetRows { get; set; }
    }

    public class FirstSheet
    {
        [ExcelColumn("Text Column")]
        public string TextColumn { get; set; }

        [ExcelColumn("Some Int")]
        public int IntColumn { get; set; }

        [ExcelColumn("Decimals")]
        public decimal DecimalColumn { get; set; }
    }
}
