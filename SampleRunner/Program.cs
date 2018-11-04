using ExcelDataReader;
using ExcelDataReader.ReadToClass;
using ExcelDataReader.ReadToClass.Mapper;
using System;
using System.Collections.Generic;
using System.IO;

namespace SampleRunner
{
    class Program
    {
        static void Main(string[] args)
        {
            var source = Resource1.Sample_OneSheet;

            using (var ms = new MemoryStream(source))
            using (var reader = ExcelReaderFactory.CreateReader(ms))
            {
                var result = reader.AsClass<OneSheetExcel>();
                Console.WriteLine("Rows: " + result.FirstSheetRows.Count.ToString());
            }

            Console.WriteLine("== Done! ==");
            Console.ReadKey();
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
    }
}
