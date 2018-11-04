using ExcelDataReader.ReadToClass.Mapper;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Shouldly;
using System.Collections.Generic;
using System.IO;

namespace ExcelDataReader.ReadToClass.Tests
{
    [TestClass]
    public class OneSheetFileTest
    {
        [TestMethod]
        public void ProcessOneSheetFile_Has3Rows()
        {
            var source = TestSampleFiles.Sample_OneSheet;

            using (var ms = new MemoryStream(source))
            using (var reader = ExcelReaderFactory.CreateReader(ms))
            {
                var result = reader.AsClass<OneSheetExcel>();
                result.FirstSheetRows.Count.ShouldBe(3);
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
        [ExcelColumn("Text Column", 1)]
        public string TextColumn { get; set; }

        [ExcelColumn("Some Int", 2)]
        public string IntColumn { get; set; }

        [ExcelColumn("Decimals", 3)]
        public string DecimalColumn { get; set; }
    }
}
