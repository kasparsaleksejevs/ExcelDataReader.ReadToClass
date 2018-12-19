using ExcelDataReader.ReadToClass.AttributeMapping;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Shouldly;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ExcelDataReader.ReadToClass.Tests.AttributeTests
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

        [TestMethod]
        public void ProcessOneSheetFile_HasCorrectData()
        {
            var source = TestSampleFiles.Sample_OneSheet;

            using (var ms = new MemoryStream(source))
            using (var reader = ExcelReaderFactory.CreateReader(ms))
            {
                var result = reader.AsClass<OneSheetExcel>();

                var targetText = new List<string> { "Data 1", "Data 2", "Other Data" };
                result.FirstSheetRows.Select(s => s.TextColumn).ShouldBe(targetText);

                var targetInts = new List<int> { 1, 2, 3 };
                result.FirstSheetRows.Select(s => s.IntColumn).ShouldBe(targetInts);

                var targetDecimals = new List<decimal> { 1.5m, 3m, 4.5m };
                result.FirstSheetRows.Select(s => s.DecimalColumn).ShouldBe(targetDecimals);

                var targetEnums = new List<MyEnum> { MyEnum.Val1, MyEnum.Val1, MyEnum.Val2 };
                result.FirstSheetRows.Select(s => s.EnumColumn).ShouldBe(targetEnums);
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

            [ExcelColumn("Enums")]
            public MyEnum EnumColumn { get; set; }
        }

        public enum MyEnum
        {
            Val1 = 1,
            Val2 = 2,
        }
    }
}
