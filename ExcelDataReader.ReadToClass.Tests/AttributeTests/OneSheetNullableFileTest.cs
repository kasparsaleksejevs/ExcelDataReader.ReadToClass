using ExcelDataReader.ReadToClass.AttributeMapping;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Shouldly;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ExcelDataReader.ReadToClass.Tests.AttributeTests
{
    [TestClass]
    public class OneSheetNullableFileTest
    {
        [TestMethod]
        public void ProcessOneSheetNullableFile_Has8Rows()
        {
            var source = TestSampleFiles.Sample_OneSheet_Nullable;

            using (var ms = new MemoryStream(source))
            using (var reader = ExcelReaderFactory.CreateReader(ms))
            {
                var result = reader.AsClass<OneSheetExcel>();
                result.FirstSheetRows.Count.ShouldBe(7);
            }
        }

        [TestMethod]
        public void ProcessOneSheetNullableFile_HasCorrectNullableSequence()
        {
            var source = TestSampleFiles.Sample_OneSheet_Nullable;

            using (var ms = new MemoryStream(source))
            using (var reader = ExcelReaderFactory.CreateReader(ms))
            {
                var result = reader.AsClass<OneSheetExcel>();

                var targetResult = new List<int?> { 1, 2, null, 4, 5, null, 7 };
                result.FirstSheetRows.Select(s => s.NullableIntColumn).ShouldBe(targetResult);

                var targetEnumResult = new List<MyEnum?> { MyEnum.Val1, MyEnum.Val2, MyEnum.Val2, null, MyEnum.Val1, MyEnum.Val1, MyEnum.Val1 };
                result.FirstSheetRows.Select(s => s.NullableEnumColumn).ShouldBe(targetEnumResult);
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

            [ExcelColumn("Nullable ints")]
            public int? NullableIntColumn { get; set; }

            [ExcelColumn("Nullable Enums")]
            public MyEnum? NullableEnumColumn { get; set; }
        }

        public enum MyEnum
        {
            Val1 = 1,
            Val2 = 2,
        }
    }
}
