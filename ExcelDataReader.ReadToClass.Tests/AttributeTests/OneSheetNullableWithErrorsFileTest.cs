using ExcelDataReader.ReadToClass.AttributeMapping;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Shouldly;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ExcelDataReader.ReadToClass.Tests.AttributeTests
{
    [TestClass]
    public class OneSheetNullableWithErrorsFileTest
    {
        [TestMethod]
        public void ProcessOneSheetNullableFile_Has8Rows()
        {
            var source = TestSampleFiles.Sample_OneSheet_Nullable_WithErrors;

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
            var source = TestSampleFiles.Sample_OneSheet_Nullable_WithErrors;

            using (var ms = new MemoryStream(source))
            using (var reader = ExcelReaderFactory.CreateReader(ms))
            {
                var result = reader.AsClass<OneSheetExcel>(out List<string> errors);

                var targetResult = new List<int?> { 1, null, 3, 4, 5, 6, 7 };
                result.FirstSheetRows.Select(s => s.NullableIntWithErrColumn).ShouldBe(targetResult);

                errors.Count.ShouldBe(1);
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

            [ExcelColumn("Nullable ints with err")]
            public int? NullableIntWithErrColumn { get; set; }
        }
    }
}
