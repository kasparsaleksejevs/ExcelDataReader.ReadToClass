﻿using ExcelDataReader.ReadToClass.Mapper;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Shouldly;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ExcelDataReader.ReadToClass.Tests
{
    [TestClass]
    public class OneSheetFileSkipColumnsTest
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
        public void ProcessOneSheetFile_OnlyDecimalsColumn()
        {
            var source = TestSampleFiles.Sample_OneSheet;

            using (var ms = new MemoryStream(source))
            using (var reader = ExcelReaderFactory.CreateReader(ms))
            {
                var result = reader.AsClass<OneSheetExcel>();

                var targetResult = new List<decimal> { 1.5m, 3m, 4.5m };
                result.FirstSheetRows.Select(s => s.DecimalColumn).ShouldBe(targetResult);
            }
        }

        public class OneSheetExcel
        {
            [ExcelTable("My Sheet 1")]
            public List<FirstSheet> FirstSheetRows { get; set; }
        }

        public class FirstSheet
        {
            //[ExcelColumn("Text Column", 1)]
            //public string TextColumn { get; set; }

            //[ExcelColumn("Some Int", 2)]
            //public int IntColumn { get; set; }

            [ExcelColumn("Decimals", 3)]
            public decimal DecimalColumn { get; set; }
        }
    }
}
