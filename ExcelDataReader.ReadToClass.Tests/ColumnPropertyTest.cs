using Microsoft.VisualStudio.TestTools.UnitTesting;
using Shouldly;

namespace ExcelDataReader.ReadToClass.Tests
{
    [TestClass]
    public class ColumnPropertyTest
    {
        [TestMethod]
        public void IndexConverter_A()
        {
            var index = ColumnPropertyData.GetColumnIndexFromCellAddress("A");
            index.ShouldBe(1);
        }

        [TestMethod]
        public void IndexConverter_BDJ()
        {
            var index = ColumnPropertyData.GetColumnIndexFromCellAddress("BDJ");
            index.ShouldBe(1466);
        }

        [TestMethod]
        public void IndexConverter_IV()
        {
            var index = ColumnPropertyData.GetColumnIndexFromCellAddress("IV");
            index.ShouldBe(256);
        }

        [TestMethod]
        public void IndexConverter_IVWithRow()
        {
            var index = ColumnPropertyData.GetColumnIndexFromCellAddress("IV22334");
            index.ShouldBe(256);
        }

        [TestMethod]
        public void IndexConverterRow_IVWithRow()
        {
            var index = ColumnPropertyData.GetRowIndexFromCellAddress("IV22334");
            index.ShouldBe(22334);
        }

        [TestMethod]
        public void IndexConverter_PC()
        {
            var index = ColumnPropertyData.GetColumnIndexFromCellAddress("PC");
            index.ShouldBe(419);
        }

        [TestMethod]
        public void IndexConverter_XFD()
        {
            var index = ColumnPropertyData.GetColumnIndexFromCellAddress("XFD");
            index.ShouldBe(16384);
        }

        [TestMethod]
        public void IndexConverter_XFE()
        {
            var index = ColumnPropertyData.GetColumnIndexFromCellAddress("XFE");
            index.ShouldBe(16385);
        }

        [TestMethod]
        public void IndexConverter_XFEWithRow()
        {
            var index = ColumnPropertyData.GetColumnIndexFromCellAddress("XFE234");
            index.ShouldBe(16385);
        }

        [TestMethod]
        public void IndexConverterRow_XFEWithRow()
        {
            var index = ColumnPropertyData.GetRowIndexFromCellAddress("XFE234");
            index.ShouldBe(234);
        }
    }
}
