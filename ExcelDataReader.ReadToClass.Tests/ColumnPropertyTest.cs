using ExcelDataReader.ReadToClass.Mapper;
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
            var index = ColumnPropertyData.GetColumnIndexFromLetters("A");
            index.ShouldBe(1);
        }

        [TestMethod]
        public void IndexConverter_BDJ()
        {
            var index = ColumnPropertyData.GetColumnIndexFromLetters("BDJ");
            index.ShouldBe(1466);
        }

        [TestMethod]
        public void IndexConverter_IV()
        {
            var index = ColumnPropertyData.GetColumnIndexFromLetters("IV");
            index.ShouldBe(256);
        }

        [TestMethod]
        public void IndexConverter_PC()
        {
            var index = ColumnPropertyData.GetColumnIndexFromLetters("PC");
            index.ShouldBe(419);
        }

        [TestMethod]
        public void IndexConverter_XFD()
        {
            var index = ColumnPropertyData.GetColumnIndexFromLetters("XFD");
            index.ShouldBe(16384);
        }

        [TestMethod]
        public void IndexConverter_XFE()
        {
            var index = ColumnPropertyData.GetColumnIndexFromLetters("XFE");
            index.ShouldBe(16385);
        }
    }
}
