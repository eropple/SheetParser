using System;
using System.IO;
using System.Linq;
using System.Reflection;
using EdCanHack.SheetParser.SpreadsheetML;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace EdCanHack.SheetParser.Tests
{
    [TestClass]
    public class SpreadsheetMLTest
    {
        private static SpreadsheetMLReader _reader;

        [ClassInitialize]
        public static void TestInitialize(TestContext context)
        {
            var dataPath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "SpreadsheetMLTestData.xml");

            _reader = new SpreadsheetMLReader(dataPath, true);
        }


        [TestMethod]
        public void CorrectSheetNames()
        {
            var sheetNames = _reader.SheetNames.ToList();

            Assert.AreEqual(2, sheetNames.Count, "Should have two sheets.");
            Assert.AreEqual("Sheet1", sheetNames[0]);
            Assert.AreEqual("Sheet2", sheetNames[1]);
        }

        [TestMethod]
        public void CanIterateSheet()
        {
            var list = _reader.ReadSheet("Sheet1").EnumerateRows().ToList();

            Assert.AreEqual(5, list.Count, "Should have 5 items from Sheet1.");
        }
    }
}
