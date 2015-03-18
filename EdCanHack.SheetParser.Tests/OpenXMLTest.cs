using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using EdCanHack.SheetParser.OpenXml;
using EdCanHack.SheetParser.SpreadsheetML;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace EdCanHack.SheetParser.Tests
{
    [TestClass]
    public class OpenXMLTest
    {
        private static OpenXMLReader _reader;

        [ClassInitialize]
        public static void TestInitialize(TestContext context)
        {
            var dataPath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "OpenXMLTestData.xlsx");

            _reader = new OpenXMLReader(dataPath, true);
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
            var list = _reader.ReadSheet("Sheet1").EnumerateRows().Select(r => r.ToArray()).ToList();

            Assert.AreEqual(3, list.Count, "Should have 3 items from Sheet1.");
            Assert.AreEqual("foo", list[0][0]);
            Assert.AreEqual("1", list[0][1]);
            Assert.AreEqual("10", list[0][2]);

            Assert.AreEqual("bar", list[1][0]);
            Assert.IsNull(list[1][1]);
            Assert.AreEqual("20", list[1][2]);

            Assert.AreEqual("baz", list[2][0]);
            Assert.AreEqual("3", list[2][1]);
            Assert.IsNull(list[2][2]);
        }
    }
}
