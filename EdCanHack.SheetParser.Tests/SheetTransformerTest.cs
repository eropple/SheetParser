using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using EdCanHack.SheetParser.SpreadsheetML;
using EdCanHack.SheetParser.Transforms;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace EdCanHack.SheetParser.Tests
{
    [TestClass]
    public class SheetTransformerTest
    {
        private static SpreadsheetMLReader _reader;

        [ClassInitialize]
        public static void TestInitialize(TestContext context)
        {
            var dataPath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "SpreadsheetMLTestData.xml");

            _reader = new SpreadsheetMLReader(dataPath, true);
        }

        [TestMethod]
        public void TransformRow()
        {
            var transformer = new TestSheetTransformer();
            var enumerable = transformer.TransformRows(_reader.ReadSheet("Sheet1"));
            var list = enumerable.ToList();

            Assert.AreEqual(5, list.Count, "Should be 5 entries.");
            var recordA = list[4];
            Assert.AreEqual("boop", recordA.Key);
            var recordB = list[3];
            Assert.AreEqual("4", recordB.ColumnB);
            var recordC = list[2];
            Assert.AreEqual("6", recordC.ColumnC);
        }


        public class Sheet1Record : IKeyed<String>
        {
            public Sheet1Record(string key, string columnB, string columnC)
            {
                ColumnC = columnC;
                ColumnB = columnB;
                Key = key;
            }

            public string Key { get; private set; }

            public String ColumnB { get; private set; }
            public String ColumnC { get; private set; }
        }

        public class TestSheetTransformer : KeyedSheetTransformer<String, Sheet1Record>
        {
            protected override Sheet1Record FromRow(IList<String> row)
            {
                return new Sheet1Record(row[0], row[1], row[2]);
            }
        }
    }
}
