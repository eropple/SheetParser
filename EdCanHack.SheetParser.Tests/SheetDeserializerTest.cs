using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using EdCanHack.SheetParser.SpreadsheetML;
using EdCanHack.SheetParser.Transforms;
using EdCanHack.SheetParser.Transforms.Serialization;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace EdCanHack.SheetParser.Tests
{
    [TestClass]
    public class SheetDeserializerTest
    {
        private static SpreadsheetMLReader _reader;

        [ClassInitialize]
        public static void TestInitialize(TestContext context)
        {
            var dataPath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "SpreadsheetMLTestData.xml");

            _reader = new SpreadsheetMLReader(dataPath, true);
        }



        [TestMethod]
        public void Deserialize()
        {
            var deserializer = new SheetDeserializer<Sheet1Record>();

            var enumerable = deserializer.TransformRows(_reader.ReadSheet("Sheet1"));
            var list = enumerable.ToList();

            Assert.AreEqual(5, list.Count);

            Assert.AreEqual("foo", list[0].Key);
            Assert.AreEqual(2, list[1].Int32Column);
            Assert.AreEqual("boo", list[0].StringColumn);
            Assert.IsNull(list[0].NopeColumn);
            Assert.AreEqual("e", list[4].JumpedColumn);

            Assert.AreEqual(8282.82m, list[4].DecimalColumn);
        }


        [SheetDeserializable(ignoreExtraFields: true, ignoreMissingFields: false)]
        public class Sheet1Record : IKeyed<String>
        {
            public String Key { get; private set; }

            [SheetProperty("Int32Column")]
            public Int32 Int32Column { get; private set; }
            [SheetProperty("UInt32Column")]
            public UInt32 UInt32Column { get; private set; }

            [SheetProperty("Int64Column")]
            public Int64 Int64Column { get; private set; }
            [SheetProperty("UInt64Column")]
            public UInt64 UInt64Column { get; private set; }

            [SheetProperty("Int16Column")]
            public Int16 Int16Column { get; private set; }
            [SheetProperty("UInt16Column")]
            public UInt16 UInt16Column { get; private set; }

            [SheetProperty("StringColumn")]
            public String StringColumn { get; private set; }
            [SheetProperty("NopeColumn")]
            public String NopeColumn { get; private set; }
            [SheetProperty("JumpedColumn")]
            public String JumpedColumn { get; private set; }
            [SheetProperty("DecimalColumn")]
            public Decimal DecimalColumn { get; private set; }

        }
    }
}
