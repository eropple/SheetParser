using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;


namespace EdCanHack.SheetParser.OpenXml
{
    public class OpenXMLReader : INamedSheetReader
    {
        private readonly bool _hasHeaderRows;
        private XLWorkbook _workbook;

        public OpenXMLReader(String filename, bool hasHeaderRows)
            : this(File.OpenRead(filename), hasHeaderRows)
        {
        }

        public OpenXMLReader(Stream stream, bool hasHeaderRows)
        {
            _workbook = new XLWorkbook(stream);
            _hasHeaderRows = hasHeaderRows;
        }


        public IEnumerable<string> SheetNames => _workbook.Worksheets.Select(w => w.Name);

        public Sheet ReadSheet(string name)
        {
            return new OpenXMLSheet(_workbook.Worksheet(name), _hasHeaderRows);
        }
    }
}
