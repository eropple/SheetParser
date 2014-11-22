using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using System.Xml.Serialization;

namespace EdCanHack.SheetParser.SpreadsheetML
{
    public class SpreadsheetMLReader : INamedSheetReader
    {
        public static readonly XNamespace SSNamespace = "urn:schemas-microsoft-com:office:spreadsheet";
        private static readonly XName WorksheetName = SSNamespace + "Worksheet";
        private static readonly XName NameAttr = SSNamespace + "Name";

        private readonly Dictionary<String, SpreadsheetMLSheet> _sheetNodes;

        public readonly Boolean HasHeaderRow;

        public SpreadsheetMLReader(XDocument doc, bool hasHeaderRow)
        {
            HasHeaderRow = hasHeaderRow;
            _sheetNodes = ParseDocument(doc, hasHeaderRow);
        }

        public SpreadsheetMLReader(String uri, bool hasHeaderRow) : this(XDocument.Load(uri), hasHeaderRow) { }

        public IEnumerable<String> SheetNames { get { return _sheetNodes.Keys; } }

        public Sheet ReadSheet(String name)
        {
            SpreadsheetMLSheet sheet;
            if (_sheetNodes.TryGetValue(name, out sheet))
            {
                return sheet;
            }

            throw new SheetParserException("Could not find sheet '{0}'.", name);
        }


        private static Dictionary<String, SpreadsheetMLSheet> ParseDocument(XDocument doc, Boolean hasHeaderRow)
        {
            var root = doc.Root;

            return root.Elements(WorksheetName)
                .ToDictionary(ws => ws.Attribute(NameAttr).Value, ws => new SpreadsheetMLSheet(ws, hasHeaderRow));
        }
    }
}
