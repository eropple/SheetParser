using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace EdCanHack.SheetParser.SpreadsheetML
{
    public class SpreadsheetMLSheet : Sheet
    {
        private readonly XElement _sheetRoot;
        public readonly Boolean HasHeaderRow;

        internal SpreadsheetMLSheet(XElement sheetRoot, bool hasHeaderRow) : base(hasHeaderRow)
        {
            _sheetRoot = sheetRoot;
        }

        protected override IEnumerable<IList<String>> EnumerateRowsImpl()
        {
            var table = _sheetRoot.Element(SpreadsheetMLReader.SSNamespace + "Table");


            foreach (var row in table.Elements(SpreadsheetMLReader.SSNamespace + "Row"))
            {
                var cellIndex = 0;
                var cells = new List<String>();

                foreach (var cell in row.Elements(SpreadsheetMLReader.SSNamespace + "Cell"))
                {
                    var indexAttr = cell.Attribute(SpreadsheetMLReader.SSNamespace + "Index");
                    if (indexAttr != null) cellIndex = Int32.Parse(indexAttr.Value) - 1;

                    if (cellIndex != cells.Count)
                    {
                        var difference = cellIndex - cells.Count;
                        for (var i = 0; i < difference; ++i) cells.Add(null);
                    }

                    var data = cell.Element(SpreadsheetMLReader.SSNamespace + "Data");
                    if (data == null)
                    {
                        cells.Add(null);
                    }
                    else
                    {
                        cells.Add(!String.IsNullOrWhiteSpace(data.Value) ? data.Value : null);
                    }

                    cellIndex += 1;
                }


                if (cells.All(c => c == null)) continue;
                    
                yield return cells;
            }
        }
    }
}
