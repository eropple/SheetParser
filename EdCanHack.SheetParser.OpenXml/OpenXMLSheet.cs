using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace EdCanHack.SheetParser.OpenXml
{
    public class OpenXMLSheet : Sheet
    {
        private readonly IXLWorksheet _worksheet;

        public OpenXMLSheet(IXLWorksheet worksheet, bool hasHeaderRow) : base(hasHeaderRow)
        {
            _worksheet = worksheet;
        }

        protected override IEnumerable<IList<string>> EnumerateRowsImpl()
        {
            var rows = _worksheet.Rows().ToList();
            var cellCount = rows.Max(r => r.Cells().Max(c => c.Address.ColumnNumber));

            foreach (var row in _worksheet.Rows())
            {
                var cells = row.CellsUsed().ToList();
                if (cells.Count == 0) continue;

                var content = new String[cellCount];

                foreach (var cell in cells)
                {
                    content[cell.Address.ColumnNumber - 1] = cell.Value.ToString();
                }

                yield return content;
            }
        }
    }
}
