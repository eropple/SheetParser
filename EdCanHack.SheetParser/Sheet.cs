using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EdCanHack.SheetParser
{
    public abstract class Sheet
    {
        public readonly Boolean HasHeaderRow;

        protected Sheet(bool hasHeaderRow)
        {
            HasHeaderRow = hasHeaderRow;
        }

        /// <summary>
        /// Enumerates the rows of the requested sheet of your document.
        /// </summary>
        /// <returns>
        /// Returns a list of potentially nullable String objects that correspond to the
        /// columnar positions of each parsed cell. Cells should never contain empty or
        /// whitespace strings.
        /// </returns>
        public IEnumerable<IList<String>> EnumerateRows()
        {
            var implRows = EnumerateRowsImpl();
            return HasHeaderRow ? implRows.Skip(1) : implRows;
        }

        /// <summary>
        /// Gets the header row of a sheet with a header row. If no such header exists, returns null.
        /// </summary>
        public IList<String> HeaderRow
        {
            get { return HasHeaderRow ? EnumerateRowsImpl().First() : null; }
        }

        protected abstract IEnumerable<IList<String>> EnumerateRowsImpl();

        /// <summary>
        /// Transforms the contents of the sheet into a lookup table with a key based on
        /// the contents of the first non-empty cell.
        /// </summary>
        /// <returns>
        /// Returns the rows of your document, keyed to the contents of the first non-empty
        /// cell.
        /// </returns>
        public ILookup<String, IList<String>> GetKeyedRows()
        {
            return EnumerateRows().ToLookup(t => t.First(c => c != null));
        }
    }
}
