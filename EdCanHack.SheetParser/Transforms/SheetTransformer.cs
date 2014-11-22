using System;
using System.Collections.Generic;
using System.Linq;

namespace EdCanHack.SheetParser.Transforms
{
    public abstract class SheetTransformer<TTransformType>
    {
        public virtual IEnumerable<TTransformType> TransformRows(Sheet sheet)
        {
            var sheetRows = sheet.EnumerateRows();
            var transformRows = sheetRows.Select(FromRow);
            return transformRows;
        }

        protected abstract TTransformType FromRow(IList<String> row); 
    }

    public abstract class KeyedSheetTransformer<TKeyType, TTransformType>
            : SheetTransformer<TTransformType>
        where TTransformType : IKeyed<TKeyType>
    {
        public ILookup<TKeyType, TTransformType> TransformRowsToLookup(Sheet sheet)
        {
            return TransformRows(sheet).ToLookup(t => t.Key);
        }
    }
}
