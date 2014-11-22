using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EdCanHack.SheetParser
{
    public interface INamedSheetReader
    {
        IEnumerable<String> SheetNames { get; }
        Sheet ReadSheet(String name);
    }
}
