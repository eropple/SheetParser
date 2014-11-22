using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EdCanHack.SheetParser
{
    public class SheetParserException : Exception
    {
        public SheetParserException() { }
        public SheetParserException(string message) : base(message) { }
        public SheetParserException(string message, Exception innerException) : base(message, innerException) { }
        public SheetParserException(string message, params Object[] args) : base(String.Format(message, args), null) { }
        public SheetParserException(string message, Exception innerException, params Object[] args) : base(String.Format(message, args), innerException) { }
    }
}
