using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EdCanHack.SheetParser.Transforms.Serialization
{
    public class SheetDeserializationException : SheetParserException
    {
        public SheetDeserializationException()
        {
        }

        public SheetDeserializationException(string message) : base(message)
        {
        }

        public SheetDeserializationException(string message, Exception innerException) : base(message, innerException)
        {
        }

        public SheetDeserializationException(string message, params object[] args) : base(message, args)
        {
        }

        public SheetDeserializationException(string message, Exception innerException, params object[] args) : base(message, innerException, args)
        {
        }
    }
}
