using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordExport
{
    public class WordParserException : Exception
    {
        public WordParserException() : base() { }

        public WordParserException(string s) : base(s) { }
    }
}
