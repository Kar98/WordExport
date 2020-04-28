using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordExport.COMObjects
{
    public static class WordStringExtensions
    {
        /// <summary>
        /// Strips /a and /r
        /// </summary>
        /// <param name="s"></param>
        /// <returns></returns>
        public static string TrimJunk(this string s)
        {
            return s.Replace("\a", "").Replace("\r", "").Trim();
        }
    }
}
