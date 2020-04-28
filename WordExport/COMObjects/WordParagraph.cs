using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Remoting.Messaging;
using System.Text;
using System.Threading.Tasks;

namespace WordExport.COMObjects
{
    public class WordParagraph
    {
        public string Text { get; set; }
        public WdListType ListType { get; set; }
        public int ListLevel { get; set; } // All text starts at 1 even if there is no List applied.
        public List<int> RowNumber { get; set; }

        public int ParagraphHeight { get; set; } // Calcualted height of the paragraph

        public int ColumnWidth { get; set; }

        public WordParagraph(Range range, int columnWidth)
        {
            ColumnWidth = columnWidth;
            var tmpText = "";
            if(range.ListFormat.ListType == WdListType.wdListBullet && range.ListFormat.ListLevelNumber == 2)
            {
                tmpText = "\t*" + range.Text.TrimJunk();
            }
            else if (range.ListFormat.ListType == WdListType.wdListBullet && range.ListFormat.ListLevelNumber == 3)
            {
                tmpText = "\t\t**" + range.Text.TrimJunk();
            }
            else
            {
                tmpText = range.Text.TrimJunk();
            }
            Text = tmpText;
            ListType = range.ListFormat.ListType;
            ListLevel = range.ListFormat.ListLevelNumber;
            ParagraphHeight = CalculateHeight();
        }

        private int CalculateHeight()
        {
            // Magic number based on entering in random characters in word inside the column, getting the column width, then dividing col width by 
            // number of chars.  

            double len = Text.Length;
            double basefactor = ColumnWidth / 4.6f; // Magic number
            double factor;

            //basefactor = 46;

            if(len == 0)
            {
                return 1;
            }
            if (ListType == WdListType.wdListNoNumbering)
            {
                factor = basefactor; //46
            }else if (ListType == WdListType.wdListBullet && ListLevel == 1)
            {
                factor = basefactor - (6 * ListLevel); //39
            }else if (ListType == WdListType.wdListBullet && ListLevel == 2)
            {
                factor = basefactor - (6* ListLevel); //33
            }
            else if (ListType == WdListType.wdListBullet && ListLevel == 3)
            {
                factor = basefactor - (6 * ListLevel); //27
            }
            else
            {
                throw new WordParserException("Unhandled wdListBullet exception");
            }
            double calc = len / factor;
            return (int)Math.Ceiling(calc);
        }


    }
}
