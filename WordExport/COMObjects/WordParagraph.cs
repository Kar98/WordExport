﻿using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Remoting.Messaging;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;

namespace WordExport.COMObjects
{
    public class WordParagraph
    {
        public string Text { get; private set; }
        public WdListType ListType { get; private set; }
        public float FontSize { get; private set; }
        public int ListLevel { get; private set; } // All text starts at 1 even if there is no List applied.
        public List<int> RowNumber { get; set; }

        public int ParagraphHeight { get; private set; } // Calculated height of the paragraph

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
            FontSize = range.Font.Size;
            Text = tmpText;
            ListType = range.ListFormat.ListType;
            ListLevel = range.ListFormat.ListLevelNumber;
            ParagraphHeight = range.ComputeStatistics(WdStatistic.wdStatisticLines);
        }

    }
}
