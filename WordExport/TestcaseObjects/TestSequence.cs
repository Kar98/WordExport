using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WordExport.COMObjects;

namespace WordExport.TestcaseObjects
{
    /// <summary>
    /// A test sequence is a single paragraph, test step. With either 0, 1, or many expected results. The description can be a bulleted point or a sub-bulleted point
    /// </summary>
    public class TestSequence
    {
        public WordParagraph Description { get; set; }
        public List<string> Expected { get; set; }

        public string Step { get; set; }

        public TestSequence(WordParagraph d)
        {
            Description = d;
        }



    }
}
