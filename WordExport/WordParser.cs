using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using WordExport.TestcaseObjects;

namespace WordExport
{
    public class WordParser : IDisposable
    {
        Application ap;
        Document doc;
        public WordParser(Application ap) { this.ap = ap; }

        public void Dispose()
        {
            ap.Quit();
        }

        public TestCase Parse(string filepath, int testcaseId)
        {
            /*
             * Load doc
             * Get the table
             * Get the row and get the description and expected result
             * Generate height per WordParagraph for Description
             * Attempt to join the Description and Expected
             */

            var doc = ap.Documents.Open(filepath);
            this.doc = doc;
            Thread.Sleep(500); // Word needs to load the document and there is no way of doing a waitfor.

            var testcase = new TestCase
            {
                Id = testcaseId
            };
            Globals.Log("Total word paragraphs found : " + doc.Paragraphs.Count);

            var tables = doc.Tables.Cast<Table>().ToList();
            var allParagraphs = doc.Paragraphs.Cast<Paragraph>().ToList();
            testcase.AddTitle(allParagraphs);
            testcase.AddDescription(allParagraphs);
            testcase.AddTestSteps(tables);
            doc.Close();

            return testcase;
        }

        public TestCase ParseSpecificTable(string filepath, int tableIndex)
        {
            var doc = ap.Documents.Open(filepath);
            this.doc = doc;
            Console.WriteLine(doc.Tables.Count);
            Thread.Sleep(500); // Word needs to load the document and there is no way of doing a waitfor.
            Console.WriteLine(doc.Tables.Count);

            var rawList = doc.Tables.Cast<Table>().ToList();
            var passIn = new List<Table>() { rawList[tableIndex] };

            var testcase = new TestCase();
            testcase.AddTestSteps(passIn);

            return testcase;
        }

        public void CloseCurrentDoc()
        {
            doc.Close();
        }


    }
}
