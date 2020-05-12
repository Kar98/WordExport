using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using ALMOctaneExport;
using Microsoft.Office.Interop.Word;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using TestRailExport;
using WordExport;
using WordExport.ALMTestExporter;
using WordExport.COMObjects;
using WordExport.TestcaseObjects;
using WordExport.TestrailConverter;

namespace WordExportUnittests
{
    [TestClass]
    public class Tests
    {
        const string root = @"C:\Wordfiles\";

        const string big = @"TC_01 - Bulk XML Upload to AusTender and CNID Upload to migrate.docx";
        const string file2 = @"TC_01 - PRF - Status Management - to migrate.docx";
        const string test = @"test.docx";
        const string file = @"TC_01 - WIBS - ZRFX ZRSP ZCTR ZNPO.docx";
        const string filepath = root + file;
        static Application ap;
        static WordParser parser;

        [TestInitialize]
        public void Init()
        {
            ap = new Application();
            parser = new WordParser(ap);
        }

        [TestCleanup]
        public void Teardown()
        {
            try
            {
                parser.CloseCurrentDoc();
            }
            catch (Exception)
            {
                //silently fail
            }
            parser.Dispose();
        }

        [TestMethod]
        [TestCategory("Debug")]
        public void Description()
        {
            Console.WriteLine("File : " + filepath);
            var testcase = parser.Parse(filepath,1);
            Console.WriteLine(testcase.Description);
            ALMConvertor export = new ALMConvertor();
            File.WriteAllText("testcase.txt", testcase.Description);
            File.WriteAllText("testcaseformat.txt", export.ConvertDescriptionToString(testcase));
        }

        [TestMethod]
        [TestCategory("Test")]
        public void TestALMOctaneUpload()
        {
            Console.WriteLine("File : " + filepath);
            
            var testcase = parser.Parse(filepath, 1);
            Console.WriteLine(testcase.Description);
            File.WriteAllText("testcase.txt", testcase.Description);

            ALMConvertor export = new ALMConvertor();
            var configFile = File.ReadAllLines("Config\\config.txt");
            ALMOctaneConnection con = new ALMOctaneConnection(
                configFile[0], configFile[1], configFile[2], configFile[3], configFile[4]);

            ALMOctaneAPI api = new ALMOctaneAPI(con) { UserId = configFile[5] };
            var str = export.ConvertToString(testcase);
            var testid = api.CreateTest(testcase.Title, export.ConvertDescriptionToString(testcase));
            
            api.UpdateTest(testid, str);
        }

        [TestMethod]
        [TestCategory("Test")]
        public void TestExpectedParsing()
        {
            var file = Path.GetFullPath(@"Files\test.docx");
            Console.WriteLine(file);
            var testcase = parser.Parse(file, 1);
            Assert.AreEqual(testcase.TestSteps.Count, 4);
            Assert.AreEqual(testcase.TestSteps[0].Expected.Count, 5);
        }

        [TestMethod]
        [TestCategory("Debug")]
        public void OutputColumnWidth()
        {
            var doc = ap.Documents.Open(filepath);
            var stepWidth = (int)Math.Ceiling(doc.Tables[1].Cell(1, 2).Width);
            Console.WriteLine(stepWidth);
        }

        [TestMethod]
        [TestCategory("Debug")]
        public void OutputHeightCalculations()
        {
            Console.WriteLine("File : " + filepath);
            var testcase = parser.Parse(filepath, 1);
            int total = 0;
            foreach(var step in testcase.TestSteps)
            {
                foreach(var seq in step.Sequences)
                {
                    Console.WriteLine(seq.Key.Text + " - " + seq.Key.ParagraphHeight);
                    total += seq.Key.ParagraphHeight;
                }
            }
            Console.WriteLine(total);
        }

        [TestMethod]
        [TestCategory("Debug")]
        public void OutputParagraphStats()
        {
            var doc = ap.Documents.Open(filepath);
            var tables = doc.Tables.Cast<Table>().ToList();
            var allParagraphs = doc.Paragraphs.Cast<Paragraph>().ToList();
            foreach(var p in tables[0].Range.Paragraphs.Cast<Paragraph>().ToList())
            {
                Console.WriteLine(p.Range.Text.TrimJunk() + " - " + p.Range.ComputeStatistics(WdStatistic.wdStatisticLines));
            }
        }

        [TestMethod]
        [TestCategory("Test")]
        public void TestStepToExpectedMapping()
        {
            var file = Path.GetFullPath(@"Files\test.docx");
            Console.WriteLine(file);
            var testcase = parser.Parse(file, 1);
            Assert.IsTrue(testcase.TestSteps[0].Sequences.First(a => a.Key.Text.Contains("PRF with SON")).Value.Count > 0);
            Assert.IsTrue(testcase.TestSteps[0].Sequences.First(a => a.Key.Text.Contains("After indent")).Value.Count > 0);
            Assert.IsTrue(testcase.TestSteps[0].Sequences.First(a => a.Key.Text.Contains("Extra indent2")).Value.Count > 0);
            Assert.IsTrue(testcase.TestSteps[0].Sequences.First(a => a.Key.Text.Contains("Normal indent2")).Value.Count == 0);
            Assert.IsTrue(testcase.TestSteps[0].Sequences.First(a => a.Key.Text.Contains("Back to normal indent2")).Value.Count == 0);
        }

        [TestMethod]
        [TestCategory("Test")]
        public void TestHeightGeneration()
        {
            var file = Path.GetFullPath(@"Files\test.docx");
            var testcase = parser.Parse(file, 1);
            
            var list = testcase.TestSteps[0].Sequences.ToList();
            var count = list.Count;
            Assert.IsTrue(list[count-2].Key.ParagraphHeight == 1); // Second last
            Assert.IsTrue(list[count - 1].Key.ParagraphHeight == 2); //  last
        }

        [TestMethod]
        [TestCategory("Test")]
        public void TestTestRailConverter()
        {
            var file = Path.GetFullPath(@"Files\test.docx");
            var testcase = parser.Parse(file, 1);
            var json = TRConverter.Convert(testcase);
            TRExporter tre = new TRExporter("https://environment.testrail.net", "rory.crickmore@kjr.com.au", "TESTsatellite11!");
            tre.CreateTest("2429", json);
        }


    }
}
