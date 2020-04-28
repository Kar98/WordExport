using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using ALMOctaneExport;
using Microsoft.Office.Interop.Word;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using WordExport;
using WordExport.ALMTestExporter;
using WordExport.TestcaseObjects;

namespace WordExportUnittests
{
    [TestClass]
    public class Tests
    {
        const string big = @"TC_01 - Bulk XML Upload to AusTender and CNID Upload to migrate.docx";
        const string file2 = @"TC_01 - PRF - Status Management - to migrate.docx";
        const string test = @"test.docx";
        const string root = @"C:\Users\Kar98\Documents\Work\ALMOctane\";
        const string filepath = @"C:\Users\Kar98\Documents\Work\ALMOctane\"+ test;
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
        public void Description()
        {
            Console.WriteLine("File : " + filepath);
            var testcase = parser.Parse(filepath,1);
            Console.WriteLine(testcase.Description);
            ALMExporter export = new ALMExporter();
            File.WriteAllText("testcase.txt", testcase.Description);
            File.WriteAllText("testcaseformat.txt", export.ConvertDescriptionToString(testcase));
        }

        [TestMethod]
        public void TestDescriptionUpload()
        {
            Console.WriteLine("File : " + filepath);
            
            var testcase = parser.Parse(filepath, 1);
            Console.WriteLine(testcase.Description);
            File.WriteAllText("testcase.txt", testcase.Description);

            ALMExporter export = new ALMExporter();
            var configFile = File.ReadAllLines("Config\\config.txt");
            ALMOctaneConnection con = new ALMOctaneConnection(
                configFile[0], configFile[1], configFile[2], configFile[3], configFile[4]);

            ALMOctaneAPI api = new ALMOctaneAPI(con) { UserId = configFile[5] };
            var str = export.ConvertToString(testcase);
            var testid = api.CreateTest(testcase.Title, export.ConvertDescriptionToString(testcase));
            
            api.UpdateTest(testid, str);
        }

        [TestMethod]
        public void TestExpectedParsing()
        {
            var file = Path.GetFullPath(@"Files\test.docx");
            var testcase = parser.Parse(file, 1);
            Assert.AreEqual(testcase.TestSteps.Count, 4);
            Assert.AreEqual(testcase.TestSteps[0].Expected.Count, 3);
        }

        [TestMethod]
        public void OutputHeightCalculations()
        {
            Console.WriteLine("File : " + filepath);
            var testcase = parser.Parse(filepath, 1);
            int total = 0;
            foreach(var step in testcase.TestSteps)
            {
                foreach(var seq in step.Sequences)
                {
                    Console.WriteLine(seq.Key.ParagraphHeight);
                    total += seq.Key.ParagraphHeight;
                }
            }
            Console.WriteLine(total);
        }

        [TestMethod]
        public void TestStepToExpectedMapping()
        {
            var file = Path.GetFullPath(@"Files\test.docx");
            var testcase = parser.Parse(file, 1);
            Assert.IsTrue(testcase.TestSteps[0].Sequences.First(a => a.Key.Text.Contains("PRF with SON")).Value.Count > 0);
            Assert.IsTrue(testcase.TestSteps[0].Sequences.First(a => a.Key.Text.Contains("After indent")).Value.Count > 0);
            Assert.IsTrue(testcase.TestSteps[0].Sequences.First(a => a.Key.Text.Contains("Extra indent2")).Value.Count > 0);
            Assert.IsTrue(testcase.TestSteps[0].Sequences.First(a => a.Key.Text.Contains("Normal indent2")).Value.Count == 0);
        }

        [TestMethod]
        public void TestHeightGeneration()
        {
            var file = Path.GetFullPath(@"Files\test.docx");
            var testcase = parser.Parse(file, 1);
            
            var list = testcase.TestSteps[0].Sequences.ToList();
            var count = list.Count;
            Assert.IsTrue(list[count-2].Key.ParagraphHeight == 1); // Second last
            Assert.IsTrue(list[count - 1].Key.ParagraphHeight == 2); //  last
        }


    }
}
