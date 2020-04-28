using System;
using System.Diagnostics;
using System.IO;
using ALMOctaneExport;
using Microsoft.Office.Interop.Word;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using WordExport;
using WordExport.ALMTestExporter;

namespace WordExportUnittests
{
    [TestClass]
    public class Tests
    {
        const string big = @"TC_01 - Bulk XML Upload to AusTender and CNID Upload to migrate.docx";
        const string file2 = @"TC_01 - PRF - Status Management - to migrate.docx";
        const string test = @"test.docx";
        const string root = @"C:\Users\Kar98\Documents\Work\ALMOctane\";
        const string filepath = @"C:\Users\Kar98\Documents\Work\ALMOctane\"+ file2;
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
        public void Expected()
        {
            Console.WriteLine("File : " + filepath);
            var testcase = parser.Parse(filepath,1);
            Assert.AreEqual(testcase.TestSteps.Count, 2);
            Assert.AreEqual(testcase.TestSteps[0].Expected.Count, 1);
        }

        [TestMethod]
        public void TestHeightCalculations()
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

        



    }
}
