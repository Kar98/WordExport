using ALMOctaneExport;
using MicroFocus.Adm.Octane.Api.Core.Connector;
using MicroFocus.Adm.Octane.Api.Core.Connector.Authentication;
using MicroFocus.Adm.Octane.Api.Core.Entities;
using MicroFocus.Adm.Octane.Api.Core.Services;
using MicroFocus.Adm.Octane.Api.Core.Services.RequestContext;
using Microsoft.Office.Interop.Word;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using TestRailExport;
using WordExport.ALMTestExporter;
using WordExport.ExcelExporter;
using WordExport.TestcaseObjects;
using static WordExport.Globals;

namespace WordExport
{
    class Program
    {
        static void Main(string[] args)
        {
            File.WriteAllText(Globals.LogPath, "");
            File.WriteAllText(Globals.ErrorPath, "");
            Globals.Log("Start");

            MainProgram();
            //Test();

            Globals.Log("Done!");
            Console.WriteLine("Press enter to exit");
            Console.Read();
        }

        private static void Test()
        {
            TRExporter tre = new TRExporter("https://environment.testrail.net", "***", "***");
            var jsonobj = new TestrailJSON() { Title = "jsnobj title 2", TemplateId = 2, TypeId = 6, PriorityId = 1, 
                CustomTestscenario = "description new as obj", CustomStepsSeparated = new CustomStepsSeparated[] 
                { new CustomStepsSeparated() { Content = "step1", Expected = "exp1" },
                new CustomStepsSeparated() { Content = "step2", Expected = "exp2" }} };

            var json = JsonConvert.SerializeObject(jsonobj);
            File.WriteAllText("json.txt", json);
            var id = tre.CreateTest("2429", jsonobj);
            Console.WriteLine(id);
        }

        private static void MainProgram()
        {
            Stopwatch watch = new Stopwatch();

            var ap = new Application();
            var folderpath = ConfigurationManager.AppSettings["folderpath"];
            var files = Directory.GetFiles(folderpath, "*");

            List<TestCase> testcases = new List<TestCase>();
            int iCurrentTestCase = 1;

            using (WordParser parser = new WordParser(ap))
            {
                foreach (var f in files)
                {
                    try
                    {
                        if (!f.Contains(".docx") || !f.Contains(".doc"))
                        {
                            throw new FileLoadException($"File '{f}' is not a doc or docx");
                        }

                        watch.Start();
                        Globals.Log("Loading tescase " + iCurrentTestCase + " " + f);

                        testcases.Add(parser.Parse(f, iCurrentTestCase));
                        watch.Stop();
                        Globals.Log("Time taken to parse : " + watch.ElapsedMilliseconds);
                        watch.Reset();

                        //test.PrintStats();
                        //test.PrintSequences();
                        //test.PrintHeights();

                    }
                    catch (FileLoadException fex)
                    {
                        Globals.Log("Error loading " + f);
                        Globals.Error(fex.Message);
                        Globals.Error(fex.StackTrace);
                    }
                    catch (WordParserException wex)
                    {
                        Globals.Log("Error parsing " + f);
                        Globals.Error(wex.Message);
                        Globals.Error(wex.StackTrace);
                        parser.CloseCurrentDoc();
                    }
                    catch (Exception ex)
                    {
                        Globals.Error(ex.Message);
                        Globals.Error(ex.StackTrace);
                        Console.WriteLine($"Unknown exception occurred on {f}. Press enter to close");
                        Console.Read();
                        break;
                    }
                    finally
                    {
                        watch.Stop();
                    }
                    iCurrentTestCase++;
                }
            }   

            ALMConvertor export = new ALMConvertor();
            ALMOctaneConnection con = new ALMOctaneConnection(
                ConfigurationManager.AppSettings["webAppUrl"],
                ConfigurationManager.AppSettings["userName"],
                ConfigurationManager.AppSettings["password"],
                ConfigurationManager.AppSettings["sharedSpaceId"],
                ConfigurationManager.AppSettings["workspaceId"]);

            ALMOctaneAPI api = new ALMOctaneAPI(con) { UserId = ConfigurationManager.AppSettings["userId"] };



            foreach (var tcs in testcases)
            {
                var str = export.ConvertToString(tcs);
                var testid = api.CreateTest(tcs.Title, tcs.Description);
                api.UpdateTest(testid, str);
                Globals.Log(str, true);
            }

            
        }
    }
}

