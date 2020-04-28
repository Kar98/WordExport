
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WordExport.COMObjects;
using WordExport.TestcaseObjects;
using static WordExport.Globals;

namespace WordExport.ExcelExporter
{
    public class ExcelExport : IDisposable
    {
        Application ap;
        Workbook wb;
        List<TestCase> allTestcases;

        int currentid = 1; //  unique_id
        int r = 1; // current row

        // Excel column mapping
        const int c_Id = 1;
        const int c_type = 2;
        const int c_Name = 3;
        const int c_steptype=4;
        const int c_stepdescription=5;
        const int c_testtype=6;
        const int c_description=7;
        const int c_phase=8;


        public ExcelExport()
        {
            // : Create a new excel file. 
            // : Delete all the sheets except 1 called maunal tests
            // : Export a TestCase obj to the excel file in the correct format.
            allTestcases = new List<TestCase>();
        }

        public void Export(TestCase testcase)
        {
            ap = new Application();
            ap.DisplayAlerts = false;
            wb = ap.Workbooks.Add();

            ap.Worksheets["Sheet1"].Name = "manual tests";
            ap.Worksheets["Sheet2"].Delete();
            ap.Worksheets["Sheet3"].Delete();

            ap.Worksheets["manual tests"].Columns[5].ColumnWidth = 50.00;

            allTestcases.Add(testcase);
            
            WriteOctaneHeader();

            foreach(var t in allTestcases)
            {
                WriteTestCaseToFile(t);
            }

            wb.SaveAs(@"C:\Users\Kar98\Documents\Work\ALMOctane\testexcel.xlsx", XlFileFormat.xlOpenXMLWorkbook, ReadOnlyRecommended: false, ConflictResolution: XlSaveConflictResolution.xlLocalSessionChanges, AccessMode: XlSaveAsAccessMode.xlExclusive);
        }

        private void WriteOctaneHeader()
        {
            ap.Cells[r, c_Id] = "unique_id";
            ap.Cells[r, c_type] = "type";
            ap.Cells[r, c_Name] = "name";
            ap.Cells[r, c_steptype] = "step_type";
            ap.Cells[r, c_stepdescription] = "step_description";
            ap.Cells[r, c_testtype] = "test_type";
            ap.Cells[r, c_description] = "description";
            ap.Cells[r, c_phase] = "phase";
            r++;
        }


        private void WriteTestCaseToFile(TestCase testcase)
        {
            //Write the test case header
            IdRow();
            ap.Cells[r, c_type] = UploadType.TESTMANUAL.AsString();
            ap.Cells[r, c_Name] = testcase.Title;
            ap.Cells[r, c_description] = testcase.Description;
            ap.Cells[r, c_testtype] = UploadTestType.E2E.AsString();
            ap.Cells[r, c_phase] = UploadPhase.NEW.AsString();
            r++;

            var almSteps = CreateALMTestStep(testcase.TestSteps);
            
            foreach(var step in almSteps)
            {
                if (string.IsNullOrEmpty(step.Item2))
                {
                    WriteStep(step);
                }
                else
                {
                    WriteValidation(step);
                }
                
            }
        }

        public List<Tuple<string, string>> CreateALMTestStep(List<TestStep> steps)
        {
            //StringBuilder sb = new StringBuilder();
            List<Tuple<string, string>> outList = new List<Tuple<string, string>>();
            //Tuple<string, string> outList = new Tuple<string, string>();

            StringBuilder stepBuilder = new StringBuilder();
            StringBuilder expBuilder = new StringBuilder();

            foreach (var step in steps)
            {
                // If pre-req then do special steps
                if (step.IsPrerequisite)
                {
                    if (!string.IsNullOrEmpty(stepBuilder.ToString()))
                    {
                        outList.Add(new Tuple<string, string>(stepBuilder.ToString(), expBuilder.ToString()));
                    }
                    stepBuilder.Clear();
                    expBuilder.Clear();
                    outList.Add(AddPrereqStep(step));
                }
                else
                {
                    foreach (var seq in step.Sequences)
                    {
                        // Else if it's a sub item, add it to the main test step
                        if (seq.Key.ListType == Microsoft.Office.Interop.Word.WdListType.wdListBullet && seq.Key.ListLevel >= 2)
                        {
                            stepBuilder.AppendLine(seq.Key.Text);
                        }
                        // Otherwise it's assumed to be a test step.
                        else
                        {
                            // IF the previous step is not null, add it to the main output and continue.
                            if (!string.IsNullOrEmpty(stepBuilder.ToString()))
                            {
                                outList.Add(new Tuple<string, string>(stepBuilder.ToString(), expBuilder.ToString()));
                                stepBuilder.Clear();
                                expBuilder.Clear();
                            }

                            stepBuilder.AppendLine(seq.Key.Text);
                        }
                        // Test expected
                        if (seq.Value.Count > 0)
                        {
                            seq.Value.ForEach(a => expBuilder.AppendLine(a));
                        }
                    }
                }
               
            }
            if (!string.IsNullOrEmpty(stepBuilder.ToString()))
            {
                outList.Add(new Tuple<string, string>(stepBuilder.ToString(), expBuilder.ToString()));
            }
            return outList;
        }


        private Tuple<string, string> AddPrereqStep(TestStep step)
        {
            // Append any steps added, then add all the pre-req stesps
            StringBuilder stepBuilder = new StringBuilder();
            StringBuilder expBuilder = new StringBuilder();

            foreach (var seq in step.Sequences)
            {
                if(seq.Key.ListType == Microsoft.Office.Interop.Word.WdListType.wdListBullet)
                {
                    stepBuilder.AppendLine("\t*"+seq.Key.Text);
                }
                else
                {
                    stepBuilder.AppendLine(seq.Key.Text);
                }
            }
            return new Tuple<string, string>(stepBuilder.ToString(), expBuilder.ToString());
        }

        private void WriteStep(Tuple<string, string> key)
        {
            IdRow();
            ap.Cells[r, c_type] = UploadType.STEP.AsString();
            ap.Cells[r, c_steptype] = UploadStepType.SIMPLE.AsString();
            ap.Cells[r, c_stepdescription] = key.Item1;
            r++;
        }

        private void WriteValidation(Tuple<string, string> key)
        {
            IdRow();
            ap.Cells[r, c_type] = UploadType.STEP.AsString();
            ap.Cells[r, c_steptype] = UploadStepType.SIMPLE.AsString();
            ap.Cells[r, c_stepdescription] = key.Item1;
            r++;
            IdRow();
            ap.Cells[r, c_type] = UploadType.STEP.AsString();
            ap.Cells[r, c_steptype] = UploadStepType.VALIDATION.AsString();
            ap.Cells[r, c_stepdescription] = key.Item2;
            r++;
        }

        private void IdRow()
        {
            ap.Cells[r, c_Id] = currentid.ToString();
            currentid++;
        }

        public void Dispose()
        {
            if(wb != null)
            {
                wb.Close();
            }
            
            if(ap != null)
            {
                ap.Quit();
            }
            
        }
    }
}
