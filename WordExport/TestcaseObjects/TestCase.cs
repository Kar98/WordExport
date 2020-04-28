using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WordExport.COMObjects;

namespace WordExport.TestcaseObjects
{
    public class TestCase
    {
        const int testStepNumColIndex = 1;
        const int descriptionColIndex = 2;
        const int expectedColIndex = 3;

        public int Id { get; set; }

        public string Title { get; set; }
        public string Description { get; set; }
        public List<TestStep> TestSteps { get; set; }
        public int TotalSteps
        {
            get { 
            if(TestSteps == null)
                {
                    return 0;
                }
                else
                {
                    return TestSteps.Count;
                }
            }
        }

        public TestCase() 
        {
        }

        public void AddTestSteps(List<Table> tables)
        {
            TestSteps = new List<TestStep>();

            var tableIter = 0;
            var currentStep = 1;

            // For each table
            foreach (var table in tables)
            {
                //Globals.Log("Table iteration : " + tableIter);
                // Go through all the rows
                for (int i = 1; i <= table.Rows.Count; i++)
                {
                    //Globals.Log("Processing row " + i);
                    // If there is 1 cell, it could be a prereqo or a mini title  No consistency across the files
                    if (table.Rows[i].Cells.Count == 1)
                    {
                        var description = table.Cell(i, 1).Range.Paragraphs;
                        //Assume it's a prereq step
                        if (description[1].Range.Text.Contains("Pre-requisites"))
                        {
                            var stepWidth = (int)Math.Ceiling(table.Cell(1, 1).Width);
                            //Globals.Log("Prerequisite found " + currentStep);
                            var test = new TestStep(currentStep, stepWidth, description);
                            test.IsPrerequisite = true;
                            TestSteps.Add(test);
                            currentStep++;
                        }
                        else
                        {
                            //Globals.Log($"1 liner found at {currentStep}. Skipping step");
                        }
                    }
                    else if (table.Cell(i, 2).Range.Text.TrimJunk() == "Description" || table.Cell(i, 3).Range.Text.TrimJunk() == "Expected Result")
                    {
                        // Skip
                        //Globals.Log("Header found. Skipping " + currentStep);
                    }
                    // If a proper column, then do the standard logic.
                    else if (table.Rows[i].Cells.Count == 5)
                    {
                        var stepWidth = (int)Math.Ceiling(table.Cell(i, descriptionColIndex).Width);
                        var expectedWidth = (int)Math.Ceiling(table.Cell(i, expectedColIndex).Width);
                        //Globals.Log("stepWidth : " + stepWidth);
                        //Globals.Log("expectedWidth : " + expectedWidth);
                        //var num = int.Parse(table.Cell(i, testStepNumColIndex).Range.Text.TrimJunk());
                        var description = table.Cell(i, descriptionColIndex).Range.Paragraphs;
                        var expected = table.Cell(i, expectedColIndex).Range.Paragraphs;
                        TestSteps.Add(new TestStep(currentStep, stepWidth, expectedWidth, description, expected));
                        currentStep++;

                    }


                }
                tableIter++;

            }
        }

        /// <summary>
        /// Assumes that the first paragraph is the title.
        /// </summary>
        /// <param name="paragraphs"></param>
        public void AddTitle(List<Paragraph> paragraphs)
        {
            Title = paragraphs[0].Range.Text.TrimJunk();
        }

        /// <summary>
        /// Will iterate through all the paragraphs until a Table border is hit. For some reason 5 is no table, and 7 is a table. JustWordThings.
        /// </summary>
        /// <param name="paragraphs"></param>
        public void AddDescription(List<Paragraph> paragraphs)
        {
            StringBuilder sb = new StringBuilder();
            // Skip the first paragraph since it's assumed to be the title.
            for(int i = 1;i < paragraphs.Count; i++)
            {
                if (paragraphs[i].Range.Borders.Count < 7)
                {
                    var para = paragraphs[i].Range;
                    if(para.ListFormat.ListType == WdListType.wdListBullet && para.ListFormat.ListLevelNumber == 1)
                    {
                        sb.AppendLine(para.Text.TrimJunk());
                    }else if (para.ListFormat.ListType == WdListType.wdListBullet && para.ListFormat.ListLevelNumber == 2)
                    {
                        sb.AppendLine("\t*"+para.Text.TrimJunk());
                    }
                    else if (para.ListFormat.ListType == WdListType.wdListBullet && para.ListFormat.ListLevelNumber == 3)
                    {
                        sb.AppendLine("\t\t**" + para.Text.TrimJunk());
                    }
                    else
                    {
                        sb.AppendLine(paragraphs[i].Range.Text.TrimJunk());
                    }
                    
                }
                else
                {
                    break;
                }
            }
            Description = sb.ToString();
        }

        public void PrintStats()
        {
            //Log("Print stats");
            Log("Title : "+Title);
            Log("Description : "+Description);
            Log("Total test steps: " + TotalSteps);
            int iStep = 1;
            int iDesc = 1;
            if(TestSteps != null)
            {
                foreach (var step in TestSteps)
                {
                    //Log("iStep " + iStep);
                    foreach (var s in step.Description)
                    {
                        //Log("iDesc " + iDesc);
                        //Log(s);
                        iDesc++;
                    }
                    iStep++;
                    iDesc = 1;
                }
            }
            
        }

        public void PrintSequences()
        {
            if(TestSteps != null)
            {
                foreach (var step in TestSteps)
                {
                    Log("Step : " + step.StepNumber);
                    StringBuilder sb = new StringBuilder();
                    foreach (var seq in step.Sequences)
                    {
                        seq.Value.ForEach(a => sb.Append("," + a));
                        Log(seq.Key.Text + "|" + sb.ToString());
                        sb.Clear();
                    }
                }
            }
        }

        public void PrintHeights()
        {
            int totalD = 0;
            int totalE = 0;
            if (TestSteps != null)
            {
                foreach (var step in TestSteps)
                {
                    Log("Step : "+step.StepNumber);
                    Log("*Test step*");
                    foreach(var seq in step.COMDescription)
                    {
                        Log($"{seq.Text} - {seq.ParagraphHeight}");
                        totalD += seq.ParagraphHeight;
                    }
                    Log("*Test Expected*");
                    foreach (var seq in step.COMExpected)
                    {
                        Log($"{seq.Text} - {seq.ParagraphHeight}");
                        totalE += seq.ParagraphHeight;
                    }
                    Log("Total D : " + totalD);
                    Log("Total E : " + totalE);
                }
                
            }
            
        }

        private void Log(string s)
        {
            Console.WriteLine(s);
            File.AppendAllText(Globals.LogPath, s+"\r\n",Encoding.UTF8);
        }


        /// <summary>
        /// Iterates through the rows to find the Header
        /// </summary>
        /// <param name="t"></param>
        /// <param name="header"></param>
        /// <param name="headerColIdx"></param>
        /// <returns></returns>
        private int GetTableHeader(Table t, string header, int headerColIdx)
        {
            for (int i = 1; i <= t.Rows.Count; i++)
            {
                if (t.Cell(i, headerColIdx).Range.Text.Trim().Contains(header))
                {
                    return i;
                }
            }
            throw new WordParserException("Could not find header " + header);
        }

    }
}
