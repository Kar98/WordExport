using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WordExport.COMObjects;

namespace WordExport.TestcaseObjects
{
    /// <summary>
    /// Contains: Test step, Expected result,Pass/Fail, Comments.
    /// Note that this is a TestStep from the word doc and not what a true test step would be. A TestSequence is a proper test step.
    /// </summary>
    public class TestStep
    {
        public int StepNumber { get; set; }
        public bool IsPrerequisite { get; set; }
        public List<string> Description { get; set; } // Formatted list of descriptions
        public List<string> Expected { get; set; } // Formatted list of expected results

        public Dictionary<WordParagraph,List<string>> Sequences { get; set; } // This is used to link the Description and Expected result together.
        
        public List<WordParagraph> COMDescription { get; set; } // word paragraphs for description
        public List<WordParagraph> COMExpected { get; set; }// word paragraphs for expected results


        public TestStep(int stepNum, int stepWidth, int expectedWidth, Paragraphs stepParagraphs, Paragraphs expectedParagraphs)
        {
            /*
             * Load in the COM elements from the Doc
             * Arrange the Description into a Step
             * Set the expected result to the calculated spot
             */
            StepNumber = stepNum;
            COMDescription = new List<WordParagraph>();
            COMExpected = new List<WordParagraph>();
            foreach(var p in stepParagraphs.Cast<Paragraph>())
            {
                COMDescription.Add(new WordParagraph(p.Range,stepWidth));
            }
            foreach (var p in expectedParagraphs.Cast<Paragraph>())
            {
                COMExpected.Add(new WordParagraph(p.Range,expectedWidth));
            }

            SetRowNumbers(COMDescription);
            SetRowNumbers(COMExpected);
            SetStepDescription();
            SetExpected();
            SetTestSequence();

        }

        public TestStep(int stepNum, int stepWidth, Paragraphs stepParagraphs)
        {
            /*
             * Load in the COM elements from the Doc
             * Determine the test step
             * Determine the expected results
             * If it's a prerequisite step, then it mark it accordingly.
             */
            StepNumber = stepNum;
            COMDescription = new List<WordParagraph>();
            COMExpected = new List<WordParagraph>();
            foreach (var p in stepParagraphs.Cast<Paragraph>())
            {
                COMDescription.Add(new WordParagraph(p.Range,stepWidth));
            }

            SetRowNumbers(COMDescription);
            SetStepDescription();
            SetTestSequence();

        }

        /// <summary>
        /// Sets the row numbers for the Wordparagraphs provided
        /// </summary>
        private void SetRowNumbers(List<WordParagraph> comList)
        {
            int currentIdx = 1;
            foreach(var desc in comList)
            {
                desc.RowNumber = new List<int>();
                if(currentIdx == 1)
                {
                    //Set the first row
                    desc.RowNumber.Add(currentIdx);
                }
                if(desc.ParagraphHeight > 1)
                {
                    for(int i = currentIdx;i < currentIdx + desc.ParagraphHeight; i++)
                    {
                        desc.RowNumber.Add(i);
                    }
                }
                else
                {
                    desc.RowNumber.Add(currentIdx);
                }
                currentIdx += desc.ParagraphHeight;
            }
        }

        /// <summary>
        /// If the text is not indented higher than the base, it will add it as a new step. Otherwise it will append to the existing step with its own indentation.
        /// </summary>
        private void SetStepDescription()
        {
            Description = new List<string>();
            Sequences = new Dictionary<WordParagraph, List<string>>();

            StringBuilder sb = new StringBuilder();
            sb.Append("");
            //Format the text 
            foreach (var com in COMDescription)
            {
                // This will set the sequence so I can pair the Step and Expected together
                if (!string.IsNullOrEmpty(com.Text))
                {
                    Sequences.Add(com,new List<string>());
                }
                // This will format the text so when I need to output the result, it will look neat.
                if(com.ListType == WdListType.wdListBullet && com.ListLevel == 2)
                {
                    sb.AppendLine(com.Text);
                }
                else if (com.ListType == WdListType.wdListBullet && com.ListLevel == 3)
                {
                    sb.AppendLine(com.Text);
                }
                else if (com.ListLevel == 1)
                {
                    if (!string.IsNullOrEmpty(sb.ToString()))
                    {
                        Description.Add(sb.ToString());
                        sb.Clear();
                    }
                    if (!string.IsNullOrEmpty(com.Text))
                    {
                        sb.AppendLine(com.Text);
                    }
                }
                else
                {
                    throw new WordParserException("Unknown case in SetStepDescription");
                }

            }
            if (!string.IsNullOrEmpty(sb.ToString()))
            {
                Description.Add(sb.ToString()); // Load last step into description
            }
        }

        private void SetExpected()
        {
            Expected = new List<string>();
            StringBuilder sb = new StringBuilder();
            sb.Append("");
            //Format the text 
            foreach (var com in COMExpected)
            {
                // This will format the text so when I need to output the result, it will look neat.
                if (com.ListType == WdListType.wdListBullet && com.ListLevel == 2)
                {
                    sb.AppendLine(com.Text);
                }
                else if (com.ListType == WdListType.wdListBullet && com.ListLevel == 3)
                {
                    sb.AppendLine(com.Text);
                }
                else if (com.ListLevel == 1)
                {
                    if (!string.IsNullOrEmpty(sb.ToString()))
                    {
                        Expected.Add(sb.ToString());
                        sb.Clear();
                    }
                    if (!string.IsNullOrEmpty(com.Text))
                    {
                        sb.AppendLine(com.Text);
                    }
                }
                else
                {
                    throw new WordParserException("Unknown case in SetStepDescription");
                }

            }
            // After the loop is done, check the last result and see if it was added.
            if(sb.Length > 0)
            {
                Expected.Add(sb.ToString());
                sb.Clear();
            }
        }

        private void SetTestSequence()
        {
            var validExpecteds = from com in COMExpected
                                 where !string.IsNullOrEmpty(com.Text)
                                 select com;

            foreach (var ecom in validExpecteds)
            {
                var comNum = ecom.RowNumber.First();

                // If the Expected row number is higher than the last possible Description row number, then the rest of the Expected are part of the last Description.
                if(ecom.RowNumber.First() > COMDescription.Last().RowNumber.Last())
                {
                    var validDescription = from com in COMDescription
                                         where !string.IsNullOrEmpty(com.Text)
                                         select com;

                    Sequences[validDescription.Last()].Add(ecom.Text);
                }
                else
                {
                    var linq = from desc in COMDescription
                               where desc.RowNumber.Contains(comNum)
                               select desc;

                    var res = linq.First();
                    //If the matching row number is blank, then move up the table until text is found.
                    if (string.IsNullOrEmpty(res.Text))
                    {
                        var resIdx = COMDescription.IndexOf(res);
                        for (int i = resIdx; i > 0; i--)
                        {
                            if (!string.IsNullOrEmpty(COMDescription[i].Text))
                            {
                                Sequences[COMDescription[i]].Add(ecom.Text);
                                break;
                            }
                        }
                    }
                    else
                    {
                        Sequences[res].Add(ecom.Text);
                    }
                }
            }
        }

        /// <summary>
        /// Pass in the rownumber of the expected result and it will try to find the closest match of the Description
        /// </summary>
        /// <param name="expectedResultRowNum"></param>
        /// <returns></returns>
        private WordParagraph CalculatePosition(int expectedResultRowNum)
        {
            int dResCount = 1;
            foreach(var dcom in COMDescription)
            {
                if(dResCount <= expectedResultRowNum)
                {
                    return dcom;
                }
                dResCount += dcom.ParagraphHeight;
            }
            throw new WordParserException("Could not find a match for Expected Result to Description");
        }

    }
}
