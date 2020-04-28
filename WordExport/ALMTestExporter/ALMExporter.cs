using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WordExport.ExcelExporter;
using WordExport.TestcaseObjects;

namespace WordExport.ALMTestExporter
{
    public class ALMExporter
    {

        public ALMExporter() { }

        /// <summary>
        /// Converts the testcase to a ALM Octane 'script' parameter for a API upload
        /// </summary>
        /// <param name="test"></param>
        /// <returns></returns>
        public string ConvertToString(TestCase test)
        {
            ExcelExport exp = new ExcelExport();
            var steps = exp.CreateALMTestStep(test.TestSteps);
            StringBuilder sb = new StringBuilder();

            sb.Append("- "+steps[0].Item1);
            if (!string.IsNullOrEmpty(steps[0].Item2))
            {
                sb.Append(steps[0].Item2);
            }
            for (int i = 1;i < steps.Count; i++)
            {
                sb.Append("\n- " + steps[i].Item1);
                if (!string.IsNullOrEmpty(steps[i].Item2))
                {
                    sb.Append("\n- ?" + steps[i].Item2);
                }
            }
            return sb.ToString();
        }

        /// <summary>
        /// This will get the description from the Testcase object and convert it to the format that ALM expects in the API. It's in html format
        /// </summary>
        /// <param name="descriptionFromTestCase"></param>
        /// <returns></returns>
        public string ConvertDescriptionToString(TestCase testcase)
        {
            var desc = testcase.Description.Replace("\r\n","\n");
            var splits = desc.Split('\n');
            StringBuilder sb = new StringBuilder();
            foreach (var s in splits)
            {
                sb.Append($"<p>{s.Replace("\t", "&nbsp;")}</p>\n"); // Octane doens't have \t chars and so we replace them with a space
            }
            desc = $"<html><body>\n{sb}</body></html>";
            return desc;
        }

    }
}
