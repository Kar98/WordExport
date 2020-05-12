using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TestRailExport;
using WordExport.TestcaseObjects;

namespace WordExport.TestrailConverter
{
    public class TRConverter
    {

        public static TestrailJSON Convert(TestCase tc)
        {
            TestrailJSON rtn = new TestrailJSON()
            {
                Title = tc.Title,
                TemplateId = 2,
                TypeId = 6,
                PriorityId = 1,
                CustomTestscenario = tc.Description
            };

            List<CustomStepsSeparated> customstep = new List<CustomStepsSeparated>();
            foreach(var step in tc.TestSteps)
            {
                foreach(var s in step.Sequences)
                {
                    customstep.Add(new CustomStepsSeparated() { Content = s.Key.Text, Expected = string.Join("\n", s.Value.ToArray())});
                }
            }
            rtn.CustomStepsSeparated = customstep.ToArray();
            return rtn;
        }



    }
}
