using System;
using System.Collections.Generic;
using System.Dynamic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VCO
{
   
    public class DNA
    {
        private Random rand;
        public List<double> Genes { get; set; }
        public double Fitness { get; set; }
        public double TotalCost { get; set; }
        private string FilePath { get; set; }

        private VzwCostOptimization costCalc;
        public DNA(int geneSize, string path)
        {
            FilePath = path;
            costCalc = new VzwCostOptimization();
            rand = new Random(DateTime.Now.Millisecond);
            Genes = new List<double>();
            for (var i = 0; i < geneSize; i++)
            {
                Genes.Add(1 + rand.NextDouble());
            }
            costCalc.LoadPlanInformation(Directory.GetCurrentDirectory() + "\\plan.xml");
        }

        public void CalculateFitness()
        {
           
        }
    }
}
