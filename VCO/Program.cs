using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using VCO.Properties;

namespace VCO
{
    class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        static void Main(string[] args)
        {
            VzwCostOptimization costOpt = new VzwCostOptimization();
            if (File.Exists(Directory.GetCurrentDirectory() + "\\plan.xml"))
            {
                costOpt.LoadPlanInformation(Directory.GetCurrentDirectory() + "\\plan.xml");
            }
            else
            {
                costOpt.CreateTable();
            }

            costOpt.ReadFile(Settings.Default.FilePath);
            var optimalPlan = costOpt.SearchPlans(Settings.Default.FilePath);
            try
            {
                costOpt.UpdateFile(Settings.Default.FilePath, optimalPlan);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            Console.WriteLine("Done...");
            Console.ReadKey();

        }
    }
}
