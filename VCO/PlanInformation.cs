using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VCO
{
    public class Plan
    {
        public int Size { get; set; }
        public double Cost { get; set; }
        public double OverageCost { get; set; }


    }
    public static class PlanInformation
    {
        public static Plan GetInfo(int planId)
        {

            switch (planId)
            {
                case 0:
                    {
                        return new Plan { Size = 3, Cost = 1, OverageCost = 0.7 };
                    }
                case 1:
                    {
                        return new Plan { Size = 25, Cost = 7, OverageCost = 0.009 };
                    }
                case 2:
                    {
                        return new Plan { Size = 250, Cost = 8, OverageCost = 0.009 };
                    }
                case 3:
                    {
                        return new Plan { Size = 500, Cost = 10, OverageCost = 0.009 };
                    }
                case 4:
                    {
                        return new Plan { Size = 1024, Cost = 15, OverageCost = 0.009 }; //old cost of 1gb is 20*
                    }
                case 5:
                    {
                        return new Plan { Size = 5120, Cost = 35, OverageCost = 0.009 };
                    }
                case 6:
                    {
                        return new Plan { Size = 10240, Cost = 60, OverageCost = 0.009 };
                    }
                case 7:
                    {
                        return new Plan { Size = 20480, Cost = 125, OverageCost = 0.009 };
                    }
                case 8:
                    {
                        return new Plan { Size = 30720, Cost = 235, OverageCost = 0.009 };
                    }


            }


            return null;
        }
    }
}
