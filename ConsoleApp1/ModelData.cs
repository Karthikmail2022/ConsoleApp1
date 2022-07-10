using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApp1
{


    public class TestModel
    {
        public int StockLength { get; set; }
        public int Waste { get; set; }
        public Dictionary<CutData, float> Cuts { get; set; }
        public string Material { get; set; }
        public string Type { get; set; }
        public int Diameter { get; set; }


    }


    public class CutData
    {
        public int PosNumber { get; set; }
        public string Cut { get; set; }
    }
}
