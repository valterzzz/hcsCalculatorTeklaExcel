using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Drawing;

namespace hcsCalculatorTeklaExcel
{
    public class Calculation
    {
        public int ID { get; set; }
        public string Name { get; set; }
        public int Assemblies { get; set; }
        public Color ColorCode { get; set; }  // new property for storing color

    }

}
