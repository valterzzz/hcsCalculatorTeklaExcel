using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace hcsCalculatorTeklaExcel
{
    public partial class Form3 : Form
    {
        NotifyEvent notifyDel;

        public Form3(NotifyEvent notify)
        {
            InitializeComponent();
            notifyDel = notify;

            string elementDirectory = GenerateModelInfo.ReturnModelCalcTypeDirect();

            RegenerateControlItems.regenarateListOfTypes(calculationlistBox, elementDirectory, "calculation");


    }

        private void addCalcButton_Click(object sender, EventArgs e)
        {
            //Generate a random colour
            Random random = new Random();
            Color color = Color.FromArgb(random.Next(256), random.Next(256), random.Next(256));

            //ID and Assemblies do not matter here, I am not taking them from here.
            Calculation addedCalculation = new Calculation() { ID = 3, Name = calculationlistBox.Text, Assemblies = 0, ColorCode = color };

            notifyDel.Invoke(addedCalculation);

        }
    }
}
