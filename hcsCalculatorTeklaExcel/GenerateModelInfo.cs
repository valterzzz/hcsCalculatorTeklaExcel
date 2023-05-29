using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Tekla.Structures.Model;

namespace hcsCalculatorTeklaExcel
{
    public static class GenerateModelInfo
    {
        public static bool CheckIfModelIsConnected(Tekla.Structures.Model.Model model)
        {
            bool isConnected = true;

            if (!model.GetConnectionStatus())
            {
                MessageBox.Show("Tekla Structures not connected!");
                isConnected = false;
            }
            return isConnected;
        }

        //Function to check if array has only Assembly
        public static bool CheckIfOnlyAssembly(ModelObjectEnumerator elementArray)
        {
            bool isAssembly = true;
            foreach (Tekla.Structures.Model.Object obj in elementArray)
            {
                if (obj.GetType().Name != "Assembly")
                {
                    MessageBox.Show("Non-Assemblies Selected!");
                    MessageBox.Show(obj.GetType().Name);
                    isAssembly = false;
                }

            }

            return isAssembly;
        }

        public static string ReturnModelCalcTypeDirect()
        {
            Model model = new Model();

            if (CheckIfModelIsConnected(model) == false)
            {
                return "";
            }

            string modelPath = model.GetInfo().ModelPath;

            string elementDirectory = modelPath + @"\hcsCheckerData\calcTypes";

            return elementDirectory;

        }

        public static string ReturnModelDirect()
        {
            Model model = new Model();

            if (CheckIfModelIsConnected(model) == false)
            {
                return "";
            }

            string modelPath = model.GetInfo().ModelPath;

            string elementDirectory = modelPath + @"\hcsCheckerData";

            return elementDirectory;

        }

        public static string ReturnModelAttributesDirect()
        {
            Model model = new Model();

            if (CheckIfModelIsConnected(model) == false)
            {
                return "";
            }

            string modelPath = model.GetInfo().ModelPath;

            string elementDirectory = modelPath + @"\attributes";

            return elementDirectory;

        }

    }
}
