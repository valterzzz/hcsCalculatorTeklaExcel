using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace hcsCalculatorTeklaExcel
{
    public static class RegenerateControlItems
    {
        //Function to return all selection types from calcType folder
        private static List<string> returnListofAvalibleTypeSelections(string elementDirectory, string typeOfItem)
        {

            List<string> listOfTypes = new List<string>();

            if (typeOfItem == "calculation")
            {
                System.IO.Directory.CreateDirectory(elementDirectory);
                System.IO.DirectoryInfo di = new System.IO.DirectoryInfo(elementDirectory);
                System.IO.FileInfo[] Files = di.GetFiles("*.txt"); //Getting Text files

                foreach (System.IO.FileInfo file in Files)
                {
                    string[] lastFile = file.Name.Split('.');
                    listOfTypes.Add(lastFile[0]);

                }
            }

            if (typeOfItem == "strand")
            {
                System.IO.DirectoryInfo di = new System.IO.DirectoryInfo(elementDirectory);
                System.IO.FileInfo[] Files = di.GetFiles("*.m140000060"); //Getting Strand files

                foreach (System.IO.FileInfo file in Files)
                {
                    string lastFile = file.Name.Replace(".m140000060", "");
                    listOfTypes.Add(lastFile);

                }
            }
            

            return listOfTypes;

        }



        //Delete all comboBox Items and generate a new list of items with just Saved included from calcTypes folder
        public static void regenarateListOfTypes(Control x, string elementDirectory, string typeOfItem)
        {

            if (x is ComboBox comboBox)
            {
                comboBox.Items.Clear();

                List<string> listOfTypes = new List<string>();

                listOfTypes = returnListofAvalibleTypeSelections(elementDirectory, typeOfItem);

                foreach (string type in listOfTypes)
                {
                    comboBox.Items.Add(type);
                }

            }
            else if (x is ListBox listBox)
            {
                listBox.Items.Clear();

                List<string> listOfTypes = new List<string>();

                listOfTypes = returnListofAvalibleTypeSelections(elementDirectory, typeOfItem);

                foreach (string type in listOfTypes)
                {
                    listBox.Items.Add(type);
                }

            }    

        }
    }
}
