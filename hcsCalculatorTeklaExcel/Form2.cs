using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Tekla.Structures.Model;
using static System.Net.WebRequestMethods;

namespace hcsCalculatorTeklaExcel
{
    public partial class Form2 : Form
    {

        public Form2()
        {
            InitializeComponent();

            string elementDirectory = GenerateModelInfo.ReturnModelCalcTypeDirect();

            if (elementDirectory == "")
            {
                return;
            }

            RegenerateControlItems.regenarateListOfTypes(typeListCombobox, elementDirectory, "calculation");

            //Set up default parameters when opening the Form

            partialFactorCTextBox.Text = "1.5";
            partialFactorSTextBox.Text = "1.15";
            accTextBox.Text = "1";
            allowableDeflectionComboBox.Text = "1/250";
            allowableCamberTextBox.Text = "20";

            loadCategoryComboBox.Text = "Category A : domestic; residential areas";
            γGTextBox.Text = "1.35";
            γQTextBox.Text = "1.5";
            gkTextBox.Text = "2.1";
            qkTextBox.Text = "5";

            supportLenghtTextBox.Text = "0.1";

            concreteClassStrikingComboBox.Text = "C30/37";
            concreteClassFinalComboBox.Text = "C45/55";
            cementClassComboBox.Text = "N";

            strandClassComboBox.Text = "Y1860S7";

            string elementAttributesDirectory = GenerateModelInfo.ReturnModelAttributesDirect();

            RegenerateControlItems.regenarateListOfTypes(strandPaternComboBox, elementAttributesDirectory, "strand");


        }
            

        private void saveTypebutton_Click(object sender, EventArgs e)
        {
            string elementDirectory = GenerateModelInfo.ReturnModelCalcTypeDirect();

            if (elementDirectory == "")
            {
                return;
            }

            System.IO.Directory.CreateDirectory(elementDirectory);

            string calcTypeName = typeListCombobox.Text;

            //Create a Dict containing information

            IDictionary<string, string> set_names = new Dictionary<string, string>();

            set_names.Add("partialFactorCTextBox", partialFactorCTextBox.Text);
            set_names.Add("partialFactorSTextBox", partialFactorSTextBox.Text);
            set_names.Add("accTextBox", accTextBox.Text);
            set_names.Add("allowableDeflectionComboBox", allowableDeflectionComboBox.Text);
            set_names.Add("allowableCamberTextBox", allowableCamberTextBox.Text);

            set_names.Add("loadCategoryComboBox", loadCategoryComboBox.Text);
            set_names.Add("γGTextBox", γGTextBox.Text);
            set_names.Add("γQTextBox", γQTextBox.Text);
            set_names.Add("gkTextBox", gkTextBox.Text);
            set_names.Add("qkTextBox", qkTextBox.Text);

            set_names.Add("supportLenghtTextBox", supportLenghtTextBox.Text);

            set_names.Add("concreteClassStrikingComboBox", concreteClassStrikingComboBox.Text);
            set_names.Add("concreteClassFinalComboBox", concreteClassFinalComboBox.Text);
            set_names.Add("cementClassComboBox", cementClassComboBox.Text);

            set_names.Add("strandClassComboBox", strandClassComboBox.Text);
            set_names.Add("strandPaternComboBox", strandPaternComboBox.Text);


            //Check if input is correct
            foreach (KeyValuePair<string, string> kvp in set_names)
            {
                //Check if , is not used in value field, becuase it it used as a seperator between Key and Value in the settings file
                if (kvp.Value.Contains(",") == true)
                {
                    MessageBox.Show("Do not use , (comma) in field!");
                    return;
                }

                if (kvp.Value == "")
                {
                    MessageBox.Show("Fill all the input fields!");
                    return;
                }

            }


            //Do not let files with out a name be saved

            if (calcTypeName == "")
            {
                MessageBox.Show("Add a name for calculation Type!");
                return;
            }

            //Do not let files with dot in name to be saved, probobly need to add more non allowed characters
            else if (calcTypeName.Contains(".") == true)
            {
                MessageBox.Show("Do not use . (dot) in name!");
                return;
            }

            //Create a file or update it with all the Keys and Values from dialog.
            else
            {
                string filePath = elementDirectory + string.Format(@"\{0}.txt", (calcTypeName));
                var newFile = System.IO.File.Create(filePath);
                newFile.Close();

                using (FileStream fs = new FileStream(filePath, FileMode.OpenOrCreate))
                {
                    using (TextWriter tw = new StreamWriter(fs))
                    {
                        foreach (KeyValuePair<string, string>kvp in set_names)
                        {
                            tw.WriteLine(string.Format("{0},{1}", kvp.Key, kvp.Value));
                        }
                    }
                }
            }

            RegenerateControlItems.regenarateListOfTypes(typeListCombobox, elementDirectory, "calculation");

        }

        private void loadTypeButton_Click(object sender, EventArgs e)
        {
            string elementDirectory = GenerateModelInfo.ReturnModelCalcTypeDirect();

            if (elementDirectory == "")
            {
                return;
            }

            string calcTypeName = typeListCombobox.Text + ".txt";

            string filePath = elementDirectory + @"\" + calcTypeName;


            //Read values from txt file and fill in the form
            try
            {
                using (FileStream fs = new FileStream(filePath, FileMode.Open))
                {
                    string ln;

                    using (TextReader tr = new StreamReader(fs))
                    {
                        
                        while ((ln = tr.ReadLine()) != null)
                        {
                            string[] attributeKeyAndValue = ln.Split(',');
                            string attribueKey = attributeKeyAndValue[0];
                            string attribueValue = attributeKeyAndValue[1];

                            //Iterate trought all the textBoxies in form and fill the correct one with data from txt file

                            foreach(Control x in this.Controls)
                            {
                                if (x is TextBox & x.Name == attribueKey)
                                {
                                    x.Text = attribueValue;
                                }

                                if (x is ComboBox & x.Name == attribueKey)
                                {
                                    x.Text = attribueValue;
                                }

                            }


                        }


                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occured: " + ex.Message);
            }



        }
    }
}
