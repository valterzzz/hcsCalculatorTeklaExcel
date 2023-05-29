using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Tekla.Structures.Model;
using System.Collections;

using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

using System.Data.SQLite;
using Microsoft.Office.Interop.Excel;
using System.IO;

namespace hcsCalculatorTeklaExcel
{
    public delegate void NotifyEvent(Calculation data);

    

    public partial class Form1 : Form
    {
        public NotifyEvent notifyDelegate;

        List<Calculation> calcList = new List<Calculation>();

        //Create a SQLite Database for storing calculations + rest of the Data accociated with that
        public void CreateDatabaseFile(string elementDirectory)
        {
            string connectionString = "Data Source=" + elementDirectory + @"\myDatabase.sqlite;Version=3;";
            SQLiteConnection connection = new SQLiteConnection(connectionString);

            //Creates a DB, but if there is one, just opens and closes it.
            connection.Open();

            //Try to create database tables if they are not created yet
            try
            {
                string sqlCommand = "CREATE TABLE Calculations (ID INTEGER PRIMARY KEY, Name TEXT, ColorCode TEXT)";
                // Create a SQLiteCommand object to execute the SQL command
                SQLiteCommand command = new SQLiteCommand(sqlCommand, connection);
                command.ExecuteNonQuery();

                sqlCommand = "CREATE TABLE Slabs (GUID TEXT PRIMARY KEY, calculationID INTEGER, FOREIGN KEY (calculationID) REFERENCES Calculations(ID) ON DELETE CASCADE)";
                // Create a SQLiteCommand object to execute the SQL command
                command = new SQLiteCommand(sqlCommand, connection);
                command.ExecuteNonQuery();


            }
            catch (Exception ex)
            {
                
            }

            connection.Close();

        }

        public Form1()
        {
            string elementDirectory = GenerateModelInfo.ReturnModelDirect();

            CreateDatabaseFile(elementDirectory);

            InitializeComponent();
            selectDisplaySettingComboBox.SelectedIndex = 0;

            // Add the event handler for CellContentClick event of the DataGridView
            dataGridView1.CellContentClick += DataGridView1_CellContentClick;

            notifyDelegate += new NotifyEvent(ButtonClickedOnForm3);

            utilizationPictureBox.Hide();


        }

        private List<Calculation> GetCalculations(string elementDirectory)
        {
            // Create a list for dataGridView
            var list = new List<Calculation>();

            // Create a connection to the database
            string connectionString = "Data Source=" + elementDirectory + @"\myDatabase.sqlite;Version=3;";
            SQLiteConnection connection = new SQLiteConnection(connectionString);
            connection.Open();

            // Define the SQL command to select all rows from the table
            string sqlCommand = "SELECT * FROM Calculations";
            SQLiteCommand command = new SQLiteCommand(sqlCommand, connection);

            // Execute the command and get the results
            SQLiteDataReader reader = command.ExecuteReader();

            
            // Loop through the results and do something with each row
            while (reader.Read())
            {
                int id = Convert.ToInt32(reader["ID"]);
                string name = reader["Name"].ToString();
                string colorCode = reader["ColorCode"].ToString();
                Color color = ColorTranslator.FromHtml(colorCode);


                //Find how many Slabs there are for each Calculation and display in table
                
                list.Add(new Calculation()
                {
                    ID = id,
                    Name = name,
                    Assemblies = GetSlabGuidsByCalculationID(id).Count,
                    ColorCode = color
                });

            }

            // Close the reader and the connection to the database
            reader.Close();
            connection.Close();

            // Set the DataSource of the DataGridView to the list of Calculations
            dataGridView1.DataSource = list;

            // Create a new DataGridViewButtonColumn
            DataGridViewButtonColumn buttonColumn = new DataGridViewButtonColumn();
            buttonColumn.HeaderText = "Button";
            buttonColumn.Name = "Button";
            buttonColumn.Text = "Change Color";
            buttonColumn.UseColumnTextForButtonValue = true;
            buttonColumn.DefaultCellStyle.NullValue = null;
            buttonColumn.DefaultCellStyle.Padding = new Padding(0);



            // Since the button is not part of the Calculation but is added leiter, you need to manually adjust its location as it does not get delated when DGV data resets.
            // Also you need to create it only the first time.
            // I should probobly include this button in class defenition somehow

            if (dataGridView1.Columns.Contains(buttonColumn.HeaderText))
            {
                dataGridView1.Columns["Button"].DisplayIndex = dataGridView1.ColumnCount - 1;
            }
            else
            {
                // Add the button column to the DataGridView
                dataGridView1.Columns.Add(buttonColumn);
            }

            foreach (DataGridViewRow row in dataGridView1.Rows)
            {

                Calculation calculation = (Calculation)row.DataBoundItem;

                row.Cells["ColorCode"].Style.BackColor = calculation.ColorCode;

            }

            return list;
        }

        private void AddCalculation(string elementDirectory, Calculation data)
        {


            // Create a connection to the database
            string connectionString = "Data Source=" + elementDirectory + @"\myDatabase.sqlite;Version=3;";
            SQLiteConnection connection = new SQLiteConnection(connectionString);
            connection.Open();

            // Define the SQL command to insert rows into the table

            string sqlCommand = string.Format("INSERT INTO Calculations (Name, ColorCode) VALUES ('{0}', '{1}')", data.Name, ColorTranslator.ToHtml(data.ColorCode));
            SQLiteCommand command = new SQLiteCommand(sqlCommand, connection);
            command.ExecuteNonQuery();
            // Close the connection to the database
            connection.Close();


        }

        private void UpdateCalculation(string elementDirectory, Calculation data)
        {


            // Create a connection to the database
            string connectionString = "Data Source=" + elementDirectory + @"\myDatabase.sqlite;Version=3;";
            SQLiteConnection connection = new SQLiteConnection(connectionString);
            connection.Open();

            // Define the SQL command to insert rows into the table

            string sqlCommand = string.Format("UPDATE Calculations SET Name = '{0}', ColorCode = '{1}' WHERE ID = {2}", data.Name, ColorTranslator.ToHtml(data.ColorCode), data.ID);
            SQLiteCommand command = new SQLiteCommand(sqlCommand, connection);
            command.ExecuteNonQuery();
            // Close the connection to the database
            connection.Close();


        }


        public void DeleteCalculation(Calculation data, SQLiteConnection connection)
        {
                // Define the SQL command to delete a row from the table using a parameterized query
                string sqlCommand = "DELETE FROM Calculations WHERE ID = @idToDelete";
                SQLiteCommand command = new SQLiteCommand(sqlCommand, connection);

                // Add the parameter to the command
                command.Parameters.AddWithValue("@idToDelete", data.ID);

                // Execute the command
                int rowsAffected = command.ExecuteNonQuery();
        }

        public List<string> GetSlabGuidsByCalculationID(int calculationID)
        {
            List<string> slabGuids = new List<string>();

            string elementDirectory = GenerateModelInfo.ReturnModelDirect();

            // Create a connection to the database
            string connectionString = "Data Source=" + elementDirectory + @"\myDatabase.sqlite;Version=3;";
            SQLiteConnection connection = new SQLiteConnection(connectionString);

            // Open the database connection
            connection.Open();

            // Define the SQL command to select all rows from the "slabs" table for a given calculationID
            string sqlCommand = "SELECT GUID FROM Slabs WHERE calculationID = @calculationID";

            // Create a SQLiteCommand object to execute the SQL command
            SQLiteCommand command = new SQLiteCommand(sqlCommand, connection);

            // Add the parameter to the command
            command.Parameters.AddWithValue("@calculationID", calculationID);

            // Execute the command and read the results
            using (SQLiteDataReader reader = command.ExecuteReader())
            {
                while (reader.Read())
                {
                    string guid = reader["GUID"].ToString();
                    slabGuids.Add(guid);
                }
            }

            // Close the connection to the database
            connection.Close();

            return slabGuids;
        }

        public ArrayList GetSlabModelObjectsByCalculationID(int calculationID, Tekla.Structures.Model.Model model)
        {
  
            ArrayList slabModelObjects = new ArrayList();

            List<string> slabGuids = GetSlabGuidsByCalculationID(calculationID);

            // Loop through the GUIDs and get the corresponding model objects
            foreach (string guid in slabGuids)
            {

                ModelObject modelObject = model.SelectModelObject(new Tekla.Structures.Identifier(guid));
                if (modelObject != null)
                {
                    slabModelObjects.Add(modelObject);
                }
            }

            return slabModelObjects;
        }

        public string FillExcelSheetWithCalcData(string calcName, Excel._Worksheet xlWorksheet)
        {

            string calcStrandSettings = "";

            string elementDirectory = GenerateModelInfo.ReturnModelCalcTypeDirect();

            string calcTypeName = calcName + ".txt";

            string filePath = elementDirectory + @"\" + calcTypeName;


            //Read values from txt file and fill in the form
            try
            {
                using (FileStream fs = new FileStream(filePath, FileMode.Open))
                {
                    string ln;

                    using (TextReader tr = new StreamReader(fs))
                    {
                        //Iterate trought all settings from the file and fill the coresponging excel cells
                        while ((ln = tr.ReadLine()) != null)
                        {
                            string[] attributeKeyAndValue = ln.Split(',');
                            string attribueKey = attributeKeyAndValue[0];
                            string attribueValue = attributeKeyAndValue[1];

                            if (attribueKey == "partialFactorCTextBox")
                            {
                                Range cellToChange = xlWorksheet.Range["D5"];
                                cellToChange.Value = attribueValue;
                            }

                            if (attribueKey == "partialFactorSTextBox")
                            {
                                Range cellToChange = xlWorksheet.Range["D6"];
                                cellToChange.Value = attribueValue;
                            }

                            if(attribueKey == "accTextBox")
                            {
                                Range cellToChange = xlWorksheet.Range["D7"];
                                cellToChange.Value = attribueValue;
                            }

                            if (attribueKey == "allowableDeflectionComboBox")
                            {
                                Range cellToChange = xlWorksheet.Range["D9"];
                                cellToChange.Value = attribueValue;
                            }

                            if (attribueKey == "allowableCamberTextBox")
                            {
                                Range cellToChange = xlWorksheet.Range["D10"];
                                cellToChange.Value = attribueValue;
                            }

                            if (attribueKey == "loadCategoryComboBox")
                            {
                                Range cellToChange = xlWorksheet.Range["D14"];
                                cellToChange.Value = attribueValue;
                            }

                            if (attribueKey == "γGTextBox")
                            {
                                Range cellToChange = xlWorksheet.Range["D16"];
                                cellToChange.Value = attribueValue;
                            }

                            if (attribueKey == "γQTextBox")
                            {
                                Range cellToChange = xlWorksheet.Range["D17"];
                                cellToChange.Value = attribueValue;
                            }

                            if (attribueKey == "gkTextBox")
                            {
                                Range cellToChange = xlWorksheet.Range["D22"];
                                cellToChange.Value = attribueValue;
                            }

                            if (attribueKey == "qkTextBox")
                            {
                                Range cellToChange = xlWorksheet.Range["D23"];
                                cellToChange.Value = attribueValue;
                            }

                            if (attribueKey == "supportLenghtTextBox")
                            {
                                Range cellToChange = xlWorksheet.Range["D29"];
                                cellToChange.Value = attribueValue;
                            }

                            if (attribueKey == "concreteClassStrikingComboBox")
                            {
                                Range cellToChange = xlWorksheet.Range["D38"];
                                cellToChange.Value = attribueValue;
                            }

                            if (attribueKey == "concreteClassFinalComboBox")
                            {
                                Range cellToChange = xlWorksheet.Range["D39"];
                                cellToChange.Value = attribueValue;
                            }

                            if (attribueKey == "cementClassComboBox")
                            {
                                Range cellToChange = xlWorksheet.Range["D41"];
                                cellToChange.Value = attribueValue;
                            }

                            if (attribueKey == "strandClassComboBox")
                            {
                                Range cellToChange = xlWorksheet.Range["D45"];
                                cellToChange.Value = attribueValue;
                            }

                            //SPECIAL CASE TO RETURN CALC STRAND SETTINGS, SO THEY CAN BE USED TO GET ACTUAL DATA FOR HALLOW CORE SLAB
                            if (attribueKey == "strandPaternComboBox")
                            {
                                calcStrandSettings = attribueValue;
                            }

                        }


                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occured: " + ex.Message);
            }

            return calcStrandSettings;


        }

        public void FillExcelSheetWithStrandData(ArrayList strandData, Excel._Worksheet xlWorksheet, double hallowCoreSlabHeight)
        {
            //Set up for next itteration to know if there was a first row

            bool thereIsSecondRow = false;
            double firstRowHeight = 0.00;

            //Reset all strand data cells

            Range cellToChange = xlWorksheet.Range["D53"];
            cellToChange.Value = 0.00;
            cellToChange = xlWorksheet.Range["D54"];
            cellToChange.Value = 9.30;
            cellToChange = xlWorksheet.Range["D55"];
            cellToChange.Value = 0.00;
            cellToChange = xlWorksheet.Range["D57"];
            cellToChange.Value = 0.00;

            cellToChange = xlWorksheet.Range["D61"];
            cellToChange.Value = 0.00;
            cellToChange = xlWorksheet.Range["D62"];
            cellToChange.Value = 9.30;
            cellToChange = xlWorksheet.Range["D63"];
            cellToChange.Value = 0.00;
            cellToChange = xlWorksheet.Range["D65"];
            cellToChange.Value = 0.00;

            cellToChange = xlWorksheet.Range["D69"];
            cellToChange.Value = 0.00;
            cellToChange = xlWorksheet.Range["D70"];
            cellToChange.Value = 9.30;
            cellToChange = xlWorksheet.Range["D71"];
            cellToChange.Value = 0.00;
            cellToChange = xlWorksheet.Range["D73"];
            cellToChange.Value = 0.00;


            foreach (ArrayList strandRowData in strandData)
            {

                //Read values from txt file and fill in the form
                try
                {
                    if ((double)strandRowData[0] > hallowCoreSlabHeight / 2)
                    {
                        cellToChange = xlWorksheet.Range["D53"];
                        cellToChange.Value = strandRowData[1];

                        cellToChange = xlWorksheet.Range["D54"];
                        cellToChange.Value = strandRowData[3];

                        cellToChange = xlWorksheet.Range["D55"];
                        cellToChange.Value = strandRowData[2];

                        cellToChange = xlWorksheet.Range["D57"];
                        cellToChange.Value = strandRowData[0];


                    }
                    else
                    {
                        if (thereIsSecondRow == true)
                        {
                            cellToChange = xlWorksheet.Range["D61"];
                            cellToChange.Value = strandRowData[1];

                            cellToChange = xlWorksheet.Range["D62"];
                            cellToChange.Value = strandRowData[3];

                            cellToChange = xlWorksheet.Range["D63"];
                            cellToChange.Value = strandRowData[2];

                            cellToChange = xlWorksheet.Range["D65"];
                            cellToChange.Value = (double)strandRowData[0] - firstRowHeight;
                        }
                        else
                        {
                            
                            cellToChange = xlWorksheet.Range["D69"];
                            cellToChange.Value = strandRowData[1];

                            cellToChange = xlWorksheet.Range["D70"];
                            cellToChange.Value = strandRowData[3];

                            cellToChange = xlWorksheet.Range["D71"];
                            cellToChange.Value = strandRowData[2];

                            cellToChange = xlWorksheet.Range["D73"];
                            cellToChange.Value = strandRowData[0];

                            //Make sure to let next loop know that there it is a second row.
                            thereIsSecondRow = true;
                            firstRowHeight = (double)strandRowData[0];


                        }

                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("An error occured in strand setting: " + ex.Message);
                }
            }

        }

        public void FillExcelSheetWithGeometryData(double hallowCoreSlabLength, double hallowCoreSlabWidth, double hallowCoreSlabHeight, Excel._Worksheet xlWorksheet)
        {
            try
            {

                Range cellToChange = xlWorksheet.Range["D28"];
                cellToChange.Value = hallowCoreSlabLength / 1000;

                cellToChange = xlWorksheet.Range["D32"];
                cellToChange.Value = hallowCoreSlabWidth / 1000;

                cellToChange = xlWorksheet.Range["D34"];
                cellToChange.Value = "HCS " + hallowCoreSlabHeight.ToString();


            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occured: " + ex.Message);
            }
        }




        private void defineTypesbutton_Click(object sender, EventArgs e)
        {
            Form2 f2 = new Form2();
            f2.ShowDialog();

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            string elementDirectory = GenerateModelInfo.ReturnModelDirect();
            calcList = GetCalculations(elementDirectory);
            //dataGridView1.DataSource = calcList;

        }

        public void ButtonClickedOnForm3(Calculation data)
        {
            string elementDirectory = GenerateModelInfo.ReturnModelDirect();

            AddCalculation(elementDirectory, data);

            calcList = GetCalculations(elementDirectory);


        }

        private void addCalcButton_Click(object sender, EventArgs e)
        {
            Form3 f3 = new Form3(notifyDelegate);
            f3.ShowDialog();
        }

        private void deleteCalcButton_Click(object sender, EventArgs e)
        {

            string elementDirectory = GenerateModelInfo.ReturnModelDirect();

            // Enable foreign key constraints
            string pragmaCommand = "PRAGMA foreign_keys=ON";

            // Create a connection to the database
            string connectionString = "Data Source=" + elementDirectory + @"\myDatabase.sqlite;Version=3;";
            SQLiteConnection connection = new SQLiteConnection(connectionString);

            // Open the database connection
            connection.Open();

            // Enable foreign key constraints
            using (SQLiteCommand pragma = new SQLiteCommand(pragmaCommand, connection))
            {
                pragma.ExecuteNonQuery();

                foreach (DataGridViewRow row in this.dataGridView1.SelectedRows)
                {
                    Calculation calc = row.DataBoundItem as Calculation;
                    if (calc != null)
                    {

                        DeleteCalculation(calc, connection);
                    }
                }

                connection.Close();
            }

            // Close the connection to the database

            calcList = GetCalculations(elementDirectory);


        }

        private void assignToCalcButton_Click(object sender, EventArgs e)
        {

            Tekla.Structures.Model.Model model = new Tekla.Structures.Model.Model();

            Calculation selectedCalc = new Calculation();

            string elementDirectory = GenerateModelInfo.ReturnModelDirect();

            Tekla.Structures.Model.UI.ModelObjectSelector mS = new Tekla.Structures.Model.UI.ModelObjectSelector();

            Tekla.Structures.Model.ModelObjectEnumerator selectedObjectList = mS.GetSelectedObjects();
            Tekla.Structures.Model.ModelObjectEnumerator selectedObjectList2 = mS.GetSelectedObjects();

            if (GenerateModelInfo.CheckIfModelIsConnected(model) == false)
            {
                return;
            }

            if (GenerateModelInfo.CheckIfOnlyAssembly(selectedObjectList2) == false)
            {
                return;
            }

            // Create a connection to the database
            string connectionString = "Data Source=" + elementDirectory + @"\myDatabase.sqlite;Version=3;";
            SQLiteConnection connection = new SQLiteConnection(connectionString);
            connection.Open();

            foreach (DataGridViewRow row in this.dataGridView1.SelectedRows)
            {
                Calculation calc = row.DataBoundItem as Calculation;
                if (calc != null)
                {
                    selectedCalc = calc;
                }
            }

            // Define the SQL command to insert a new row into the "slabs" table using a parameterized query
            string sqlCommand = "INSERT OR REPLACE INTO Slabs (GUID, calculationID) VALUES (@guid, @calculationID)";


            // Create a single SQLiteCommand object to execute the batch insert
            SQLiteCommand command = new SQLiteCommand(sqlCommand, connection);

            // Add the parameters to the command
            command.Parameters.Add("@guid", DbType.String);
            command.Parameters.Add("@calculationID", DbType.Int32);

            // Begin a transaction
            using (var transaction = connection.BeginTransaction())
            {
                foreach (Assembly obj in selectedObjectList)
                {
                    if (obj != null)
                    {
                        // Set the parameter values for the current record
                        command.Parameters["@guid"].Value = obj.Identifier.GUID.ToString();
                        command.Parameters["@calculationID"].Value = selectedCalc.ID;


                        // Queue the command for execution
                        command.ExecuteNonQuery();
                    }
                }

                // Commit the transaction
                transaction.Commit();
            }

            // Close the connection to the database
            connection.Close();

            //Update Table in Form
            calcList = GetCalculations(elementDirectory);


        }

        private void unasignFromCalcButton_Click(object sender, EventArgs e)
        {

            Tekla.Structures.Model.Model model = new Tekla.Structures.Model.Model();

            string elementDirectory = GenerateModelInfo.ReturnModelDirect();

            Tekla.Structures.Model.UI.ModelObjectSelector mS = new Tekla.Structures.Model.UI.ModelObjectSelector();

            Tekla.Structures.Model.ModelObjectEnumerator selectedObjectList = mS.GetSelectedObjects();
            Tekla.Structures.Model.ModelObjectEnumerator selectedObjectList2 = mS.GetSelectedObjects();

            if (GenerateModelInfo.CheckIfModelIsConnected(model) == false)
            {
                return;
            }

            if (GenerateModelInfo.CheckIfOnlyAssembly(selectedObjectList2) == false)
            {
                return;
            }

            // Create a connection to the database
            string connectionString = "Data Source=" + elementDirectory + @"\myDatabase.sqlite;Version=3;";
            SQLiteConnection connection = new SQLiteConnection(connectionString);
            connection.Open();

            // Create a list of GUIDs to delete
            List<string> guidsToDelete = new List<string>();
            foreach (Assembly obj in selectedObjectList)
            {
                if (obj != null)
                {
                    guidsToDelete.Add(obj.Identifier.GUID.ToString());
                }
            }

            // Delete all rows with the specified GUIDs
            string sqlCommand = "DELETE FROM Slabs WHERE GUID IN (" + string.Join(",", guidsToDelete.Select(x => "'" + x + "'")) + ")";
            SQLiteCommand command = new SQLiteCommand(sqlCommand, connection);
            int rowsAffected = command.ExecuteNonQuery();
            

            // Close the connection to the database
            connection.Close();

            //Update Table in Form
            calcList = GetCalculations(elementDirectory);

        }

        private void selectByCalcButton_Click(object sender, EventArgs e)
        {
            Tekla.Structures.Model.Model model = new Tekla.Structures.Model.Model();

            Calculation selectedCalc = dataGridView1.SelectedRows[0].DataBoundItem as Calculation;

            // Get the ID of the selected calculation
            int calculationID = Convert.ToInt32(selectedCalc.ID);
            
            // Get the slabs for the selected calculation
            ArrayList slabs = GetSlabModelObjectsByCalculationID(calculationID, model);

            //Select the ob
            Tekla.Structures.Model.UI.ModelObjectSelector MS = new Tekla.Structures.Model.UI.ModelObjectSelector();
            MS.Select(slabs);


        }

        // Define the event handler for CellContentClick event
        private void DataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            string elementDirectory = GenerateModelInfo.ReturnModelDirect();

            // Check if the click event occurred in the ButtonColumn
            if (e.ColumnIndex == dataGridView1.Columns["Button"].Index && e.RowIndex >= 0)
            {

                ColorDialog colorDlg = new ColorDialog();

                if (colorDlg.ShowDialog() == DialogResult.OK)
                {

                    foreach (DataGridViewRow row in this.dataGridView1.SelectedRows)
                    {
                        Calculation calc = row.DataBoundItem as Calculation;
                        if (calc != null)
                        {
                            calc.ColorCode = colorDlg.Color;

                            UpdateCalculation(elementDirectory, calc);
                            calcList = GetCalculations(elementDirectory);
                        }
                    }


                }

            }
        }

        private void selectDisplaySettingComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            string elementDirectory = GenerateModelInfo.ReturnModelDirect();
            Tekla.Structures.Model.Model model = new Tekla.Structures.Model.Model();

            if (selectDisplaySettingComboBox.SelectedIndex == 0)
            {
                Tekla.Structures.Model.UI.ModelObjectVisualization.ClearAllTemporaryStates();
                utilizationPictureBox.Hide();

            }

            if (selectDisplaySettingComboBox.SelectedIndex == 1)
            {
                Tekla.Structures.Model.UI.ModelObjectVisualization.ClearAllTemporaryStates();
                utilizationPictureBox.Show();

                Excel.Application xlApp = new Excel.Application();
                Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(elementDirectory + @"\hcsCheckDemo.xlsx");
                Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];


                foreach (Calculation calc in calcList)
                {
                    //Because some of the API uses ArrayLists, some uses strongyly type list du-uh
                    var oldArrayList = GetSlabModelObjectsByCalculationID(calc.ID, model);
                    List<ModelObject> newList = oldArrayList.Cast<ModelObject>().ToList();

                    string calcStrandSettings = FillExcelSheetWithCalcData(calc.Name, xlWorksheet);


                    foreach (ModelObject obj in newList)
                    {
                        Tekla.Structures.Model.Assembly assembly = (Assembly)obj;

                        double hallowCoreSlabLength = 0.00;
                        double hallowCoreSlabWidth = 0.00;
                        double hallowCoreSlabHeight = 0.00;

                        assembly.GetReportProperty("CAST_UNIT_LENGTH_ONLY_CONCRETE_PARTS", ref hallowCoreSlabLength);
                        assembly.GetReportProperty("CAST_UNIT_WIDTH_ONLY_CONCRETE_PARTS", ref hallowCoreSlabWidth);
                        assembly.GetReportProperty("CAST_UNIT_HEIGHT_ONLY_CONCRETE_PARTS", ref hallowCoreSlabHeight);

                        //Get Geometry Data from ModelObject

                        ArrayList strandData = GenerateStrandInfo.GetStrandInfo(obj, calcStrandSettings, model);

                        FillExcelSheetWithStrandData(strandData, xlWorksheet, hallowCoreSlabHeight);

                        FillExcelSheetWithGeometryData(hallowCoreSlabLength, hallowCoreSlabWidth, hallowCoreSlabHeight, xlWorksheet);

                        Range resultCell = xlWorksheet.Range["K27"];


                        if (resultCell.Value < 0.9)
                        {
                            Color color = Color.Green;

                            double red = color.R / 255.00;
                            double green = color.G / 255.00;
                            double blue = color.B / 255.00;

                            Tekla.Structures.Model.UI.Color teklaColor = new Tekla.Structures.Model.UI.Color(red, green, blue);

                            List<ModelObject> objectList = new List<ModelObject>();

                            objectList.Add(obj);


                            Tekla.Structures.Model.UI.ModelObjectVisualization.SetTemporaryState(objectList, teklaColor);

                        }
                        if (resultCell.Value >= 0.9 & resultCell.Value <= 1)
                        {
                            Color color = Color.Yellow;

                            double red = color.R / 255.00;
                            double green = color.G / 255.00;
                            double blue = color.B / 255.00;

                            Tekla.Structures.Model.UI.Color teklaColor = new Tekla.Structures.Model.UI.Color(red, green, blue);

                            List<ModelObject> objectList = new List<ModelObject>();

                            objectList.Add(obj);

                            Tekla.Structures.Model.UI.ModelObjectVisualization.SetTemporaryState(objectList, teklaColor);

                        }
                        if (resultCell.Value > 1)
                        {
                            Color color = Color.Red;

                            double red = color.R / 255.00;
                            double green = color.G / 255.00;
                            double blue = color.B / 255.00;

                            Tekla.Structures.Model.UI.Color teklaColor = new Tekla.Structures.Model.UI.Color(red, green, blue);

                            List<ModelObject> objectList = new List<ModelObject>();

                            objectList.Add(obj);


                            Tekla.Structures.Model.UI.ModelObjectVisualization.SetTemporaryState(objectList, teklaColor);

                        }

                    }

                }

                //Disable this to not get Save/Dont save excel pop up after all calculations have been done. If enabled it is easy to debug by saving sheet and cheking data.
                xlApp.DisplayAlerts = false;
                 
                xlWorkbook.Close();

                //cleanup
                GC.Collect();
                GC.WaitForPendingFinalizers();

                //release com objects to fully kill excel process from running in the background
                Marshal.ReleaseComObject(xlWorksheet);

                //close and release
                Marshal.ReleaseComObject(xlWorkbook);

                //quit and release
                xlApp.Quit();
                Marshal.ReleaseComObject(xlApp);

                MessageBox.Show("All HCS have been checked!");

            }

            if (selectDisplaySettingComboBox.SelectedIndex == 2)
            {
                Tekla.Structures.Model.UI.ModelObjectVisualization.ClearAllTemporaryStates();
                utilizationPictureBox.Hide();

                calcList = GetCalculations(elementDirectory);

                foreach(Calculation calc in calcList)
                {
                    //Because some of the API uses ArrayLists, some uses strongyly type list du-uh
                    var oldArrayList = GetSlabModelObjectsByCalculationID(calc.ID, model);
                    List<ModelObject> newList = oldArrayList.Cast<ModelObject>().ToList();

                    //Create a Tekla colour

                    double red = calc.ColorCode.R / 255.00;
                    double green = calc.ColorCode.G / 255.00;
                    double blue = calc.ColorCode.B / 255.00;

                    Tekla.Structures.Model.UI.Color teklaColor = new Tekla.Structures.Model.UI.Color(red, green, blue);

                    Tekla.Structures.Model.UI.ModelObjectVisualization.SetTemporaryState(newList, teklaColor);



                }

                
            }
            
            
        }
    }

}
