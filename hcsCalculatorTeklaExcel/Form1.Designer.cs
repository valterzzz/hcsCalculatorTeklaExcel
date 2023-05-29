namespace hcsCalculatorTeklaExcel
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.defineTypesbutton = new System.Windows.Forms.Button();
            this.addDeleteCalcButton = new System.Windows.Forms.Button();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.deleteCalcButton = new System.Windows.Forms.Button();
            this.assignToCalcButton = new System.Windows.Forms.Button();
            this.unasignFromCalcButton = new System.Windows.Forms.Button();
            this.selectByCalcButton = new System.Windows.Forms.Button();
            this.selectDisplaySettingComboBox = new System.Windows.Forms.ComboBox();
            this.utilizationPictureBox = new System.Windows.Forms.PictureBox();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.utilizationPictureBox)).BeginInit();
            this.SuspendLayout();
            // 
            // defineTypesbutton
            // 
            this.defineTypesbutton.Location = new System.Drawing.Point(638, 12);
            this.defineTypesbutton.Name = "defineTypesbutton";
            this.defineTypesbutton.Size = new System.Drawing.Size(139, 34);
            this.defineTypesbutton.TabIndex = 0;
            this.defineTypesbutton.Text = "Define Types";
            this.defineTypesbutton.UseVisualStyleBackColor = true;
            this.defineTypesbutton.Click += new System.EventHandler(this.defineTypesbutton_Click);
            // 
            // addDeleteCalcButton
            // 
            this.addDeleteCalcButton.Location = new System.Drawing.Point(638, 52);
            this.addDeleteCalcButton.Name = "addDeleteCalcButton";
            this.addDeleteCalcButton.Size = new System.Drawing.Size(139, 34);
            this.addDeleteCalcButton.TabIndex = 1;
            this.addDeleteCalcButton.Text = "Add Calc";
            this.addDeleteCalcButton.UseVisualStyleBackColor = true;
            this.addDeleteCalcButton.Click += new System.EventHandler(this.addCalcButton_Click);
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToDeleteRows = false;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(12, 12);
            this.dataGridView1.MultiSelect = false;
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.ReadOnly = true;
            this.dataGridView1.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dataGridView1.Size = new System.Drawing.Size(620, 249);
            this.dataGridView1.TabIndex = 2;
            // 
            // deleteCalcButton
            // 
            this.deleteCalcButton.Location = new System.Drawing.Point(638, 92);
            this.deleteCalcButton.Name = "deleteCalcButton";
            this.deleteCalcButton.Size = new System.Drawing.Size(139, 34);
            this.deleteCalcButton.TabIndex = 3;
            this.deleteCalcButton.Text = "Delete Calc";
            this.deleteCalcButton.UseVisualStyleBackColor = true;
            this.deleteCalcButton.Click += new System.EventHandler(this.deleteCalcButton_Click);
            // 
            // assignToCalcButton
            // 
            this.assignToCalcButton.Location = new System.Drawing.Point(12, 267);
            this.assignToCalcButton.Name = "assignToCalcButton";
            this.assignToCalcButton.Size = new System.Drawing.Size(94, 34);
            this.assignToCalcButton.TabIndex = 4;
            this.assignToCalcButton.Text = "Assign To Calculation";
            this.assignToCalcButton.UseVisualStyleBackColor = true;
            this.assignToCalcButton.Click += new System.EventHandler(this.assignToCalcButton_Click);
            // 
            // unasignFromCalcButton
            // 
            this.unasignFromCalcButton.Location = new System.Drawing.Point(112, 267);
            this.unasignFromCalcButton.Name = "unasignFromCalcButton";
            this.unasignFromCalcButton.Size = new System.Drawing.Size(94, 34);
            this.unasignFromCalcButton.TabIndex = 5;
            this.unasignFromCalcButton.Text = "Unassigned (Selected)";
            this.unasignFromCalcButton.UseVisualStyleBackColor = true;
            this.unasignFromCalcButton.Click += new System.EventHandler(this.unasignFromCalcButton_Click);
            // 
            // selectByCalcButton
            // 
            this.selectByCalcButton.Location = new System.Drawing.Point(638, 145);
            this.selectByCalcButton.Name = "selectByCalcButton";
            this.selectByCalcButton.Size = new System.Drawing.Size(139, 34);
            this.selectByCalcButton.TabIndex = 6;
            this.selectByCalcButton.Text = "Select by Calc";
            this.selectByCalcButton.UseVisualStyleBackColor = true;
            this.selectByCalcButton.Click += new System.EventHandler(this.selectByCalcButton_Click);
            // 
            // selectDisplaySettingComboBox
            // 
            this.selectDisplaySettingComboBox.FormattingEnabled = true;
            this.selectDisplaySettingComboBox.Items.AddRange(new object[] {
            "No Colouring",
            "By Utilization",
            "By Calculation"});
            this.selectDisplaySettingComboBox.Location = new System.Drawing.Point(638, 240);
            this.selectDisplaySettingComboBox.Name = "selectDisplaySettingComboBox";
            this.selectDisplaySettingComboBox.Size = new System.Drawing.Size(139, 21);
            this.selectDisplaySettingComboBox.TabIndex = 7;
            this.selectDisplaySettingComboBox.SelectedIndexChanged += new System.EventHandler(this.selectDisplaySettingComboBox_SelectedIndexChanged);
            // 
            // utilizationPictureBox
            // 
            this.utilizationPictureBox.Image = global::hcsCalculatorTeklaExcel.Properties.Resources.utilizationBar;
            this.utilizationPictureBox.Location = new System.Drawing.Point(212, 267);
            this.utilizationPictureBox.Name = "utilizationPictureBox";
            this.utilizationPictureBox.Size = new System.Drawing.Size(447, 34);
            this.utilizationPictureBox.TabIndex = 8;
            this.utilizationPictureBox.TabStop = false;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(789, 313);
            this.Controls.Add(this.utilizationPictureBox);
            this.Controls.Add(this.selectDisplaySettingComboBox);
            this.Controls.Add(this.selectByCalcButton);
            this.Controls.Add(this.unasignFromCalcButton);
            this.Controls.Add(this.assignToCalcButton);
            this.Controls.Add(this.deleteCalcButton);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.addDeleteCalcButton);
            this.Controls.Add(this.defineTypesbutton);
            this.Name = "Form1";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.utilizationPictureBox)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button defineTypesbutton;
        private System.Windows.Forms.Button addDeleteCalcButton;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Button deleteCalcButton;
        private System.Windows.Forms.Button assignToCalcButton;
        private System.Windows.Forms.Button unasignFromCalcButton;
        private System.Windows.Forms.Button selectByCalcButton;
        private System.Windows.Forms.ComboBox selectDisplaySettingComboBox;
        private System.Windows.Forms.PictureBox utilizationPictureBox;
    }
}

