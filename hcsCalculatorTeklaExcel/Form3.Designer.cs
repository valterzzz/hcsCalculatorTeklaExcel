namespace hcsCalculatorTeklaExcel
{
    partial class Form3
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
            this.calculationlistBox = new System.Windows.Forms.ListBox();
            this.addCalcButton = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // calculationlistBox
            // 
            this.calculationlistBox.FormattingEnabled = true;
            this.calculationlistBox.Location = new System.Drawing.Point(12, 12);
            this.calculationlistBox.Name = "calculationlistBox";
            this.calculationlistBox.Size = new System.Drawing.Size(241, 199);
            this.calculationlistBox.TabIndex = 0;
            // 
            // addCalcButton
            // 
            this.addCalcButton.Location = new System.Drawing.Point(72, 222);
            this.addCalcButton.Name = "addCalcButton";
            this.addCalcButton.Size = new System.Drawing.Size(117, 23);
            this.addCalcButton.TabIndex = 1;
            this.addCalcButton.Text = "Add Calculation";
            this.addCalcButton.UseVisualStyleBackColor = true;
            this.addCalcButton.Click += new System.EventHandler(this.addCalcButton_Click);
            // 
            // Form3
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(268, 257);
            this.Controls.Add(this.addCalcButton);
            this.Controls.Add(this.calculationlistBox);
            this.Name = "Form3";
            this.Text = "Form3";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.ListBox calculationlistBox;
        private System.Windows.Forms.Button addCalcButton;
    }
}