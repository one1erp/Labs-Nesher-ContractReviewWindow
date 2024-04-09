namespace TestForm
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components;

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
            this.radDropDownList1 = new Telerik.WinControls.UI.RadDropDownList();
            this.button1 = new System.Windows.Forms.Button();
            this.radThemeManager1 = new Telerik.WinControls.RadThemeManager();
            this.button2 = new System.Windows.Forms.Button();
            this.userControl11 = new TestForm.UserControl1();
            ((System.ComponentModel.ISupportInitialize)(this.radDropDownList1)).BeginInit();
            this.SuspendLayout();
            // 
            // radDropDownList1
            // 
            this.radDropDownList1.Location = new System.Drawing.Point(402, 318);
            this.radDropDownList1.Name = "radDropDownList1";
            this.radDropDownList1.Size = new System.Drawing.Size(125, 20);
            this.radDropDownList1.TabIndex = 1;
            this.radDropDownList1.Text = "radDropDownList1";
            this.radDropDownList1.Visible = false;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(297, 334);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 2;
            this.button1.Text = "button1";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Visible = false;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(343, 13);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(75, 23);
            this.button2.TabIndex = 3;
            this.button2.Text = "button2";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // userControl11
            // 
            this.userControl11.Location = new System.Drawing.Point(158, 94);
            this.userControl11.Name = "userControl11";
            this.userControl11.Size = new System.Drawing.Size(684, 190);
            this.userControl11.TabIndex = 4;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1317, 814);
            this.Controls.Add(this.userControl11);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.radDropDownList1);
            this.Name = "Form1";
            this.Text = "Form1";
            ((System.ComponentModel.ISupportInitialize)(this.radDropDownList1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private Telerik.WinControls.UI.RadDropDownList radDropDownList1;
        private System.Windows.Forms.Button button1;
        private Telerik.WinControls.RadThemeManager radThemeManager1;
        private System.Windows.Forms.Button button2;
        private UserControl1 userControl11;

    }
}

