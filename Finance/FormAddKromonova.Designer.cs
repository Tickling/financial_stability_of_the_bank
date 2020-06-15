namespace Finance
{
    partial class FormAddKromonova
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
            this.components = new System.ComponentModel.Container();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.label8 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.label10 = new System.Windows.Forms.Label();
            this.label11 = new System.Windows.Forms.Label();
            this.label12 = new System.Windows.Forms.Label();
            this.label13 = new System.Windows.Forms.Label();
            this.label14 = new System.Windows.Forms.Label();
            this.label16 = new System.Windows.Forms.Label();
            this.label15 = new System.Windows.Forms.Label();
            this.numericUpDown1 = new System.Windows.Forms.NumericUpDown();
            this.button1 = new System.Windows.Forms.Button();
            this.panel1 = new System.Windows.Forms.Panel();
            this.button2 = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDown1)).BeginInit();
            this.SuspendLayout();
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(12, 222);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(21, 13);
            this.label8.TabIndex = 21;
            this.label8.Text = "ЗК";
            this.toolTip1.SetToolTip(this.label8, "Защищенный капитал");
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(12, 196);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(21, 13);
            this.label9.TabIndex = 20;
            this.label9.Text = "АР";
            this.toolTip1.SetToolTip(this.label9, "Активы работающие");
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(12, 170);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(22, 13);
            this.label10.TabIndex = 19;
            this.label10.Text = "ЛА";
            this.toolTip1.SetToolTip(this.label10, "Ликвидные активы");
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(12, 144);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(22, 13);
            this.label11.TabIndex = 18;
            this.label11.Text = "СО";
            this.toolTip1.SetToolTip(this.label11, "Совокупность(суммарных) обязательств");
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Location = new System.Drawing.Point(12, 118);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(22, 13);
            this.label12.TabIndex = 17;
            this.label12.Text = "ОВ";
            this.toolTip1.SetToolTip(this.label12, "Обязательства до востребования");
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Location = new System.Drawing.Point(12, 92);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(14, 13);
            this.label13.TabIndex = 16;
            this.label13.Text = "К";
            this.toolTip1.SetToolTip(this.label13, "Собственный капитал");
            // 
            // label14
            // 
            this.label14.AutoSize = true;
            this.label14.Location = new System.Drawing.Point(12, 66);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(26, 13);
            this.label14.TabIndex = 15;
            this.label14.Text = "УФ";
            this.toolTip1.SetToolTip(this.label14, "Уставной фонд");
            // 
            // label16
            // 
            this.label16.AutoSize = true;
            this.label16.Location = new System.Drawing.Point(13, 41);
            this.label16.Name = "label16";
            this.label16.Size = new System.Drawing.Size(25, 13);
            this.label16.TabIndex = 26;
            this.label16.Text = "Год";
            this.toolTip1.SetToolTip(this.label16, "Уставной фонд");
            // 
            // label15
            // 
            this.label15.AutoSize = true;
            this.label15.Location = new System.Drawing.Point(12, 14);
            this.label15.Name = "label15";
            this.label15.Size = new System.Drawing.Size(98, 13);
            this.label15.TabIndex = 22;
            this.label15.Text = "Количество годов";
            // 
            // numericUpDown1
            // 
            this.numericUpDown1.Location = new System.Drawing.Point(116, 12);
            this.numericUpDown1.Minimum = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.numericUpDown1.Name = "numericUpDown1";
            this.numericUpDown1.Size = new System.Drawing.Size(46, 20);
            this.numericUpDown1.TabIndex = 23;
            this.numericUpDown1.Value = new decimal(new int[] {
            1,
            0,
            0,
            0});
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(168, 9);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(94, 23);
            this.button1.TabIndex = 24;
            this.button1.Text = "Сформировать";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // panel1
            // 
            this.panel1.AutoScroll = true;
            this.panel1.Location = new System.Drawing.Point(44, 38);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(385, 222);
            this.panel1.TabIndex = 25;
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(187, 266);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(75, 23);
            this.button2.TabIndex = 27;
            this.button2.Text = "Добавить";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // FormAddKromonova
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(439, 295);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.label16);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.numericUpDown1);
            this.Controls.Add(this.label15);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.label9);
            this.Controls.Add(this.label10);
            this.Controls.Add(this.label11);
            this.Controls.Add(this.label12);
            this.Controls.Add(this.label13);
            this.Controls.Add(this.label14);
            this.MaximizeBox = false;
            this.MaximumSize = new System.Drawing.Size(455, 334);
            this.MinimumSize = new System.Drawing.Size(455, 334);
            this.Name = "FormAddKromonova";
            this.ShowIcon = false;
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Добавить";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.FormAddKromonova_FormClosing);
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDown1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.ToolTip toolTip1;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.Label label14;
        private System.Windows.Forms.Label label15;
        private System.Windows.Forms.NumericUpDown numericUpDown1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label label16;
        private System.Windows.Forms.Button button2;
    }
}