namespace Finance
{
    partial class FormRisk
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
            this.label_H = new System.Windows.Forms.Label();
            this.label_size = new System.Windows.Forms.Label();
            this.label_profit = new System.Windows.Forms.Label();
            this.numericUpDown_H = new System.Windows.Forms.NumericUpDown();
            this.numericUpDown_size = new System.Windows.Forms.NumericUpDown();
            this.numericUpDown_profit = new System.Windows.Forms.NumericUpDown();
            this.label_risk = new System.Windows.Forms.Label();
            this.label_pr = new System.Windows.Forms.Label();
            this.label_pr_str = new System.Windows.Forms.Label();
            this.textBox_risk = new System.Windows.Forms.TextBox();
            this.textBox_pr = new System.Windows.Forms.TextBox();
            this.textBox_pr_str = new System.Windows.Forms.TextBox();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.fileToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.saveExcelToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.saveAsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.loadToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDown_H)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDown_size)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDown_profit)).BeginInit();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // label_H
            // 
            this.label_H.Location = new System.Drawing.Point(12, 24);
            this.label_H.Name = "label_H";
            this.label_H.Size = new System.Drawing.Size(202, 29);
            this.label_H.TabIndex = 0;
            this.label_H.Text = "Норматив достаточности собственных средств (капитала) H1.0";
            // 
            // label_size
            // 
            this.label_size.AutoSize = true;
            this.label_size.Location = new System.Drawing.Point(12, 66);
            this.label_size.Name = "label_size";
            this.label_size.Size = new System.Drawing.Size(159, 13);
            this.label_size.TabIndex = 1;
            this.label_size.Text = "Размер собственных средств";
            // 
            // label_profit
            // 
            this.label_profit.AutoSize = true;
            this.label_profit.Location = new System.Drawing.Point(12, 99);
            this.label_profit.Name = "label_profit";
            this.label_profit.Size = new System.Drawing.Size(207, 13);
            this.label_profit.TabIndex = 2;
            this.label_profit.Text = "Прибыль (убыток)до налого облажения";
            // 
            // numericUpDown_H
            // 
            this.numericUpDown_H.DecimalPlaces = 8;
            this.numericUpDown_H.Location = new System.Drawing.Point(220, 29);
            this.numericUpDown_H.Minimum = new decimal(new int[] {
            2,
            0,
            0,
            0});
            this.numericUpDown_H.Name = "numericUpDown_H";
            this.numericUpDown_H.Size = new System.Drawing.Size(130, 20);
            this.numericUpDown_H.TabIndex = 3;
            this.numericUpDown_H.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.numericUpDown_H.Value = new decimal(new int[] {
            2,
            0,
            0,
            0});
            this.numericUpDown_H.ValueChanged += new System.EventHandler(this.numericUpDown_H_ValueChanged);
            this.numericUpDown_H.KeyUp += new System.Windows.Forms.KeyEventHandler(this.numericUpDown_H_KeyUp);
            // 
            // numericUpDown_size
            // 
            this.numericUpDown_size.DecimalPlaces = 8;
            this.numericUpDown_size.Location = new System.Drawing.Point(220, 64);
            this.numericUpDown_size.Maximum = new decimal(new int[] {
            1241513984,
            370409800,
            542101,
            0});
            this.numericUpDown_size.Minimum = new decimal(new int[] {
            1241513984,
            370409800,
            542101,
            -2147483648});
            this.numericUpDown_size.Name = "numericUpDown_size";
            this.numericUpDown_size.Size = new System.Drawing.Size(130, 20);
            this.numericUpDown_size.TabIndex = 4;
            this.numericUpDown_size.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.numericUpDown_size.ValueChanged += new System.EventHandler(this.numericUpDown_H_ValueChanged);
            this.numericUpDown_size.KeyUp += new System.Windows.Forms.KeyEventHandler(this.numericUpDown_H_KeyUp);
            // 
            // numericUpDown_profit
            // 
            this.numericUpDown_profit.DecimalPlaces = 8;
            this.numericUpDown_profit.Location = new System.Drawing.Point(220, 97);
            this.numericUpDown_profit.Maximum = new decimal(new int[] {
            1241513984,
            370409800,
            542101,
            0});
            this.numericUpDown_profit.Minimum = new decimal(new int[] {
            1241513984,
            370409800,
            542101,
            -2147483648});
            this.numericUpDown_profit.Name = "numericUpDown_profit";
            this.numericUpDown_profit.Size = new System.Drawing.Size(130, 20);
            this.numericUpDown_profit.TabIndex = 5;
            this.numericUpDown_profit.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.numericUpDown_profit.ValueChanged += new System.EventHandler(this.numericUpDown_H_ValueChanged);
            this.numericUpDown_profit.KeyUp += new System.Windows.Forms.KeyEventHandler(this.numericUpDown_H_KeyUp);
            // 
            // label_risk
            // 
            this.label_risk.AutoSize = true;
            this.label_risk.Location = new System.Drawing.Point(12, 153);
            this.label_risk.Name = "label_risk";
            this.label_risk.Size = new System.Drawing.Size(155, 13);
            this.label_risk.TabIndex = 6;
            this.label_risk.Text = "Величина совокупного риска";
            // 
            // label_pr
            // 
            this.label_pr.AutoSize = true;
            this.label_pr.Location = new System.Drawing.Point(12, 179);
            this.label_pr.Name = "label_pr";
            this.label_pr.Size = new System.Drawing.Size(86, 13);
            this.label_pr.TabIndex = 7;
            this.label_pr.Text = "Показатель PR";
            // 
            // label_pr_str
            // 
            this.label_pr_str.AutoSize = true;
            this.label_pr_str.Location = new System.Drawing.Point(12, 205);
            this.label_pr_str.Name = "label_pr_str";
            this.label_pr_str.Size = new System.Drawing.Size(143, 13);
            this.label_pr_str.TabIndex = 8;
            this.label_pr_str.Text = "Финансовая устойчивость";
            // 
            // textBox_risk
            // 
            this.textBox_risk.Location = new System.Drawing.Point(173, 150);
            this.textBox_risk.Name = "textBox_risk";
            this.textBox_risk.ReadOnly = true;
            this.textBox_risk.Size = new System.Drawing.Size(177, 20);
            this.textBox_risk.TabIndex = 9;
            // 
            // textBox_pr
            // 
            this.textBox_pr.Location = new System.Drawing.Point(173, 176);
            this.textBox_pr.Name = "textBox_pr";
            this.textBox_pr.ReadOnly = true;
            this.textBox_pr.Size = new System.Drawing.Size(177, 20);
            this.textBox_pr.TabIndex = 10;
            // 
            // textBox_pr_str
            // 
            this.textBox_pr_str.Location = new System.Drawing.Point(173, 202);
            this.textBox_pr_str.Name = "textBox_pr_str";
            this.textBox_pr_str.ReadOnly = true;
            this.textBox_pr_str.Size = new System.Drawing.Size(177, 20);
            this.textBox_pr_str.TabIndex = 11;
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.fileToolStripMenuItem,
            this.loadToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(363, 24);
            this.menuStrip1.TabIndex = 12;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // fileToolStripMenuItem
            // 
            this.fileToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.saveExcelToolStripMenuItem,
            this.saveAsToolStripMenuItem});
            this.fileToolStripMenuItem.Name = "fileToolStripMenuItem";
            this.fileToolStripMenuItem.Size = new System.Drawing.Size(48, 20);
            this.fileToolStripMenuItem.Text = "Файл";
            // 
            // saveExcelToolStripMenuItem
            // 
            this.saveExcelToolStripMenuItem.Name = "saveExcelToolStripMenuItem";
            this.saveExcelToolStripMenuItem.Size = new System.Drawing.Size(171, 22);
            this.saveExcelToolStripMenuItem.Text = "Созранить в Excel";
            this.saveExcelToolStripMenuItem.Click += new System.EventHandler(this.saveExcelToolStripMenuItem_Click);
            // 
            // saveAsToolStripMenuItem
            // 
            this.saveAsToolStripMenuItem.Name = "saveAsToolStripMenuItem";
            this.saveAsToolStripMenuItem.Size = new System.Drawing.Size(171, 22);
            this.saveAsToolStripMenuItem.Text = "Сохранить как...";
            this.saveAsToolStripMenuItem.Click += new System.EventHandler(this.saveAsToolStripMenuItem_Click);
            // 
            // loadToolStripMenuItem
            // 
            this.loadToolStripMenuItem.Name = "loadToolStripMenuItem";
            this.loadToolStripMenuItem.Size = new System.Drawing.Size(103, 20);
            this.loadToolStripMenuItem.Text = "Загрузить Excel";
            this.loadToolStripMenuItem.Click += new System.EventHandler(this.loadToolStripMenuItem_Click);
            // 
            // FormRisk
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(363, 229);
            this.Controls.Add(this.textBox_pr_str);
            this.Controls.Add(this.textBox_pr);
            this.Controls.Add(this.textBox_risk);
            this.Controls.Add(this.label_pr_str);
            this.Controls.Add(this.label_pr);
            this.Controls.Add(this.label_risk);
            this.Controls.Add(this.numericUpDown_profit);
            this.Controls.Add(this.numericUpDown_size);
            this.Controls.Add(this.numericUpDown_H);
            this.Controls.Add(this.label_profit);
            this.Controls.Add(this.label_size);
            this.Controls.Add(this.label_H);
            this.Controls.Add(this.menuStrip1);
            this.MainMenuStrip = this.menuStrip1;
            this.MaximizeBox = false;
            this.MaximumSize = new System.Drawing.Size(379, 268);
            this.MinimumSize = new System.Drawing.Size(379, 268);
            this.Name = "FormRisk";
            this.ShowIcon = false;
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Расчёт рисков финансовой устойчивости";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.FormRisk_FormClosing);
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDown_H)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDown_size)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDown_profit)).EndInit();
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label_H;
        private System.Windows.Forms.Label label_size;
        private System.Windows.Forms.Label label_profit;
        private System.Windows.Forms.NumericUpDown numericUpDown_H;
        private System.Windows.Forms.NumericUpDown numericUpDown_size;
        private System.Windows.Forms.Label label_risk;
        private System.Windows.Forms.Label label_pr;
        private System.Windows.Forms.Label label_pr_str;
        private System.Windows.Forms.TextBox textBox_risk;
        private System.Windows.Forms.TextBox textBox_pr;
        private System.Windows.Forms.TextBox textBox_pr_str;
        private System.Windows.Forms.NumericUpDown numericUpDown_profit;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem fileToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem saveExcelToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem saveAsToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem loadToolStripMenuItem;
    }
}