namespace Finance
{
    partial class FormMultipliers
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
            System.Windows.Forms.DataVisualization.Charting.ChartArea chartArea1 = new System.Windows.Forms.DataVisualization.Charting.ChartArea();
            System.Windows.Forms.DataVisualization.Charting.Series series1 = new System.Windows.Forms.DataVisualization.Charting.Series();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.mK_File_ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.mk_saveExcel_ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.mK_saveAs_ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.mK_load_Excel_ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.mK_input_ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripComboBox_m = new System.Windows.Forms.ToolStripComboBox();
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.panel1 = new System.Windows.Forms.Panel();
            this.chart_m = new System.Windows.Forms.DataVisualization.Charting.Chart();
            this.panel2 = new System.Windows.Forms.Panel();
            this.button_legend_m = new System.Windows.Forms.Button();
            this.dataGridView_m = new System.Windows.Forms.DataGridView();
            this.menuStrip1.SuspendLayout();
            this.tableLayoutPanel1.SuspendLayout();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.chart_m)).BeginInit();
            this.panel2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView_m)).BeginInit();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.mK_File_ToolStripMenuItem,
            this.mK_load_Excel_ToolStripMenuItem,
            this.mK_input_ToolStripMenuItem,
            this.toolStripComboBox_m});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(964, 27);
            this.menuStrip1.TabIndex = 1;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // mK_File_ToolStripMenuItem
            // 
            this.mK_File_ToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.mk_saveExcel_ToolStripMenuItem,
            this.mK_saveAs_ToolStripMenuItem});
            this.mK_File_ToolStripMenuItem.Name = "mK_File_ToolStripMenuItem";
            this.mK_File_ToolStripMenuItem.Size = new System.Drawing.Size(48, 23);
            this.mK_File_ToolStripMenuItem.Text = "Файл";
            // 
            // mk_saveExcel_ToolStripMenuItem
            // 
            this.mk_saveExcel_ToolStripMenuItem.Name = "mk_saveExcel_ToolStripMenuItem";
            this.mk_saveExcel_ToolStripMenuItem.Size = new System.Drawing.Size(172, 22);
            this.mk_saveExcel_ToolStripMenuItem.Text = "Сохранить в Excel";
            this.mk_saveExcel_ToolStripMenuItem.Click += new System.EventHandler(this.mk_saveExcel_ToolStripMenuItem_Click);
            // 
            // mK_saveAs_ToolStripMenuItem
            // 
            this.mK_saveAs_ToolStripMenuItem.Name = "mK_saveAs_ToolStripMenuItem";
            this.mK_saveAs_ToolStripMenuItem.Size = new System.Drawing.Size(172, 22);
            this.mK_saveAs_ToolStripMenuItem.Text = "Сохранить как...";
            this.mK_saveAs_ToolStripMenuItem.Click += new System.EventHandler(this.mK_saveAs_ToolStripMenuItem_Click);
            // 
            // mK_load_Excel_ToolStripMenuItem
            // 
            this.mK_load_Excel_ToolStripMenuItem.Name = "mK_load_Excel_ToolStripMenuItem";
            this.mK_load_Excel_ToolStripMenuItem.Size = new System.Drawing.Size(103, 23);
            this.mK_load_Excel_ToolStripMenuItem.Text = "Загрузить Excel";
            this.mK_load_Excel_ToolStripMenuItem.Click += new System.EventHandler(this.mK_load_Excel_ToolStripMenuItem_Click);
            // 
            // mK_input_ToolStripMenuItem
            // 
            this.mK_input_ToolStripMenuItem.Name = "mK_input_ToolStripMenuItem";
            this.mK_input_ToolStripMenuItem.Size = new System.Drawing.Size(115, 23);
            this.mK_input_ToolStripMenuItem.Text = "Добавить данные";
            this.mK_input_ToolStripMenuItem.Click += new System.EventHandler(this.mK_input_ToolStripMenuItem_Click);
            // 
            // toolStripComboBox_m
            // 
            this.toolStripComboBox_m.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.toolStripComboBox_m.DropDownWidth = 280;
            this.toolStripComboBox_m.Name = "toolStripComboBox_m";
            this.toolStripComboBox_m.Size = new System.Drawing.Size(180, 23);
            this.toolStripComboBox_m.SelectedIndexChanged += new System.EventHandler(this.toolStripComboBox_m_SelectedIndexChanged);
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.ColumnCount = 1;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.tableLayoutPanel1.Controls.Add(this.panel1, 0, 1);
            this.tableLayoutPanel1.Controls.Add(this.panel2, 0, 0);
            this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel1.Location = new System.Drawing.Point(0, 27);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 2;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 150F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(964, 371);
            this.tableLayoutPanel1.TabIndex = 2;
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.chart_m);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(0, 221);
            this.panel1.Margin = new System.Windows.Forms.Padding(0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(964, 150);
            this.panel1.TabIndex = 0;
            // 
            // chart_m
            // 
            chartArea1.Name = "ChartArea1";
            this.chart_m.ChartAreas.Add(chartArea1);
            this.chart_m.Dock = System.Windows.Forms.DockStyle.Fill;
            this.chart_m.Location = new System.Drawing.Point(0, 0);
            this.chart_m.Name = "chart_m";
            series1.ChartArea = "ChartArea1";
            series1.Name = "Series2";
            this.chart_m.Series.Add(series1);
            this.chart_m.Size = new System.Drawing.Size(964, 150);
            this.chart_m.TabIndex = 1;
            this.chart_m.Text = "ch";
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.button_legend_m);
            this.panel2.Controls.Add(this.dataGridView_m);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel2.Location = new System.Drawing.Point(0, 0);
            this.panel2.Margin = new System.Windows.Forms.Padding(0);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(964, 221);
            this.panel2.TabIndex = 1;
            // 
            // button_legend_m
            // 
            this.button_legend_m.Location = new System.Drawing.Point(0, 0);
            this.button_legend_m.Name = "button_legend_m";
            this.button_legend_m.Size = new System.Drawing.Size(21, 21);
            this.button_legend_m.TabIndex = 4;
            this.button_legend_m.Text = "?";
            this.button_legend_m.UseVisualStyleBackColor = true;
            this.button_legend_m.Click += new System.EventHandler(this.button_legend_m_Click);
            // 
            // dataGridView_m
            // 
            this.dataGridView_m.AllowUserToAddRows = false;
            this.dataGridView_m.AllowUserToDeleteRows = false;
            this.dataGridView_m.AllowUserToResizeColumns = false;
            this.dataGridView_m.AllowUserToResizeRows = false;
            this.dataGridView_m.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView_m.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridView_m.Location = new System.Drawing.Point(0, 0);
            this.dataGridView_m.MultiSelect = false;
            this.dataGridView_m.Name = "dataGridView_m";
            this.dataGridView_m.ReadOnly = true;
            this.dataGridView_m.RowHeadersWidth = 60;
            this.dataGridView_m.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing;
            this.dataGridView_m.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullColumnSelect;
            this.dataGridView_m.Size = new System.Drawing.Size(964, 221);
            this.dataGridView_m.TabIndex = 3;
            this.dataGridView_m.RowPostPaint += new System.Windows.Forms.DataGridViewRowPostPaintEventHandler(this.dataGridView_m_RowPostPaint);
            // 
            // FormMultipliers
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(964, 398);
            this.Controls.Add(this.tableLayoutPanel1);
            this.Controls.Add(this.menuStrip1);
            this.MaximizeBox = false;
            this.MaximumSize = new System.Drawing.Size(980, 437);
            this.MinimumSize = new System.Drawing.Size(980, 437);
            this.Name = "FormMultipliers";
            this.ShowIcon = false;
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Мультипликаторы";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.FormMultipliers_FormClosing);
            this.Load += new System.EventHandler(this.FormMultipliers_Load);
            this.Move += new System.EventHandler(this.FormMultipliers_Move);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.tableLayoutPanel1.ResumeLayout(false);
            this.panel1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.chart_m)).EndInit();
            this.panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView_m)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem mK_File_ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem mk_saveExcel_ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem mK_saveAs_ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem mK_load_Excel_ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem mK_input_ToolStripMenuItem;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.DataVisualization.Charting.Chart chart_m;
        private System.Windows.Forms.Button button_legend_m;
        private System.Windows.Forms.DataGridView dataGridView_m;
        private System.Windows.Forms.ToolStripComboBox toolStripComboBox_m;
    }
}