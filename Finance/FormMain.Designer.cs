namespace Finance
{
    partial class FormMain
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
            this.button_Methodology = new System.Windows.Forms.Button();
            this.button_risk = new System.Windows.Forms.Button();
            this.button_Analysis = new System.Windows.Forms.Button();
            this.button_Multipliers = new System.Windows.Forms.Button();
            this.button_get_data = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // button_Methodology
            // 
            this.button_Methodology.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.button_Methodology.Location = new System.Drawing.Point(12, 12);
            this.button_Methodology.Name = "button_Methodology";
            this.button_Methodology.Size = new System.Drawing.Size(260, 70);
            this.button_Methodology.TabIndex = 0;
            this.button_Methodology.Text = "Методика";
            this.button_Methodology.UseVisualStyleBackColor = true;
            this.button_Methodology.Click += new System.EventHandler(this.button_Methodology_Click);
            // 
            // button_risk
            // 
            this.button_risk.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.button_risk.Location = new System.Drawing.Point(12, 88);
            this.button_risk.Name = "button_risk";
            this.button_risk.Size = new System.Drawing.Size(260, 90);
            this.button_risk.TabIndex = 1;
            this.button_risk.Text = "Расчёт рисков финансовой устойчивости";
            this.button_risk.UseVisualStyleBackColor = true;
            this.button_risk.Click += new System.EventHandler(this.button_risk_Click);
            // 
            // button_Analysis
            // 
            this.button_Analysis.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.button_Analysis.Location = new System.Drawing.Point(12, 184);
            this.button_Analysis.Name = "button_Analysis";
            this.button_Analysis.Size = new System.Drawing.Size(260, 65);
            this.button_Analysis.TabIndex = 2;
            this.button_Analysis.Text = "Анализ пассивов";
            this.button_Analysis.UseVisualStyleBackColor = true;
            this.button_Analysis.Click += new System.EventHandler(this.button_Analysis_Click);
            // 
            // button_Multipliers
            // 
            this.button_Multipliers.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.button_Multipliers.Location = new System.Drawing.Point(12, 326);
            this.button_Multipliers.Name = "button_Multipliers";
            this.button_Multipliers.Size = new System.Drawing.Size(260, 65);
            this.button_Multipliers.TabIndex = 3;
            this.button_Multipliers.Text = "Мультипликаторы";
            this.button_Multipliers.UseVisualStyleBackColor = true;
            this.button_Multipliers.Click += new System.EventHandler(this.button_Multipliers_Click);
            // 
            // button_get_data
            // 
            this.button_get_data.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.button_get_data.Location = new System.Drawing.Point(12, 255);
            this.button_get_data.Name = "button_get_data";
            this.button_get_data.Size = new System.Drawing.Size(260, 65);
            this.button_get_data.TabIndex = 4;
            this.button_get_data.Text = "Получить данные";
            this.button_get_data.UseVisualStyleBackColor = true;
            this.button_get_data.Click += new System.EventHandler(this.button_get_data_Click);
            // 
            // FormMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(284, 396);
            this.Controls.Add(this.button_get_data);
            this.Controls.Add(this.button_Multipliers);
            this.Controls.Add(this.button_Analysis);
            this.Controls.Add(this.button_risk);
            this.Controls.Add(this.button_Methodology);
            this.MaximizeBox = false;
            this.MaximumSize = new System.Drawing.Size(300, 435);
            this.MinimumSize = new System.Drawing.Size(300, 435);
            this.Name = "FormMain";
            this.ShowIcon = false;
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Главное меню";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button button_Methodology;
        private System.Windows.Forms.Button button_risk;
        private System.Windows.Forms.Button button_Analysis;
        private System.Windows.Forms.Button button_Multipliers;
        private System.Windows.Forms.Button button_get_data;
    }
}

