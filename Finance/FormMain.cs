using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Finance
{
    public partial class FormMain : Form
    {
        FormMethodology _FormMethodology;   // Форма
        FormRisk _FormRisk;                 // Форма
        FormAnalysis _FormAnalysis;         // Форма
        FormGetData _FormGetData;           // Форма
        FormMultipliers _FormMultipliers;   // Форма

        public FormMain()
        {
            InitializeComponent();
        }

        private void button_Methodology_Click(object sender, EventArgs e)
        {
            if (_FormMethodology == null)
                _FormMethodology = new FormMethodology(this);       // создаем форму
            _FormMethodology.Show();                                // Показываем форму
            _FormMethodology.WindowState = FormWindowState.Normal;  // разворачиваем если свернута
            _FormMethodology.Focus();                               // Переводим фокус на форму
        }

        private void button_risk_Click(object sender, EventArgs e)
        {
            if (_FormRisk == null)
                _FormRisk = new FormRisk(this);             // создаем форму
            _FormRisk.Show();                               // Показываем форму
            _FormRisk.WindowState = FormWindowState.Normal; // разворачиваем если свернута
            _FormRisk.Focus();                              // Переводим фокус на форму
        }

        private void button_Analysis_Click(object sender, EventArgs e)
        {
            if (_FormAnalysis == null)
                _FormAnalysis = new FormAnalysis(this);         // создаем форму
            _FormAnalysis.Show();                               // Показываем форму
            _FormAnalysis.WindowState = FormWindowState.Normal; // разворачиваем если свернута
            _FormAnalysis.Focus();                              // Переводим фокус на форму
        }

        public void Close_Methodology()
        {
            _FormMethodology = null;    // когда закрыли форму
        }

        public void Close_Risk()
        {
            _FormRisk = null;           // когда закрыли форму
        }

        public void Close_Analysis()
        {
            _FormAnalysis = null;       // когда закрыли форму
        }

        private void button_Multipliers_Click(object sender, EventArgs e)
        {
            if (_FormMultipliers == null)
                _FormMultipliers = new FormMultipliers(this);       // создаем форму
            _FormMultipliers.Show();                                // Показываем форму
            _FormMultipliers.WindowState = FormWindowState.Normal;  // разворачиваем если свернута
            _FormMultipliers.Focus();                               // Переводим фокус на форму
        }

        public void Close_Multipliers()
        {
            _FormMultipliers = null;       // когда закрыли форму
        }

        private void button_get_data_Click(object sender, EventArgs e)
        {
            if (_FormGetData == null)
                _FormGetData = new FormGetData(this);       // создаем форму
            _FormGetData.Show();                                // Показываем форму
            _FormGetData.WindowState = FormWindowState.Normal;  // разворачиваем если свернута
            _FormGetData.Focus();                               // Переводим фокус на форму
        }

        public void Close_get_data()
        {
            _FormGetData = null;       // когда закрыли форму
        }
    }
}
