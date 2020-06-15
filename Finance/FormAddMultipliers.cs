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
    public partial class FormAddMultipliers : Form
    {
        // элементы управления
        class M
        {
            public TextBox _year;
            public NumericUpDown _P_E;
            public NumericUpDown _E_P;
            public NumericUpDown _P_S;
            public NumericUpDown _P_BV;
            public NumericUpDown _P_CF;
            public NumericUpDown _CF_P;
            public NumericUpDown _P_FCE;
            public NumericUpDown _FCE_P;
        }

        FormMultipliers form_parent = null; // родительская форма

        List<M> l_M = new List<M>();     // список элементов  

        public FormAddMultipliers(FormMultipliers _form_parent)
        {
            InitializeComponent();

            form_parent = _form_parent;
            button2.Enabled = false;
        }

        // установка значение для элемента управления
        void setDef(NumericUpDown _n, int x, int y)
        {
            _n.TextAlign = HorizontalAlignment.Right;
            _n.DecimalPlaces = 3;
            _n.Minimum = decimal.MinValue;
            _n.Maximum = decimal.MaxValue;
            _n.Size = new Size(180, 20);

            panel1.Controls.Add(_n);
            _n.Location = new Point(x, y);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < (int)numericUpDown1.Value; i++)
            {
                TextBox _year = new TextBox();
                NumericUpDown _P_E = new NumericUpDown();
                NumericUpDown _E_P = new NumericUpDown();
                NumericUpDown _P_S = new NumericUpDown();
                NumericUpDown _P_BV = new NumericUpDown();
                NumericUpDown _P_CF = new NumericUpDown();
                NumericUpDown _CF_P = new NumericUpDown();
                NumericUpDown _P_FCE = new NumericUpDown();
                NumericUpDown _FCE_P = new NumericUpDown();

                _year.Text = (i + 1).ToString() + " год";
                _year.Size = new Size(180, 20);
                panel1.Controls.Add(_year);
                _year.Location = new Point(190 * i, 0);

                setDef(_P_E, 190 * i, 26);
                setDef(_E_P, 190 * i, 26 * 2);
                setDef(_P_S, 190 * i, 26 * 3);
                setDef(_P_BV, 190 * i, 26 * 4);
                setDef(_P_CF, 190 * i, 26 * 5);
                setDef(_CF_P, 190 * i, 26 * 6);
                setDef(_P_FCE, 190 * i, 26 * 7);
                setDef(_FCE_P, 190 * i, 26 * 8);

                l_M.Add(new M()
                {
                    _year = _year,
                    _P_E = _P_E,
                    _E_P = _E_P,
                    _P_S = _P_S,
                    _P_BV = _P_BV,
                    _P_CF = _P_CF,
                    _CF_P = _CF_P,
                    _P_FCE = _P_FCE,
                    _FCE_P = _FCE_P
                });
            }

            button1.Enabled = false;
            button2.Enabled = true;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            foreach (M m in l_M)
            {
                if (form_parent != null)
                {
                    form_parent.Add_(m._year.Text
                        , m._P_E.Value
                        , m._E_P.Value
                        , m._P_S.Value
                        , m._P_BV.Value
                        , m._P_CF.Value
                        , m._CF_P.Value
                        , m._P_FCE.Value
                        , m._FCE_P.Value);
                }
            }
            Close();
        }

        private void FormAddMultipliers_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (form_parent != null)
                form_parent.Close_addM(); // объявляем что та форма закрывается
        }
    }
}
