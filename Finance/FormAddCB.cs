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
    public partial class FormAddCB : Form
    {
        // элементы управления
        class MCB
        {
            public TextBox _year;
            public NumericUpDown numericUpDown_K;
            public NumericUpDown numericUpDown_Ar;
            public NumericUpDown numericUpDown_Lat;
            public NumericUpDown numericUpDown_Ovt;
            public NumericUpDown numericUpDown_Lam;
            public NumericUpDown numericUpDown_Ovm;
            public NumericUpDown numericUpDown_Krd;
            public NumericUpDown numericUpDown_Ob;
            public NumericUpDown numericUpDown_A;
            public NumericUpDown numericUpDown_KrKr;
            public NumericUpDown numericUpDown_Vkl;
            public NumericUpDown numericUpDown_Inv;
            public NumericUpDown numericUpDown_Veks;
        }

        FormMethodology form_parent = null; // родительская форма

        List<MCB> l_MCB = new List<MCB>();  // список элементов

        public FormAddCB(FormMethodology _form_parent)
        {
            InitializeComponent();

            form_parent = _form_parent;
            button_add.Enabled = false;
        }

        // создаем наборы элементов управления
        private void button1_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < (int)numericUpDown1.Value; i++)
            {
                TextBox _year = new TextBox();
                NumericUpDown numericUpDown_K = new NumericUpDown();
                NumericUpDown numericUpDown_Ar = new NumericUpDown();
                NumericUpDown numericUpDown_Lat = new NumericUpDown();
                NumericUpDown numericUpDown_Ovt = new NumericUpDown();
                NumericUpDown numericUpDown_Lam = new NumericUpDown();
                NumericUpDown numericUpDown_Ovm = new NumericUpDown();
                NumericUpDown numericUpDown_Krd = new NumericUpDown();
                NumericUpDown numericUpDown_Ob = new NumericUpDown();
                NumericUpDown numericUpDown_A = new NumericUpDown();
                NumericUpDown numericUpDown_KrKr = new NumericUpDown();
                NumericUpDown numericUpDown_Vkl = new NumericUpDown();
                NumericUpDown numericUpDown_Inv = new NumericUpDown();
                NumericUpDown numericUpDown_Veks = new NumericUpDown();

                _year.Text = (i + 1).ToString() + " год";
                _year.Size = new Size(180, 20);
                panel1.Controls.Add(_year);
                _year.Location = new Point(190 * i, 0);

                setDef(numericUpDown_K, 190 * i, 26);
                setDef(numericUpDown_Ar, 190 * i, 26 * 2);
                setDef(numericUpDown_Lat, 190 * i, 26 * 3);
                setDef(numericUpDown_Ovt, 190 * i, 26 * 4);
                setDef(numericUpDown_Lam, 190 * i, 26 * 5);
                setDef(numericUpDown_Ovm, 190 * i, 26 * 6);
                setDef(numericUpDown_Krd, 190 * i, 26 * 7);
                setDef(numericUpDown_Ob, 190 * i, 26 * 8);
                setDef(numericUpDown_A, 190 * i, 26 * 9);
                setDef(numericUpDown_KrKr, 190 * i, 26 * 10);
                setDef(numericUpDown_Vkl, 190 * i, 26 * 11);
                setDef(numericUpDown_Inv, 190 * i, 26 * 12);
                setDef(numericUpDown_Veks, 190 * i, 26 * 13);

                l_MCB.Add(new MCB()
                {
                    _year = _year,
                    numericUpDown_K = numericUpDown_K,
                    numericUpDown_Ar = numericUpDown_Ar,
                    numericUpDown_Lat = numericUpDown_Lat,
                    numericUpDown_Ovt = numericUpDown_Ovt,
                    numericUpDown_Lam = numericUpDown_Lam,
                    numericUpDown_Ovm = numericUpDown_Ovm,
                    numericUpDown_Krd = numericUpDown_Krd,
                    numericUpDown_Ob = numericUpDown_Ob,
                    numericUpDown_A = numericUpDown_A,
                    numericUpDown_KrKr = numericUpDown_KrKr,
                    numericUpDown_Vkl = numericUpDown_Vkl,
                    numericUpDown_Inv = numericUpDown_Inv,
                    numericUpDown_Veks = numericUpDown_Veks
                });
            }

            button1.Enabled = false;
            button_add.Enabled = true;
        }

        // установка значение для элемента управления
        void setDef(NumericUpDown _n, int x, int y)
        {
            _n.TextAlign = HorizontalAlignment.Right;
            _n.DecimalPlaces = 8;
            _n.Minimum = decimal.MinValue;
            _n.Maximum = decimal.MaxValue;
            _n.Size = new Size(180, 20);

            panel1.Controls.Add(_n);
            _n.Location = new Point(x, y);
        }

        // кнопка добавить
        private void button_add_Click(object sender, EventArgs e)
        {
            foreach (MCB m in l_MCB)
            {
                if (form_parent != null)
                {
                    form_parent.Add_CB(m._year.Text
                        , m.numericUpDown_K.Value
                        , m.numericUpDown_Ar.Value
                        , m.numericUpDown_Lat.Value
                        , m.numericUpDown_Ovt.Value
                        , m.numericUpDown_Lam.Value
                        , m.numericUpDown_Ovm.Value
                        , m.numericUpDown_Krd.Value
                        , m.numericUpDown_Ob.Value
                        , m.numericUpDown_A.Value
                        , m.numericUpDown_KrKr.Value
                        , m.numericUpDown_Vkl.Value
                        , m.numericUpDown_Inv.Value
                        , m.numericUpDown_Veks.Value);
                }
            }
            Close();
        }

        private void FormAddCB_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (form_parent != null)
                form_parent.Close_addMCB(); // объявляем что та форма закрывается
        }
    }
}
