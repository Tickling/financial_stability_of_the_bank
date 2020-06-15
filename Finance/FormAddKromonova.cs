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
    public partial class FormAddKromonova : Form
    {
        // элементы управления
        class MK
        {
            public TextBox _year;
            public NumericUpDown _YF;
            public NumericUpDown _K;
            public NumericUpDown _OV;
            public NumericUpDown _CO;
            public NumericUpDown _LA;
            public NumericUpDown _AR;
            public NumericUpDown _ZK;
        }

        FormMethodology form_parent = null; // родительская форма

        List<MK> l_MK = new List<MK>();     // список элементов  

        public FormAddKromonova(FormMethodology _form_parent)
        {
            InitializeComponent();

            form_parent = _form_parent;
            button2.Enabled = false;
        }

        // создаем наборы элементов управления
        private void button1_Click(object sender, EventArgs e)
        {
            for(int i = 0; i < (int)numericUpDown1.Value; i++)
            {
                TextBox _year = new TextBox();
                NumericUpDown _YF = new NumericUpDown();
                NumericUpDown _K = new NumericUpDown();
                NumericUpDown _OV = new NumericUpDown();
                NumericUpDown _CO = new NumericUpDown();
                NumericUpDown _LA = new NumericUpDown();
                NumericUpDown _AR = new NumericUpDown();
                NumericUpDown _ZK = new NumericUpDown();

                _year.Text = (i + 1).ToString() + " год";
                _year.Size = new Size(180, 20);
                panel1.Controls.Add(_year);
                _year.Location = new Point(190 * i, 0);

                setDef(_YF, 190 * i, 26);
                setDef(_K, 190 * i, 26 * 2);
                setDef(_OV, 190 * i, 26 * 3);
                setDef(_CO, 190 * i, 26 * 4);
                setDef(_LA, 190 * i, 26 * 5);
                setDef(_AR, 190 * i, 26 * 6);
                setDef(_ZK, 190 * i, 26 * 7);

                l_MK.Add(new MK()
                {
                    _year = _year,
                    _YF = _YF,
                    _K = _K,
                    _OV = _OV,
                    _CO = _CO,
                    _LA = _LA,
                    _AR = _AR,
                    _ZK = _ZK
                });
            }

            button1.Enabled = false;
            button2.Enabled = true;
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
        private void button2_Click(object sender, EventArgs e)
        {
            foreach(MK m in l_MK)
            {
                if (form_parent != null)
                {
                    form_parent.Add_MK(m._year.Text
                        , m._YF.Value
                        , m._K.Value
                        , m._OV.Value
                        , m._CO.Value
                        , m._LA.Value
                        , m._AR.Value
                        , m._ZK.Value);
                }
            }
            Close();
        }

        private void FormAddKromonova_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (form_parent != null)
                form_parent.Close_addMK(); // объявляем что та форма закрывается
        }
    }
}
