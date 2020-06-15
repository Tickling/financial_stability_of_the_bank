using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;
using Syncfusion.XlsIO;
using System.Xml.Serialization;
using System.IO;

namespace Finance
{
    public partial class FormMethodology : Form
    {
        // класс для сохранения в xml
        public class ks
        {
            public decimal ks_K1;
            public decimal ks_K2;
            public decimal ks_K3;
            public decimal ks_K4;
            public decimal ks_K5;
            public decimal ks_K6;

            public decimal ks_d_K1;
            public decimal ks_d_K2;
            public decimal ks_d_K3;
            public decimal ks_d_K4;
            public decimal ks_d_K5;
            public decimal ks_d_K6;
        }

        public static decimal k_K1 = 45;
        public static decimal k_K2 = 20;
        public static decimal k_K3 = 10;
        public static decimal k_K4 = 15;
        public static decimal k_K5 = 5;
        public static decimal k_K6 = 5;

        public static decimal k_d_K1 = 1;
        public static decimal k_d_K2 = 1;
        public static decimal k_d_K3 = 1;
        public static decimal k_d_K4 = 1;
        public static decimal k_d_K5 = 1;
        public static decimal k_d_K6 = 1;


        public static string filename_k = "ks.xml";

        // метод Кромонова
        public class cMK
        {
            public string name;

            public decimal YF;
            public decimal K;
            public decimal OV;
            public decimal CO;
            public decimal LA;
            public decimal AR;
            public decimal ZK;

            public decimal? K1 { get { return AR == 0 ? (decimal?)null : (K / AR); } }
            public decimal? K2 { get { return OV == 0 ? (decimal?)null : (LA / OV); } }
            public decimal? K3 { get { return AR == 0 ? (decimal?)null : (CO / AR); } }
            public decimal? K4 { get { return CO == 0 ? (decimal?)null : ((LA + ZK) / CO); } }
            public decimal? K5 { get { return K == 0 ? (decimal?)null : (ZK / K); } }
            public decimal? K6 { get { return YF == 0 ? (decimal?)null : (K / YF); } }

            public decimal? N
            {
                get
                {
                    if (K1 == null || K2 == null || K3 == null || K4 == null || K5 == null || K6 == null)
                    {
                        return (decimal?)null;
                    }
                    else
                    {
                        return k_K1 * K1 / k_d_K1
                            + k_K2 * K2 / k_d_K2
                            + k_K3 * K3 / k_d_K3
                            + k_K4 * K4 / k_d_K4
                            + k_K5 * K5 / k_d_K5
                            + k_K6 * K6 / k_d_K6;
                    }
                }
            }

            public string N_rez
            {
                get
                {
                    string rez = "";

                    if (N != null)
                    {
                        if (N < (decimal)33.7)
                            rez = "кризисное";
                        else if (N >= (decimal)33.7 && N < (decimal)67.3)
                            rez = "проблемное";
                        else if (N >= (decimal)67.3 && N < (decimal)100.9)
                            rez = "с некоторыми признаками проблемности";
                        else if (N >= (decimal)100.9 && N < (decimal)134.5)
                            rez = "хорошее";
                        else if (N >= (decimal)134.5)
                            rez = "отличное";
                    }

                    return rez;
                }
            }
        }

        // Метод ЦБ
        public class cCB
        {
            public string name;

            public decimal K;
            public decimal Ar;
            public decimal Lat;
            public decimal Ovt;
            public decimal Lam;
            public decimal Ovm;
            public decimal Krd;
            public decimal Ob;
            public decimal A;
            public decimal KrKr;
            public decimal Vkl;
            public decimal Inv;
            public decimal Veks;

            public decimal? H1 { get { return Ar == 0 ? (decimal?)null : (K / Ar); } }
            public decimal? H2 { get { return Ovt == 0 ? (decimal?)null : (Lat / Ovt); } }
            public decimal? H3 { get { return Ovm == 0 ? (decimal?)null : (Lam / Ovm); } }
            public decimal? H4 { get { return Ob == 0 ? (decimal?)null : (Krd / Ob); } }
            public decimal? H5 { get { return A == 0 ? (decimal?)null : (Lat / A); } }
            public decimal? H7 { get { return K == 0 ? (decimal?)null : (KrKr / K); } }
            public decimal? H11 { get { return K == 0 ? (decimal?)null : (Vkl / K); } }
            public decimal? H12 { get { return K == 0 ? (decimal?)null : (Inv / K); } }
            public decimal? H13 { get { return K == 0 ? (decimal?)null : (Veks / K); } }
        }

        List<cMK> list_mk = new List<cMK>();        // список метода Кромонова
        string filename_load_mk = "";               // путь к отрытому/сохраненному файлу метода Кромонова
        FormLegendMK legendMK = new FormLegendMK(); // форма легенды
        FormAddKromonova addMK = null;              // форма добавления

        List<cCB> list_cb = new List<cCB>();            // список метода ЦБ
        string filename_load_mcb = "";                  // путь к отрытому/сохраненному файлу метода ЦБ
        FormLegendMCB legendMCB = new FormLegendMCB();  // форма легенды
        FormAddCB addMCB = null;                        // форма добавления

        static string decimal_format = "##0.00000000";  // формат числа
        FormMain _FormMain = null;                      // форма родителя


        public FormMethodology(FormMain f)
        {
            InitializeComponent();

            _FormMain = f;
        }

        private void FormMethodology_Load(object sender, EventArgs e)
        {
            if (filename_load_mk == "")
                mk_saveExcel_ToolStripMenuItem.Enabled = false; // деактивируем кнопку если нет пути к файлу
            mk_showTable();                                     // пересчитываем метод Кромонова
            k_getDef();                                         // загрузка файла или значения по-умолчанию с коэффициентами для N

            legendMK.Show();            // показываем форму легенды
            legendMK.Visible = false;   // скрываем легенду

            legendMCB.Show();           // показываем форму легенды
            legendMCB.Visible = false;  // скрываем легенду
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            legendMK.Visible = false;   // скрываем легенду
            legendMCB.Visible = false;  // скрываем легенду
            switch (tabControl1.SelectedTab.Name)
            {
                case "tabPage_Kromonova":
                    if (filename_load_mk == "")
                        mk_saveExcel_ToolStripMenuItem.Enabled = false;     // деактивируем кнопку если нет пути к файлу
                    mk_showTable();                                         // пересчитываем метод Кромонова
                    k_getDef();                                             // загрузка файла или значения по-умолчанию с коэффициентами для N
                    break;
                case "tabPage_CB":
                    if (filename_load_mcb == "")
                        mCB_saveExcel_ToolStripMenuItem.Enabled = false;    // деактивируем кнопку если нет пути к файлу
                    cb_showTable();                                          // пересчитываем метод ЦБ
                    break;
            }
        }

        #region tabPage_Kromonova

        // загрузка файла или значения по-умолчанию с коэффициентами для N
        void k_getDef()
        {
            try
            {
                if (File.Exists(filename_k)) // Только если есть файл
                {
                    XmlSerializer reader = new XmlSerializer(typeof(ks));   // Создаем экземпляр класса XmlSerializer; указываем тип для десериализации.
                    FileStream input = File.OpenRead(filename_k);           // Открываем файл для чтения из него
                    ks _ks = reader.Deserialize(input) as ks;               // Вычитываем из файла
                    input.Close();                                          // Прекращаем работу с файлом и закрываем его

                    k_K1 = _ks.ks_K1;
                    k_K2 = _ks.ks_K2;
                    k_K3 = _ks.ks_K3;
                    k_K4 = _ks.ks_K4;
                    k_K5 = _ks.ks_K5;
                    k_K6 = _ks.ks_K6;

                    k_d_K1 = _ks.ks_d_K1;
                    k_d_K2 = _ks.ks_d_K2;
                    k_d_K3 = _ks.ks_d_K3;
                    k_d_K4 = _ks.ks_d_K4;
                    k_d_K5 = _ks.ks_d_K5;
                    k_d_K6 = _ks.ks_d_K6;
                }
            }
            catch
            {
                k_K1 = 45;
                k_K2 = 20;
                k_K3 = 10;
                k_K4 = 15;
                k_K5 = 5;
                k_K6 = 5;

                k_d_K1 = 1;
                k_d_K2 = 1;
                k_d_K3 = 1;
                k_d_K4 = 1;
                k_d_K5 = 1;
                k_d_K6 = 1;
            }

            numericUpDown_k1.Value = k_K1;
            numericUpDown_k2.Value = k_K2;
            numericUpDown_k3.Value = k_K3;
            numericUpDown_k4.Value = k_K4;
            numericUpDown_k5.Value = k_K5;
            numericUpDown_k6.Value = k_K6;

            numericUpDown_d_k1.Value = k_d_K1;
            numericUpDown_d_k2.Value = k_d_K2;
            numericUpDown_d_k3.Value = k_d_K3;
            numericUpDown_d_k4.Value = k_d_K4;
            numericUpDown_d_k5.Value = k_d_K5;
            numericUpDown_d_k6.Value = k_d_K6;
        }

        // сохранение файла с коэффициентами для N
        void k_save()
        {
            try
            {
                if (File.Exists(filename_k))   // Если есть файл удаляем его т.к. будем его заменять
                {
                    File.Delete(filename_k);   // Удаляем файл
                }

                ks _ks = new ks()
                {
                    ks_K1 = k_K1,
                    ks_K2 = k_K2,
                    ks_K3 = k_K3,
                    ks_K4 = k_K4,
                    ks_K5 = k_K5,
                    ks_K6 = k_K6,

                    ks_d_K1 = k_d_K1,
                    ks_d_K2 = k_d_K2,
                    ks_d_K3 = k_d_K3,
                    ks_d_K4 = k_d_K4,
                    ks_d_K5 = k_d_K5,
                    ks_d_K6 = k_d_K6
                };

                XmlSerializer writer = new XmlSerializer(typeof(ks));   // Создаем экземпляр класса XmlSerializer; указываем тип для десериализации.
                FileStream file = File.OpenWrite(filename_k);           // Создаем файл для записи
                writer.Serialize(file, _ks);                            // Записываем в файл данные
                file.Close();                                           // Прекращаем работу с файлом и закрываем его
            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.ToString(), "Ошибка сохранения файла с коэффициентами для N", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void numericUpDown_k1_ValueChanged(object sender, EventArgs e)
        {
            k_K1 = numericUpDown_k1.Value;  // устанавливаем коэффициент для N
            k_save();                       // сохранение файла с коэффициентами для N
            mk_showTable();                 // пересчитываем метод Кромонова
        }

        private void numericUpDown_k2_ValueChanged(object sender, EventArgs e)
        {
            k_K2 = numericUpDown_k2.Value;  // устанавливаем коэффициент для N
            k_save();                       // сохранение файла с коэффициентами для N
            mk_showTable();                 // пересчитываем метод Кромонова
        }

        private void numericUpDown_k3_ValueChanged(object sender, EventArgs e)
        {
            k_K3 = numericUpDown_k3.Value;  // устанавливаем коэффициент для N
            k_save();                       // сохранение файла с коэффициентами для N
            mk_showTable();                 // пересчитываем метод Кромонова
        }

        private void numericUpDown_k4_ValueChanged(object sender, EventArgs e)
        {
            k_K4 = numericUpDown_k4.Value;  // устанавливаем коэффициент для N
            k_save();                       // сохранение файла с коэффициентами для N
            mk_showTable();                 // пересчитываем метод Кромонова
        }

        private void numericUpDown_k5_ValueChanged(object sender, EventArgs e)
        {
            k_K5 = numericUpDown_k5.Value;  // устанавливаем коэффициент для N
            k_save();                       // сохранение файла с коэффициентами для N
            mk_showTable();                 // пересчитываем метод Кромонова
        }

        private void numericUpDown_k6_ValueChanged(object sender, EventArgs e)
        {
            k_K6 = numericUpDown_k6.Value;  // устанавливаем коэффициент для N
            k_save();                       // сохранение файла с коэффициентами для N
            mk_showTable();                 // пересчитываем метод Кромонова
        }
        
        private void numericUpDown_k1_KeyUp(object sender, KeyEventArgs e)
        {
            k_K1 = numericUpDown_k1.Value;  // устанавливаем коэффициент для N
            k_save();                       // сохранение файла с коэффициентами для N
            mk_showTable();                 // пересчитываем метод Кромонова
        }

        private void numericUpDown_k2_KeyUp(object sender, KeyEventArgs e)
        {
            k_K2 = numericUpDown_k2.Value;  // устанавливаем коэффициент для N
            k_save();                       // сохранение файла с коэффициентами для N
            mk_showTable();                 // пересчитываем метод Кромонова
        }

        private void numericUpDown_k3_KeyUp(object sender, KeyEventArgs e)
        {
            k_K3 = numericUpDown_k3.Value;  // устанавливаем коэффициент для N
            k_save();                       // сохранение файла с коэффициентами для N
            mk_showTable();                 // пересчитываем метод Кромонова
        }

        private void numericUpDown_k4_KeyUp(object sender, KeyEventArgs e)
        {
            k_K4 = numericUpDown_k4.Value;  // устанавливаем коэффициент для N
            k_save();                       // сохранение файла с коэффициентами для N
            mk_showTable();                 // пересчитываем метод Кромонова
        }

        private void numericUpDown_k5_KeyUp(object sender, KeyEventArgs e)
        {
            k_K5 = numericUpDown_k5.Value;  // устанавливаем коэффициент для N
            k_save();                       // сохранение файла с коэффициентами для N
            mk_showTable();                 // пересчитываем метод Кромонова
        }

        private void numericUpDown_k6_KeyUp(object sender, KeyEventArgs e)
        {
            k_K6 = numericUpDown_k6.Value;  // устанавливаем коэффициент для N
            k_save();                       // сохранение файла с коэффициентами для N
            mk_showTable();                 // пересчитываем метод Кромонова
        }


        private void numericUpDown_d_k1_ValueChanged(object sender, EventArgs e)
        {
            k_d_K1 = numericUpDown_d_k1.Value;  // устанавливаем коэффициент для N
            k_save();                           // сохранение файла с коэффициентами для N
            mk_showTable();                     // пересчитываем метод Кромонова
        }

        private void numericUpDown_d_k2_ValueChanged(object sender, EventArgs e)
        {
            k_d_K2 = numericUpDown_d_k2.Value;  // устанавливаем коэффициент для N
            k_save();                           // сохранение файла с коэффициентами для N
            mk_showTable();                     // пересчитываем метод Кромонова
        }

        private void numericUpDown_d_k3_ValueChanged(object sender, EventArgs e)
        {
            k_d_K3 = numericUpDown_d_k3.Value;  // устанавливаем коэффициент для N
            k_save();                           // сохранение файла с коэффициентами для N
            mk_showTable();                     // пересчитываем метод Кромонова
        }

        private void numericUpDown_d_k4_ValueChanged(object sender, EventArgs e)
        {
            k_d_K4 = numericUpDown_d_k4.Value;  // устанавливаем коэффициент для N
            k_save();                           // сохранение файла с коэффициентами для N
            mk_showTable();                     // пересчитываем метод Кромонова
        }

        private void numericUpDown_d_k5_ValueChanged(object sender, EventArgs e)
        {
            k_d_K5 = numericUpDown_d_k5.Value;  // устанавливаем коэффициент для N
            k_save();                           // сохранение файла с коэффициентами для N
            mk_showTable();                     // пересчитываем метод Кромонова
        }

        private void numericUpDown_d_k6_ValueChanged(object sender, EventArgs e)
        {
            k_d_K6 = numericUpDown_d_k6.Value;  // устанавливаем коэффициент для N
            k_save();                           // сохранение файла с коэффициентами для N
            mk_showTable();                     // пересчитываем метод Кромонова
        }

        private void numericUpDown_d_k1_KeyUp(object sender, KeyEventArgs e)
        {
            k_d_K1 = numericUpDown_d_k1.Value;  // устанавливаем коэффициент для N
            k_save();                           // сохранение файла с коэффициентами для N
            mk_showTable();                     // пересчитываем метод Кромонова
        }

        private void numericUpDown_d_k2_KeyUp(object sender, KeyEventArgs e)
        {
            k_d_K2 = numericUpDown_d_k2.Value;  // устанавливаем коэффициент для N
            k_save();                           // сохранение файла с коэффициентами для N
            mk_showTable();                     // пересчитываем метод Кромонова
        }

        private void numericUpDown_d_k3_KeyUp(object sender, KeyEventArgs e)
        {
            k_d_K3 = numericUpDown_d_k3.Value;  // устанавливаем коэффициент для N
            k_save();                           // сохранение файла с коэффициентами для N
            mk_showTable();                     // пересчитываем метод Кромонова
        }

        private void numericUpDown_d_k4_KeyUp(object sender, KeyEventArgs e)
        {
            k_d_K4 = numericUpDown_d_k4.Value;  // устанавливаем коэффициент для N
            k_save();                           // сохранение файла с коэффициентами для N
            mk_showTable();                     // пересчитываем метод Кромонова
        }

        private void numericUpDown_d_k5_KeyUp(object sender, KeyEventArgs e)
        {
            k_d_K5 = numericUpDown_d_k5.Value;  // устанавливаем коэффициент для N
            k_save();                           // сохранение файла с коэффициентами для N
            mk_showTable();                     // пересчитываем метод Кромонова
        }

        private void numericUpDown_d_k6_KeyUp(object sender, KeyEventArgs e)
        {
            k_d_K6 = numericUpDown_d_k6.Value;  // устанавливаем коэффициент для N
            k_save();                           // сохранение файла с коэффициентами для N
            mk_showTable();                     // пересчитываем метод Кромонова
        }

        // пересчет таблицы по методу Кромонова
        void mk_showTable()
        {
            DataGridViewCellStyle _DataGridViewCellStyle = new DataGridViewCellStyle(dataGridView_MCB.DefaultCellStyle);
            _DataGridViewCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;


            dataGridView_mk.Rows.Clear();
            dataGridView_mk.Columns.Clear();

            dataGridView_mk.Columns.Add(new DataGridViewTextBoxColumn()
            {
                Name = "ColumnIndex",
                HeaderText = "Показатель",
                Visible = false,
                Resizable = DataGridViewTriState.False,
                SortMode = DataGridViewColumnSortMode.NotSortable,
                Width = 300
            });
            for (int i = 0; i < list_mk.Count; i++)
            {
                dataGridView_mk.Columns.Add(new DataGridViewTextBoxColumn()
                {
                    Name = "Column_i_" + i.ToString(),
                    HeaderText = list_mk[i].name,
                    Resizable = DataGridViewTriState.False,
                    SortMode = DataGridViewColumnSortMode.NotSortable,
                    Width = 150,
                    DefaultCellStyle = _DataGridViewCellStyle
                });
            }

            dataGridView_mk.Rows.Add(15);

            dataGridView_mk.Rows[0].Cells[0].Value = "Уставной фонд";
            dataGridView_mk.Rows[0].Cells[0].ToolTipText = "УФ";
            dataGridView_mk.Rows[1].Cells[0].Value = "Собственный капитал";
            dataGridView_mk.Rows[1].Cells[0].ToolTipText = "К";
            dataGridView_mk.Rows[2].Cells[0].Value = "Обязательства до востребования";
            dataGridView_mk.Rows[2].Cells[0].ToolTipText = "ОВ";
            dataGridView_mk.Rows[3].Cells[0].Value = "Совокупность(суммарных) обязательств";
            dataGridView_mk.Rows[3].Cells[0].ToolTipText = "СО";
            dataGridView_mk.Rows[4].Cells[0].Value = "Ликвидные активы";
            dataGridView_mk.Rows[4].Cells[0].ToolTipText = "ЛА";
            dataGridView_mk.Rows[5].Cells[0].Value = "Активы работающие";
            dataGridView_mk.Rows[5].Cells[0].ToolTipText = "АР";
            dataGridView_mk.Rows[6].Cells[0].Value = "Защищенный капитал";
            dataGridView_mk.Rows[6].Cells[0].ToolTipText = "ЗК";

            dataGridView_mk.Rows[7].Cells[0].Value = "генеральный коэффициент надежности";
            dataGridView_mk.Rows[7].Cells[0].ToolTipText = "К1";
            dataGridView_mk.Rows[8].Cells[0].Value = "генеральный мгновенной надежности";
            dataGridView_mk.Rows[8].Cells[0].ToolTipText = "К2";
            dataGridView_mk.Rows[9].Cells[0].Value = "кросс-коэффициент";
            dataGridView_mk.Rows[9].Cells[0].ToolTipText = "К3";
            dataGridView_mk.Rows[10].Cells[0].Value = "генеральный коэффициент ликвидности";
            dataGridView_mk.Rows[10].Cells[0].ToolTipText = "К4";
            dataGridView_mk.Rows[11].Cells[0].Value = "коэффициент защищенности капитала";
            dataGridView_mk.Rows[11].Cells[0].ToolTipText = "К5";
            dataGridView_mk.Rows[12].Cells[0].Value = "коэффициент капитализации прибыли";
            dataGridView_mk.Rows[12].Cells[0].ToolTipText = "К6";

            dataGridView_mk.Rows[13].Cells[0].Value = "свободный коэффициент надежности";
            dataGridView_mk.Rows[13].Cells[0].ToolTipText = "N";

            dataGridView_mk.Rows[14].Cells[0].Value = "состояние";
            dataGridView_mk.Rows[14].Cells[0].ToolTipText = "";

            int i_col = 1;
            foreach (cMK _mk in list_mk)
            {
                dataGridView_mk.Rows[0].Cells[i_col].Value = _mk.YF.ToString(decimal_format).Trim();
                dataGridView_mk.Rows[1].Cells[i_col].Value = _mk.K.ToString(decimal_format).Trim();
                dataGridView_mk.Rows[2].Cells[i_col].Value = _mk.OV.ToString(decimal_format).Trim();
                dataGridView_mk.Rows[3].Cells[i_col].Value = _mk.CO.ToString(decimal_format).Trim();
                dataGridView_mk.Rows[4].Cells[i_col].Value = _mk.LA.ToString(decimal_format).Trim();
                dataGridView_mk.Rows[5].Cells[i_col].Value = _mk.AR.ToString(decimal_format).Trim();
                dataGridView_mk.Rows[6].Cells[i_col].Value = _mk.ZK.ToString(decimal_format).Trim();

                dataGridView_mk.Rows[7].Cells[i_col].Value = _mk.K1 == null ? "" : _mk.K1.Value.ToString(decimal_format).Trim();
                dataGridView_mk.Rows[8].Cells[i_col].Value = _mk.K2 == null ? "" : _mk.K2.Value.ToString(decimal_format).Trim();
                dataGridView_mk.Rows[9].Cells[i_col].Value = _mk.K3 == null ? "" : _mk.K3.Value.ToString(decimal_format).Trim();
                dataGridView_mk.Rows[10].Cells[i_col].Value = _mk.K4 == null ? "" : _mk.K4.Value.ToString(decimal_format).Trim();
                dataGridView_mk.Rows[11].Cells[i_col].Value = _mk.K5 == null ? "" : _mk.K5.Value.ToString(decimal_format).Trim();
                dataGridView_mk.Rows[12].Cells[i_col].Value = _mk.K6 == null ? "" : _mk.K6.Value.ToString(decimal_format).Trim();

                dataGridView_mk.Rows[13].Cells[i_col].Value = _mk.N == null ? "" : _mk.N.Value.ToString("##0.000").Trim();
                dataGridView_mk.Rows[14].Cells[i_col].Value = _mk.N_rez;

                i_col++;
            }

            if (list_mk.Count <= 1)
                chart_mk.Visible = false;
            else
            {
                chart_mk.Visible = true;

                chart_mk.Series.Clear();
                var series1 = new Series
                {
                    Name = "Series1",
                    Color = Color.Green,
                    IsVisibleInLegend = false,
                    IsXValueIndexed = true,
                    ChartType = SeriesChartType.Column
                };

                chart_mk.Series.Add(series1);

                series1.Points.Clear();
                foreach (cMK _mk in list_mk)
                {
                    series1.Points.AddXY(_mk.name, _mk.N);
                }
                chart_mk.Invalidate();
            }
        }

        // добавляем метод Кромонова
        public void Add_MK(string _name, decimal _YF, decimal _K, decimal _OV, decimal _CO, decimal _LA, decimal _AR, decimal _ZK)
        {
            list_mk.Add(new cMK()
            {
                name = _name,
                YF = _YF,
                K = _K,
                OV = _OV,
                CO = _CO,
                LA = _LA,
                AR = _AR,
                ZK = _ZK
            });

            mk_showTable();
        }

        // условные обозначения в таблице
        private void dataGridView_mk_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            using (SolidBrush b = new SolidBrush(dataGridView_mk.RowHeadersDefaultCellStyle.ForeColor))
            {
                e.Graphics.DrawString(dataGridView_mk.Rows[e.RowIndex].Cells[0].ToolTipText
                    , dataGridView_mk.RowHeadersDefaultCellStyle.Font
                    , b
                    , e.RowBounds.Location.X + 12
                    , e.RowBounds.Location.Y + 4);
            }
        }

        // Кнопка Добавить по методу Кромонова
        private void mK_input_ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (addMK == null)
                addMK = new FormAddKromonova(this);
            addMK.Show();
            addMK.WindowState = FormWindowState.Normal;
            addMK.Focus();
        }

        // Кнопка Сохранить в Excel
        private void mk_saveExcel_ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (filename_load_mk != "" && MessageBox.Show("Хотите перезаписать файл \"" + filename_load_mk + "\" ?", "Вопрос", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                SaveExcel_mk(filename_load_mk);
        }

        // Кнопка Сохранить как...
        private void mK_saveAs_ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //Создаем диалоговое окно длля сохранения
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.CheckPathExists = true;
            saveFileDialog1.Filter = "Excel (*.xlsx)|*.xlsx|Excel (*.xls)|*.xls";
            saveFileDialog1.FilterIndex = 1;
            saveFileDialog1.RestoreDirectory = true;

            // Вызываем диалоговое окно для сохранения файла Excel
            if (saveFileDialog1.ShowDialog() != DialogResult.Cancel)
                SaveExcel_mk(saveFileDialog1.FileName);
        }

        // Кнопка загружаем из Excel метод Кромонова
        private void mK_load_Excel_ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.CheckPathExists = true;
            openFileDialog1.Filter = "Excel (*.xlsx)|*.xlsx|Excel (*.xls)|*.xls";
            openFileDialog1.RestoreDirectory = true;

            if (openFileDialog1.ShowDialog() != DialogResult.Cancel)
            {
                try
                {
                    FileInfo fi = new FileInfo(openFileDialog1.FileName);

                    ExcelEngine excelEngine = new ExcelEngine();
                    IApplication _excel = excelEngine.Excel;

                    // Тип файла 
                    _excel.DefaultVersion = ExcelVersion.Excel97to2003; // .xls
                    if (fi.Extension == ".xlsx")
                        _excel.DefaultVersion = ExcelVersion.Excel2010; // .xlsx

                    IWorkbook wBook = _excel.Workbooks.Open(openFileDialog1.FileName);
                    IWorksheet wSheet = wBook.Worksheets[0];

                    list_mk.Clear();

                    for (int i = 2; i <= wSheet.Columns.Count(); i++)
                    {
                        list_mk.Add(new cMK()
                        {
                            name = wSheet.GetText(1, i),
                            YF = (decimal)wSheet.GetNumber(2, i),
                            K = (decimal)wSheet.GetNumber(3, i),
                            OV = (decimal)wSheet.GetNumber(4, i),
                            CO = (decimal)wSheet.GetNumber(5, i),
                            LA = (decimal)wSheet.GetNumber(6, i),
                            AR = (decimal)wSheet.GetNumber(7, i),
                            ZK = (decimal)wSheet.GetNumber(8, i)
                        });
                    }

                    filename_load_mk = openFileDialog1.FileName;
                }
                catch (Exception ee)
                {
                    MessageBox.Show(ee.ToString(), "Ошибка открытия файла Excel", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }


                if (filename_load_mk != "")
                    mk_saveExcel_ToolStripMenuItem.Enabled = true;
                mk_showTable();
            }
        }

        // сохраняем в Excel метод Кромонова
        private void SaveExcel_mk(string PathToSave_file)
        {
            try
            {
                FileInfo fi = new FileInfo(PathToSave_file);

                ExcelEngine excelEngine = new ExcelEngine();
                IApplication _excel = excelEngine.Excel;

                // Тип файла 
                _excel.DefaultVersion = ExcelVersion.Excel97to2003; // .xls
                if (fi.Extension == ".xlsx")
                    _excel.DefaultVersion = ExcelVersion.Excel2010; // .xlsx

                // Создаем книгу
                IWorkbook wBook = _excel.Workbooks.Create(1);
                // Создаем лист
                IWorksheet wSheet = wBook.Worksheets[0];

                IStyle style = wBook.Styles.Add("FillColor");
                style.ColorIndex = ExcelKnownColors.Grey_25_percent;
                // заполняем заголовок (1-я строка)
                for (int i = 0; i < dataGridView_mk.Columns.Count; i++)
                {
                    wSheet.Range[1, i + 1].Text = dataGridView_mk.Columns[i].HeaderText;
                    wSheet.Range[1, i + 1].CellStyle = style;
                    wSheet.Range[1, i + 1].BorderAround(ExcelLineStyle.Thin, Color.Black);
                }

                IStyle style2 = wBook.Styles.Add("FillColor2");
                style2.ColorIndex = ExcelKnownColors.Light_yellow;
                // Заполняем данными (начинаем заполянть со 2-й строки)
                for (int i = 0; i < dataGridView_mk.Rows.Count; i++)
                {
                    for (int j = 0; j < dataGridView_mk.Columns.Count; j++)
                    {
                        wSheet.Range[i + 2, j + 1].Value = dataGridView_mk.Rows[i].Cells[j].Value.ToString() 
                            + (j == 0 && !string.IsNullOrEmpty(dataGridView_mk.Rows[i].Cells[j].ToolTipText) ? (" (" + dataGridView_mk.Rows[i].Cells[j].ToolTipText + ")") : "");

                        if (i >= 7)
                            wSheet.Range[i + 2, j + 1].CellStyle = style2;
                        wSheet.Range[i + 2, j + 1].BorderAround(ExcelLineStyle.Thin, Color.Black);
                    }
                }

                for (int i = 0; i < dataGridView_mk.Columns.Count; i++)
                {
                    wSheet.AutofitColumn(i + 1);
                }

                // Сохраняем файл Excel
                _excel.Save(PathToSave_file);

                filename_load_mk = PathToSave_file;
            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.ToString(), "Ошибка сохранения файла Excel", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            if (filename_load_mk != "")
                mk_saveExcel_ToolStripMenuItem.Enabled = true;
        }

        private void button_legend_mk_Click(object sender, EventArgs e)
        {
            legendMK.Visible = !legendMK.Visible;   // показываем/скрываем легенду
            posLegends();                           // пересчет положения форм легенды
        }

        #endregion

        #region tabPage_CB

        // пересчет таблицы по методу ЦБ
        void cb_showTable()
        {
            DataGridViewCellStyle _DataGridViewCellStyle = new DataGridViewCellStyle(dataGridView_MCB.DefaultCellStyle);
            _DataGridViewCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;


            dataGridView_MCB.Rows.Clear();
            dataGridView_MCB.Columns.Clear();

            dataGridView_MCB.Columns.Add(new DataGridViewTextBoxColumn()
            {
                Name = "ColumnCBIndex",
                HeaderText = "Показатель",
                Visible = false,
                Resizable = DataGridViewTriState.False,
                SortMode = DataGridViewColumnSortMode.NotSortable,
                Width = 300
            });
            for (int i = 0; i < list_cb.Count; i++)
            {
                dataGridView_MCB.Columns.Add(new DataGridViewTextBoxColumn()
                {
                    Name = "ColumnCB_i_" + i.ToString(),
                    HeaderText = list_cb[i].name,
                    Resizable = DataGridViewTriState.False,
                    SortMode = DataGridViewColumnSortMode.NotSortable,
                    Width = 150,
                    DefaultCellStyle = _DataGridViewCellStyle
                });
            }

            dataGridView_MCB.Rows.Add(22);

            dataGridView_MCB.Rows[0].Cells[0].ToolTipText = "K";
            dataGridView_MCB.Rows[0].Cells[0].Value = "Собственный капитал";
            dataGridView_MCB.Rows[1].Cells[0].ToolTipText = "Ар";
            dataGridView_MCB.Rows[1].Cells[0].Value = "Активы, взвешенные с учетом риска";
            dataGridView_MCB.Rows[2].Cells[0].ToolTipText = "Лат";
            dataGridView_MCB.Rows[2].Cells[0].Value = "Ликвидные активы";
            dataGridView_MCB.Rows[3].Cells[0].ToolTipText = "Овт";
            dataGridView_MCB.Rows[3].Cells[0].Value = "Обязательства до востребования";
            dataGridView_MCB.Rows[4].Cells[0].ToolTipText = "Лам";
            dataGridView_MCB.Rows[4].Cells[0].Value = "Ликвидные активы";
            dataGridView_MCB.Rows[5].Cells[0].ToolTipText = "Овм";
            dataGridView_MCB.Rows[5].Cells[0].Value = "Обязательства до востребования";
            dataGridView_MCB.Rows[6].Cells[0].ToolTipText = "Крд";
            dataGridView_MCB.Rows[6].Cells[0].Value = "Сумма выходных кредитов";
            dataGridView_MCB.Rows[7].Cells[0].ToolTipText = "Об";
            dataGridView_MCB.Rows[7].Cells[0].Value = "Суммарные обязательства";
            dataGridView_MCB.Rows[8].Cells[0].ToolTipText = "А";
            dataGridView_MCB.Rows[8].Cells[0].Value = "Нетто-активы";
            dataGridView_MCB.Rows[9].Cells[0].ToolTipText = "КрКр";
            dataGridView_MCB.Rows[9].Cells[0].Value = "Сумма крупных кредитов";
            dataGridView_MCB.Rows[10].Cells[0].ToolTipText = "Вкл";
            dataGridView_MCB.Rows[10].Cells[0].Value = "Сумма вкладов населения";
            dataGridView_MCB.Rows[11].Cells[0].ToolTipText = "Инв";
            dataGridView_MCB.Rows[11].Cells[0].Value = "Инвестиции в акции других юредических лиц";
            dataGridView_MCB.Rows[12].Cells[0].ToolTipText = "Векс";
            dataGridView_MCB.Rows[12].Cells[0].Value = "Собственные вексельные обязательства";

            dataGridView_MCB.Rows[13].Cells[0].ToolTipText = "Н1";
            dataGridView_MCB.Rows[13].Cells[0].Value = "достаточность капитала";
            dataGridView_MCB.Rows[14].Cells[0].ToolTipText = "Н2";
            dataGridView_MCB.Rows[14].Cells[0].Value = "текущая ликвидность";
            dataGridView_MCB.Rows[15].Cells[0].ToolTipText = "Н3";
            dataGridView_MCB.Rows[15].Cells[0].Value = "мгновенная ликвидность";
            dataGridView_MCB.Rows[16].Cells[0].ToolTipText = "Н4";
            dataGridView_MCB.Rows[16].Cells[0].Value = "долгосрочная ликвидность";
            dataGridView_MCB.Rows[17].Cells[0].ToolTipText = "Н5";
            dataGridView_MCB.Rows[17].Cells[0].Value = "доля ликвидных активов";
            dataGridView_MCB.Rows[18].Cells[0].ToolTipText = "Н7";
            dataGridView_MCB.Rows[18].Cells[0].Value = "размер крупных кредитных рисков";
            dataGridView_MCB.Rows[19].Cells[0].ToolTipText = "Н11";
            dataGridView_MCB.Rows[19].Cells[0].Value = "размер вкладов населения";
            dataGridView_MCB.Rows[20].Cells[0].ToolTipText = "Н12";
            dataGridView_MCB.Rows[20].Cells[0].Value = "акции других юредических лиц";
            dataGridView_MCB.Rows[21].Cells[0].ToolTipText = "Н13";
            dataGridView_MCB.Rows[21].Cells[0].Value = "риск собственных вексельных обязательств";

            int i_col = 1;
            foreach (cCB _cb in list_cb)
            {
                dataGridView_MCB.Rows[0].Cells[i_col].Value = _cb.K.ToString(decimal_format).Trim();
                dataGridView_MCB.Rows[1].Cells[i_col].Value = _cb.Ar.ToString(decimal_format).Trim();
                dataGridView_MCB.Rows[2].Cells[i_col].Value = _cb.Lat.ToString(decimal_format).Trim();
                dataGridView_MCB.Rows[3].Cells[i_col].Value = _cb.Ovt.ToString(decimal_format).Trim();
                dataGridView_MCB.Rows[4].Cells[i_col].Value = _cb.Lam.ToString(decimal_format).Trim();
                dataGridView_MCB.Rows[5].Cells[i_col].Value = _cb.Ovm.ToString(decimal_format).Trim();
                dataGridView_MCB.Rows[6].Cells[i_col].Value = _cb.Krd.ToString(decimal_format).Trim();
                dataGridView_MCB.Rows[7].Cells[i_col].Value = _cb.Ob.ToString(decimal_format).Trim();
                dataGridView_MCB.Rows[8].Cells[i_col].Value = _cb.A.ToString(decimal_format).Trim();
                dataGridView_MCB.Rows[9].Cells[i_col].Value = _cb.KrKr.ToString(decimal_format).Trim();
                dataGridView_MCB.Rows[10].Cells[i_col].Value = _cb.Vkl.ToString(decimal_format).Trim();
                dataGridView_MCB.Rows[11].Cells[i_col].Value = _cb.Inv.ToString(decimal_format).Trim();
                dataGridView_MCB.Rows[12].Cells[i_col].Value = _cb.Veks.ToString(decimal_format).Trim();

                dataGridView_MCB.Rows[13].Cells[i_col].Value = _cb.H1 == null ? "" : _cb.H1.Value.ToString(decimal_format).Trim();
                dataGridView_MCB.Rows[14].Cells[i_col].Value = _cb.H2 == null ? "" : _cb.H2.Value.ToString(decimal_format).Trim();
                dataGridView_MCB.Rows[15].Cells[i_col].Value = _cb.H3 == null ? "" : _cb.H3.Value.ToString(decimal_format).Trim();
                dataGridView_MCB.Rows[16].Cells[i_col].Value = _cb.H4 == null ? "" : _cb.H4.Value.ToString(decimal_format).Trim();
                dataGridView_MCB.Rows[17].Cells[i_col].Value = _cb.H5 == null ? "" : _cb.H5.Value.ToString(decimal_format).Trim();
                dataGridView_MCB.Rows[18].Cells[i_col].Value = _cb.H7 == null ? "" : _cb.H7.Value.ToString(decimal_format).Trim();
                dataGridView_MCB.Rows[19].Cells[i_col].Value = _cb.H11 == null ? "" : _cb.H11.Value.ToString(decimal_format).Trim();
                dataGridView_MCB.Rows[20].Cells[i_col].Value = _cb.H12 == null ? "" : _cb.H12.Value.ToString(decimal_format).Trim();
                dataGridView_MCB.Rows[21].Cells[i_col].Value = _cb.H13 == null ? "" : _cb.H13.Value.ToString(decimal_format).Trim();

                i_col++;
            }
        }

        // добавляем метод ЦБ
        public void Add_CB(string _name, decimal _K, decimal _Ar, decimal _Lat, decimal _Ovt, decimal _Lam, decimal _Ovm, decimal _Krd, decimal _Ob, decimal _A, decimal _KrKr, decimal _Vkl, decimal _Inv, decimal _Veks)
        {
            list_cb.Add(new cCB()
            {
                name = _name,
                K = _K,
                Ar = _Ar,
                Lat = _Lat,
                Ovt = _Ovt,
                Lam = _Lam,
                Ovm = _Ovm,
                Krd = _Krd,
                Ob = _Ob,
                A = _A,
                KrKr = _KrKr,
                Vkl = _Vkl,
                Inv = _Inv,
                Veks = _Veks
            });

            cb_showTable();
        }

        // условные обозначения в таблице
        private void dataGridView_MCB_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            using (SolidBrush b = new SolidBrush(dataGridView_MCB.RowHeadersDefaultCellStyle.ForeColor))
            {
                e.Graphics.DrawString(dataGridView_MCB.Rows[e.RowIndex].Cells[0].ToolTipText
                    , dataGridView_MCB.RowHeadersDefaultCellStyle.Font
                    , b
                    , e.RowBounds.Location.X + 12
                    , e.RowBounds.Location.Y + 4);
            }
        }

        // Кнопка Добавить по методу ЦБ
        private void mCB_input_ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (addMCB == null)
                addMCB = new FormAddCB(this);
            addMCB.Show();
            addMCB.WindowState = FormWindowState.Normal;
            addMCB.Focus();
        }

        // Кнопка Сохранить в Excel
        private void mCB_saveExcel_ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (filename_load_mcb != "" && MessageBox.Show("Хотите перезаписать файл \"" + filename_load_mcb + "\" ?", "Вопрос", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                SaveExcel_mcb(filename_load_mcb);
        }

        // Кнопка Сохранить как...
        private void mCB_saveAs_ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //Создаем диалоговое окно длля сохранения
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.CheckPathExists = true;
            saveFileDialog1.Filter = "Excel (*.xlsx)|*.xlsx|Excel (*.xls)|*.xls";
            saveFileDialog1.FilterIndex = 1;
            saveFileDialog1.RestoreDirectory = true;

            // Вызываем диалоговое окно для сохранения файла Excel
            if (saveFileDialog1.ShowDialog() != DialogResult.Cancel)
                SaveExcel_mcb(saveFileDialog1.FileName);
        }

        // Кнопка загружаем из Excel метод ЦБ
        private void mCB_loadExcel_ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.CheckPathExists = true;
            openFileDialog1.Filter = "Excel (*.xlsx)|*.xlsx|Excel (*.xls)|*.xls";
            openFileDialog1.RestoreDirectory = true;

            if (openFileDialog1.ShowDialog() != DialogResult.Cancel)
            {
                try
                {
                    FileInfo fi = new FileInfo(openFileDialog1.FileName);

                    ExcelEngine excelEngine = new ExcelEngine();
                    IApplication _excel = excelEngine.Excel;

                    // Тип файла 
                    _excel.DefaultVersion = ExcelVersion.Excel97to2003; // .xls
                    if (fi.Extension == ".xlsx")
                        _excel.DefaultVersion = ExcelVersion.Excel2010; // .xlsx

                    IWorkbook wBook = _excel.Workbooks.Open(openFileDialog1.FileName);
                    IWorksheet wSheet = wBook.Worksheets[0];

                    list_cb.Clear();

                    for (int i = 2; i <= wSheet.Columns.Count(); i++)
                    {
                        list_cb.Add(new cCB()
                        {
                            name = wSheet.GetText(1, i),
                            K = (decimal)wSheet.GetNumber(2, i),
                            Ar = (decimal)wSheet.GetNumber(3, i),
                            Lat = (decimal)wSheet.GetNumber(4, i),
                            Ovt = (decimal)wSheet.GetNumber(5, i),
                            Lam = (decimal)wSheet.GetNumber(6, i),
                            Ovm = (decimal)wSheet.GetNumber(7, i),
                            Krd = (decimal)wSheet.GetNumber(8, i),
                            Ob = (decimal)wSheet.GetNumber(9, i),
                            A = (decimal)wSheet.GetNumber(10, i),
                            KrKr = (decimal)wSheet.GetNumber(11, i),
                            Vkl = (decimal)wSheet.GetNumber(12, i),
                            Inv = (decimal)wSheet.GetNumber(13, i),
                            Veks = (decimal)wSheet.GetNumber(14, i),
                        });
                    }

                    filename_load_mcb = openFileDialog1.FileName;
                }
                catch (Exception ee)
                {
                    MessageBox.Show(ee.ToString(), "Ошибка открытия файла Excel", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }


                if (filename_load_mcb != "")
                    mCB_saveExcel_ToolStripMenuItem.Enabled = true;
                cb_showTable();
            }
        }

        // сохраняем в Excel метод ЦБ
        private void SaveExcel_mcb(string PathToSave_file)
        {
            try
            {
                FileInfo fi = new FileInfo(PathToSave_file);

                ExcelEngine excelEngine = new ExcelEngine();
                IApplication _excel = excelEngine.Excel;

                // Тип файла 
                _excel.DefaultVersion = ExcelVersion.Excel97to2003; // .xls
                if (fi.Extension == ".xlsx")
                    _excel.DefaultVersion = ExcelVersion.Excel2010; // .xlsx

                // Создаем книгу
                IWorkbook wBook = _excel.Workbooks.Create(1);
                // Создаем лист
                IWorksheet wSheet = wBook.Worksheets[0];

                IStyle style = wBook.Styles.Add("FillColor");
                style.ColorIndex = ExcelKnownColors.Grey_25_percent;
                // заполняем заголовок (1-я строка)
                for (int i = 0; i < dataGridView_MCB.Columns.Count; i++)
                {
                    wSheet.Range[1, i + 1].Text = dataGridView_MCB.Columns[i].HeaderText;
                    wSheet.Range[1, i + 1].CellStyle = style;
                    wSheet.Range[1, i + 1].BorderAround(ExcelLineStyle.Thin, Color.Black);
                }

                IStyle style2 = wBook.Styles.Add("FillColor2");
                style2.ColorIndex = ExcelKnownColors.Light_yellow;
                // Заполняем данными (начинаем заполянть со 2-й строки)
                for (int i = 0; i < dataGridView_MCB.Rows.Count; i++)
                {
                    for (int j = 0; j < dataGridView_MCB.Columns.Count; j++)
                    {
                        wSheet.Range[i + 2, j + 1].Value = dataGridView_MCB.Rows[i].Cells[j].Value.ToString() + (j == 0 ? (" (" + dataGridView_MCB.Rows[i].Cells[j].ToolTipText + ")") : "");

                        if (i >= 13)
                            wSheet.Range[i + 2, j + 1].CellStyle = style2;
                        wSheet.Range[i + 2, j + 1].BorderAround(ExcelLineStyle.Thin, Color.Black);
                    }
                }

                for (int i = 0; i < dataGridView_MCB.Columns.Count; i++)
                {
                    wSheet.AutofitColumn(i + 1);
                }

                // Сохраняем файл Excel
                _excel.Save(PathToSave_file);

                filename_load_mcb = PathToSave_file;
            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.ToString(), "Ошибка сохранения файла Excel", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            if (filename_load_mcb != "")
                mCB_saveExcel_ToolStripMenuItem.Enabled = true;
        }

        private void buttonLegendMCB_Click(object sender, EventArgs e)
        {
            legendMCB.Visible = !legendMCB.Visible; // показываем/скрываем легенду
            posLegends();                           // пересчет положения форм легенды
        }

        #endregion

        private void FormMethodology_FormClosing(object sender, FormClosingEventArgs e)
        {
            legendMK.Close();   // закрываем форму легенды
            legendMCB.Close();  // закрываем форму легенды

            if (addMK != null)
                addMK.Close();  // закрываем форму добавления

            if (addMCB != null)
                addMCB.Close(); // закрываем форму добавления

            if (_FormMain != null)
                _FormMain.Close_Methodology();  // объявляем что та форма закрывается
        }

        // пересчет положения форм легенды
        void posLegends()
        {
            legendMK.Location = new Point(this.Location.X - legendMK.Size.Width + 15, this.Location.Y + 90);    // пересчитываем положение легенды
            legendMCB.Location = new Point(this.Location.X - legendMCB.Size.Width + 15, this.Location.Y + 86);  // пересчитываем положение легенды
        }

        private void FormMethodology_Move(object sender, EventArgs e)
        {
            posLegends();   // пересчет положения форм легенды
        }

        public void Close_addMK()
        {
            addMK = null;   // когда закрыли форму
        }

        public void Close_addMCB()
        {
            addMCB = null;  // когда закрыли форму
        }

        private void dataGridView_mk_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}
