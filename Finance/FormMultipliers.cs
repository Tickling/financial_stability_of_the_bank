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
    public partial class FormMultipliers : Form
    {
        // Метод ЦБ
        public class mul
        {
            public string name;

            public decimal P_E;
            public decimal E_P;
            public decimal P_S;
            public decimal P_BV;
            public decimal P_CF;
            public decimal CF_P;
            public decimal P_FCE;
            public decimal FCE_P;
        }

        List<mul> list = new List<mul>();               // список
        string filename_load = "";                      // путь к отрытому/сохраненному файлу
        FormLegendM legendM = new FormLegendM();      // форма легенды
        FormAddMultipliers addM = null;                   // форма добавления

        static string decimal_format = "##0.000";  // формат числа
        FormMain _FormMain = null;                      // форма родителя

        public FormMultipliers(FormMain f)
        {
            InitializeComponent();

            _FormMain = f;
        }

        private void FormMultipliers_Load(object sender, EventArgs e)
        {
            toolStripComboBox_m.Items.Add("P/E {Цена / Прибыль}");
            toolStripComboBox_m.Items.Add("E/P {Прибыль / цена}");
            toolStripComboBox_m.Items.Add("P/S {Цена / Выручка}");
            toolStripComboBox_m.Items.Add("P/BV {Цена / Собственный капитал}");
            toolStripComboBox_m.Items.Add("P/CF {Цена / Денежный поток}");
            toolStripComboBox_m.Items.Add("CF/P {Денежный поток / Цена}");
            toolStripComboBox_m.Items.Add("P/FCE {Цена / Свободный денежный поток}");
            toolStripComboBox_m.Items.Add("FCE/P {Свободный денежный поток / Цена}");

            toolStripComboBox_m.SelectedIndex = 0;

            if (filename_load == "")
                mk_saveExcel_ToolStripMenuItem.Enabled = false; // деактивируем кнопку если нет пути к файлу
            showTable();                                        //  пересчет таблицы

            legendM.Show();            // показываем форму легенды
            legendM.Visible = false;   // скрываем легенду
        }

        // Кнопка Сохранить в Excel
        private void mk_saveExcel_ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (filename_load != "" && MessageBox.Show("Хотите перезаписать файл \"" + filename_load + "\" ?", "Вопрос", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                SaveExcel_m(filename_load);
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
                SaveExcel_m(saveFileDialog1.FileName);
        }

        // Кнопка загружаем из Excel 
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

                    list.Clear();

                    for (int i = 2; i <= wSheet.Columns.Count(); i++)
                    {
                        list.Add(new mul()
                        {
                            name = string.IsNullOrEmpty(wSheet.GetValueRowCol(1, i).ToString().Replace("-", "").Trim()) ? (i - 1).ToString() : wSheet.GetValueRowCol(1, i).ToString().Replace("-", "").Trim(),
                            P_E = string.IsNullOrEmpty(wSheet.GetValueRowCol(2, i).ToString().Replace("-", "").Trim()) ? 0 : (decimal)wSheet.GetNumber(2, i),
                            E_P = string.IsNullOrEmpty(wSheet.GetValueRowCol(3, i).ToString().Replace("-", "").Trim()) ? 0 : (decimal)wSheet.GetNumber(3, i),
                            P_S = string.IsNullOrEmpty(wSheet.GetValueRowCol(4, i).ToString().Replace("-", "").Trim()) ? 0 : (decimal)wSheet.GetNumber(4, i),
                            P_BV = string.IsNullOrEmpty(wSheet.GetValueRowCol(5, i).ToString().Replace("-", "").Trim()) ? 0 : (decimal)wSheet.GetNumber(5, i),
                            P_CF = string.IsNullOrEmpty(wSheet.GetValueRowCol(6, i).ToString().Replace("-", "").Trim()) ? 0 : (decimal)wSheet.GetNumber(6, i),
                            CF_P = string.IsNullOrEmpty(wSheet.GetValueRowCol(7, i).ToString().Replace("-", "").Trim()) ? 0 : (decimal)wSheet.GetNumber(7, i),
                            P_FCE = string.IsNullOrEmpty(wSheet.GetValueRowCol(8, i).ToString().Replace("-", "").Trim()) ? 0 : (decimal)wSheet.GetNumber(8, i),
                            FCE_P = string.IsNullOrEmpty(wSheet.GetValueRowCol(9, i).ToString().Replace("-", "").Trim()) ? 0 : (decimal)wSheet.GetNumber(9, i)
                        });
                    }

                    filename_load = openFileDialog1.FileName;
                }
                catch (Exception ee)
                {
                    MessageBox.Show(ee.ToString(), "Ошибка открытия файла Excel", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }


                if (filename_load != "")
                    mk_saveExcel_ToolStripMenuItem.Enabled = true;
                showTable();
            }
        }

        // Кнопка Добавить 
        private void mK_input_ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (addM == null)
                addM = new FormAddMultipliers(this);
            addM.Show();
            addM.WindowState = FormWindowState.Normal;
            addM.Focus();
        }

        // пересчет таблицы
        void showTable()
        {
            DataGridViewCellStyle _DataGridViewCellStyle = new DataGridViewCellStyle(dataGridView_m.DefaultCellStyle);
            _DataGridViewCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;


            dataGridView_m.Rows.Clear();
            dataGridView_m.Columns.Clear();

            dataGridView_m.Columns.Add(new DataGridViewTextBoxColumn()
            {
                Name = "ColumnIndex",
                HeaderText = "Показатель",
                Visible = false,
                Resizable = DataGridViewTriState.False,
                SortMode = DataGridViewColumnSortMode.NotSortable,
                Width = 300
            });
            for (int i = 0; i < list.Count; i++)
            {
                dataGridView_m.Columns.Add(new DataGridViewTextBoxColumn()
                {
                    Name = "Column_i_" + i.ToString(),
                    HeaderText = list[i].name,
                    Resizable = DataGridViewTriState.False,
                    SortMode = DataGridViewColumnSortMode.NotSortable,
                    Width = 150,
                    DefaultCellStyle = _DataGridViewCellStyle
                });
            }

            dataGridView_m.Rows.Add(8);

            dataGridView_m.Rows[0].Cells[0].Value = "{Цена / Прибыль}";
            dataGridView_m.Rows[0].Cells[0].ToolTipText = "P/E";
            dataGridView_m.Rows[1].Cells[0].Value = "{Прибыль / цена}";
            dataGridView_m.Rows[1].Cells[0].ToolTipText = "E/P";
            dataGridView_m.Rows[2].Cells[0].Value = "{Цена / Выручка}";
            dataGridView_m.Rows[2].Cells[0].ToolTipText = "P/S";
            dataGridView_m.Rows[3].Cells[0].Value = "{Цена / Собственный капитал}";
            dataGridView_m.Rows[3].Cells[0].ToolTipText = "P/BV";
            dataGridView_m.Rows[4].Cells[0].Value = "{Цена / Денежный поток}";
            dataGridView_m.Rows[4].Cells[0].ToolTipText = "P/CF";
            dataGridView_m.Rows[5].Cells[0].Value = "{Денежный поток / Цена}";
            dataGridView_m.Rows[5].Cells[0].ToolTipText = "CF/P";
            dataGridView_m.Rows[6].Cells[0].Value = "{Цена / Свободный денежный поток}";
            dataGridView_m.Rows[6].Cells[0].ToolTipText = "P/FCE";
            dataGridView_m.Rows[7].Cells[0].Value = "{Свободный денежный поток / Цена}";
            dataGridView_m.Rows[7].Cells[0].ToolTipText = "FCE/P";

            int i_col = 1;
            foreach (mul _m in list)
            {
                dataGridView_m.Rows[0].Cells[i_col].Value = _m.P_E.ToString(decimal_format).Trim();
                dataGridView_m.Rows[1].Cells[i_col].Value = _m.E_P.ToString(decimal_format).Trim();
                dataGridView_m.Rows[2].Cells[i_col].Value = _m.P_S.ToString(decimal_format).Trim();
                dataGridView_m.Rows[3].Cells[i_col].Value = _m.P_BV.ToString(decimal_format).Trim();
                dataGridView_m.Rows[4].Cells[i_col].Value = _m.P_CF.ToString(decimal_format).Trim();
                dataGridView_m.Rows[5].Cells[i_col].Value = _m.CF_P.ToString(decimal_format).Trim();
                dataGridView_m.Rows[6].Cells[i_col].Value = _m.P_FCE.ToString(decimal_format).Trim();
                dataGridView_m.Rows[7].Cells[i_col].Value = _m.FCE_P.ToString(decimal_format).Trim();

                i_col++;
            }

            if (list.Count <= 1)
                chart_m.Visible = false;
            else
            {
                chart_m.Visible = true;

                chart_m.Series.Clear();
                var series1 = new Series
                {
                    Name = "Series1",
                    Color = Color.Green,
                    IsVisibleInLegend = false,
                    IsXValueIndexed = true,
                    ChartType = SeriesChartType.Column
                };

                chart_m.Series.Add(series1);

                series1.Points.Clear();

                foreach (mul _m in list)
                {
                    switch (toolStripComboBox_m.SelectedIndex)
                    {
                        case 0:
                            series1.Points.AddXY(_m.name, _m.P_E);
                            break;
                        case 1:
                            series1.Points.AddXY(_m.name, _m.E_P);
                            break;
                        case 2:
                            series1.Points.AddXY(_m.name, _m.P_S);
                            break;
                        case 3:
                            series1.Points.AddXY(_m.name, _m.P_BV);
                            break;
                        case 4:
                            series1.Points.AddXY(_m.name, _m.P_CF);
                            break;
                        case 5:
                            series1.Points.AddXY(_m.name, _m.CF_P);
                            break;
                        case 6:
                            series1.Points.AddXY(_m.name, _m.P_FCE);
                            break;
                        case 7:
                            series1.Points.AddXY(_m.name, _m.FCE_P);
                            break;
                    }
                }
                chart_m.Invalidate();
            }
        }

        // добавляем метод Кромонова
        public void Add_(string _name, decimal _P_E, decimal _E_P, decimal _P_S, decimal _P_BV, decimal _P_CF, decimal _CF_P, decimal _P_FCE, decimal _FCE_P)
        {
            list.Add(new mul()
            {
                name = _name,
                P_E = _P_E,
                E_P = _E_P,
                P_S = _P_S,
                P_BV = _P_BV,
                P_CF = _P_CF,
                CF_P = _CF_P,
                P_FCE = _P_FCE,
                FCE_P = _FCE_P
            });

            showTable();
        }

        // пересчет положения форм легенды
        void posLegends()
        {
            legendM.Location = new Point(this.Location.X - legendM.Size.Width + 7, this.Location.Y + 68);    // пересчитываем положение легенды
        }

        // сохраняем в Excel
        private void SaveExcel_m(string PathToSave_file)
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
                for (int i = 0; i < dataGridView_m.Columns.Count; i++)
                {
                    wSheet.Range[1, i + 1].Text = dataGridView_m.Columns[i].HeaderText;
                    wSheet.Range[1, i + 1].CellStyle = style;
                    wSheet.Range[1, i + 1].BorderAround(ExcelLineStyle.Thin, Color.Black);
                }

                IStyle style2 = wBook.Styles.Add("FillColor2");
                style2.ColorIndex = ExcelKnownColors.Light_yellow;
                // Заполняем данными (начинаем заполянть со 2-й строки)
                for (int i = 0; i < dataGridView_m.Rows.Count; i++)
                {
                    for (int j = 0; j < dataGridView_m.Columns.Count; j++)
                    {
                        if (j == 0)
                            wSheet.Range[i + 2, j + 1].Value = dataGridView_m.Rows[i].Cells[j].ToolTipText
                                + (!string.IsNullOrEmpty(dataGridView_m.Rows[i].Cells[j].Value.ToString()) ? dataGridView_m.Rows[i].Cells[j].Value.ToString() : "");
                        else
                            wSheet.Range[i + 2, j + 1].Value = dataGridView_m.Rows[i].Cells[j].Value.ToString();

                        wSheet.Range[i + 2, j + 1].BorderAround(ExcelLineStyle.Thin, Color.Black);
                    }
                }

                // выравниваем колонки
                for (int i = 0; i < dataGridView_m.Columns.Count; i++)
                {
                    wSheet.AutofitColumn(i + 1);
                }

                // создаем график
                IChartShape chart = wSheet.Charts.Add();
                chart.DataRange = wSheet.Range[1, 1, 2, dataGridView_m.Columns.Count];
                chart.ChartType = ExcelChartType.Column_Clustered;
                chart.IsSeriesInRows = true;

                chart.TopRow = dataGridView_m.Rows.Count + 3;
                chart.LeftColumn = 1;
                chart.RightColumn = 20;
                chart.BottomRow = dataGridView_m.Rows.Count + 30;

                chart.ChartTitle = "";

                // Сохраняем файл Excel
                _excel.Save(PathToSave_file);

                filename_load = PathToSave_file;
            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.ToString(), "Ошибка сохранения файла Excel", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            if (filename_load != "")
                mk_saveExcel_ToolStripMenuItem.Enabled = true;
        }

        public void Close_addM()
        {
            addM = null;   // когда закрыли форму
        }

        private void button_legend_m_Click(object sender, EventArgs e)
        {
            legendM.Visible = !legendM.Visible;   // показываем/скрываем легенду
            posLegends();                           // пересчет положения форм легенды
        }

        private void dataGridView_m_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            using (SolidBrush b = new SolidBrush(dataGridView_m.RowHeadersDefaultCellStyle.ForeColor))
            {
                e.Graphics.DrawString(dataGridView_m.Rows[e.RowIndex].Cells[0].ToolTipText
                    , dataGridView_m.RowHeadersDefaultCellStyle.Font
                    , b
                    , e.RowBounds.Location.X + 12
                    , e.RowBounds.Location.Y + 4);
            }
        }

        private void FormMultipliers_Move(object sender, EventArgs e)
        {
            posLegends();   // пересчет положения форм легенды
        }

        private void FormMultipliers_FormClosing(object sender, FormClosingEventArgs e)
        {
            legendM.Close();   // закрываем форму легенды

            if (addM != null)
                addM.Close();  // закрываем форму добавления

            if (_FormMain != null)
                _FormMain.Close_Multipliers();  // объявляем что та форма закрывается
        }

        private void toolStripComboBox_m_SelectedIndexChanged(object sender, EventArgs e)
        {
            showTable(); // пересчет таблицы
        }
    }
}
