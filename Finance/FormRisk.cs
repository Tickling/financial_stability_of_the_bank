using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Syncfusion.XlsIO;
using System.IO;

namespace Finance
{
    public partial class FormRisk : Form
    {
        decimal _risk { get { return numericUpDown_size.Value / numericUpDown_H.Value; } }
        decimal? _pr { get { return _risk == 0 ? (decimal?)null : (12 * numericUpDown_profit.Value / _risk); } }
        string _pr_str
        {
            get
            {
                string str = "";
                if (_pr == null)
                    str = "-";
                else if (_pr <= 0)
                    str = "низкая";
                else if (_pr > 0 && _pr <= (decimal)1.2)
                    str = "сомнительная";
                else if (_pr > (decimal)1.2 && _pr <= (decimal)2.4)
                    str = "удовлетворительная";
                else if (_pr > (decimal)2.4 && _pr <= (decimal)3.6)
                    str = "хорошая";
                else if (_pr > (decimal)3.6)
                    str = "высокая";
                return str;
            }
        }
        string filename_load = "";  // формат числа
        FormMain _FormMain = null;  // форма родителя


        public FormRisk(FormMain f)
        {
            InitializeComponent();

            if (filename_load == "")
                saveExcelToolStripMenuItem.Enabled = false;

            _FormMain = f;
        }

        // перерасчет
        void update_result()
        {
            string decimal_format = "##0.00000000";

            textBox_risk.Text = _risk.ToString(decimal_format).Trim();
            textBox_pr.Text = (_pr == null) ? "-" : _pr.Value.ToString(decimal_format).Trim();
            textBox_pr_str.Text = _pr_str;
        }

        // нажали клавишу в элементе
        private void numericUpDown_H_KeyUp(object sender, KeyEventArgs e)
        {
            update_result();
        }

        // изменили значение переменной
        private void numericUpDown_H_ValueChanged(object sender, EventArgs e)
        {
            update_result();
        }

        // Кнопка Сохранить в Excel
        private void saveExcelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (filename_load != "" && MessageBox.Show("Хотите перезаписать файл \"" + filename_load + "\" ?", "Вопрос", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                SaveExcel(filename_load);
        }

        // Кнопка Сохранить как...
        private void saveAsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //Создаем диалоговое окно длля сохранения
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.CheckPathExists = true;
            saveFileDialog1.Filter = "Excel (*.xlsx)|*.xlsx|Excel (*.xls)|*.xls";
            saveFileDialog1.FilterIndex = 1;
            saveFileDialog1.RestoreDirectory = true;

            // Вызываем диалоговое окно для сохранения файла Excel
            if (saveFileDialog1.ShowDialog() != DialogResult.Cancel)
                SaveExcel(saveFileDialog1.FileName);
        }

        // загружаем из Excel
        private void loadToolStripMenuItem_Click(object sender, EventArgs e)
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

                    decimal _H = (decimal)wSheet.GetNumber(1, 2);
                    decimal _size = (decimal)wSheet.GetNumber(2, 2);
                    decimal _profit = (decimal)wSheet.GetNumber(3, 2);

                    if (_H < 2 || _H > 100)
                    {
                        MessageBox.Show("У «Норматив достаточности собственных средств (капитала) H1.0» число не может быть меньше 2 и больше 100", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    else if (_size < numericUpDown_size.Minimum || _size > numericUpDown_size.Maximum)
                    {
                        MessageBox.Show("У «Размер собственных средств» число не может быть меньше " + numericUpDown_size.Minimum + " и больше " + numericUpDown_size.Maximum, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    else if (_profit < numericUpDown_profit.Minimum || _profit > numericUpDown_profit.Maximum)
                    {
                        MessageBox.Show("У «Прибыль (убыток)до налого облажения» число не может быть меньше " + numericUpDown_profit.Minimum + " и больше " + numericUpDown_profit.Maximum, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    else
                    {
                        numericUpDown_H.Value = _H;
                        numericUpDown_size.Value = _size;
                        numericUpDown_profit.Value = _profit;

                        filename_load = openFileDialog1.FileName;
                    }
                }
                catch (Exception ee)
                {
                    MessageBox.Show(ee.ToString(), "Ошибка открытия файла Excel", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }


                if (filename_load != "")
                    saveExcelToolStripMenuItem.Enabled = true;
            }
        }

        // сохраняем в Excel
        private void SaveExcel(string PathToSave_file)
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

                int i = 1;
                setCellExcel(wSheet, style, i, 1, label_H.Text);
                i++;
                setCellExcel(wSheet, style, i, 1, label_size.Text);
                i++;
                setCellExcel(wSheet, style, i, 1, label_profit.Text);
                i++;
                setCellExcel(wSheet, style, i, 1, label_risk.Text);
                i++;
                setCellExcel(wSheet, style, i, 1, label_pr.Text);
                i++;
                setCellExcel(wSheet, style, i, 1, label_pr_str.Text);


                IStyle style2 = wBook.Styles.Add("FillColor2");
                style2.ColorIndex = ExcelKnownColors.Light_yellow;

                i = 1;
                setCellExcel(wSheet, null, i, 2, numericUpDown_H.Value.ToString());
                i++;
                setCellExcel(wSheet, null, i, 2, numericUpDown_size.Value.ToString());
                i++;
                setCellExcel(wSheet, null, i, 2, numericUpDown_profit.Value.ToString());
                i++;
                setCellExcel(wSheet, style2, i, 2, textBox_risk.Text);
                i++;
                setCellExcel(wSheet, style2, i, 2, textBox_pr.Text);
                i++;
                setCellExcel(wSheet, style2, i, 2, textBox_pr_str.Text);


                wSheet.AutofitColumn(1);
                wSheet.AutofitColumn(2);

                // Сохраняем файл Excel
                _excel.Save(PathToSave_file);

                filename_load = PathToSave_file;
            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.ToString(), "Ошибка сохранения файла Excel", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            if (filename_load != "")
                saveExcelToolStripMenuItem.Enabled = true;
        }

        // заполняем строку в Excel
        void setCellExcel(IWorksheet wSheet, IStyle style, int row, int col, string c)
        {
            wSheet.Range[row, col].Text = c;
            if(style != null)
                wSheet.Range[row, col].CellStyle = style;
            wSheet.Range[row, col].BorderAround(ExcelLineStyle.Thin, Color.Black);
        }

        private void FormRisk_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (_FormMain != null)
            {
                _FormMain.Close_Risk(); // объявляем что та форма закрывается
            }
        }
    }
}
