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
    public partial class FormAnalysis : Form
    {
        string filename_load = "";  // формат числа
        FormMain _FormMain = null;  // форма родителя


        public FormAnalysis(FormMain f)
        {
            InitializeComponent();

            if (filename_load == "")
                saveExcelToolStripMenuItem.Enabled = false;

            updateTable();

            _FormMain = f;
        }

        // форматирование в строку
        string ToStr(decimal d)
        {
            return d == 0 ? "-" : d.ToString("##0.00000000").Trim();
        }

        // перерасчет
        void updateTable()
        {
            #region 1 - На начало периода - тыс. руб.
            decimal ss_n_all = numericUpDown_s_n_yk.Value + numericUpDown_s_n_dk.Value + numericUpDown_s_n_rk.Value + numericUpDown_s_n_fss.Value + numericUpDown_s_n_cfp.Value;

            decimal zs_n_all = numericUpDown_z_n_dk.Value + numericUpDown_z_n_kk.Value + numericUpDown_z_n_kz.Value + numericUpDown_z_n_p.Value;

            decimal n_all = ss_n_all + zs_n_all;


            label_ss_all_n_r.Text = ToStr(ss_n_all);
            label_zs_all_n_r.Text = ToStr(zs_n_all);
            label_all_n_r.Text = ToStr(n_all);
            #endregion

            #region 2 - На начало периода - % к итогу
            label_s_n_yk_prc.Text = n_all == 0 ? "-" : ToStr(numericUpDown_s_n_yk.Value / n_all * 100);
            label_s_n_dk_prc.Text = n_all == 0 ? "-" : ToStr(numericUpDown_s_n_dk.Value / n_all * 100);
            label_s_n_rk_prc.Text = n_all == 0 ? "-" : ToStr(numericUpDown_s_n_rk.Value / n_all * 100);
            label_s_n_fss_prc.Text = n_all == 0 ? "-" : ToStr(numericUpDown_s_n_fss.Value / n_all * 100);
            label_s_n_cfp_prc.Text = n_all == 0 ? "-" : ToStr(numericUpDown_s_n_cfp.Value / n_all * 100);

            label_ss_all_n_prc.Text = n_all == 0 ? "-" : ToStr(ss_n_all / n_all * 100);


            label_z_n_dk_prc.Text = n_all == 0 ? "-" : ToStr(numericUpDown_z_n_dk.Value / n_all * 100);
            label_z_n_kk_prc.Text = n_all == 0 ? "-" : ToStr(numericUpDown_z_n_kk.Value / n_all * 100);
            label_z_n_kz_prc.Text = n_all == 0 ? "-" : ToStr(numericUpDown_z_n_kz.Value / n_all * 100);
            label_z_n_p_prc.Text = n_all == 0 ? "-" : ToStr(numericUpDown_z_n_p.Value / n_all * 100);

            label_zs_all_n_prc.Text = n_all == 0 ? "-" : ToStr(zs_n_all / n_all * 100);

            label_all_n_prc.Text = n_all == 0 ? "-" : ToStr(n_all / n_all * 100);
            #endregion

            #region 3 - На конец периода - тыс. руб.
            decimal ss_k_all = numericUpDown_s_k_yk.Value + numericUpDown_s_k_dk.Value + numericUpDown_s_k_rk.Value + numericUpDown_s_k_fss.Value + numericUpDown_s_k_cfp.Value;

            decimal zs_k_all = numericUpDown_z_k_dk.Value + numericUpDown_z_k_kk.Value + numericUpDown_z_k_kz.Value + numericUpDown_z_k_p.Value;

            decimal k_all = ss_k_all + zs_k_all;


            label_ss_all_k_r.Text = ToStr(ss_k_all);
            label_zs_all_k_r.Text = ToStr(zs_k_all);
            label_all_k_r.Text = ToStr(k_all);
            #endregion

            #region 4 - На конец периода - % к итогу
            label_s_k_yk_prc.Text = k_all == 0 ? "-" : ToStr(numericUpDown_s_k_yk.Value / k_all * 100);
            label_s_k_dk_prc.Text = k_all == 0 ? "-" : ToStr(numericUpDown_s_k_dk.Value / k_all * 100);
            label_s_k_rk_prc.Text = k_all == 0 ? "-" : ToStr(numericUpDown_s_k_rk.Value / k_all * 100);
            label_s_k_fss_prc.Text = k_all == 0 ? "-" : ToStr(numericUpDown_s_k_fss.Value / k_all * 100);
            label_s_k_cfp_prc.Text = k_all == 0 ? "-" : ToStr(numericUpDown_s_k_cfp.Value / k_all * 100);

            label_ss_all_k_prc.Text = k_all == 0 ? "-" : ToStr(ss_k_all / k_all * 100);


            label_z_k_dk_prc.Text = k_all == 0 ? "-" : ToStr(numericUpDown_z_k_dk.Value / k_all * 100);
            label_z_k_kk_prc.Text = k_all == 0 ? "-" : ToStr(numericUpDown_z_k_kk.Value / k_all * 100);
            label_z_k_kz_prc.Text = k_all == 0 ? "-" : ToStr(numericUpDown_z_k_kz.Value / k_all * 100);
            label_z_k_p_prc.Text = k_all == 0 ? "-" : ToStr(numericUpDown_z_k_p.Value / k_all * 100);

            label_zs_all_k_prc.Text = k_all == 0 ? "-" : ToStr(zs_k_all / k_all * 100);

            label_all_k_prc.Text = k_all == 0 ? "-" : ToStr(k_all / k_all * 100);
            #endregion

            #region 5 - Изменения - тыс. руб.
            decimal s_i_yk = numericUpDown_s_k_yk.Value - numericUpDown_s_n_yk.Value;
            decimal s_i_dk = numericUpDown_s_k_dk.Value - numericUpDown_s_n_dk.Value;
            decimal s_i_rk = numericUpDown_s_k_rk.Value - numericUpDown_s_n_rk.Value;
            decimal s_i_fss = numericUpDown_s_k_fss.Value - numericUpDown_s_n_fss.Value;
            decimal s_i_cfp = numericUpDown_s_k_cfp.Value - numericUpDown_s_n_cfp.Value;

            decimal ss_all_i_r = s_i_yk + s_i_dk + s_i_rk + s_i_fss + s_i_cfp;

            decimal z_i_dk = numericUpDown_z_k_dk.Value - numericUpDown_z_n_dk.Value;
            decimal z_i_kk = numericUpDown_z_k_kk.Value - numericUpDown_z_n_kk.Value;
            decimal z_i_kz = numericUpDown_z_k_kz.Value - numericUpDown_z_n_kz.Value;
            decimal z_i_p = numericUpDown_z_k_p.Value - numericUpDown_z_n_p.Value;

            decimal zs_all_i_r = z_i_dk + z_i_kk + z_i_kz + z_i_p;

            decimal all_i_r = ss_all_i_r + zs_all_i_r;

            label_s_i_yk.Text = ToStr(s_i_yk);
            label_s_i_dk.Text = ToStr(s_i_dk);
            label_s_i_rk.Text = ToStr(s_i_rk);
            label_s_i_fss.Text = ToStr(s_i_fss);
            label_s_i_cfp.Text = ToStr(s_i_cfp);

            label_ss_all_i_r.Text = ToStr(ss_all_i_r);

            label_z_i_dk.Text = ToStr(z_i_dk);
            label_z_i_kk.Text = ToStr(z_i_kk);
            label_z_i_kz.Text = ToStr(z_i_kz);
            label_z_i_p.Text = ToStr(z_i_p);

            label_zs_all_i_r.Text = ToStr(zs_all_i_r);

            label_all_i_r.Text = ToStr(all_i_r);
            #endregion

            #region 6 - Изменения - % к величине на начало периода
            label_s_i_yk_prc_n.Text = numericUpDown_s_n_yk.Value == 0 ? "-" : ToStr(s_i_yk / numericUpDown_s_n_yk.Value * 100);
            label_s_i_dk_prc_n.Text = numericUpDown_s_n_dk.Value == 0 ? "-" : ToStr(s_i_dk / numericUpDown_s_n_dk.Value * 100);
            label_s_i_rk_prc_n.Text = numericUpDown_s_n_rk.Value == 0 ? "-" : ToStr(s_i_rk / numericUpDown_s_n_rk.Value * 100);
            label_s_i_fss_prc_n.Text = numericUpDown_s_n_fss.Value == 0 ? "-" : ToStr(s_i_fss / numericUpDown_s_n_fss.Value * 100);
            label_s_i_cfp_prc_n.Text = numericUpDown_s_n_cfp.Value == 0 ? "-" : ToStr(s_i_cfp / numericUpDown_s_n_cfp.Value * 100);

            label_ss_all_i_prc_n.Text = ss_n_all == 0 ? "-" : ToStr(ss_all_i_r / ss_n_all * 100);


            label_z_i_dk_prc_n.Text = numericUpDown_z_n_dk.Value == 0 ? "-" : ToStr(z_i_dk / numericUpDown_z_n_dk.Value * 100);
            label_z_i_kk_prc_n.Text = numericUpDown_z_n_kk.Value == 0 ? "-" : ToStr(z_i_kk / numericUpDown_z_n_kk.Value * 100);
            label_z_i_kz_prc_n.Text = numericUpDown_z_n_kz.Value == 0 ? "-" : ToStr(z_i_kz / numericUpDown_z_n_kz.Value * 100);
            label_z_i_p_prc_n.Text = numericUpDown_z_n_p.Value == 0 ? "-" : ToStr(z_i_p / numericUpDown_z_n_p.Value * 100);

            label_zs_all_i_prc_n.Text = zs_n_all == 0 ? "-" : ToStr(zs_all_i_r / zs_n_all * 100);

            label_all_i_prc_n.Text = n_all == 0 ? "-" : ToStr(all_i_r / n_all * 100);
            #endregion

            #region 7 - Изменения - % к изменению итога баланса
            label_s_i_yk_prc_all.Text = all_i_r == 0 ? "-" : ToStr(s_i_yk / all_i_r * 100);
            label_s_i_dk_prc_all.Text = all_i_r == 0 ? "-" : ToStr(s_i_dk / all_i_r * 100);
            label_s_i_rk_prc_all.Text = all_i_r == 0 ? "-" : ToStr(s_i_rk / all_i_r * 100);
            label_s_i_fss_prc_all.Text = all_i_r == 0 ? "-" : ToStr(s_i_fss / all_i_r * 100);
            label_s_i_cfp_prc_all.Text = all_i_r == 0 ? "-" : ToStr(s_i_cfp / all_i_r * 100);

            label_ss_all_i_prc_all.Text = all_i_r == 0 ? "-" : ToStr(ss_all_i_r / all_i_r * 100);


            label_z_i_dk_prc_all.Text = all_i_r == 0 ? "-" : ToStr(z_i_dk / all_i_r * 100);
            label_z_i_kk_prc_all.Text = all_i_r == 0 ? "-" : ToStr(z_i_kk / all_i_r * 100);
            label_z_i_kz_prc_all.Text = all_i_r == 0 ? "-" : ToStr(z_i_kz / all_i_r * 100);
            label_z_i_p_prc_all.Text = all_i_r == 0 ? "-" : ToStr(z_i_p / all_i_r * 100);

            label_zs_all_i_prc_all.Text = all_i_r == 0 ? "-" : ToStr(zs_all_i_r / all_i_r * 100);

            label_all_i_prc_all.Text = all_i_r == 0 ? "-" : ToStr(all_i_r / all_i_r * 100);
            #endregion


            textBox_ss_ib_n.Text = n_all == 0 ? "-" : ToStr(ss_n_all / n_all);
            textBox_ss_ib_k.Text = k_all == 0 ? "-" : ToStr(ss_k_all / k_all);
            textBox_ss_ib_i.Text = all_i_r == 0 ? "-" : ToStr(ss_all_i_r / all_i_r);

            textBox_sk_zk_n.Text = zs_n_all == 0 ? "-" : ToStr(ss_n_all / zs_n_all);
            textBox_sk_zk_k.Text = zs_k_all == 0 ? "-" : ToStr(ss_k_all / zs_k_all);
            textBox_sk_zk_i.Text = zs_all_i_r == 0 ? "-" : ToStr(ss_all_i_r / zs_all_i_r);

            textBox_ib_sk_n.Text = ss_n_all == 0 ? "-" : ToStr(n_all / ss_n_all);
            textBox_ib_sk_k.Text = ss_k_all == 0 ? "-" : ToStr(k_all / ss_k_all);
            textBox_ib_sk_i.Text = ss_all_i_r == 0 ? "-" : ToStr(all_i_r / ss_all_i_r);
        }

        // нажали клавишу в элементе
        private void numericUpDown1_KeyUp(object sender, KeyEventArgs e)
        {
            updateTable();  // перерасчет
        }

        // изменили значение переменной
        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {
            updateTable();  // перерасчет
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

                    decimal _s_n_yk = (decimal)wSheet.GetNumber(4, 2);
                    decimal _s_n_dk = (decimal)wSheet.GetNumber(5, 2);
                    decimal _s_n_rk = (decimal)wSheet.GetNumber(6, 2);
                    decimal _s_n_fss = (decimal)wSheet.GetNumber(7, 2);
                    decimal _s_n_cfp = (decimal)wSheet.GetNumber(8, 2);

                    decimal _z_n_dk = (decimal)wSheet.GetNumber(11, 2);
                    decimal _z_n_kk = (decimal)wSheet.GetNumber(12, 2);
                    decimal _z_n_kz = (decimal)wSheet.GetNumber(13, 2);
                    decimal _z_n_p = (decimal)wSheet.GetNumber(14, 2);


                    decimal _s_k_yk = (decimal)wSheet.GetNumber(4, 4);
                    decimal _s_k_dk = (decimal)wSheet.GetNumber(5, 4);
                    decimal _s_k_rk = (decimal)wSheet.GetNumber(6, 4);
                    decimal _s_k_fss = (decimal)wSheet.GetNumber(7, 4);
                    decimal _s_k_cfp = (decimal)wSheet.GetNumber(8, 4);

                    decimal _z_k_dk = (decimal)wSheet.GetNumber(11, 4);
                    decimal _z_k_kk = (decimal)wSheet.GetNumber(12, 4);
                    decimal _z_k_kz = (decimal)wSheet.GetNumber(13, 4);
                    decimal _z_k_p = (decimal)wSheet.GetNumber(14, 4);


                    if (_s_n_yk < numericUpDown_s_n_yk.Minimum || _s_n_yk > numericUpDown_s_n_yk.Maximum)
                        MessageBox.Show("У «" + label_col1_2.Text + "» число не может быть меньше " + numericUpDown_s_n_yk.Minimum + " и больше " + numericUpDown_s_n_yk.Maximum, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    else if (_s_n_dk < numericUpDown_s_n_dk.Minimum || _s_n_dk > numericUpDown_s_n_dk.Maximum)
                        MessageBox.Show("У «" + label_col1_3.Text + "» число не может быть меньше " + numericUpDown_s_n_dk.Minimum + " и больше " + numericUpDown_s_n_dk.Maximum, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    else if (_s_n_rk < numericUpDown_s_n_rk.Minimum || _s_n_rk > numericUpDown_s_n_rk.Maximum)
                        MessageBox.Show("У «" + label_col1_4.Text + "» число не может быть меньше " + numericUpDown_s_n_rk.Minimum + " и больше " + numericUpDown_s_n_rk.Maximum, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    else if (_s_n_fss < numericUpDown_s_n_fss.Minimum || _s_n_fss > numericUpDown_s_n_fss.Maximum)
                        MessageBox.Show("У «" + label_col1_5.Text + "» число не может быть меньше " + numericUpDown_s_n_fss.Minimum + " и больше " + numericUpDown_s_n_fss.Maximum, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    else if (_s_n_cfp < numericUpDown_s_n_cfp.Minimum || _s_n_cfp > numericUpDown_s_n_cfp.Maximum)
                        MessageBox.Show("У «" + label_col1_6.Text + "» число не может быть меньше " + numericUpDown_s_n_cfp.Minimum + " и больше " + numericUpDown_s_n_cfp.Maximum, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);

                    else if (_z_n_dk < numericUpDown_z_n_dk.Minimum || _z_n_dk > numericUpDown_z_n_dk.Maximum)
                        MessageBox.Show("У «" + label_col1_9.Text + "» число не может быть меньше " + numericUpDown_z_n_dk.Minimum + " и больше " + numericUpDown_z_n_dk.Maximum, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    else if (_z_n_kk < numericUpDown_z_n_kk.Minimum || _z_n_kk > numericUpDown_z_n_kk.Maximum)
                        MessageBox.Show("У «" + label_col1_10.Text + "» число не может быть меньше " + numericUpDown_z_n_kk.Minimum + " и больше " + numericUpDown_z_n_kk.Maximum, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    else if (_z_n_kz < numericUpDown_z_n_kz.Minimum || _z_n_kz > numericUpDown_z_n_kz.Maximum)
                        MessageBox.Show("У «" + label_col1_11.Text + "» число не может быть меньше " + numericUpDown_z_n_kz.Minimum + " и больше " + numericUpDown_z_n_kz.Maximum, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    else if (_z_n_p < numericUpDown_z_n_p.Minimum || _z_n_p > numericUpDown_z_n_p.Maximum)
                        MessageBox.Show("У «" + label_col1_12.Text + "» число не может быть меньше " + numericUpDown_z_n_p.Minimum + " и больше " + numericUpDown_z_n_p.Maximum, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);

                    else if (_s_k_yk < numericUpDown_s_k_yk.Minimum || _s_k_yk > numericUpDown_s_k_yk.Maximum)
                        MessageBox.Show("У «" + label_col1_2.Text + "» число не может быть меньше " + numericUpDown_s_k_yk.Minimum + " и больше " + numericUpDown_s_k_yk.Maximum, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    else if (_s_k_dk < numericUpDown_s_k_dk.Minimum || _s_k_dk > numericUpDown_s_k_dk.Maximum)
                        MessageBox.Show("У «" + label_col1_3.Text + "» число не может быть меньше " + numericUpDown_s_k_dk.Minimum + " и больше " + numericUpDown_s_k_dk.Maximum, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    else if (_s_k_rk < numericUpDown_s_k_rk.Minimum || _s_k_rk > numericUpDown_s_k_rk.Maximum)
                        MessageBox.Show("У «" + label_col1_4.Text + "» число не может быть меньше " + numericUpDown_s_k_rk.Minimum + " и больше " + numericUpDown_s_k_rk.Maximum, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    else if (_s_k_fss < numericUpDown_s_k_fss.Minimum || _s_k_fss > numericUpDown_s_k_fss.Maximum)
                        MessageBox.Show("У «" + label_col1_5.Text + "» число не может быть меньше " + numericUpDown_s_k_fss.Minimum + " и больше " + numericUpDown_s_k_fss.Maximum, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    else if (_s_k_cfp < numericUpDown_s_k_cfp.Minimum || _s_k_cfp > numericUpDown_s_k_cfp.Maximum)
                        MessageBox.Show("У «" + label_col1_6.Text + "» число не может быть меньше " + numericUpDown_s_k_cfp.Minimum + " и больше " + numericUpDown_s_k_cfp.Maximum, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);

                    else if (_z_k_dk < numericUpDown_z_k_dk.Minimum || _z_k_dk > numericUpDown_z_k_dk.Maximum)
                        MessageBox.Show("У «" + label_col1_9.Text + "» число не может быть меньше " + numericUpDown_z_k_dk.Minimum + " и больше " + numericUpDown_z_k_dk.Maximum, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    else if (_z_k_kk < numericUpDown_z_k_kk.Minimum || _z_k_kk > numericUpDown_z_k_kk.Maximum)
                        MessageBox.Show("У «" + label_col1_10.Text + "» число не может быть меньше " + numericUpDown_z_k_kk.Minimum + " и больше " + numericUpDown_z_k_kk.Maximum, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    else if (_z_k_kz < numericUpDown_z_k_kz.Minimum || _z_k_kz > numericUpDown_z_k_kz.Maximum)
                        MessageBox.Show("У «" + label_col1_11.Text + "» число не может быть меньше " + numericUpDown_z_k_kz.Minimum + " и больше " + numericUpDown_z_k_kz.Maximum, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    else if (_z_k_p < numericUpDown_z_k_p.Minimum || _z_k_p > numericUpDown_z_k_p.Maximum)
                        MessageBox.Show("У «" + label_col1_12.Text + "» число не может быть меньше " + numericUpDown_z_k_p.Minimum + " и больше " + numericUpDown_z_k_p.Maximum, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    else
                    {
                        numericUpDown_s_n_yk.Value = _s_n_yk;
                        numericUpDown_s_n_dk.Value = _s_n_dk;
                        numericUpDown_s_n_rk.Value = _s_n_rk;
                        numericUpDown_s_n_fss.Value = _s_n_fss;
                        numericUpDown_s_n_cfp.Value = _s_n_cfp;

                        numericUpDown_z_n_dk.Value = _z_n_dk;
                        numericUpDown_z_n_kk.Value = _z_n_kk;
                        numericUpDown_z_n_kz.Value = _z_n_kz;
                        numericUpDown_z_n_p.Value = _z_n_p;

                        numericUpDown_s_k_yk.Value = _s_k_yk;
                        numericUpDown_s_k_dk.Value = _s_k_dk;
                        numericUpDown_s_k_rk.Value = _s_k_rk;
                        numericUpDown_s_k_fss.Value = _s_k_fss;
                        numericUpDown_s_k_cfp.Value = _s_k_cfp;

                        numericUpDown_z_k_dk.Value = _z_k_dk;
                        numericUpDown_z_k_kk.Value = _z_k_kk;
                        numericUpDown_z_k_kz.Value = _z_k_kz;
                        numericUpDown_z_k_p.Value = _z_k_p;

                        filename_load = openFileDialog1.FileName;
                    }
                }
                catch (Exception ee)
                {
                    MessageBox.Show(ee.Message, "Ошибка открытия файла Excel", MessageBoxButtons.OK, MessageBoxIcon.Error);

                    numericUpDown_s_n_yk.Value = 0;
                    numericUpDown_s_n_dk.Value = 0;
                    numericUpDown_s_n_rk.Value = 0;
                    numericUpDown_s_n_fss.Value = 0;
                    numericUpDown_s_n_cfp.Value = 0;

                    numericUpDown_z_n_dk.Value = 0;
                    numericUpDown_z_n_kk.Value = 0;
                    numericUpDown_z_n_kz.Value = 0;
                    numericUpDown_z_n_p.Value = 0;

                    numericUpDown_s_k_yk.Value = 0;
                    numericUpDown_s_k_dk.Value = 0;
                    numericUpDown_s_k_rk.Value = 0;
                    numericUpDown_s_k_fss.Value = 0;
                    numericUpDown_s_k_cfp.Value = 0;

                    numericUpDown_z_k_dk.Value = 0;
                    numericUpDown_z_k_kk.Value = 0;
                    numericUpDown_z_k_kz.Value = 0;
                    numericUpDown_z_k_p.Value = 0;
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

                #region Шапка

                wSheet.Range[1, 1].Text = label_head_1_1_2.Text;
                wSheet.Range[1, 1, 2, 1].Merge();
                wSheet.Range[1, 1, 2, 1].CellStyle = style;
                wSheet.Range[1, 1, 2, 1].BorderAround(ExcelLineStyle.Thin, Color.Black);

                wSheet.Range[1, 2].Text = label_head_2_2_1.Text;
                wSheet.Range[1, 2, 1, 3].Merge();
                wSheet.Range[1, 2, 1, 3].CellStyle = style;
                wSheet.Range[1, 2, 1, 3].BorderAround(ExcelLineStyle.Thin, Color.Black);

                wSheet.Range[1, 4].Text = label_head_3_2_1.Text;
                wSheet.Range[1, 4, 1, 5].Merge();
                wSheet.Range[1, 4, 1, 5].CellStyle = style;
                wSheet.Range[1, 4, 1, 5].BorderAround(ExcelLineStyle.Thin, Color.Black);

                wSheet.Range[1, 6].Text = label_head_4_3_1.Text;
                wSheet.Range[1, 6, 1, 8].Merge();
                wSheet.Range[1, 6, 1, 8].CellStyle = style;
                wSheet.Range[1, 6, 1, 8].BorderAround(ExcelLineStyle.Thin, Color.Black);


                wSheet.Range[2, 2].Text = label_head_5_1_1.Text;
                wSheet.Range[2, 2].CellStyle = style;
                wSheet.Range[2, 2].BorderAround(ExcelLineStyle.Thin, Color.Black);

                wSheet.Range[2, 3].Text = label_head_6_1_1.Text;
                wSheet.Range[2, 3].CellStyle = style;
                wSheet.Range[2, 3].BorderAround(ExcelLineStyle.Thin, Color.Black);

                wSheet.Range[2, 4].Text = label_head_7_1_1.Text;
                wSheet.Range[2, 4].CellStyle = style;
                wSheet.Range[2, 4].BorderAround(ExcelLineStyle.Thin, Color.Black);

                wSheet.Range[2, 5].Text = label_head_8_1_1.Text;
                wSheet.Range[2, 5].CellStyle = style;
                wSheet.Range[2, 5].BorderAround(ExcelLineStyle.Thin, Color.Black);

                wSheet.Range[2, 6].Text = label_head_9_1_1.Text;
                wSheet.Range[2, 6].CellStyle = style;
                wSheet.Range[2, 6].BorderAround(ExcelLineStyle.Thin, Color.Black);

                wSheet.Range[2, 7].Text = label_head_10_1_1.Text;
                wSheet.Range[2, 7].CellStyle = style;
                wSheet.Range[2, 7].BorderAround(ExcelLineStyle.Thin, Color.Black);

                wSheet.Range[2, 8].Text = label_head_11_1_1.Text;
                wSheet.Range[2, 8].CellStyle = style;
                wSheet.Range[2, 8].BorderAround(ExcelLineStyle.Thin, Color.Black);

                #endregion


                IStyle style2 = wBook.Styles.Add("FillColor2");
                style2.ColorIndex = ExcelKnownColors.Light_yellow;


                #region cnhjrb

                int i = 3;
                setRowExcel(wSheet, style, style2, i, label_col1_1.Text, "", "", "", "", "", "", "");
                i++;
                setRowExcel(wSheet, style, style2, i, label_col1_2.Text, ToStr(numericUpDown_s_n_yk.Value), label_s_n_yk_prc.Text, ToStr(numericUpDown_s_k_yk.Value), label_s_k_yk_prc.Text, label_s_i_yk.Text, label_s_i_yk_prc_n.Text, label_s_i_yk_prc_all.Text);
                i++;
                setRowExcel(wSheet, style, style2, i, label_col1_3.Text, ToStr(numericUpDown_s_n_dk.Value), label_s_n_dk_prc.Text, ToStr(numericUpDown_s_k_dk.Value), label_s_k_dk_prc.Text, label_s_i_dk.Text, label_s_i_dk_prc_n.Text, label_s_i_dk_prc_all.Text);
                i++;
                setRowExcel(wSheet, style, style2, i, label_col1_4.Text, ToStr(numericUpDown_s_n_rk.Value), label_s_n_rk_prc.Text, ToStr(numericUpDown_s_k_rk.Value), label_s_k_rk_prc.Text, label_s_i_rk.Text, label_s_i_rk_prc_n.Text, label_s_i_rk_prc_all.Text);
                i++;
                setRowExcel(wSheet, style, style2, i, label_col1_5.Text, ToStr(numericUpDown_s_n_fss.Value), label_s_n_fss_prc.Text, ToStr(numericUpDown_s_k_fss.Value), label_s_k_fss_prc.Text, label_s_i_fss.Text, label_s_i_fss_prc_n.Text, label_s_i_fss_prc_all.Text);
                i++;
                setRowExcel(wSheet, style, style2, i, label_col1_6.Text, ToStr(numericUpDown_s_n_cfp.Value), label_s_n_cfp_prc.Text, ToStr(numericUpDown_s_k_cfp.Value), label_s_k_cfp_prc.Text, label_s_i_cfp.Text, label_s_i_cfp_prc_n.Text, label_s_i_cfp_prc_all.Text);

                i++;
                setRowExcel(wSheet, style, style2, i, label_col1_7.Text, label_ss_all_n_r.Text, label_ss_all_n_prc.Text, label_ss_all_k_r.Text, label_ss_all_k_prc.Text, label_ss_all_i_r.Text, label_ss_all_i_prc_n.Text, label_ss_all_i_prc_all.Text);

                i++;
                setRowExcel(wSheet, style, style2, i, label_col1_8.Text, "", "", "", "", "", "", "");

                i++;
                setRowExcel(wSheet, style, style2, i, label_col1_9.Text, ToStr(numericUpDown_z_n_dk.Value), label_z_n_dk_prc.Text, ToStr(numericUpDown_z_k_dk.Value), label_z_k_dk_prc.Text, label_z_i_dk.Text, label_z_i_dk_prc_n.Text, label_z_i_dk_prc_all.Text);
                i++;
                setRowExcel(wSheet, style, style2, i, label_col1_10.Text, ToStr(numericUpDown_z_n_kk.Value), label_z_n_kk_prc.Text, ToStr(numericUpDown_z_k_kk.Value), label_z_k_kk_prc.Text, label_z_i_kk.Text, label_z_i_kk_prc_n.Text, label_z_i_kk_prc_all.Text);
                i++;
                setRowExcel(wSheet, style, style2, i, label_col1_11.Text, ToStr(numericUpDown_z_n_kz.Value), label_z_n_kz_prc.Text, ToStr(numericUpDown_z_k_kz.Value), label_z_k_kz_prc.Text, label_z_i_kz.Text, label_z_i_kz_prc_n.Text, label_z_i_kz_prc_all.Text);
                i++;
                setRowExcel(wSheet, style, style2, i, label_col1_12.Text, ToStr(numericUpDown_z_n_p.Value), label_z_n_p_prc.Text, ToStr(numericUpDown_z_k_p.Value), label_z_k_p_prc.Text, label_z_i_p.Text, label_z_i_p_prc_n.Text, label_z_i_p_prc_all.Text);

                i++;
                setRowExcel(wSheet, style, style2, i, label_col1_13.Text, label_zs_all_n_r.Text, label_zs_all_n_prc.Text, label_zs_all_k_r.Text, label_zs_all_k_prc.Text, label_zs_all_i_r.Text, label_zs_all_i_prc_n.Text, label_zs_all_i_prc_all.Text);

                i++;
                setRowExcel(wSheet, style, style2, i, label_col1_14.Text, label_all_n_r.Text, label_all_n_prc.Text, label_all_k_r.Text, label_all_k_prc.Text, label_all_i_r.Text, label_all_i_prc_n.Text, label_all_i_prc_all.Text);


                i+=2;
                setRowExcel(wSheet, style, style2, i, label1.Text, textBox_ss_ib_n.Text, textBox_ss_ib_k.Text, textBox_ss_ib_i.Text);

                i++;
                setRowExcel(wSheet, style, style2, i, label2.Text, textBox_sk_zk_n.Text, textBox_sk_zk_k.Text, textBox_sk_zk_i.Text);

                i++;
                setRowExcel(wSheet, style, style2, i, label3.Text, textBox_ib_sk_n.Text, textBox_ib_sk_k.Text, textBox_ib_sk_i.Text);

                #endregion

                for (i = 1; i <= 8; i++)
                {
                    wSheet.AutofitColumn(i);
                }

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
        void setRowExcel(IWorksheet wSheet, IStyle style, IStyle style2, int row, string c1, string c2, string c3, string c4, string c5, string c6, string c7, string c8)
        {
            int i = 1;
            wSheet.Range[row, i].Text = c1;
            wSheet.Range[row, i].CellStyle = style;
            wSheet.Range[row, i].BorderAround(ExcelLineStyle.Thin, Color.Black);

            i++;
            wSheet.Range[row, i].Value = c2;
            if (row == 3 || row == 9 || row == 10 || row == 15 || row == 16)
                wSheet.Range[row, i].CellStyle = style2;
            wSheet.Range[row, i].BorderAround(ExcelLineStyle.Thin, Color.Black);

            i++;
            wSheet.Range[row, i].Value = c3;
            wSheet.Range[row, i].CellStyle = style2;
            wSheet.Range[row, i].BorderAround(ExcelLineStyle.Thin, Color.Black);

            i++;
            wSheet.Range[row, i].Value = c4;
            if (row == 3 || row == 9 || row == 10 || row == 15 || row == 16)
                wSheet.Range[row, i].CellStyle = style2;
            wSheet.Range[row, i].BorderAround(ExcelLineStyle.Thin, Color.Black);

            i++;
            wSheet.Range[row, i].Value = c5;
            wSheet.Range[row, i].CellStyle = style2;
            wSheet.Range[row, i].BorderAround(ExcelLineStyle.Thin, Color.Black);

            i++;
            wSheet.Range[row, i].Value = c6;
            wSheet.Range[row, i].CellStyle = style2;
            wSheet.Range[row, i].BorderAround(ExcelLineStyle.Thin, Color.Black);

            i++;
            wSheet.Range[row, i].Value = c7;
            wSheet.Range[row, i].CellStyle = style2;
            wSheet.Range[row, i].BorderAround(ExcelLineStyle.Thin, Color.Black);

            i++;
            wSheet.Range[row, i].Value = c8;
            wSheet.Range[row, i].CellStyle = style2;
            wSheet.Range[row, i].BorderAround(ExcelLineStyle.Thin, Color.Black);
        }

        // заполняем строку в Excel
        void setRowExcel(IWorksheet wSheet, IStyle style, IStyle style2, int row, string c1, string c2, string c3, string c4)
        {
            int i = 1;
            wSheet.Range[row, i].Text = c1;
            wSheet.Range[row, i].CellStyle = style;
            wSheet.Range[row, i].BorderAround(ExcelLineStyle.Thin, Color.Black);

            i++;
            wSheet.Range[row, i].Value = c2;
            wSheet.Range[row, i].CellStyle = style2;
            wSheet.Range[row, i].BorderAround(ExcelLineStyle.Thin, Color.Black);

            i += 2;
            wSheet.Range[row, i].Value = c3;
            wSheet.Range[row, i].CellStyle = style2;
            wSheet.Range[row, i].BorderAround(ExcelLineStyle.Thin, Color.Black);

            i += 2;
            wSheet.Range[row, i].Value = c4;
            wSheet.Range[row, i].CellStyle = style2;
            wSheet.Range[row, i].BorderAround(ExcelLineStyle.Thin, Color.Black);
        }

        private void FormAnalysis_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (_FormMain != null)
            {
                _FormMain.Close_Analysis(); // объявляем что та форма закрывается
            }
        }
    }
}
