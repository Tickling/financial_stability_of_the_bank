using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using HtmlAgilityPack;
using System.Net;
using System.IO;
using Syncfusion.XlsIO;

namespace Finance
{
    public partial class FormGetData : Form
    {
        List<string> urls_years_802 = new List<string>(); // список ссылок на форму 802
        List<string> urls_years_101 = new List<string>(); // список ссылок на форму 101
        string FileName = "";


        FormMain _FormMain = null;                      // форма родителя

        public FormGetData(FormMain f)
        {
            InitializeComponent();

            _FormMain = f;

            comboBox_bank.Items.Add("Сбербанк");
            comboBox_bank.Items.Add("Тинькофф");
            comboBox_bank.Items.Add("Втб");
            comboBox_bank.Items.Add("Уралсиб");
            comboBox_bank.Items.Add("МОСКОВСКИЙ ОБЛАСТНОЙ БАНК");
        }

        private void FormGetData_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (backgroundWorker1.IsBusy)           // если запущен процесс формирования excel
                backgroundWorker1.CancelAsync();    // останавливаем формирование excel

            if (_FormMain != null)
                _FormMain.Close_get_data();  // объявляем что форма закрывается
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                urls_years_802.Clear(); // очищаем форму 802
                urls_years_101.Clear(); // очищаем форму 101

                // проверяем какой выбран банк и берем его ссылку
                switch (comboBox_bank.SelectedIndex)
                {
                    case 0:
                        ParserYears(GetHTML("https://www.cbr.ru/banking_sector/credit/coinfo/?id=350000004"));
                        break;
                    case 1:
                        ParserYears(GetHTML("https://www.cbr.ru/banking_sector/credit/coinfo/?id=450000562"));
                        break;
                    case 2:
                        ParserYears(GetHTML("https://www.cbr.ru/banking_sector/credit/coinfo/?id=350000008"));
                        break;
                    case 3:
                        ParserYears(GetHTML("https://www.cbr.ru/banking_sector/credit/coinfo/?id=800000002"));
                        break;
                    case 4:
                        ParserYears(GetHTML("https://www.cbr.ru/banking_sector/credit/coinfo/?id=820000042"));
                        break;
                }

                if (urls_years_802.Count > 0 || urls_years_101.Count > 0)
                {

                    //Создаем диалоговое окно длля сохранения
                    SaveFileDialog saveFileDialog1 = new SaveFileDialog();
                    saveFileDialog1.CheckPathExists = true;
                    saveFileDialog1.Filter = "Excel (*.xlsx)|*.xlsx|Excel (*.xls)|*.xls";
                    saveFileDialog1.FilterIndex = 1;
                    saveFileDialog1.RestoreDirectory = true;

                    // Вызываем диалоговое окно для сохранения файла Excel
                    if (saveFileDialog1.ShowDialog() != DialogResult.Cancel)
                    {
                        FileName = saveFileDialog1.FileName;
                        button1.Enabled = false;
                        progressBar1.Value = 0;
                        backgroundWorker1.RunWorkerAsync();
                    }
                }
                else
                {
                    MessageBox.Show("Нет ссылок на формы для сохранения (может что-то поменялось на сайте)", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.ToString(), "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // функция возвращает код HTML
        string GetHTML(string url)
        {
            string html = "";

            try
            {
                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
                request.Credentials = CredentialCache.DefaultCredentials;
                request.UserAgent = @"Mozilla/5.0 (compatible; MSIE 10.0; Windows NT 6.1; Trident/6.0)";
                HttpWebResponse response = (HttpWebResponse)request.GetResponse();
                using (StreamReader sr = new StreamReader(response.GetResponseStream()))    // выполняем запрос по ссылке
                {
                    html = sr.ReadToEnd(); // получаем код HTML
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString(), "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return html;
        }

        // парсим ссылки по годам на 1 января
        void ParserYears(string html)
        {
            try
            {
                if (!string.IsNullOrEmpty(html))
                {
                    HtmlAgilityPack.HtmlDocument htmlDocument = new HtmlAgilityPack.HtmlDocument(); // создпем парсер
                    htmlDocument.LoadHtml(html);     // заполняем парсер полученным кодом HTML

                    // парсим и находим элементы "div" с классом "reports"
                    var rezult = from x in htmlDocument.DocumentNode.DescendantNodes()
                                 where x.Name == "div" && x.Attributes["class"] != null && x.Attributes["class"].Value == "reports"
                                 select x;

                    if (rezult.Count() > 0) // если есть результаты парсинга
                    {
                        HtmlNode row = rezult.First(); // берем первый HtmlNode из парсинга

                        for (int i = 0; i < row.ChildNodes.Count; i++) // пробегаем по результатам парсера rezult
                        {
                            if (row.ChildNodes[i].InnerText.ToLower().Trim() == "форма 802") // находим форму 802
                            {
                                int ip = 1;
                                if (row.ChildNodes[i + 1].Name == "#text") // если слудующий элемент "#text" переход еще на один дальше
                                    ip += 1;

                                // парсим и находим ссылки на 1 января по всем годам и добавляем их в список urls_years_802
                                urls_years_802.AddRange(
                                    (from x in row.ChildNodes[i + ip].DescendantNodes()
                                     where x.Name == "a" 
                                        && x.Attributes["class"] != null && x.Attributes["class"].Value == "versions_item"
                                        && x.InnerText.IndexOf("на 1 января") >= 0
                                     select x.Attributes["href"].Value.Replace("&amp;", "&")).ToList()
                                             );
                            }
                            else if (row.ChildNodes[i].InnerText.ToLower().Trim() == "форма 101") // находим форму 101
                            {
                                int ip = 1;
                                if (row.ChildNodes[i + 1].Name == "#text") // если слудующий элемент "#text" переход еще на один дальше
                                    ip += 1;

                                // парсим и находим ссылки на 1 января по всем годам и добавляем их в список urls_years_101
                                urls_years_101.AddRange(
                                    (from x in row.ChildNodes[i + ip].DescendantNodes()
                                     where x.Name == "a"
                                        && x.Attributes["class"] != null && x.Attributes["class"].Value == "versions_item"
                                        && x.InnerText.IndexOf("на 1 января") >= 0
                                     select x.Attributes["href"].Value.Replace("&amp;", "&")).ToList()
                                             );
                            }
                        }
                    }
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString(), "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // парсим таблицы
        void ParserTable(string html, IWorksheet wSheet, bool is_101)
        {
            try
            {
                if (!string.IsNullOrEmpty(html))
                {
                    HtmlAgilityPack.HtmlDocument htmlDocument = new HtmlAgilityPack.HtmlDocument(); // создпем парсер
                    htmlDocument.LoadHtml(html);     // заполняем парсер полученным кодом HTML

                    // парсим и находим элементы "table" с классом "data" или "data spaced", а родительски элемент должен быть "div" и его класс "table"
                    var rezult = from x in htmlDocument.DocumentNode.DescendantNodes()
                                 where x.Name == "table" && x.Attributes["class"] != null && (x.Attributes["class"].Value == "data" || x.Attributes["class"].Value == "data spaced")
                                    && x.ParentNode.Name == "div" && x.ParentNode.Attributes["class"] != null && x.ParentNode.Attributes["class"].Value == "table"
                                 select x;

                    int max_col = 1;
                    int col = 1;
                    int row = 1;
                    foreach (HtmlNode table in rezult) // пробегаем по результатам парсера rezult
                    {
                        var trs = from x in table.DescendantNodes() where x.Name == "tr" select x; // находим все элементы "tr" в table

                        foreach (HtmlNode tr in trs) // пробегаем по результатам trs
                        {
                            var ths = from x in tr.DescendantNodes() where x.Name == "th" || x.Name == "td" select x; // находим все элементы "th" или "td" в tr
                            col = 1;

                            foreach (HtmlNode th in ths) // пробегаем по результатам ths
                            {
                                int colspan = th.Attributes["colspan"] == null ? 1 : Convert.ToInt32(th.Attributes["colspan"].Value); // получаем colspan для ячейки
                                int rowspan = th.Attributes["rowspan"] == null ? 1 : Convert.ToInt32(th.Attributes["rowspan"].Value); // получаем rowspan для ячейки

                            ff:
                                IRange marge = wSheet.Range[row, col].MergeArea;
                                if (marge != null)  // проверяем есть ли для той ящейки объединение
                                {
                                    col = marge.LastColumn + 1; // если есть берем следующую колонку
                                    goto ff;                    // и возвращаемся снова с проверки на объединение
                                }
                                wSheet.Range[row, col, row - 1 + rowspan, col - 1 + colspan].Merge(true);                                           // объединяем ячейки
                                wSheet.Range[row, col, row - 1 + rowspan, col - 1 + colspan].BorderAround(ExcelLineStyle.Thin, Color.Black);        // делаем рамку для ячейки
                                wSheet.Range[row, col].Text = th.InnerText.Trim();                                                                  // заполняем данные для этой ячейки
                                if(th.Attributes["class"] != null && th.Attributes["class"].Value.IndexOf("right") >=0)                             // если выравнивание справа
                                    wSheet.Range[row, col, row - 1 + rowspan, col - 1 + colspan].HorizontalAlignment = ExcelHAlign.HAlignRight;
                                if (th.Attributes["class"] != null && th.Attributes["class"].Value.IndexOf("center") >= 0)                          // если выравнивание в центре
                                {
                                    wSheet.Range[row, col, row - 1 + rowspan, col - 1 + colspan].HorizontalAlignment = ExcelHAlign.HAlignCenter;
                                    wSheet.Range[row, col, row - 1 + rowspan, col - 1 + colspan].VerticalAlignment = ExcelVAlign.VAlignCenter;
                                }
                                if (th.Attributes["class"] != null && th.Attributes["class"].Value.IndexOf("bold") >= 0)                            // если выравнивание жирный шрифт значит и выравнивание в центре
                                {
                                    wSheet.Range[row, col, row - 1 + rowspan, col - 1 + colspan].HorizontalAlignment = ExcelHAlign.HAlignCenter;
                                    wSheet.Range[row, col, row - 1 + rowspan, col - 1 + colspan].CellStyle.Font.Bold = true;
                                }

                                col += colspan;
                            }

                            if (col > max_col)
                                max_col = col;

                            row++;
                        }
                        row++;
                    }

                    wSheet.UsedRange.AutofitColumns();  // автоматически выравниваем столбцы
                    if (is_101)
                    {
                        wSheet.SetColumnWidth(1, 23);   // первый столбец для формы 101 имеет длину 23
                    }
                    else
                    {
                        for (int i = 1; i <= wSheet.Columns.Count(); i++)
                        {
                            wSheet.SetColumnWidth(i, 35);   // все столбцы для формы 802 имееют длину 35
                        }
                    }
                    wSheet.UsedRange.WrapText = true;   // во всех ячейках устанавливаем перенос слов
                    wSheet.UsedRange.AutofitRows();     //  автоматически выравниваем строки
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString(), "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // если прогресс парсинга и сохраненния увеличился
        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar1.Value = e.ProgressPercentage;
        }

        // выполянем процесс парсинга и сохранения в excel
        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            var backgroundWorker = sender as BackgroundWorker;
            FileInfo fi = new FileInfo(FileName);

            ExcelEngine excelEngine = new ExcelEngine();
            IApplication _excel = excelEngine.Excel;

            // Тип файла 
            _excel.DefaultVersion = ExcelVersion.Excel97to2003; // .xls
            if (fi.Extension == ".xlsx")
                _excel.DefaultVersion = ExcelVersion.Excel2010; // .xlsx

            // Создаем книгу
            IWorkbook wBook = _excel.Workbooks.Create(urls_years_802.Count + urls_years_101.Count);
            int ws = 0;

            string mail_url = "https://www.cbr.ru/banking_sector/credit/coinfo/";
            foreach (string url in urls_years_802)
            {
                // Создаем лист
                IWorksheet wSheet = wBook.Worksheets[ws];
                wSheet.Name = "Форма 802 (" + url.Substring(url.Length - 6, 4) + ")";

                ParserTable(GetHTML(mail_url + url), wSheet, false); // парсим и заполняем лист в excel 

                ws++;

                backgroundWorker.ReportProgress((ws * 100) / (urls_years_802.Count + urls_years_101.Count)); // увеличиваем процесс сохранения

                if (backgroundWorker.CancellationPending)// если закрыли окно
                {
                    e.Cancel = true; // сообщаем что была отмена
                    return;
                }
            }
            foreach (string url in urls_years_101)
            {
                // Создаем лист
                IWorksheet wSheet = wBook.Worksheets[ws];
                wSheet.Name = "Форма 101 (" + url.Substring(url.Length - 10, 4) + ")";

                ParserTable(GetHTML(mail_url + url), wSheet, true); // парсим и заполняем лист в excel 

                ws++;

                backgroundWorker.ReportProgress((ws * 100) / (urls_years_802.Count + urls_years_101.Count)); // увеличиваем процесс сохранения

                if (backgroundWorker.CancellationPending)// если закрыли окно
                {
                    e.Cancel = true; // сообщаем что была отмена
                    return;
                }
            }

            // Сохраняем файл Excel
            _excel.Save(FileName);
        }

        // завершаем процесс парсинга и сохранения в excel
        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Error != null) // если была ощибка
            {
                MessageBox.Show(e.Error.Message);
                progressBar1.Value = 0;
                button1.Enabled = true;
                return;
            }

            if (!e.Cancelled) // если не отменили выводим сообщение
            {
                MessageBox.Show("Файл сохранен.", "Сохранено", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            progressBar1.Value = 0;
            button1.Enabled = true;
        }
    }
}
