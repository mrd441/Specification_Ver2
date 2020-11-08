using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Independentsoft.Office.Spreadsheet;
using Independentsoft.Office.Spreadsheet.Tables;
using Independentsoft.Office.Spreadsheet.Styles;
using Independentsoft.Office.Odf.Styles;
using System.IO;

namespace Specification_Ver2
{
    public partial class Form1 : Form
    {
        List<List<string>> inSheet1;
        List<List<string>> inSheet2;
        List<List<string>> spec1;
        List<List<string>> spec2;

        List<List<string>> NeOprPUList;
        Dictionary<string, Dictionary<string, Dictionary<string, int[]>>> NeOprPU;
        int NeOprPU_Count;

        List<string> filtrOpornayaPS;
        List<string> filtrTypePU;
        List<string> filtrVariantPU;
        List<string> filtrTypeUSPD;

        public Form1()
        {
            InitializeComponent();
            isLoading(false);
            textBox0.Text = "D:\\work\\VisualStudio\\Specification_Ver2\\testInput\\Дахадаевские РЭС Реестр потребителей.xlsx";

            button1.Enabled = false;
            button2.Enabled = false;
            button3.Enabled = false;
            button4.Enabled = false;
            button5.Enabled = false;
            button6.Enabled = false;
            button7.Enabled = false;
            button8.Enabled = false;
        }

        private async void readFile_Click(object sender, EventArgs e)
        {
            isLoading(true);
            await Task.Run(() => proc1());
            isLoading(false);
        }
        private void proc1()
        {
            readFileProc(textBox0.Text);
            loadFilters();            
        }

        private void loadFilters()
        {
            if (InvokeRequired)
            {
                this.Invoke(new Action(loadFilters));
                return;
            }
            listBox1.Items.Clear();
            listBoxTypePS.Items.Clear();
            listBoxOpornayaPS.Items.Clear();
            listBoxTypeUSPD.Items.Clear();
            
            listBox1.Items.AddRange(filtrVariantPU.ToArray());
            listBoxTypePS.Items.AddRange(filtrTypePU.ToArray());
            listBoxOpornayaPS.Items.AddRange(filtrOpornayaPS.ToArray());
            listBoxTypeUSPD.Items.AddRange(filtrTypeUSPD.ToArray());

            for (int i = 0; i< listBox1.Items.Count; i++)
                listBox1.SetSelected(i, true);

            for (int i = 0; i < listBoxTypePS.Items.Count; i++)
                listBoxTypePS.SetSelected(i, true);

            for (int i = 0; i < listBoxTypeUSPD.Items.Count; i++)
                listBoxTypeUSPD.SetSelected(i, true);

            if (listBoxOpornayaPS.Items.Count>0)
                listBoxOpornayaPS.SetSelected(0, true);

        }
        private void readFileProc(string file)
        {
            loging(0, "Чтение файла...");
            try
            {
                inSheet1 = new List<List<string>>();
                inSheet2 = new List<List<string>>();

                NeOprPUList = new List<List<string>>();
                NeOprPU = new Dictionary<string, Dictionary<string, Dictionary<string, int[]>>>();
                NeOprPU_Count = 0;

                filtrOpornayaPS = new List<string>();
                filtrTypePU = new List<string>();
                filtrVariantPU = new List<string>();
                filtrTypeUSPD = new List<string>();

                Workbook book = new Workbook(file);

                Sheet sheet = book.Sheets[0];
                if (sheet is Worksheet)
                {
                    Worksheet worksheet = (Worksheet)sheet;
                    for (int i = 2; i < worksheet.Rows.Count; i++)
                        if (worksheet.Rows[i] != null)
                        {
                            List<string> newRow = new List<string>();
                            for (int j = 0; j < 20; j++)
                            {
                                if (j < worksheet.Rows[i].Cells.Count)
                                    newRow.Add(getStringCell(worksheet.Rows[i].Cells[j]));
                                else
                                    newRow.Add("");
                            }
                            newRow.Add("Фидер №" + newRow[3]);    // добавить столбец "Фидер 10кВ" 
                            newRow.Add(getVariantPoFaze(newRow)); // добавить столбец "Вариант по фазе"
                            newRow.Add(getNeOprPU(newRow)); // добавить столбец "Неопрашиваемый ПУ"
                            newRow.Add(""); // добавить столбец "Условия для вариантов"
                            newRow.Add(""); // добавить столбец "Вариант по типу"
                            newRow.Add(getVariantPU(newRow)); // добавить столбец "Вариант установки ПУ"

                            if (!filtrTypePU.Contains(newRow[14])) filtrTypePU.Add(newRow[14]);
                            if (!filtrVariantPU.Contains(newRow[25])) filtrVariantPU.Add(newRow[25]);
                            add_NeOprPU(newRow);
                            inSheet1.Add(newRow);
                        }
                }
                updateNeOprPUList();
                loging(0, "Прочитано " + inSheet1.Count.ToString() + " строк из листа " + sheet.Name);

                sheet = book.Sheets[1];
                if (sheet is Worksheet)
                {
                    Worksheet worksheet = (Worksheet)sheet;
                    for (int i = 1; i < worksheet.Rows.Count; i++)
                        if (worksheet.Rows[i] != null)
                        {
                            List<string> newRow = new List<string>();
                            for (int j = 0; j < 20; j++)
                            {
                                if (j < worksheet.Rows[i].Cells.Count)
                                    newRow.Add(getStringCell(worksheet.Rows[i].Cells[j]));
                                else
                                    newRow.Add("");
                            }
                            newRow.Add(getTypeUSPD(newRow));
                            if (!filtrTypeUSPD.Contains(newRow.Last())) filtrTypeUSPD.Add(newRow.Last());
                            if (!filtrOpornayaPS.Contains(newRow[2])) filtrOpornayaPS.Add(newRow[2]);
                            inSheet2.Add(newRow);
                        }
                }
                loging(0, "Прочитано " + inSheet2.Count.ToString() + " строк из листа " + sheet.Name);
                loging(0, "Чтение файла завершено");
                if (NeOprPU_Count < inSheet2.Count)
                    loging(2, "Ошибка: на втором листе есть записи не для всех ПУ из первого листа");

                updateTexboxes();
            }
            catch (Exception ex)
            {
                loging(2, "Ошибка: " + ex.Message);
            }
        }
        private void updateNeOprPUList()
        {
            foreach (List<string> aRow in NeOprPUList)
            {
                int a0 = NeOprPU[aRow[0]][aRow[1]][aRow[2]][0];
                int a1 = NeOprPU[aRow[0]][aRow[1]][aRow[2]][1];
                string a2 = ((a1 * 100) / a0).ToString() + "%";
                aRow.Add(a0.ToString());
                aRow.Add(a1.ToString());
                aRow.Add(a2);
            }
        }
        private void updateTexboxes()
        {
            if (InvokeRequired)
            {
                this.Invoke(new Action(updateTexboxes));
                return;
            }
            int s = textBox0.Text.LastIndexOf('\\') + 1;
            int p = textBox0.Text.LastIndexOf('.');
            string path = textBox0.Text.Substring(0, s);
            string fileName = textBox0.Text.Substring(s, p-s);
            textBox1.Text = path + fileName + "\\Приложение1.xlsx";
            textBox2.Text = path + fileName + "\\Приложение2.xlsx";
            textBox3.Text = path + fileName + "\\Приложение3.xlsx";
            textBox4.Text = path + fileName + "\\Приложение4.xlsx";
            textBox5.Text = path + fileName + "\\Приложение5.xlsx";
            textBox6.Text = path + fileName + "\\Приложение6.xlsx";
            textBox7.Text = path + fileName + "\\Приложение7.xlsx";
            textBox8.Text = path + fileName + "\\Спецификация.xlsx";
            button1.Enabled = true;
            button2.Enabled = true;
            button3.Enabled = true;
            button4.Enabled = true;
            button5.Enabled = true;
            button6.Enabled = true; 
            button7.Enabled = true;
            button8.Enabled = true;
        }
        private void generateSpec2()
        {
            
            spec2 = new List<List<string>>();
            foreach (List<string> aRow in inSheet1)
            {
                if (!filterSpec2(aRow)) continue;
                List<string> newRow = new List<string>();
                newRow.Add(aRow[2]);
                newRow.Add(aRow[4]);
                newRow.Add(aRow[5]);
                newRow.Add(aRow[6]);
                newRow.Add(aRow[7]);
                newRow.Add(aRow[8]);
                newRow.Add(aRow[9]);
                newRow.Add(aRow[10]);
                newRow.Add(aRow[11]);
                newRow.Add(aRow[12]);
                newRow.Add(aRow[13]);
                newRow.Add(aRow[14]);
                newRow.Add(aRow[16]);
                newRow.Add(convertCoord(aRow[17]));
                newRow.Add(convertCoord(aRow[18]));
                newRow.Add(aRow[25]);
                spec2.Add(newRow);
            }
            //loging(1, "Генерация Приложения №1 завершена");
        }
        private void generateSpec1()
        {
            
            spec1 = new List<List<string>>();
            foreach (List<string> aRow in inSheet2)
            {
                if (!filterSpec1(aRow)) continue;
                List<string> newRow = new List<string>();
                newRow.Add(aRow[2]); //1  Опорная ПС
                newRow.Add(aRow[3]); //2
                newRow.Add(aRow[4]); //3
                newRow.Add(aRow[5]); //4
                newRow.Add(aRow[6]); //5
                newRow.Add(aRow[7]); //6
                newRow.Add(aRow[8]); //7
                newRow.Add(aRow[9]); //8
                newRow.Add(aRow[13]);//9
                newRow.Add(aRow[14]);//10
                newRow.Add(aRow[15]);//11
                newRow.Add(aRow[16]);//12
                newRow.Add(convertCoord(aRow[17]));//13
                newRow.Add(convertCoord(aRow[18]));//14 
                newRow.Add(getTypeTT(aRow));  //15
                newRow.Add(getTypeUSPD(aRow));//16
                spec1.Add(newRow);
            }
            //loging(1, "Генерация Приложения №1 завершена");
        }
        private bool filterSpec1(List<string> aRow)
        {
            if (InvokeRequired)
            {
                bool foo = false;
                this.Invoke((MethodInvoker)delegate
                {
                    foo = filterSpec1(aRow);
                });
                return foo;
            }
            decimal neopr = getPercentNeopr(aRow[1], aRow[2], aRow[4]);
            return 
                listBoxOpornayaPS.SelectedItems.Contains(aRow[2]) && 
                listBoxTypeUSPD.SelectedItems.Contains(aRow[20]) &&
                neopr >= neOprFrom.Value &&
                neopr <= neOprTo.Value;
        }
        private bool filterSpec2(List<string> aRow)
        {
            if (InvokeRequired)
            {
                bool foo = false;
                this.Invoke((MethodInvoker)delegate
                {
                    foo = filterSpec2(aRow);
                });
                return foo;
            }
            decimal neopr = getPercentNeopr(aRow[1], aRow[2], aRow[4]);
            return
                listBoxOpornayaPS.SelectedItems.Contains(aRow[2]) &&
                listBoxTypePS.SelectedItems.Contains(aRow[14]) &&
                listBox1.SelectedItems.Contains(aRow[25]) &&
                neopr >= neOprFrom.Value &&
                neopr <= neOprTo.Value;
        }
        private bool filterSpec3(List<string> aRow)
        {
            if (InvokeRequired)
            {
                bool foo = false;
                this.Invoke((MethodInvoker)delegate
                {
                    foo = filterSpec3(aRow);
                });
                return foo;
            }
            decimal neopr = getPercentNeopr(aRow[0], aRow[1], aRow[2]);
            int neopCount = getCountNeopr(aRow[0], aRow[1], aRow[2]);
            return
                listBoxOpornayaPS.SelectedItems.Contains(aRow[1]) &&
                neopr >= neOprFrom.Value &&
                neopr <= neOprTo.Value &&
                neopCount >= neOprCountFrom.Value &&
                neopCount <= neOprCountTo.Value;
        }
        private bool filterSpec4(List<string> aRow)
        {
            if (InvokeRequired)
            {
                bool foo = false;
                this.Invoke((MethodInvoker)delegate
                {
                    foo = filterSpec4(aRow);
                });
                return foo;
            }
            decimal neopr = getPercentNeopr(aRow[1], aRow[2], aRow[4]);
            int neopCount = getCountNeopr(aRow[1], aRow[2], aRow[4]);
            return
                listBoxOpornayaPS.SelectedItems.Contains(aRow[2]) &&
                neopr >= neOprFrom.Value &&
                neopr <= neOprTo.Value &&
                neopCount >= neOprCountFrom.Value &&
                neopCount <= neOprCountTo.Value;
        }
        private string convertCoord(string coord)
        {
            string newCoord = coord.Replace('.', ',');
            int pos = newCoord.IndexOf(',');
            if (pos > -1 && coord.Length > pos + 1 + 6) //6 символов после точки
                newCoord = newCoord.Substring(0, pos + 1 + 6);
            return newCoord;
        }
        private void add_NeOprPU(List<string> aRow)
        {
            string res = aRow[1]; //Наименование РЭС
            string opor = aRow[2];//Опорная ПС
            string tp = aRow[4];  //№ ТП

            if (!NeOprPU.ContainsKey(res)) NeOprPU.Add(res, new Dictionary<string, Dictionary<string, int[]>>()); //
            if (!NeOprPU[res].ContainsKey(opor)) NeOprPU[res].Add(opor, new Dictionary<string, int[]>());
            if (!NeOprPU[res][opor].ContainsKey(tp))
            {
                NeOprPU_Count++;
                NeOprPU[res][opor].Add(tp, new int[] { 0, 0 });

                List<string> NeOprPURow = new List<string>();
                NeOprPURow.Add(aRow[1]);
                NeOprPURow.Add(aRow[2]);
                NeOprPURow.Add(aRow[4]);
                NeOprPURow.Add(aRow[6]);
                NeOprPURow.Add(aRow[20]);                
                NeOprPUList.Add(NeOprPURow);
            }
            NeOprPU[res][opor][tp][0]++;
            if (aRow[22] == "1")
                NeOprPU[res][opor][tp][1]++;
        }
        private decimal getPercentNeopr(string res, string ps, string tp)
        {
            decimal r = -1;
            try
            {
                r = ((decimal)NeOprPU[res][ps][tp][1] / (decimal)NeOprPU[res][ps][tp][0]) * (decimal)100;
            }
            catch{}

            if (r==-1)
                loging(2, "Ошибка. Не найдет процент неопроса для комбинации: " + res + "; " + ps + "; " + tp);
            return r;
        }
        private int getCountNeopr(string res, string ps, string tp)
        {
            int r = -1;
            try
            {
                r = NeOprPU[res][ps][tp][1];
            }
            catch { }

            if (r == -1)
                loging(2, "Ошибка. Не найдет процент неопроса для комбинации: " + res + "; " + ps + "; " + tp);
            return r;
        }

        private string getVariantPoFaze(List<string> aRow) //Вариант ПУ
        {
            string value14 = aRow[14]; //Тип прибора учёта(ПУ)            

            string var1 = "";
            if (value14.Contains("1"))
                return "1";
            else if (value14.Contains("3"))
                return "2";
            else
            {
                loging(2, "Ошибка: не удалось определить Вариант ПУ строки № п/п " + aRow[0] + ". \"Тип прибора учета\" не содержит 1 или 3");
                return "Ошибка !!!";
            }
        }

        private string getVariantPU(List<string> aRow) //Вариант ПУ
        {
            string value12 = aRow[12]; //Тип ввода
            string value13 = aRow[13]; //Крепление отвода абонента
            string value19 = aRow[19]; //Магистраль

            string value14 = aRow[14]; //Тип прибора учёта(ПУ)            

            string var1 = "";
            if (value14.Contains("1"))
                var1 = "1";
            else if (value14.Contains("3"))
                var1 = "2";
            else
            {
                loging(2, "Ошибка: не удалось определить Вариант ПУ строки № п/п " + aRow[0] + ". \"Тип прибора учета\" не содержит 1 или 3");
                return "Ошибка !!!";
            }

            if (value12.Contains("СИП") && value13.Contains("фасад-кирпич") && value19.Contains("СИП"))
                return "Вариант №" + var1 + ".6";
            else if (value12.Contains("СИП") && value13.Contains("фасад-дерево") && value19.Contains("СИП"))
                return "Вариант №" + var1 + ".6";
            else if (value12.Contains("СИП") && value13.Contains("фасад-дерево") && value19.Contains("АС"))
                return "Вариант №" + var1 + ".5";
            else if (value12.Contains("СИП") && value13.Contains("фасад-кирпич") && value19.Contains("АС"))
                return "Вариант №" + var1 + ".5";
            else if (value12.Contains("АС") && value13.Contains("фасад-дерево") && value19.Contains("СИП"))
                return "Вариант №" + var1 + ".4";
            else if (value12.Contains("АС") && value13.Contains("фасад-дерево") && value19.Contains("АС"))
                return "Вариант №" + var1 + ".2";
            else if (value12.Contains("АС") && value13.Contains("фасад-кирпич") && value19.Contains("СИП"))
                return "Вариант №" + var1 + ".3";
            else if (value12.Contains("АС") && value13.Contains("фасад-кирпич") && value19.Contains("АС"))
                return "Вариант №" + var1 + ".1";
            else
            {
                loging(2, "Ошибка: не удалось определить Вариант ПУ строки № п/п " + aRow[0] + ". Ошибка в поле \"Тип ввода\" или \"Крепление отвода абонента\" или \"Магистраль\"");
                return "Ошибка !!!";
            }
        }
        private string getNeOprPU(List<string> aRow) //"НеОпрашиваемый ПУ" по столбцу "Тип прибора учёта (ПУ)"
        {
            string value = aRow[14];
            if (value.Contains("Н"))
                return "1";
            else
                return "0";
        }
        private string getTypeUSPD(List<string> aRow) //Тип УСПД 
        {
            string variant = "";
            try
            {
                if (NeOprPU[aRow[1]][aRow[2]][aRow[4]][0] > 1)
                    variant = "1";
                else
                    variant = "2";
            }
            catch (Exception ex)
            {
                loging(2, "Ошибка: не удалось определить Тип УСПД строки № п/п " + aRow[0] + ". На листе 1 нет ни одной записи для комбинации РЭС, ПУ, ТП " + aRow[1] + "; " + aRow[2] + "; " + aRow[4]);
                return "Ошибка !!!";
            }

            if (aRow[10].Contains("Вариант А") || aRow[11].Contains("Вариант А"))
                return "Вариант А" + variant;
            else if (aRow[10].Contains("Вариант Б") || aRow[11].Contains("Вариант Б"))
                return "Вариант Б" + variant;
            else if (aRow[10].Contains("Вариант В") || aRow[11].Contains("Вариант В"))
                return "Вариант В" + variant;
            else if (aRow[10].Contains("Вариант Г") || aRow[11].Contains("Вариант Г"))
                return "Вариант Г" + variant;
            else if (aRow[10] == "РУНН 0,4 кВ" && aRow[11] == "РУНН 0,4 кВ")
                return "Вариант А" + variant;
            else if (aRow[6].Contains("Столбовая") || aRow[6].Contains("Мачтовая"))
                return "Вариант Б" + variant;
            else if (aRow[10] == "В выносном шкафу" && aRow[11] == "В выносном шкафу" && aRow[16] == "Потребитель")
                return "Вариант В" + variant;
            else if (aRow[10] == "В выносном шкафу" && aRow[11] == "В выносном шкафу" && aRow[16] == "АО \"ДСК\"")
                return "Вариант Г" + variant;
            else
            {
                loging(2, "Ошибка: не удалось определить Тип УСПД строки № п/п " + aRow[0]);
                return "Ошибка !!!";
            }
        }
        private string getTypeTT(List<string> aRow) //Тип ТТ 
        {
            string numberTT = aRow[4];
            if (numberTT.Contains("/35") || (numberTT.Contains("/40")))
                return "75/5";
            else if ((numberTT.Contains("/50")) || (numberTT.Contains("/60")) || (numberTT.Contains("/560")) || (numberTT.Contains("/69")))
                return "100/5";
            if (numberTT.Contains("/63") || (numberTT.Contains("/75")))
                return "150/5";
            else if ((numberTT.Contains("/100")) || (numberTT.Contains("/135")))
                return "200/5";
            if (numberTT.Contains("/160"))
                return "300/5";
            else if ((numberTT.Contains("/180")))
                return "400/5";
            if (numberTT.Contains("/240") || (numberTT.Contains("/250")))
                return "500/5";
            else if ((numberTT.Contains("/320")))
                return "600/5";
            if (numberTT.Contains("/400") || (numberTT.Contains("/410")) || (numberTT.Contains("/420")))
                return "800/5";
            else if ((numberTT.Contains("/630")) || (numberTT.Contains("/750")))
                return "1500/5";
            if (numberTT.Contains("/1000"))
                return "2000/5";
            else if ((numberTT.Contains("/1500")))
                return "3000/5";
            else if ((numberTT.Contains("/10")))
                return "20А/5";
            else if ((numberTT.Contains("/25")) || (numberTT.Contains("/16")))
                return "40А/5";
            else if ((numberTT.Contains("/40")))
                return "3000/5";
            else
                return ">=2 ТТ";
        }
        private string getStringCell(Cell data)
        {
            if (data != null)
                return data.Value;
            else
                return "";
        }
        private int getIntFromXML(object data)
        {
            int test = 0;
            try { test = Convert.ToInt32(data.ToString()); }
            catch { }
            return test;
        }
        public void isLoading(bool value)
        {
            if (InvokeRequired)
            {
                this.Invoke(new Action<bool>(isLoading), new object[] { value });
                return;
            }
            pictureBox1.Visible = value;
            readFile.Enabled = !value;
            button1.Enabled = !value;
            button2.Enabled = !value;
            button3.Enabled = !value;
            button4.Enabled = !value;
            button5.Enabled = !value;
            button6.Enabled = !value;
            button7.Enabled = !value;
            button8.Enabled = !value;
            textBox1.Enabled = !value;
            textBox2.Enabled = !value;
            textBox3.Enabled = !value;
            textBox4.Enabled = !value;
            textBox5.Enabled = !value;
            textBox6.Enabled = !value;
            textBox7.Enabled = !value;
            textBox8.Enabled = !value;
            
            listBox1.Enabled = !value;
            listBoxTypePS.Enabled = !value;
            listBoxTypeUSPD.Enabled = !value;
            listBoxOpornayaPS.Enabled = !value;
        }
        public void loging(int level, string text)
        {
            if (InvokeRequired)
            {
                this.Invoke(new Action<int, string>(loging), new object[] { level, text });
                return;
            }
            var aColor = Color.Black;
            if (level == 1)
                aColor = Color.Green;
            else if (level == 2)
                aColor = Color.Red;
            string curentTime = DateTime.Now.TimeOfDay.ToString("hh\\:mm\\:ss");
            logBox.AppendText(curentTime + ": " + text + Environment.NewLine, aColor);
        }
        private void exportToPDF(string workbookPath, string outputPath)
        {
            loging(1, "Начано экспорта в pdf.");
            Microsoft.Office.Interop.Excel.Application excelApplication = null;
            Microsoft.Office.Interop.Excel.Workbook excelWorkbook = null;
            try
            {  
                excelApplication = new Microsoft.Office.Interop.Excel.Application();
                excelApplication.ScreenUpdating = false;
                excelApplication.DisplayAlerts = false;
                excelWorkbook = excelApplication.Workbooks.Open(workbookPath);
                if (excelWorkbook == null)
                    throw new Exception("Не удалось открыть excel файл.");
                excelWorkbook.ExportAsFixedFormat(Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF, outputPath);
            }
            catch (System.Exception ex)
            {
                loging(2, "Ошибка генерации pdf файла. " + ex.Message);
            }
            finally
            {
                if (excelWorkbook!=null)
                    excelWorkbook.Close();
                if (excelApplication != null)
                    excelApplication.Quit();
                excelApplication = null;
                excelWorkbook = null;
            }
            loging(0, "pdf файл успешно сохранен.");

        }
        
        private bool checkDirectory(string dir)
        {
            string path = dir;
            if (dir.Contains('.'))
            {
                int s = dir.LastIndexOf('\\') + 1;
                path = dir.Substring(0, s);
            }
            if (path.Contains('\\') && !Directory.Exists(path))
                Directory.CreateDirectory(path);

            if (Directory.Exists(path))
                return true;
            else
                return false;
        }
        
        private void button1_Click_fuck()
        {
            try
            {
                generateSpec1();
                Worksheet sheet1 = new Worksheet();

                sheet1["A1"] = new Cell("Опорная ПС");//13
                sheet1["B1"] = new Cell("Номер фидера 6(10) кВ");
                sheet1["C1"] = new Cell("Номер ТП 6(10)/0,4 кВ"); //11
                sheet1["D1"] = new Cell("Тип ТП");
                sheet1["E1"] = new Cell("Кол-во силовых трансформаторов");
                sheet1["F1"] = new Cell("Мощность кВА");
                sheet1["G1"] = new Cell("Кол-во отходящих фидеров 0,4");//10
                sheet1["H1"] = new Cell("Тип и уставка автоматического выключателя или ток плавкой вставки предохранителя, в А");//14
                sheet1["I1"] = new Cell("Населенный пункт");
                sheet1["J1"] = new Cell("Улица");
                sheet1["K1"] = new Cell("Дом");
                sheet1["L1"] = new Cell("Балансовая принадлежность");//12
                sheet1["M1"] = new Cell("Широта");
                sheet1["N1"] = new Cell("Долгота");
                sheet1["O1"] = new Cell("Тип ТТ");
                sheet1["P1"] = new Cell("Тип УСПД");//11

                foreach (Cell c in sheet1.Rows[0].Cells)
                {
                    c.Format = new CellFormat();
                    c.Format.Alignment = new CellAlignment();                    
                    c.Format.Alignment.WrapText = true;
                    c.Format.Alignment.VerticalAlignment = Independentsoft.Office.Spreadsheet.Styles.VerticalAlignment.Center;
                    c.Format.Alignment.HorizontalAlignment = Independentsoft.Office.Spreadsheet.Styles.HorizontalAlignment.Center;
                    c.Format.Font = new Independentsoft.Office.Spreadsheet.Font();
                    c.Format.Font.Name = "Times New Roman";
                    c.Format.Font.Size = 10;
                }
                
                for (int i = 0; i < spec1.Count; i++)
                {
                    Row aRow = new Row();
                    for (int j = 0; j < spec1[i].Count; j++)
                    {
                        Cell c = new Cell(spec1[i][j]);                        
                        c.Format = new CellFormat();
                        c.Format.Alignment = new CellAlignment();
                        c.Format.Alignment.WrapText = true;
                        c.Format.Alignment.VerticalAlignment = Independentsoft.Office.Spreadsheet.Styles.VerticalAlignment.Center;
                        c.Format.Alignment.HorizontalAlignment = Independentsoft.Office.Spreadsheet.Styles.HorizontalAlignment.Center;
                        c.Format.Font = new Independentsoft.Office.Spreadsheet.Font();
                        c.Format.Font.Name = "Times New Roman";
                        c.Format.Font.Size = 10;
                        aRow.Cells.Add(c);
                    }
                        
                    sheet1.Rows.Add(aRow);
                }

                //for (int i = 0; i < spec1.Count; i++)
                //    for (char c = 'A'; c <= 'P'; c++)
                //        sheet1[c + (i + 1).ToString()] = new Cell(spec1[i][(int)c - 65]);

                Table table1 = new Table();
                table1.ID = 1;
                table1.Name = "Table1";
                table1.DisplayName = "Table1";
                table1.TotalsRowShown = false;
                table1.Reference = "A1:P" + (spec1.Count + 1).ToString();
                table1.AutoFilter = new AutoFilter("A1:P" + (spec1.Count + 1).ToString());
                table1.Style = new Independentsoft.Office.Spreadsheet.Tables.TableStyle();
                table1.Style.ShowRowStripes = true;
                table1.Style.Name = "TableStyleLight15";

                for (int i = 1; i <= 16; i++)
                {
                    table1.Columns.Add(new TableColumn(i, sheet1[(char)(i + 64) + "1"].Value));
                    
                }

                sheet1.Tables.Add(table1);


                Workbook book = new Workbook();
                book.Sheets.Add(sheet1);
                //sheet1.HeaderFooterSettings.EvenHeader = "Реестр ТП";
                sheet1.HeaderFooterSettings.OddFooter = @"стр. &P / &N";
                sheet1.HeaderFooterSettings.OddHeader = "Приложение №1";
                sheet1.HeaderFooterSettings.AlignMargins = true;
                //sheet1.VmlHeaderFooterObjects[]
                
                sheet1.PageSetupSettings.Orientation = Independentsoft.Office.Spreadsheet.Orientation.Landscape;
                //sheet1.PageSetupSettings.FitToWidth = 45;
                //sheet1.PageMargins = new PageMargins();
                //sheet1.PageMargins.Bottom = 1.5;
                //sheet1.PageMargins.Top = 1;
                //sheet1.PageMargins.Left = 2;
                //sheet1.PageMargins.Right = 1;
                //sheet1.PageMargins.Header = 1;

                sheet1.VerticalPageBreaks.Add(new Independentsoft.Office.Spreadsheet.Break());
                sheet1.VerticalPageBreaks[0].IsManual = true;
                sheet1.VerticalPageBreaks[0].Min = 16;
                sheet1.VerticalPageBreaks[0].Max = 16;
                sheet1.Columns.Add(new Column(1, 1));
                sheet1.Columns[0]. Width = 14;
                sheet1.Columns.Add(new Column(3, 3));
                sheet1.Columns[1].Width = 11;
                sheet1.Columns.Add(new Column(4, 4));
                sheet1.Columns[2].Width = 11;
                sheet1.Columns.Add(new Column(7, 7));
                sheet1.Columns[3].Width = 10;
                sheet1.Columns.Add(new Column(8, 8));
                sheet1.Columns[4].Width = 15;
                sheet1.Columns.Add(new Column(9, 9));
                sheet1.Columns[5].Width = 12;
                sheet1.Columns.Add(new Column(12, 12));
                sheet1.Columns[6].Width = 14.5;
                sheet1.Columns.Add(new Column(13, 14));
                sheet1.Columns[7].Width = 10;
                sheet1.Columns.Add(new Column(16, 16));
                sheet1.Columns[8].Width = 12;

                book.Save("D:\\VisualStudio\\source\\Specification_Ver2\\testInput\\output.xlsx", true);
            }
            catch (Exception er)
            {
                loging(2, "Ошибка при сохранении в файл. " + er.Message);
            }
            
            //exportToPDF("D:\\VisualStudio\\source\\Specification_Ver2\\testInput\\output.xlsx", "D:\\VisualStudio\\source\\Specification_Ver2\\testInput\\output.pdf");


        }

        private void printSpec1()
        {
            string s = Environment.CurrentDirectory;
            try
            {

                Workbook book = new Workbook(s + "\\Шаблон1.xlsx");
                Worksheet sheet = (Worksheet)book.Sheets[0];
                sheet.Rows.Remove(sheet.Rows.Last());
                for (int i = 0; i < spec1.Count; i++)
                {
                    Row aRow = new Row();
                    for (int j = 0; j < spec1[i].Count; j++)
                    {
                        Cell c = new Cell(spec1[i][j]);
                        c.Format = new CellFormat();
                        c.Format.Alignment = new CellAlignment();
                        c.Format.Alignment.WrapText = true;
                        c.Format.Alignment.VerticalAlignment = Independentsoft.Office.Spreadsheet.Styles.VerticalAlignment.Center;
                        c.Format.Alignment.HorizontalAlignment = Independentsoft.Office.Spreadsheet.Styles.HorizontalAlignment.Center;
                        c.Format.Font = new Independentsoft.Office.Spreadsheet.Font();
                        c.Format.Font.Name = "Times New Roman";
                        c.Format.Font.Size = 10;
                        aRow.Cells.Add(c);
                    }

                    sheet.Rows.Add(aRow);
                    sheet.Tables[0].Reference = "A1:P" + (spec1.Count + 1).ToString();
                }
                if (!checkDirectory(textBox1.Text))
                    throw new Exception("Не верный путь для сохранения файла");
                book.Save(textBox1.Text, true);
                loging(0, "Генерация Приложения1 завершено.");
            }
            catch (Exception ex)
            {
                loging(2, "Ошибка генерации Приложения1. " + ex.Message);
                return;
            }
            if (checkBox1.Checked)
                exportToPDF(textBox1.Text, textBox1.Text.Replace(".xlsx", ".pdf"));
        }


        private void printSpec22()
        {
            string s = Environment.CurrentDirectory;
            Microsoft.Office.Interop.Excel.Application excelApplication = null;
            Microsoft.Office.Interop.Excel.Workbook excelWorkbook = null;
            try
            {
                excelApplication = new Microsoft.Office.Interop.Excel.Application();
                excelApplication.ScreenUpdating = false;
                excelApplication.DisplayAlerts = false;
                excelWorkbook = excelApplication.Workbooks.Open(s + "\\Шаблон2.xlsx");
                if (excelWorkbook == null)
                    throw new Exception("Не удалось открыть excel файл.");

                Microsoft.Office.Interop.Excel.Worksheet workSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelWorkbook.Worksheets[1];


                object[,] arr = new object[spec2.Count, 16];

                for (int i = 0; i < spec2.Count; i++)
                {
                    for (int j = 0; j < spec2[i].Count; j++)
                        arr[i,j] = spec2[i][j];
                }
                
                Microsoft.Office.Interop.Excel.Range range = workSheet.get_Range("A2", "P" + (spec2.Count + 1).ToString());
                range.Value = arr;

                loging(0, "Сохранение excel файла...");
                excelWorkbook.SaveAs(textBox2.Text);

                if (checkBox2.Checked)
                {
                    loging(0, "Экспорт в pdf файл...");
                    excelWorkbook.ExportAsFixedFormat(Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF, textBox2.Text.Replace(".xlsx", ".pdf"));
                }
            }
            catch (System.Exception ex)
            {
                loging(2, "Ошибка генерации файла. " + ex.Message);
            }
            finally
            {
                if (excelWorkbook != null)
                    excelWorkbook.Close();
                if (excelApplication != null)
                    excelApplication.Quit();
                excelApplication = null;
                excelWorkbook = null;
                loging(0, "Генерация Приложения2 завершено.");
            }
        }

        private void printSpec3()
        {
            string s = Environment.CurrentDirectory;
            Microsoft.Office.Interop.Excel.Application excelApplication = null;
            Microsoft.Office.Interop.Excel.Workbook excelWorkbook = null;
            try
            {
                excelApplication = new Microsoft.Office.Interop.Excel.Application();
                excelApplication.ScreenUpdating = false;
                excelApplication.DisplayAlerts = false;
                excelWorkbook = excelApplication.Workbooks.Open(s + "\\Шаблон3.xlsx");
                if (excelWorkbook == null)
                    throw new Exception("Не удалось открыть excel файл.");

                Microsoft.Office.Interop.Excel.Worksheet workSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelWorkbook.Worksheets[2];


                object[,] arr = new object[NeOprPUList.Count, 8];

                int ii = -1;
                for (int i = 0; i < NeOprPUList.Count; i++)
                {
                    if (!filterSpec3(NeOprPUList[i])) continue;
                    ii++;
                    for (int j = 0; j < NeOprPUList[i].Count; j++)
                        arr[ii, j] = NeOprPUList[i][j];
                }

                if (ii >= 0)
                {
                    Microsoft.Office.Interop.Excel.Range range = workSheet.get_Range("A2", "H" + (ii+2).ToString());
                    range.Value = arr;
                }

                Microsoft.Office.Interop.Excel.Worksheet workSheet2 = (Microsoft.Office.Interop.Excel.Worksheet)excelWorkbook.Worksheets[1];
                Microsoft.Office.Interop.Excel.PivotTable pivot = (Microsoft.Office.Interop.Excel.PivotTable)workSheet2.PivotTables("PivotTable1");
                pivot.SourceData = "Лист1!R1C1:R" + (ii + 2).ToString() + "C8" ;
                pivot.RefreshTable();

                if (!checkDirectory(textBox3.Text))
                    throw new Exception("Не верный путь для сохранения файла");
                loging(0, "Сохранение excel файла...");
                excelWorkbook.SaveAs(textBox3.Text);

                if (checkBox3.Checked)
                {
                    loging(0, "Экспорт в pdf файл...");
                    workSheet2.ExportAsFixedFormat(Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF, textBox3.Text.Replace(".xlsx", ".pdf"));
                }
            }
            catch (System.Exception ex)
            {
                loging(2, "Ошибка генерации файла. " + ex.Message);
            }
            finally
            {
                if (excelWorkbook != null)
                    excelWorkbook.Close();
                if (excelApplication != null)
                    excelApplication.Quit();
                
                excelApplication = null;
                excelWorkbook = null;
            }
        }

        private void printSpec4()
        {
            string s = Environment.CurrentDirectory;
            Microsoft.Office.Interop.Excel.Application excelApplication = null;
            Microsoft.Office.Interop.Excel.Workbook excelWorkbook = null;
            try
            {
                excelApplication = new Microsoft.Office.Interop.Excel.Application();
                excelApplication.ScreenUpdating = false;
                excelApplication.DisplayAlerts = false;
                excelWorkbook = excelApplication.Workbooks.Open(s + "\\Шаблон4.xlsx");
                if (excelWorkbook == null)
                    throw new Exception("Не удалось открыть excel файл.");

                Microsoft.Office.Interop.Excel.Worksheet workSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelWorkbook.Worksheets[2];
                object[,] arr = new object[NeOprPUList.Count, 8];
                int ii = -1;
                for (int i = 0; i < NeOprPUList.Count; i++)
                {
                    if (!filterSpec3(NeOprPUList[i])) continue;
                    ii++;
                    for (int j = 0; j < NeOprPUList[i].Count; j++)
                        arr[ii, j] = NeOprPUList[i][j];
                }

                if (ii >= 0)
                {
                    Microsoft.Office.Interop.Excel.Range range = workSheet.get_Range("A2", "H" + (ii + 2).ToString());
                    range.Value = arr;
                }

                workSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelWorkbook.Worksheets[3];
                arr = new object[inSheet1.Count, 26];
                int ii2 = -1;
                for (int i = 0; i < inSheet1.Count; i++)
                {
                    if (!filterSpec4(inSheet1[i])) continue;
                    ii2++;
                    for (int j = 0; j < inSheet1[i].Count; j++)
                        arr[ii2, j] = inSheet1[i][j];
                }

                if (ii2 >= 0)
                {
                    Microsoft.Office.Interop.Excel.Range range = workSheet.get_Range("A2", "Z" + (ii2 + 2).ToString());
                    range.Value = arr;
                }
                
                Microsoft.Office.Interop.Excel.Worksheet workSheet2 = (Microsoft.Office.Interop.Excel.Worksheet)excelWorkbook.Worksheets[1];
                //Microsoft.Office.Interop.Excel.PivotCaches pivotCaches = (Microsoft.Office.Interop.Excel.PivotCaches)excelWorkbook.PivotCaches();
                Microsoft.Office.Interop.Excel.Connections connex = (Microsoft.Office.Interop.Excel.Connections)excelWorkbook.Connections;
                foreach (Microsoft.Office.Interop.Excel.WorkbookConnection connection in excelWorkbook.Connections)
                {
                    if (connection.Type != Microsoft.Office.Interop.Excel.XlConnectionType.xlConnectionTypeWORKSHEET) continue;
                    Microsoft.Office.Interop.Excel.WorksheetDataConnection wbcon = (Microsoft.Office.Interop.Excel.WorksheetDataConnection)connection.WorksheetDataConnection;
                    string ct = (string)wbcon.CommandText;
                    if (ct.Contains("Лист1"))
                        wbcon.CommandText = "Лист1!$A$1:$Z$" + (ii2+2).ToString();//"Лист1!$A$1:$Z$1987"
                    else if (ct.Contains("Лист2"))
                        wbcon.CommandText = "Лист2!$A$1:$H$" + (ii + 2).ToString();

                }
                //connex.cast<Microsoft.Office.Interop.Excel.WorkbookConnection>()
                Microsoft.Office.Interop.Excel.PivotTable pivot = (Microsoft.Office.Interop.Excel.PivotTable)workSheet2.PivotTables("PivotTable1");
                //pivot.SourceData = "Лист1!R1C1:R" + NeOprPUList.Count.ToString() + "C8";
                pivot.RefreshTable();

                if (!checkDirectory(textBox4.Text))
                    throw new Exception("Не верный путь для сохранения файла");
                loging(0, "Сохранение excel файла...");
                excelWorkbook.SaveAs(textBox4.Text);

                if (checkBox4.Checked)
                {
                    loging(0, "Экспорт в pdf файл...");
                    workSheet2.ExportAsFixedFormat(Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF, textBox4.Text.Replace(".xlsx", ".pdf"));
                }
            }
            catch (System.Exception ex)
            {
                loging(2, "Ошибка генерации файла. " + ex.Message);
            }
            finally
            {
                if (excelWorkbook != null)
                    excelWorkbook.Close();
                if (excelApplication != null)
                    excelApplication.Quit();

                excelApplication = null;
                excelWorkbook = null;
            }
        }

        private void printSpec5()
        {
            string s = Environment.CurrentDirectory;
            Microsoft.Office.Interop.Excel.Application excelApplication = null;
            Microsoft.Office.Interop.Excel.Workbook excelWorkbook = null;
            try
            {
                excelApplication = new Microsoft.Office.Interop.Excel.Application();
                excelApplication.ScreenUpdating = false;
                excelApplication.DisplayAlerts = false;
                excelWorkbook = excelApplication.Workbooks.Open(s + "\\Шаблон5.xlsx");
                if (excelWorkbook == null)
                    throw new Exception("Не удалось открыть excel файл.");

                Microsoft.Office.Interop.Excel.Worksheet workSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelWorkbook.Worksheets[2];
                object[,] arr = new object[NeOprPUList.Count, 8];
                int ii = -1;
                for (int i = 0; i < NeOprPUList.Count; i++)
                {
                    if (!filterSpec3(NeOprPUList[i])) continue;
                    ii++;
                    for (int j = 0; j < NeOprPUList[i].Count; j++)
                        arr[ii, j] = NeOprPUList[i][j];
                }

                if (ii >= 0)
                {
                    Microsoft.Office.Interop.Excel.Range range = workSheet.get_Range("A2", "H" + (ii + 2).ToString());
                    range.Value = arr;
                }

                workSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelWorkbook.Worksheets[3];
                arr = new object[inSheet1.Count, 26];
                int ii2 = -1;
                for (int i = 0; i < inSheet1.Count; i++)
                {
                    if (!filterSpec4(inSheet1[i])) continue;
                    ii2++;
                    for (int j = 0; j < inSheet1[i].Count; j++)
                        arr[ii2, j] = inSheet1[i][j];
                }

                if (ii2 >= 0)
                {
                    Microsoft.Office.Interop.Excel.Range range = workSheet.get_Range("A2", "Z" + (ii2 + 2).ToString());
                    range.Value = arr;
                }

                Microsoft.Office.Interop.Excel.Worksheet workSheet2 = (Microsoft.Office.Interop.Excel.Worksheet)excelWorkbook.Worksheets[1];
                //Microsoft.Office.Interop.Excel.PivotCaches pivotCaches = (Microsoft.Office.Interop.Excel.PivotCaches)excelWorkbook.PivotCaches();
                Microsoft.Office.Interop.Excel.Connections connex = (Microsoft.Office.Interop.Excel.Connections)excelWorkbook.Connections;
                foreach (Microsoft.Office.Interop.Excel.WorkbookConnection connection in excelWorkbook.Connections)
                {
                    if (connection.Type != Microsoft.Office.Interop.Excel.XlConnectionType.xlConnectionTypeWORKSHEET) continue;
                    Microsoft.Office.Interop.Excel.WorksheetDataConnection wbcon = (Microsoft.Office.Interop.Excel.WorksheetDataConnection)connection.WorksheetDataConnection;
                    string ct = (string)wbcon.CommandText;
                    if (ct.Contains("Лист1"))
                        wbcon.CommandText = "Лист1!$A$1:$Z$" + (ii2 + 2).ToString();//"Лист1!$A$1:$Z$1987"
                    else if (ct.Contains("Лист2"))
                        wbcon.CommandText = "Лист2!$A$1:$H$" + (ii + 2).ToString();

                }
                //connex.cast<Microsoft.Office.Interop.Excel.WorkbookConnection>()
                Microsoft.Office.Interop.Excel.PivotTable pivot = (Microsoft.Office.Interop.Excel.PivotTable)workSheet2.PivotTables("PivotTable1");
                //pivot.SourceData = "Лист1!R1C1:R" + NeOprPUList.Count.ToString() + "C8";
                pivot.RefreshTable();

                if (!checkDirectory(textBox5.Text))
                    throw new Exception("Не верный путь для сохранения файла");
                loging(0, "Сохранение excel файла...");
                excelWorkbook.SaveAs(textBox5.Text);

                if (checkBox5.Checked)
                {
                    loging(0, "Экспорт в pdf файл...");
                    workSheet2.ExportAsFixedFormat(Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF, textBox5.Text.Replace(".xlsx", ".pdf"));
                }
            }
            catch (System.Exception ex)
            {
                loging(2, "Ошибка генерации файла. " + ex.Message);
            }
            finally
            {
                if (excelWorkbook != null)
                    excelWorkbook.Close();
                if (excelApplication != null)
                    excelApplication.Quit();

                excelApplication = null;
                excelWorkbook = null;
            }
        }

        private void printSpec6()
        {
            string s = Environment.CurrentDirectory;
            Microsoft.Office.Interop.Excel.Application excelApplication = null;
            Microsoft.Office.Interop.Excel.Workbook excelWorkbook = null;
            try
            {
                excelApplication = new Microsoft.Office.Interop.Excel.Application();
                excelApplication.ScreenUpdating = false;
                excelApplication.DisplayAlerts = false;
                excelWorkbook = excelApplication.Workbooks.Open(s + "\\Шаблон6.xlsx");
                if (excelWorkbook == null)
                    throw new Exception("Не удалось открыть excel файл.");

                Microsoft.Office.Interop.Excel.Worksheet workSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelWorkbook.Worksheets[2];
                object[,] arr = new object[NeOprPUList.Count, 8];
                int ii = -1;
                for (int i = 0; i < NeOprPUList.Count; i++)
                {
                    if (!filterSpec3(NeOprPUList[i])) continue;
                    ii++;
                    for (int j = 0; j < NeOprPUList[i].Count; j++)
                        arr[ii, j] = NeOprPUList[i][j];
                }

                if (ii >= 0)
                {
                    Microsoft.Office.Interop.Excel.Range range = workSheet.get_Range("A2", "H" + (ii + 2).ToString());
                    range.Value = arr;
                }

                workSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelWorkbook.Worksheets[3];
                arr = new object[inSheet1.Count, 26];
                int ii2 = -1;
                for (int i = 0; i < inSheet1.Count; i++)
                {
                    if (!filterSpec4(inSheet1[i])) continue;
                    ii2++;
                    for (int j = 0; j < inSheet1[i].Count; j++)
                        arr[ii2, j] = inSheet1[i][j];
                }

                if (ii2 >= 0)
                {
                    Microsoft.Office.Interop.Excel.Range range = workSheet.get_Range("A2", "Z" + (ii2 + 2).ToString());
                    range.Value = arr;
                }

                Microsoft.Office.Interop.Excel.Worksheet workSheet2 = (Microsoft.Office.Interop.Excel.Worksheet)excelWorkbook.Worksheets[1];
                //Microsoft.Office.Interop.Excel.PivotCaches pivotCaches = (Microsoft.Office.Interop.Excel.PivotCaches)excelWorkbook.PivotCaches();
                Microsoft.Office.Interop.Excel.Connections connex = (Microsoft.Office.Interop.Excel.Connections)excelWorkbook.Connections;
                foreach (Microsoft.Office.Interop.Excel.WorkbookConnection connection in excelWorkbook.Connections)
                {
                    if (connection.Type != Microsoft.Office.Interop.Excel.XlConnectionType.xlConnectionTypeWORKSHEET) continue;
                    Microsoft.Office.Interop.Excel.WorksheetDataConnection wbcon = (Microsoft.Office.Interop.Excel.WorksheetDataConnection)connection.WorksheetDataConnection;
                    string ct = (string)wbcon.CommandText;
                    if (ct.Contains("Лист1"))
                        wbcon.CommandText = "Лист1!$A$1:$Z$" + (ii2 + 2).ToString();//"Лист1!$A$1:$Z$1987"
                    else if (ct.Contains("Лист2"))
                        wbcon.CommandText = "Лист2!$A$1:$H$" + (ii + 2).ToString();

                }
                //connex.cast<Microsoft.Office.Interop.Excel.WorkbookConnection>()
                Microsoft.Office.Interop.Excel.PivotTable pivot = (Microsoft.Office.Interop.Excel.PivotTable)workSheet2.PivotTables("PivotTable1");
                //pivot.SourceData = "Лист1!R1C1:R" + NeOprPUList.Count.ToString() + "C8";
                pivot.RefreshTable();

                if (!checkDirectory(textBox6.Text))
                    throw new Exception("Не верный путь для сохранения файла");
                loging(0, "Сохранение excel файла...");
                excelWorkbook.SaveAs(textBox6.Text);

                if (checkBox6.Checked)
                {
                    loging(0, "Экспорт в pdf файл...");
                    workSheet2.ExportAsFixedFormat(Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF, textBox6.Text.Replace(".xlsx", ".pdf"));
                }
            }
            catch (System.Exception ex)
            {
                loging(2, "Ошибка генерации файла. " + ex.Message);
            }
            finally
            {
                if (excelWorkbook != null)
                    excelWorkbook.Close();
                if (excelApplication != null)
                    excelApplication.Quit();

                excelApplication = null;
                excelWorkbook = null;
            }
        }

        private void printSpec7()
        {
            string s = Environment.CurrentDirectory;
            Microsoft.Office.Interop.Excel.Application excelApplication = null;
            Microsoft.Office.Interop.Excel.Workbook excelWorkbook = null;
            try
            {
                generateSpec1();
                excelApplication = new Microsoft.Office.Interop.Excel.Application();
                excelApplication.ScreenUpdating = false;
                excelApplication.DisplayAlerts = false;
                excelWorkbook = excelApplication.Workbooks.Open(s + "\\Шаблон7.xlsx");
                if (excelWorkbook == null)
                    throw new Exception("Не удалось открыть excel файл.");

                Microsoft.Office.Interop.Excel.Worksheet workSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelWorkbook.Worksheets[2];
                object[,] arr = new object[NeOprPUList.Count, 8];
                int ii = -1;
                for (int i = 0; i < NeOprPUList.Count; i++)
                {
                    if (!filterSpec3(NeOprPUList[i])) continue;
                    ii++;
                    for (int j = 0; j < NeOprPUList[i].Count; j++)
                        arr[ii, j] = NeOprPUList[i][j];
                }

                if (ii >= 0)
                {
                    Microsoft.Office.Interop.Excel.Range range = workSheet.get_Range("A2", "H" + (ii + 2).ToString());
                    range.Value = arr;
                }

                workSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelWorkbook.Worksheets[3];
                arr = new object[inSheet1.Count, 26];
                int ii2 = -1;
                for (int i = 0; i < spec1.Count; i++)
                {
                    //if (!filterSpec4(inSheet1[i])) continue;
                    ii2++;
                    for (int j = 0; j < spec1[i].Count; j++)
                        arr[ii2, j] = spec1[i][j];
                }

                if (ii2 >= 0)
                {
                    Microsoft.Office.Interop.Excel.Range range = workSheet.get_Range("A2", "P" + (ii2 + 2).ToString());
                    range.Value = arr;
                }

                Microsoft.Office.Interop.Excel.Worksheet workSheet2 = (Microsoft.Office.Interop.Excel.Worksheet)excelWorkbook.Worksheets[1];
                //Microsoft.Office.Interop.Excel.PivotCaches pivotCaches = (Microsoft.Office.Interop.Excel.PivotCaches)excelWorkbook.PivotCaches();
                Microsoft.Office.Interop.Excel.Connections connex = (Microsoft.Office.Interop.Excel.Connections)excelWorkbook.Connections;
                foreach (Microsoft.Office.Interop.Excel.WorkbookConnection connection in excelWorkbook.Connections)
                {
                    if (connection.Type != Microsoft.Office.Interop.Excel.XlConnectionType.xlConnectionTypeWORKSHEET) continue;
                    Microsoft.Office.Interop.Excel.WorksheetDataConnection wbcon = (Microsoft.Office.Interop.Excel.WorksheetDataConnection)connection.WorksheetDataConnection;
                    string ct = (string)wbcon.CommandText;
                    if (ct.Contains("Лист1"))
                        wbcon.CommandText = "Лист1!$A$1:$P$" + (ii2 + 2).ToString();
                    else if (ct.Contains("Лист2"))
                        wbcon.CommandText = "Лист2!$A$1:$H$" + (ii + 2).ToString();

                }
                //connex.cast<Microsoft.Office.Interop.Excel.WorkbookConnection>()
                Microsoft.Office.Interop.Excel.PivotTable pivot = (Microsoft.Office.Interop.Excel.PivotTable)workSheet2.PivotTables("PivotTable1");
                Microsoft.Office.Interop.Excel.PivotCache pivotCa = (Microsoft.Office.Interop.Excel.PivotCache)pivot.PivotCache();
                //pivot.SourceData = "Лист1!R1C1:R" + NeOprPUList.Count.ToString() + "C8";
                //pivot.RefreshDataSourceValues();
                pivotCa.Refresh();
                pivot.RefreshTable();

                if (!checkDirectory(textBox7.Text))
                    throw new Exception("Не верный путь для сохранения файла");
                loging(0, "Сохранение excel файла...");
                excelWorkbook.SaveAs(textBox7.Text);

                if (checkBox7.Checked)
                {
                    loging(0, "Экспорт в pdf файл...");
                    workSheet2.ExportAsFixedFormat(Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF, textBox7.Text.Replace(".xlsx", ".pdf"));
                }
            }
            catch (System.Exception ex)
            {
                loging(2, "Ошибка генерации файла. " + ex.Message);
            }
            finally
            {
                if (excelWorkbook != null)
                    excelWorkbook.Close();
                if (excelApplication != null)
                    excelApplication.Quit();

                excelApplication = null;
                excelWorkbook = null;
            }
        }

        private void printSpec8()
        {
            string s = Environment.CurrentDirectory;
            Microsoft.Office.Interop.Excel.Application excelApplication = null;
            Microsoft.Office.Interop.Excel.Workbook excelWorkbook = null;
            try
            {
                generateSpec1();
                excelApplication = new Microsoft.Office.Interop.Excel.Application();
                excelApplication.ScreenUpdating = false;
                excelApplication.DisplayAlerts = false;
                excelWorkbook = excelApplication.Workbooks.Open(s + "\\Шаблон8.xlsx");
                if (excelWorkbook == null)
                    throw new Exception("Не удалось открыть excel файл.");

                Microsoft.Office.Interop.Excel.Worksheet workSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelWorkbook.Worksheets[2];
                object[,] arr = new object[NeOprPUList.Count, 8];
                int ii = -1;
                for (int i = 0; i < NeOprPUList.Count; i++)
                {
                    if (!filterSpec3(NeOprPUList[i])) continue;
                    ii++;
                    for (int j = 0; j < NeOprPUList[i].Count; j++)
                        arr[ii, j] = NeOprPUList[i][j];
                }

                if (ii >= 0)
                {
                    Microsoft.Office.Interop.Excel.Range range = workSheet.get_Range("A2", "H" + (ii + 2).ToString());
                    range.Value = arr;
                }

                workSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelWorkbook.Worksheets[3];
                arr = new object[inSheet1.Count, 26];
                int ii2 = -1;
                for (int i = 0; i < spec1.Count; i++)
                {
                    //if (!filterSpec4(inSheet1[i])) continue;
                    ii2++;
                    for (int j = 0; j < spec1[i].Count; j++)
                        arr[ii2, j] = spec1[i][j];
                }

                if (ii2 >= 0)
                {
                    Microsoft.Office.Interop.Excel.Range range = workSheet.get_Range("A2", "P" + (ii2 + 2).ToString());
                    range.Value = arr;
                }

                Microsoft.Office.Interop.Excel.Worksheet workSheet2 = (Microsoft.Office.Interop.Excel.Worksheet)excelWorkbook.Worksheets[1];
                //Microsoft.Office.Interop.Excel.PivotCaches pivotCaches = (Microsoft.Office.Interop.Excel.PivotCaches)excelWorkbook.PivotCaches();
                Microsoft.Office.Interop.Excel.Connections connex = (Microsoft.Office.Interop.Excel.Connections)excelWorkbook.Connections;
                foreach (Microsoft.Office.Interop.Excel.WorkbookConnection connection in excelWorkbook.Connections)
                {
                    if (connection.Type != Microsoft.Office.Interop.Excel.XlConnectionType.xlConnectionTypeWORKSHEET) continue;
                    Microsoft.Office.Interop.Excel.WorksheetDataConnection wbcon = (Microsoft.Office.Interop.Excel.WorksheetDataConnection)connection.WorksheetDataConnection;
                    string ct = (string)wbcon.CommandText;
                    if (ct.Contains("Лист1"))
                        wbcon.CommandText = "Лист1!$A$1:$P$" + (ii2 + 2).ToString();
                    else if (ct.Contains("Лист2"))
                        wbcon.CommandText = "Лист2!$A$1:$H$" + (ii + 2).ToString();

                }
                //connex.cast<Microsoft.Office.Interop.Excel.WorkbookConnection>()
                Microsoft.Office.Interop.Excel.PivotTable pivot = (Microsoft.Office.Interop.Excel.PivotTable)workSheet2.PivotTables("PivotTable1");
                Microsoft.Office.Interop.Excel.PivotCache pivotCa = (Microsoft.Office.Interop.Excel.PivotCache)pivot.PivotCache();
                //pivot.SourceData = "Лист1!R1C1:R" + NeOprPUList.Count.ToString() + "C8";
                //pivot.RefreshDataSourceValues();
                pivotCa.Refresh();
                pivot.RefreshTable();

                if (!checkDirectory(textBox8.Text))
                    throw new Exception("Не верный путь для сохранения файла");
                loging(0, "Сохранение excel файла...");
                excelWorkbook.SaveAs(textBox8.Text);

                if (checkBox8.Checked)
                {
                    loging(0, "Экспорт в pdf файл...");
                    workSheet2.ExportAsFixedFormat(Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF, textBox8.Text.Replace(".xlsx", ".pdf"));
                }
            }
            catch (System.Exception ex)
            {
                loging(2, "Ошибка генерации файла. " + ex.Message);
            }
            finally
            {
                if (excelWorkbook != null)
                    excelWorkbook.Close();
                if (excelApplication != null)
                    excelApplication.Quit();

                excelApplication = null;
                excelWorkbook = null;
            }
        }

        private void printSpec2()
        {
            
            string s = Environment.CurrentDirectory;
            try
            {

                Workbook book = new Workbook(s + "\\Шаблон2.xlsx");
                Worksheet sheet = (Worksheet)book.Sheets[0];
                sheet.Rows.Remove(sheet.Rows.Last());
                for (int i = 0; i < spec2.Count; i++)
                {
                    Row aRow = new Row();
                    for (int j = 0; j < spec2[i].Count; j++)
                    {
                        Cell c = new Cell(spec2[i][j]);
                        c.Format = new CellFormat();
                        c.Format.Alignment = new CellAlignment();
                        c.Format.Alignment.WrapText = true;
                        c.Format.Alignment.VerticalAlignment = Independentsoft.Office.Spreadsheet.Styles.VerticalAlignment.Center;
                        c.Format.Alignment.HorizontalAlignment = Independentsoft.Office.Spreadsheet.Styles.HorizontalAlignment.Center;
                        c.Format.Font = new Independentsoft.Office.Spreadsheet.Font();
                        c.Format.Font.Name = "Times New Roman";
                        c.Format.Font.Size = 10;
                        aRow.Cells.Add(c);
                    }

                    sheet.Rows.Add(aRow);                    
                }
                sheet.Tables[0].Reference = "A1:P" + (spec2.Count + 1).ToString();
                if (!checkDirectory(textBox2.Text))
                    throw new Exception("Не верный путь для сохранения файла");
                loging(0, "Сохранение файла Excel...");
                book.Save(textBox2.Text, true);
                loging(0, "Приложения2 успешно сгенерировано.");
            }
            catch (Exception ex)
            {
                loging(2, "Ошибка генерации Приложения2. " + ex.Message);
                return;
            }
            if (checkBox2.Checked)
                exportToPDF(textBox2.Text, textBox2.Text.Replace(".xlsx", ".pdf"));
        
        }

        private async void button1_Click(object sender, EventArgs e)
        {
            isLoading(true);
            loging(1, "Генерация Приложения №1...");
            await Task.Run(() => generateSpec1());
            await Task.Run(() => printSpec1());
            loging(1, "Генерация Приложения №1 завершено.");
            isLoading(false);

        }
        private async void button2_Click(object sender, EventArgs e)
        {
            isLoading(true);
            loging(1, "Генерация Приложения №2...");
            await Task.Run(() => generateSpec2());
            await Task.Run(() => printSpec22());
            loging(1, "Генерация Приложения №2 завершено.");
            isLoading(false);
        }
        private async void button3_Click(object sender, EventArgs e)
        {
            isLoading(true);
            loging(1, "Генерация Приложения №3...");
            //await Task.Run(() => generateSpec2());
            await Task.Run(() => printSpec3());
            loging(1, "Генерация Приложения №3 завершено.");
            isLoading(false);
        }
        private async void button4_Click(object sender, EventArgs e)
        {
            isLoading(true);
            loging(1, "Генерация Приложения №4...");
            //await Task.Run(() => generateSpec2());
            await Task.Run(() => printSpec4());
            loging(1, "Генерация Приложения №4 завершено.");
            isLoading(false);
        }
        private async void button5_Click(object sender, EventArgs e)
        {
            isLoading(true);
            loging(1, "Генерация Приложения №5...");
            //await Task.Run(() => generateSpec2());
            await Task.Run(() => printSpec5());
            loging(1, "Генерация Приложения №5 завершено.");
            isLoading(false);
        }
        private async void button6_Click(object sender, EventArgs e)
        {
            isLoading(true);
            loging(1, "Генерация Приложения №6...");
            //await Task.Run(() => generateSpec2());
            await Task.Run(() => printSpec6());
            loging(1, "Генерация Приложения №6 завершено.");
            isLoading(false);
        }
        private async void button7_Click(object sender, EventArgs e)
        {
            isLoading(true);
            loging(1, "Генерация Приложения №7...");
            await Task.Run(() => printSpec7());
            loging(1, "Генерация Приложения №7 завершено.");
            isLoading(false);
        }

        private async void button8_Click(object sender, EventArgs e)
        {
            isLoading(true);
            loging(1, "Генерация Приложения №8...");
            await Task.Run(() => printSpec8());
            loging(1, "Генерация Приложения №8 завершено.");
            isLoading(false);
        }
    }
    public static class RichTextBoxExtensions
    {
        public static void AppendText(this RichTextBox box, string text, Color color)
        {
            box.SelectionStart = box.TextLength;
            box.SelectionLength = 0;

            box.SelectionColor = color;
            box.AppendText(text);
            box.SelectionColor = box.ForeColor;
            box.SelectionStart = box.Text.Length;
            box.ScrollToCaret();
        }
    }

}
