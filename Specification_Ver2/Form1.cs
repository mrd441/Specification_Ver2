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
using Independentsoft.Office.Charts;

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
            //textBox0.Text = "D:\\VisualStudio\\source\\Specification_Ver2\\testInput\\для программиста\\Дахадаевские РЭС Реестр потребителей.xlsx";

            button1.Enabled = false;
            button2.Enabled = false;
            button3.Enabled = false;
            button4.Enabled = false;
            button5.Enabled = false;
            button6.Enabled = false;
            button7.Enabled = false;
            button8.Enabled = false;
            button9.Enabled = false;

        }

        private async void readFile_Click(object sender, EventArgs e)
        {
            //openFileDialog1.InitialDirectory = "c:\\";
            //openFileDialog1.Filter = "txt files (*.xls)|*.xlsx";
            //openFileDialog1.FilterIndex = 1;
            //openFileDialog1.RestoreDirectory = true;

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                textBox0.Text = openFileDialog1.FileName;
                isLoading(true);
                await Task.Run(() => proc1());
                isLoading(false);
            }            
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

            for (int i = 0; i < listBox1.Items.Count; i++)
                listBox1.SetSelected(i, true);

            for (int i = 0; i < listBoxTypePS.Items.Count; i++)
                listBoxTypePS.SetSelected(i, true);

            for (int i = 0; i < listBoxTypeUSPD.Items.Count; i++)
                listBoxTypeUSPD.SetSelected(i, true);

            if (listBoxOpornayaPS.Items.Count > 0)
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
                            newRow.Add(getFilterName(newRow[4]));
                            if (!newRow[20].Contains(newRow[3]))
                            {
                                loging(2, "№ фидера не соответсвует значению в столбце \"ВЛ-6(10)кВ\" на Листе1 в строке " + (i + 1).ToString());
                            }
                            //newRow.Add("Фидер №" + newRow[3]);    // добавить столбец "Фидер 10кВ" 
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
                            newRow.Add(getFilterName(newRow[4])); //номер фидера
                            if (!newRow[21].Contains(newRow[3]))
                            {
                                loging(2, "№ фидера не соответсвует значению в столбце \"Номер фидера 6(10) кВ\" на Листе2 в строке " + (i + 1).ToString());
                            }
                            if (!filtrTypeUSPD.Contains(newRow[20])) filtrTypeUSPD.Add(newRow[20]);
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

        private string getFilterName(string tp)
        {
            //ТП 1 - 1 / 40
            string fideName = "-1";
            try
            {
                fideName = tp.Trim(' ').Substring(3, 1);
            }
            catch
            {
                loging(2, "Не удалось установить номер фидера по ТП " + tp);
            }

            if (fideName == "-1")
                return "Ошибка !!!";
            else
                return "Фидер №" + fideName;
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
            string fileName = textBox0.Text.Substring(s, p - s);
            string psName = "";
            if (listBoxOpornayaPS.SelectedItems.Count == 1)
                psName = listBoxOpornayaPS.SelectedItems[0].ToString();
            textBox1.Text = path + fileName + "\\Приложение №1 Реестр ТП " + psName + ".xlsx";
            textBox2.Text = path + fileName + "\\Приложение №2 Реестр потребителей " + psName + ".xlsx";
            textBox3.Text = path + fileName + "\\Приложение №3 Количество неопрашиваемых ПУ по ТП " + psName + ".xlsx";
            textBox4.Text = path + fileName + "\\Приложение №4 Варианты установки ПУ по фидерам " + psName + ".xlsx";
            textBox5.Text = path + fileName + "\\Приложение №5 Варианты установки ПУ по ТП " + psName + ".xlsx";
            textBox6.Text = path + fileName + "\\Приложение №6 Варианты установки по ПС " + psName + ".xlsx";
            textBox7.Text = path + fileName + "\\Приложение №7 Установка УСПД " + psName + ".xlsx";
            textBox8.Text = path + fileName + "\\Приложение №8 Количество ТТ по фидерам Опорных подстанций " + psName + ".xlsx";
            textBox9.Text = path + fileName + "\\Спецификация " + psName + ".xlsx";

            //string newFileFullName = tmpDirName + "\\";// + tmpFileName.Replace(".xlsx", "_" + city + ".xlsx");
            //if (ttError)
            //    newFileFullName = newFileFullName + "!!";
            //newFileFullName = newFileFullName + city + " СО.xlsx";//tmpFileName.Replace(".xlsx", "_" + city + ".xlsx");

            button1.Enabled = true;
            button2.Enabled = true;
            button3.Enabled = true;
            button4.Enabled = true;
            button5.Enabled = true;
            button6.Enabled = true;
            button7.Enabled = true;
            button8.Enabled = true;
            button9.Enabled = true;
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
                newRow.Add(aRow[21]);//16
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
            catch { }

            if (r == -1)
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
            button9.Enabled = !value;

            textBox1.Enabled = !value;
            textBox2.Enabled = !value;
            textBox3.Enabled = !value;
            textBox4.Enabled = !value;
            textBox5.Enabled = !value;
            textBox6.Enabled = !value;
            textBox7.Enabled = !value;
            textBox8.Enabled = !value;
            textBox9.Enabled = !value;

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
            loging(0, "Экспорта в pdf...");
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
                if (excelWorkbook != null)
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
                sheet1.Columns[0].Width = 14;
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
                    for (int j = 0; j < spec1[i].Count - 1; j++)
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
                //loging(0, "Генерация Приложения1 завершено.");
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
                        arr[i, j] = spec2[i][j];
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
                //loging(0, "Генерация Приложения2 завершено.");
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
                    Microsoft.Office.Interop.Excel.Range range = workSheet.get_Range("A2", "H" + (ii + 2).ToString());
                    range.Value = arr;
                }

                Microsoft.Office.Interop.Excel.Worksheet workSheet2 = (Microsoft.Office.Interop.Excel.Worksheet)excelWorkbook.Worksheets[1];
                Microsoft.Office.Interop.Excel.PivotTable pivot = (Microsoft.Office.Interop.Excel.PivotTable)workSheet2.PivotTables("PivotTable1");
                pivot.SourceData = "Лист1!R1C1:R" + (ii + 2).ToString() + "C8";
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
                        wbcon.CommandText = "Лист1!$A$1:$Z$" + (ii2 + 2).ToString();//"Лист1!$A$1:$Z$1987"
                    else if (ct.Contains("Лист2"))
                        wbcon.CommandText = "Лист2!$A$1:$H$" + (ii + 2).ToString();

                }
                //connex.cast<Microsoft.Office.Interop.Excel.WorkbookConnection>()
                Microsoft.Office.Interop.Excel.PivotTable pivot = (Microsoft.Office.Interop.Excel.PivotTable)workSheet2.PivotTables("PivotTable1");
                //pivot.SourceData = "Лист1!R1C1:R" + NeOprPUList.Count.ToString() + "C8";
                pivot.RefreshTable();
                Microsoft.Office.Interop.Excel.Range rang1 = workSheet2.get_Range("3:3");
                rang1.RowHeight = 63;

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
                Microsoft.Office.Interop.Excel.Range rang1 = workSheet2.get_Range("3:4");
                rang1.RowHeight = 56;

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

                Microsoft.Office.Interop.Excel.PivotTable pivot = (Microsoft.Office.Interop.Excel.PivotTable)workSheet2.PivotTables("PivotTable1");
                Microsoft.Office.Interop.Excel.PivotCache pivotCa = (Microsoft.Office.Interop.Excel.PivotCache)pivot.PivotCache();

                pivot.RefreshTable();
                Microsoft.Office.Interop.Excel.Range rang1 = workSheet2.get_Range("3:4");
                rang1.RowHeight = 56;


                //pivot.RowRange. ("4:4").RowHeight = 56.25;

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
                //object[,] arr = new object[NeOprPUList.Count, 8];
                //int ii = -1;
                //for (int i = 0; i < NeOprPUList.Count; i++)
                //{
                //    if (!filterSpec3(NeOprPUList[i])) continue;
                //    ii++;
                //    for (int j = 0; j < NeOprPUList[i].Count; j++)
                //        arr[ii, j] = NeOprPUList[i][j];
                //}

                //if (ii >= 0)
                //{
                //    Microsoft.Office.Interop.Excel.Range range = workSheet.get_Range("A2", "H" + (ii + 2).ToString());
                //    range.Value = arr;
                //}

                //workSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelWorkbook.Worksheets[3];
                object[,] arr = new object[spec1.Count, 22];
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
                    Microsoft.Office.Interop.Excel.Range range = workSheet.get_Range("A2", "Q" + (ii2 + 2).ToString());
                    range.Value = arr;
                }

                Microsoft.Office.Interop.Excel.Worksheet workSheet2 = (Microsoft.Office.Interop.Excel.Worksheet)excelWorkbook.Worksheets[1];


                //Microsoft.Office.Interop.Excel.Connections connex = (Microsoft.Office.Interop.Excel.Connections)excelWorkbook.Connections;
                //foreach (Microsoft.Office.Interop.Excel.WorkbookConnection connection in excelWorkbook.Connections)
                //{
                //    if (connection.Type != Microsoft.Office.Interop.Excel.XlConnectionType.xlConnectionTypeWORKSHEET) continue;
                //    Microsoft.Office.Interop.Excel.WorksheetDataConnection wbcon = (Microsoft.Office.Interop.Excel.WorksheetDataConnection)connection.WorksheetDataConnection;
                //    string ct = (string)wbcon.CommandText;
                //    if (ct.Contains("Лист1"))
                //        wbcon.CommandText = "Лист1!$A$1:$Q$" + (ii2 + 2).ToString();
                //    //else if (ct.Contains("Лист2"))
                //    //    wbcon.CommandText = "Лист2!$A$1:$H$" + (ii + 2).ToString();

                //}


                Microsoft.Office.Interop.Excel.PivotTable pivot = (Microsoft.Office.Interop.Excel.PivotTable)workSheet2.PivotTables("PivotTable1");
                //Microsoft.Office.Interop.Excel.PivotCache pivotCa = (Microsoft.Office.Interop.Excel.PivotCache)pivot.PivotCache();
                //pivot.SourceData = "Лист1!R1C1:R" + NeOprPUList.Count.ToString() + "C8";
                //pivot.RefreshDataSourceValues();
                //pivotCa.Refresh();
                pivot.SourceData = "Лист1!R1C1:R" + (ii2 + 2).ToString() + "C17";
                pivot.RefreshTable();
                Microsoft.Office.Interop.Excel.Range rang1 = workSheet2.get_Range("3:3");
                rang1.RowHeight = 56;

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

                object[,] arr = new object[spec1.Count, 22];
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
                    Microsoft.Office.Interop.Excel.Range range = workSheet.get_Range("A2", "Q" + (ii2 + 2).ToString());
                    range.Value = arr;
                }

                Microsoft.Office.Interop.Excel.Worksheet workSheet2 = (Microsoft.Office.Interop.Excel.Worksheet)excelWorkbook.Worksheets[1];

                Microsoft.Office.Interop.Excel.PivotTable pivot = (Microsoft.Office.Interop.Excel.PivotTable)workSheet2.PivotTables("PivotTable1");

                pivot.SourceData = "Лист1!R1C1:R" + (ii2 + 2).ToString() + "C17";

                pivot.RefreshTable();
                Microsoft.Office.Interop.Excel.Range rang1 = workSheet2.get_Range("3:3");
                rang1.RowHeight = 42;

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

        private void listBoxOpornayaPS_SelectedIndexChanged(object sender, EventArgs e)
        {
            updateTexboxes();
        }
        ////////////////////////////////////////////////// Спецификация

        public struct ShB_elem
        {
            public ShB_elem(string aNumber, string aName, string aColC, string aColF, int aCount, string aColI)
            {
                Number = aNumber;
                Name = aName;
                ColC = aColC;
                ColD = "";
                ColF = aColF;
                Count = aCount;
                ColI = aColI;
            }
            public string Number { get; set; }
            public string Name { get; }
            public string ColC { get; set; }
            public string ColD { get; set; }
            public string ColF { get; }
            public string ColI { get; }
            public int Count { get; set; }
        };
        public Dictionary<string, List<ShB_elem>> ListShB1;// = new Dictionary<string, ShB1_elem>();
        public Dictionary<string, List<ShB_elem>> ListShB2;// = new Dictionary<string, ShB1_elem>();

        public Dictionary<string, Dictionary<string, string>> shifrs;
        public Dictionary<string, string> tt2List;

        public struct RP_elem
        {
            public RP_elem(string aName, int aCount)
            {
                Name = aName;
                Count = aCount;
            }
            public string Name { get; set; }
            public int Count { get; set; }
        };
        public Dictionary<string, Dictionary<string, List<RP_elem>>> USPD;
        //public Dictionary<string, Dictionary<string, List<RP_elem>>> PU;
        public Dictionary<string, Dictionary<string, Dictionary<string, int>>> PU;
        public Dictionary<string, Dictionary<string, Dictionary<string, Dictionary<string, int>>>> TT;

        public void LoadSettings()
        {
            string s = Environment.CurrentDirectory;
            Microsoft.Office.Interop.Excel.Application excelApplication = null;
            Microsoft.Office.Interop.Excel.Workbook excelWorkbook = null;
            try
            {
                excelApplication = new Microsoft.Office.Interop.Excel.Application();
                excelApplication.ScreenUpdating = false;
                excelApplication.DisplayAlerts = false;

                ListShB1 = new Dictionary<string, List<ShB_elem>>();
                ListShB2 = new Dictionary<string, List<ShB_elem>>();

                excelWorkbook = excelApplication.Workbooks.Open(s + "\\Варианты устройства ТТ, Отвл.xlsx");

                Microsoft.Office.Interop.Excel.Worksheet workSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelWorkbook.Worksheets[1];
                Microsoft.Office.Interop.Excel.Range last = workSheet.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
                var arrData = (object[,])workSheet.get_Range("A1", last).Value;

                workSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelWorkbook.Worksheets[2];
                last = workSheet.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
                var arrData2 = (object[,])workSheet.get_Range("A1", last).Value;


                int rowCount = arrData.GetUpperBound(0);
                int colCount = arrData.GetUpperBound(1);
                if (colCount < 7) throw new Exception("Не верные входные данные ШБ.");

                List<ShB_elem> aList = new List<ShB_elem>();
                string variatName = "";
                for (int i = 1; i <= rowCount; i++)
                {
                    string aName = getStringFromXML(arrData[i, 2]);
                    if (aName != "")
                    {
                        if (aName.Contains("Вариант"))
                        {
                            if (i > 2)
                                ListShB1.Add(variatName, aList);
                            variatName = aName;
                            aList = new List<ShB_elem>();
                        }
                        else
                            aList.Add(new ShB_elem(
                                getStringFromXML(arrData[i, 1]),
                                aName,
                                getStringFromXML(arrData[i, 3]),
                                getStringFromXML(arrData[i, 6]),
                                getIntFromXML(arrData[i, 7]),
                                getStringFromXML(arrData[i, 9])));
                    }
                }
                ListShB1.Add(variatName, aList);

                rowCount = arrData2.GetUpperBound(0);
                colCount = arrData2.GetUpperBound(1);
                if (colCount < 7) throw new Exception("Не верные входные данные ШБ ответвл.");

                aList = new List<ShB_elem>();
                variatName = "";
                for (int i = 1; i <= rowCount; i++)
                {
                    string aName = getStringFromXML(arrData2[i, 2]);
                    if (aName != "")
                    {
                        if (aName.Contains("Вариант"))
                        {
                            if (i > 2)
                                ListShB2.Add(variatName, aList);
                            variatName = aName;
                            aList = new List<ShB_elem>();
                        }
                        else
                            aList.Add(new ShB_elem(
                                getStringFromXML(arrData2[i, 1]),
                                aName,
                                getStringFromXML(arrData2[i, 3]),
                                getStringFromXML(arrData2[i, 6]),
                                getIntFromXML(arrData2[i, 7]),
                                getStringFromXML(arrData2[i, 9])));
                    }
                }
                ListShB2.Add(variatName, aList);
                //loging(1, "Файл успешно загружен.");
            }
            catch (Exception ex)
            {
                loging(2, "Ошибка загрузки Excel файла. " + ex.Message);
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
        private string getStringFromXML(object data)
        {
            string test = "";
            try { test = data.ToString(); }
            catch { }
            return test;
        }

        public void LoadShifrs()
        {
            string s = Environment.CurrentDirectory;
            Microsoft.Office.Interop.Excel.Application excelApplication = null;
            Microsoft.Office.Interop.Excel.Workbook excelWorkbook = null;
            try
            {
                excelApplication = new Microsoft.Office.Interop.Excel.Application();
                excelApplication.ScreenUpdating = false;
                excelApplication.DisplayAlerts = false;

                shifrs = new Dictionary<string, Dictionary<string, string>>();

                excelWorkbook = excelApplication.Workbooks.Open(s + "\\Шифры для состава проекта.xlsx");

                Microsoft.Office.Interop.Excel.Worksheet workSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelWorkbook.Worksheets[6];
                Microsoft.Office.Interop.Excel.Range last = workSheet.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
                var arrData = (object[,])workSheet.get_Range("A1", last).Value;

                int rowCount = arrData.GetUpperBound(0);
                int colCount = arrData.GetUpperBound(1);
                if (colCount < 9) throw new Exception("Не верные вхрдные данные ШБ.");

                for (int i = 1; i <= rowCount; i++)
                {
                    string resName = getStringFromXML(arrData[i, 3]).Trim();
                    string psName = getStringFromXML(arrData[i, 4]).Trim();
                    string tksName = getStringFromXML(arrData[i, 9]).Trim();
                    if (!shifrs.ContainsKey(resName))
                        shifrs.Add(resName, new Dictionary<string, string>());
                    if (!shifrs[resName].ContainsKey(psName))
                        shifrs[resName].Add(psName, tksName);
                }

                //loging(1, "Файл успешно загружен: " + file + ";");
            }
            catch (Exception ex)
            {
                loging(2, "Ошибка загрузки Excel файла. " + ex.Message);
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

        private void generateTt2List()
        {
            generateSpec2();
            generateSpec1();
            tt2List = new Dictionary<string, string>();
            foreach (List<string> aRow in inSheet2)
            {
                if (!filterSpec1(aRow)) continue;
                string ttName = getTypeTT(aRow);
                if (ttName.Contains(">=2 ТТ"))
                {
                    string uspdName = (getTypeUSPD(aRow));
                    string psName = aRow[2];
                    string vaName = aRow[7];
                    string fiName = aRow[21];
                    string key = psName.Trim() + fiName.Trim() + uspdName.Trim();
                    if (!tt2List.ContainsKey(key))
                        tt2List.Add(key, vaName);
                }
            }
            PU = new Dictionary<string, Dictionary<string, Dictionary<string, int>>>();
            foreach (List<string> aRow in spec2)
            {
                string ps = aRow[0];
                string fider = getFilterName(aRow[1]);
                string variant = aRow[15];
                if (!PU.ContainsKey(ps)) PU.Add(ps, new Dictionary<string, Dictionary<string, int>>());
                if (!PU[ps].ContainsKey(fider)) PU[ps].Add(fider, new Dictionary<string, int>());
                if (!PU[ps][fider].ContainsKey(variant)) PU[ps][fider].Add(variant, 0);
                PU[ps][fider][variant] = PU[ps][fider][variant] + 1;
            }

            TT = new Dictionary<string, Dictionary<string, Dictionary<string, Dictionary<string, int>>>>();            
            foreach (List<string> aRow in spec1)
            {
                string ps = aRow[0];
                string fider = aRow[16];
                string variant = aRow[15];
                string tp = aRow[2];
                if (!TT.ContainsKey(ps)) TT.Add(ps, new Dictionary<string, Dictionary<string, Dictionary<string, int>>>());
                if (!TT[ps].ContainsKey(fider)) TT[ps].Add(fider, new Dictionary<string, Dictionary<string, int>>());
                if (!TT[ps][fider].ContainsKey(variant)) TT[ps][fider].Add(variant, new Dictionary<string, int>());
                if (!TT[ps][fider][variant].ContainsKey(tp)) TT[ps][fider][variant].Add(tp, 0);
                TT[ps][fider][variant][tp] = TT[ps][fider][variant][tp] + 1;
            }

        }
        public void GenerateData()
        {

            string s = Environment.CurrentDirectory;
            Microsoft.Office.Interop.Excel.Application excelApplication = null;
            Microsoft.Office.Interop.Excel.Workbook excelWorkbook = null;
            try
            {
                excelApplication = new Microsoft.Office.Interop.Excel.Application();
                excelApplication.ScreenUpdating = false;
                excelApplication.DisplayAlerts = false;

                string templateFileName = Directory.GetCurrentDirectory() + "\\ШаблонСпецификация.xlsx";
                if (!File.Exists(templateFileName))
                    throw new Exception("не найден шаблон выходного файла");


                string tmpFileName = textBox0.Text.Split('\\').Last();

                string resName = tmpFileName.Trim().Replace("_", " ").Replace(" Реестр потребителей.xlsx", "");
                string tmpDirName = textBox0.Text.Replace(".xlsx", "");
                if (!Directory.Exists(tmpDirName))
                    Directory.CreateDirectory(tmpDirName);

                foreach (string city in PU.Keys)
                {
                    bool ttError = false;
                    List<int> caption1List = new List<int>();
                    List<int> caption2List = new List<int>();
                    List<int> londTextList = new List<int>();
                    List<ShB_elem> result = new List<ShB_elem>();
                    foreach (string fider in PU[city].Keys)
                    {
                        result.Add(new ShB_elem("", city + " " + fider, "", "", 0, ""));
                        caption1List.Add(result.Count + incrementIndex(result.Count));

                        foreach (string varName in PU[city][fider].Keys)
                        {
                            int varCount = PU[city][fider][varName];
                            result.Add(new ShB_elem("", varName, "", "", 0, ""));
                            caption2List.Add(result.Count + incrementIndex(result.Count));

                            try
                            {
                                foreach (ShB_elem el in ListShB2[varName.Replace("№", "")])
                                {
                                    ShB_elem newEl = el;
                                    newEl.Count = newEl.Count * varCount;
                                    result.Add(newEl);
                                    if (newEl.ColC.Length > 20) londTextList.Add(result.Count + incrementIndex(result.Count));
                                }
                            }
                            catch (Exception ex)
                            {
                                throw new Exception("Не найден вариант " + varName.Replace("№", "") + "в вариантах устройства ТТ." + ex.Message);
                            }
                        }
                        if (TT.ContainsKey(city) & TT[city].ContainsKey(fider))
                            foreach (string varName2 in TT[city][fider].Keys)
                            {
                                result.Add(new ShB_elem("", varName2, "", "", 0, ""));
                                caption2List.Add(result.Count + incrementIndex(result.Count));

                                int varCount2 = 0;
                                foreach (string aTp in TT[city][fider][varName2].Keys)
                                    varCount2 = varCount2 + TT[city][fider][varName2][aTp];
                                

                                foreach (ShB_elem el in ListShB1[varName2])
                                {
                                    ShB_elem newEl = el;
                                    newEl.Count = newEl.Count * varCount2;
                                    result.Add(newEl);
                                    if (newEl.ColC.Length > 20) londTextList.Add(result.Count + incrementIndex(result.Count));
                                }
                                result.RemoveAt(result.Count - 1);
                                int index = 0;
                                foreach (string aTp in TT[city][fider][varName2].Keys)
                                {
                                    string rpName = "!00";
                                    string vaName = "";
                                    if (aTp.Contains('/'))
                                        rpName = aTp.Replace("/5", "").Replace("А", "");
                                    else
                                    {
                                        vaName = getVaName(city, fider, varName2);
                                        ttError = true;
                                    }
                                    ShB_elem TP = ListShB1[varName2].Last();
                                    TP.ColC = "ТОП-0,66 У3 " + rpName + "/ 5 0,5S";
                                    TP.Number = (Int32.Parse(TP.Number) + index).ToString();
                                    TP.Count = TP.Count * TT[city][fider][varName2][aTp];
                                    TP.ColD = vaName;
                                    result.Add(TP);
                                    if (TP.ColC.Length > 20) londTextList.Add(result.Count + incrementIndex(result.Count));
                                }
                            }
                    }
                    object[,] arr = new object[result.Count, 9];

                    int i = -1;
                    foreach (ShB_elem el in result)
                    {
                        i++;
                        arr[i, 0] = el.Number;
                        arr[i, 1] = el.Name;
                        arr[i, 2] = el.ColC;
                        arr[i, 3] = el.ColD;
                        //arr[i, 4] = el.ColE;
                        arr[i, 5] = el.ColF;
                        arr[i, 6] = el.Count;
                        //arr[i, 7] = el.Number;
                        arr[i, 8] = el.ColI;
                    }

                    double pageCount1 = (result.Count - 24) / 29;
                    double pageCount2 = Math.Ceiling(pageCount1) + 1;
                    double pageCount = (pageCount1 > 0) ? 39 + pageCount2 * 37 : 39;

                    excelWorkbook = excelApplication.Workbooks.Open(templateFileName);
                    Microsoft.Office.Interop.Excel.Worksheet workSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelWorkbook.Worksheets[2];//2

                    workSheet.Cells.ClearContents();

                    Microsoft.Office.Interop.Excel.Range range = workSheet.get_Range("A3", "I" + (result.Count + 2).ToString());
                    range.Value = arr;

                    workSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelWorkbook.Worksheets[1];//3
                    string shtName = city;
                    workSheet.get_Range("Z35").Value = (pageCount2 + 1).ToString();
                    workSheet.get_Range("R34").Value = DateTime.Now.ToString("dd.MM.yyy");
                    workSheet.get_Range("S34").Value = city;
                    workSheet.get_Range("S31").Value = "Создание системы учета в рамках «Плана (Программы) " +
                        "снижения потерь электрической энергии в электрических сетях Республики Дагестан на" +
                        " 2018-2022 годы, реализуемого на объектах филиала ПАО «МРСК Северного Кавказа» - «Дагэнерго». " + resName;
                    //workSheet.get_Range("B5").Value = resName;
                    string shifrName = getShifr(resName, city);
                    workSheet.get_Range("S29").Value = shifrName;
                    if (shifrName == "")
                        loging(2, "не найден шифр для " + resName + " " + city);

                    foreach (int rowNum in caption1List)
                    {
                        range = workSheet.get_Range("J" + rowNum.ToString());
                        range.Font.Bold = true;
                        range.Font.Size = 18;
                    }

                    foreach (int rowNum in caption2List)
                    {
                        range = workSheet.get_Range("J" + rowNum.ToString());
                        range.Font.Bold = true;
                        range.Font.Size = 14;
                    }

                    foreach (int rowNum in londTextList)
                    {
                        range = workSheet.get_Range("K" + rowNum.ToString());
                        range.Font.Size = 8;
                    }    

                    string newFileFullName = textBox9.Text;
                    if (!newFileFullName.Contains("ПС"))
                        newFileFullName = newFileFullName.Replace(".xlsx", city + ".xlsx");

                    excelWorkbook.SaveAs(newFileFullName);

                    if (checkBox9.Checked)
                    {
                        loging(0, "Экспорт в pdf файл...");
                        workSheet.ExportAsFixedFormat(Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF, newFileFullName.Replace(".xlsx", ".pdf"), Microsoft.Office.Interop.Excel.XlFixedFormatQuality.xlQualityStandard, false,true,1, (int)pageCount2 + 1);
                    }

                    if (ttError)
                        loging(2, "Файл сохранен с ошибкой: " + newFileFullName);
                    else
                        loging(0, "Файл успешно сохранен: " + newFileFullName);
                    result.Clear();
                    caption1List.Clear();
                    caption2List.Clear();
                }

            }
            catch (Exception ex)
            {
                loging(2, "Ошибка генерации Спецификации. " + ex.Message);
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

        public int incrementIndex(int rowCount)
        {            
            if (rowCount>24)
            {
                //pageCount = (int)Math.Floor(rowCount/ 29);
                int aa;
                int pageCount = Math.DivRem(rowCount, 29, out aa);
                return (pageCount - 1) * 8 + 17;
            }
            else
            {
                return 2;
            }
        }

        public string getShifr(string resName, string psName)
        {
            resName = resName.Trim();
            psName = psName.Trim();
            if (shifrs.ContainsKey(resName) && shifrs[resName].ContainsKey(psName))
                return shifrs[resName][psName];
            else
                return "";
        }

        public string getVaName(string psName, string fiName, string uspdName)
        {
            string key = psName.Trim() + fiName.Trim() + uspdName.Trim();
            if (tt2List.ContainsKey(key))
                return tt2List[key];
            else
                return "";
        }
        private async void button9_Click(object sender, EventArgs e)
        {
            isLoading(true);
            loging(1, "Генерация Спецификации...");
            loging(0, "Загрузка вариантов устройств...");
            await Task.Run(() => LoadSettings());
            loging(0, "Загрузка шифров...");
            await Task.Run(() => LoadShifrs());
            loging(0, "Сохранение Excel файла...");
            await Task.Run(() => generateTt2List());
            await Task.Run(() => GenerateData());
            loging(1, "Генерация Спецификации завершено.");
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
