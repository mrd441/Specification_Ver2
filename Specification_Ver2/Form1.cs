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

namespace Specification_Ver2
{
    public partial class Form1 : Form
    {
        List<List<string>> inSheet1;
        List<List<string>> inSheet2;
        List<List<string>> spec1;
        List<List<string>> reestr;

        Dictionary<string, Dictionary<string, Dictionary <string, int[]>>> NeOprPU;
        int NeOprPU_Count;

        List<string> filtrOpornayaPS;
        List<string> filtrTypePU;
        List<string> filtrVariantPU;
        List<string> filtrTypeUSPD;

        public Form1()
        {
            InitializeComponent();
            isLoading(false);
        }

        private async void readFile_Click(object sender, EventArgs e)
        {
            isLoading(true);
            await Task.Run(() => proc1());
            isLoading(false);
        }
        private void proc1()
        {
            readFileProc("D:\\VisualStudio\\source\\Specification_Ver2\\testInput\\для программиста\\Книга1.xlsx");
            loadFilters();
            generateSpec1();
        }

        private void loadFilters()
        {
            if (InvokeRequired)
            {
                this.Invoke(new Action(loadFilters));
                return;
            }
            listBox1.Items.AddRange(filtrVariantPU.ToArray());
            listBox2.Items.AddRange(filtrTypePU.ToArray());
            listBox3.Items.AddRange(filtrOpornayaPS.ToArray());
        }
        private void readFileProc(string file)
        {            
            loging(0, "Чтение файла...");
            try
            {
                inSheet1 = new List<List<string>>();
                inSheet2 = new List<List<string>>();

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
                            newRow.Add(""); // добавить столбец "Вариант по фазе"
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
                            if (!filtrOpornayaPS.Contains(newRow[2])) filtrOpornayaPS.Add(newRow[2]);
                            inSheet2.Add(newRow);                            
                        }                    
                }
                loging(0, "Прочитано " + inSheet2.Count.ToString() + " строк из листа " + sheet.Name);                
                loging(0, "Чтение файла завершено");
                if (NeOprPU_Count < inSheet2.Count)
                    loging(2, "Ошибка: на втором листе есть записи не для всех ПУ из первого листа");
            }
            catch(Exception ex)
            {
                loging(2, "Ошибка: " + ex.Message);
            }            
        }
        private void generateSpec2()
        {
            //loging(1, "Генерация Приложения №2...");
            reestr = new List<List<string>>();
            foreach (List<string> aRow in inSheet1)
            {
                List<string> newRow = new List<string>();
                newRow.Add(aRow[0]);
                newRow.Add(aRow[1]);
                newRow.Add(aRow[2]);
                newRow.Add(aRow[3]);
                newRow.Add(aRow[4]);
                newRow.Add(aRow[5]);
                newRow.Add(aRow[6]);
                newRow.Add(aRow[9]);
                newRow.Add(aRow[13]);
                newRow.Add(aRow[14]);
                newRow.Add(aRow[15]);
                newRow.Add(aRow[16]);
                newRow.Add(aRow[17]);
                newRow.Add(aRow[18]);
                newRow.Add(getTypeTT(aRow));
                newRow.Add(aRow[25]);
                reestr.Add(newRow);
            }
            //loging(1, "Генерация Приложения №1 завершена");
        }
        private void generateSpec1()
        {
            loging(1, "Генерация Приложения №1...");
            spec1 = new List<List<string>>();
            foreach(List<string> aRow in inSheet2)
            {
                List<string> newRow = new List<string>();
                newRow.Add(aRow[2]);
                newRow.Add(aRow[3]);
                newRow.Add(aRow[4]);
                newRow.Add(aRow[5]);
                newRow.Add(aRow[6]);
                newRow.Add(aRow[7]);
                newRow.Add(aRow[8]);
                newRow.Add(aRow[9]);
                newRow.Add(aRow[13]);
                newRow.Add(aRow[14]);
                newRow.Add(aRow[15]);
                newRow.Add(aRow[16]);
                newRow.Add(aRow[17]);
                newRow.Add(aRow[18]); 
                newRow.Add(getTypeTT(aRow));
                newRow.Add(getTypeUSPD(aRow));
                spec1.Add(newRow);
            }
            //loging(1, "Генерация Приложения №1 завершена");
        }
        private void add_NeOprPU(List<string> aRow)
        {
            

            string res = aRow[1];
            string opor = aRow[2];
            string tp = aRow[4];

            if (!NeOprPU.ContainsKey(res)) NeOprPU.Add(res,new Dictionary<string, Dictionary<string, int[]>>());
            if (!NeOprPU[res].ContainsKey(opor)) NeOprPU[res].Add(opor, new Dictionary<string, int[]>());
            if (!NeOprPU[res][opor].ContainsKey(tp))
            {
                NeOprPU_Count++;
                NeOprPU[res][opor].Add(tp, new int[] { 0, 0 });
            }
            NeOprPU[res][opor][tp][0]++;
            if (aRow[22] == "1")
                NeOprPU[res][opor][tp][1] ++;
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
            catch(Exception ex)
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

        private void button1_Click(object sender, EventArgs e)
        {
            generateSpec1();
            Worksheet sheet1 = new Worksheet();

            sheet1["A1"] = new Cell("Опорная ПС");
            sheet1["B1"] = new Cell("Номер фидера 6(10) кВ");
            sheet1["C1"] = new Cell("Номер  ТП 6(10)/0,4 кВ");
            sheet1["D1"] = new Cell("Тип ТП");
            sheet1["E1"] = new Cell("Кол-во силовых трансформаторов");
            sheet1["F1"] = new Cell("Мощность кВА");
            sheet1["G1"] = new Cell("Кол-во отходящих фидеров 0,4");
            sheet1["H1"] = new Cell("Тип и уставка автоматического выключателя или ток плавкой вставки предохранителя, в А");
            sheet1["I1"] = new Cell("Населенный пункт");
            sheet1["J1"] = new Cell("Улица");
            sheet1["K1"] = new Cell("Дом");
            sheet1["L1"] = new Cell("Балансовая принадлежность");
            sheet1["M1"] = new Cell("Широта");
            sheet1["N1"] = new Cell("Долгота");
            sheet1["O1"] = new Cell("Тип ТТ");
            sheet1["P1"] = new Cell("Тип УСПД");

            for (int i=0; i<spec1.Count; i++)
            {
                for (int j = 0; i < spec1[i].Count; i++)
                {
                    sheet1.Rows[i + 1].Cells[j] = new Cell(spec1[i][j]);
                }                
            }

            Table table1 = new Table();
            table1.ID = 1;
            table1.Name = "Table1";
            table1.DisplayName = "Table1";
            table1.Reference = "A1:D4";
            table1.AutoFilter = new AutoFilter("A1:D4");

            TableColumn tableColumn1 = new TableColumn(1, "Column1");
            TableColumn tableColumn2 = new TableColumn(2, "Column2");
            TableColumn tableColumn3 = new TableColumn(3, "Column3");
            TableColumn tableColumn4 = new TableColumn(4, "Column4");

            table1.Columns.Add(tableColumn1);
            table1.Columns.Add(tableColumn2);
            table1.Columns.Add(tableColumn3);
            table1.Columns.Add(tableColumn4);

            sheet1.Tables.Add(table1);

            //set columns width
            Column columnInfo = new Column();
            columnInfo.FirstColumn = 1; //from column A
            columnInfo.LastColumn = 4; //to column D
            columnInfo.Width = 15;

            sheet1.Columns.Add(columnInfo);

            Workbook book = new Workbook();
            book.Sheets.Add(sheet1);

            book.Save("D:\\VisualStudio\\source\\Specification_Ver2\\testInput\\output.xlsx", true);
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
