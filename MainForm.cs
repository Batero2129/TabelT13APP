using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace TabelT13APP
{
    public partial class MainForm : Form
    {
        string[,] _inOutArray;
        string[,] _sortedNames; //Массив по листам и именам
        int SheetValue { get; set; } //количество листов
        List<PersonMaker> fMakers = new List<PersonMaker>(); //список графов на каждую персону
        Excel.Application _workExcel;
        Excel.Workbook _workBook;


        public MainForm()
        {
            InitializeComponent();
            label4.Enabled = false;
            label5.Enabled = false;
            button1.Enabled = false;

        }
        void Form1FormClosing(object sender, FormClosingEventArgs e)
        {
            e.Cancel = fclose();
        }
        void Form1FormClosed(object sender, FormClosedEventArgs e)
        {

            GC.Collect(); // убрать за собой
            Application.Exit();
        }
        public static bool fclose()
        {
            var result = MessageBox.Show("Вы действительно хотите закрыть программу?", "Подтверждение", MessageBoxButtons.OKCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1);
            if (result == DialogResult.OK) return false;
            else return true;
        }
        void LoadDefaultTabel()
        {
            string defaultTabFile = @"C:\TabelT13\T13.xlsx"; // Ссылка на образец табеля------------------------------------------------------------
            Excel.Application WorkExcel = new Excel.Application();
            Excel.Workbook WorkBook = WorkExcel.Workbooks.Open(defaultTabFile);
            Excel.Worksheet WorkSheet = (Excel.Worksheet)WorkBook.Sheets[1];

            int number = 1;

            label6.Visible = true;
            label6.Text = "Подождите идет загрузка..";
            progressBar1.Visible = true;
            progressBar1.Value = 0;
            int value = 100 / _sortedNames.GetLength(0);
            for (int i = 0; i < _sortedNames.GetLength(0); i++) // перебор для каждого листа
            {
                //WorkSheet
                int numRow = 24;
                int numColl = 1;
                int R = 24;
                int C = 49;
                Excel.Worksheet WorkSheetNew = (Excel.Worksheet)WorkBook.Sheets[(i + 1)];

                WorkSheetNew.Cells[13, 149].Value = DateTime.Now.Day.ToString() + "." + DateTime.Now.Year.ToString();

                for (int j = 0; j < _sortedNames.GetLength(1); j++) //перебор для каждого имени
                {
                    PersonMaker person = fMakers.Find(n => n.Name == _sortedNames[i, j]); //нашли фамилию в списке
                    if (person.Name == "" || person.Name == null) break;
                    WorkSheetNew.Cells[numRow, numColl].Value = number;
                    number++;
                    numColl += 8;
                    WorkSheetNew.Cells[numRow, numColl].Value = person.Name;
                    numColl += 104;
                    WorkSheetNew.Cells[numRow, numColl].Value = person.Days1Month;
                    numRow += 1;
                    WorkSheetNew.Cells[numRow, numColl].Value = person.Hours1Month;
                    numRow += 1;
                    WorkSheetNew.Cells[numRow, numColl].Value = person.Days2Month;
                    numRow += 1;
                    WorkSheetNew.Cells[numRow, numColl].Value = person.Hours2Month;
                    numRow -= 3;
                    numColl += 11;
                    WorkSheetNew.Cells[numRow, numColl].Value = person.TotalDays;
                    numRow += 2;
                    WorkSheetNew.Cells[numRow, numColl].Value = person.TotalHours;
                    numRow -= 2;
                    numColl -= 123;

                    for (int k = 1; k <= person.DaysInMonth; k++) //перебор для каждого дня месяца
                    {
                        int day = k;
                        int hours = person.DaysVisit[day];
                        if (day == 16)      //перекидываем на каждое 16-е в начало строки
                        {
                            R += 2;
                            C -= 60;
                        }

                        if (hours != 0)         // запись в ячейку и под ней
                        {
                            WorkSheetNew.Cells[R, C].Value = "Я";
                            R++;
                            WorkSheetNew.Cells[R, C].Value = hours;
                            R--;
                        }
                        else
                        {
                            WorkSheetNew.Cells[R, C].Value = "H";
                        }

                        if (person.DaysInMonth == 28) //проверка на последнее число месяца
                        {
                            if (day == person.DaysInMonth)
                            {
                                C -= 52;
                                R += 2;
                            }
                        }
                        else if (person.DaysInMonth == 29)
                        {
                            if (day == person.DaysInMonth)
                            {
                                C -= 56;
                                R += 2;
                            }

                        }
                        else if (person.DaysInMonth == 30)
                        {
                            if (day == person.DaysInMonth)
                            {
                                C -= 60;
                                R += 2;
                            }
                        }
                        else
                        {
                            if (day == person.DaysInMonth)
                            {
                                C -= 64;
                                R += 2;
                            }
                        }
                        C += 4;
                    }
                    numRow += 4;
                }
                progressBar1.Value += value;
            }
            label5.Visible = true;
            progressBar1.Visible = false;
            label6.Visible = false;

            WorkExcel.Visible = true;
            WorkExcel.UserControl = true;
        }
        private int ImportTableExcel()
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.DefaultExt = "*.xls;*.xlsx";
            ofd.Filter = "файл Excel (Spisok.xlsx)|*.xlsx";
            ofd.Title = "Выберите файл посещений";
            if (!(ofd.ShowDialog() == DialogResult.OK))
                return 0;

            Excel.Application WorkExcel = new Excel.Application();
            Excel.Workbook WorkBook = WorkExcel.Workbooks.Open(ofd.FileName);
            _workExcel = WorkExcel;
            _workBook = WorkBook;
            Excel.Worksheet WorkSheet = (Excel.Worksheet)WorkBook.Sheets[1];
            var lastCell = WorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
            var arr = WorkSheet.UsedRange.Value2;

            int lastColumn = (int)lastCell.Column;
            int lastRow = (int)lastCell.Row;

            float percenBarOneItter = lastRow / 95;
            percenBarOneItter++;

            string[,] table = new string[lastRow, lastColumn];

            groupBox1.Enabled = false;
            label6.Visible = true;
            label6.Text = "Подождите, идет загрузка...";
            label7.Visible = true;
            progressBar1.Visible = true;
            progressBar1.Value = 0;

            for (int i = 0; i < lastRow; i++)
            {
                for (int j = 0; j < lastColumn; j++)
                {
                    table[i, j] = arr[(i + 1), (j + 1)].ToString();
                }
                progressBar1.Value = i / (int)(percenBarOneItter);
                label7.Text = $"Загружено {i} из {lastRow} строк.";
            }
            groupBox1.Enabled = true;
            label6.Visible = false;
            label7.Visible = false;
            progressBar1.Visible = false;
            _inOutArray = table;
            WorkBook.Close(false, Type.Missing, Type.Missing); //закрыть не сохраняя
            WorkExcel.Quit(); // выйти из Excel
            GC.Collect(); // убрать за собой
            return table.GetLength(0);
        }
        void Sorting()
        {
            int ID;
            DateTime dateTime;
            string message;
            string name;

            List<string> allTablesName = new List<string>();
            //SORTING
            for (int i = 0; i < _inOutArray.GetLength(0); i++)
            {
                if (_inOutArray[i, 2].Contains("ВХОД") || _inOutArray[i, 2].Contains("ВЫХОД"))
                {
                    string str = _inOutArray[i, 3];
                    allTablesName.Add(str);
                }
            }
            List<string> names = new List<string>();
            string first = "";
            while (true)
            {
                first = allTablesName[0];
                names.Add(first);
                while (allTablesName.Contains(first))
                {
                    for (int i = 0; i < allTablesName.Count; i++)
                    {
                        if (allTablesName.Count == 0) break;
                        else if (allTablesName[i] == first)
                        {
                            allTablesName.RemoveAt(i);
                            break;
                        }
                    }

                }
                if (allTablesName.Count == 0) break;
            }
            int count = names.Count;
            int garantValue = count / 6;
            int sheetsValue = 0;
            if (count % 6 > 0) sheetsValue = garantValue + 1;
            SheetValue = sheetsValue;
            string[,] sheetsOfNames = new string[sheetsValue, 6];
            int x = 1;
            int sheetLenght = sheetsOfNames.GetLength(0);
            for (int i = 0; i <= sheetLenght; i++) ////////////////////////
            {
                for (int j = 0; j < 6; j++)
                {
                    if (x != count + 1)
                    {
                        string tmpName = names[x - 1];
                        sheetsOfNames[i, j] = tmpName;
                        x++;
                    }
                    else break;
                }
            }
            _sortedNames = sheetsOfNames;
        }
        void GetValuesByNames()// Рассортировать значения по различным спискам согласно именам...возможно использовав класс.
        {
            string[,] array = _inOutArray;
            string[,] sorted = _sortedNames;

            for (int i = 0; i < sorted.GetLength(0); i++) //по листам
                for (int j = 0; j < sorted.GetLength(1); j++) //по имени в каждом листе
                {
                    string name = sorted[i, j]; //взяли одно имя
                    List<string[]> listByOne = new List<string[]>(); //список посещений для одной персоны

                    for (int k = 0; k < array.GetLength(0); k++) //по каждой строке общего списка
                    {
                        if (array[k, 3] == name)
                        {
                            string[] tempArray = new string[4];
                            tempArray[0] = array[k, 0];
                            tempArray[1] = array[k, 1];
                            tempArray[2] = array[k, 2];
                            tempArray[3] = array[k, 3];
                            listByOne.Add(tempArray);
                        }

                    }
                    PersonMaker fMaker = new PersonMaker(listByOne, name);
                    fMakers.Add(fMaker);
                }
        }


        private void button2_Click(object sender, EventArgs e) //Загрузить
        {
            int n = ImportTableExcel();
            if (n != 0)
            {
                label4.Enabled = true;
                button1.Enabled = true;
            }
            else MessageBox.Show("Что-то пошло не так!");

        }
        private void button1_Click(object sender, EventArgs e)  //Заполнить
        {
            Sorting();
            GetValuesByNames();
            LoadDefaultTabel();
        }
        private void button3_Click(object sender, EventArgs e)     //Закрыть

        {
            this.Close();
        }
    }
    /// <summary>
    /// Класс определения одной персоны
    /// </summary>
    public class PersonMaker
    {
        public Dictionary<int, int> DaysVisit = new Dictionary<int, int>();
        public string Name { get; private set; }
        public int Days1Month { get; private set; }
        public int Hours1Month { get; private set; }
        public int Days2Month { get; private set; }
        public int Hours2Month { get; private set; }
        public int TotalDays { get; private set; }
        public int TotalHours { get; private set; }
        public int DaysInMonth { get; private set; }

        public PersonMaker(List<string[]> allRows, string name)
        {
            int daysInMonth = GetMonth(allRows);
            if (daysInMonth != 0)
            {
                Name = name;
                int days1Month = 0;
                Days1Month = days1Month;
                int hours1Month = 0;
                Hours1Month = hours1Month;
                int days2Month = 0;
                Days2Month = days2Month;
                int hours2Month = 0;
                Hours2Month = hours2Month;
                int totalDays = 0;
                TotalDays = totalDays;
                int totalHours = 0;
                TotalHours = totalHours;

                DaysInMonth = daysInMonth;
                SetDaysToDaysVisit(daysInMonth);
                GetDaysVisit(allRows);

                int x = 15;
                int d = 1;
                foreach (var item in DaysVisit)
                {
                    if (d <= 15)
                    {
                        if (item.Value != 0)
                        {
                            Days1Month += 1;
                            Hours1Month += item.Value;
                        }
                        d++;
                    }
                    else
                    {
                        if (item.Value != 0)
                        {
                            Days2Month += 1;
                            Hours2Month += item.Value;
                        }
                        d++;
                    }
                }
                TotalDays = Days1Month + Days2Month;
                TotalHours = Hours1Month + Hours2Month;
            }
            int i = 0;
            i++;///////////////////////////////////
        }
        void SetDaysToDaysVisit(int daysInMonth)
        {
            for (int i = 1; i <= daysInMonth; i++)
            {
                DaysVisit.Add(i, 0);
            }
        }
        void GetDaysVisit(List<string[]> allRows)
        {
            int n = DaysInMonth;
            List<Day> days = new List<Day>();

            for (int i = 1; i <= n; i++)
            {
                int currentDay = i;
                DateTime timeIN = DateTime.MinValue;
                DateTime timeOUT = DateTime.MinValue;
                foreach (var item in allRows)
                {
                    double t = double.Parse(item[1]);
                    DateTime date = DateTime.FromOADate(t);
                    int d = date.Day;
                    if (d == i)
                    {
                        if (item[2].Contains("ВХОД"))
                        {
                            double q = double.Parse(item[1]);
                            timeIN = DateTime.FromOADate(q);
                        }
                        else if (item[2].Contains("ВЫХОД"))
                        {
                            double o = double.Parse(item[1]);
                            timeOUT = DateTime.FromOADate(o);
                        }
                    }
                }
                Day day = new Day(i, timeIN, timeOUT);
                days.Add(day);
                timeIN = DateTime.MinValue;
                timeOUT = DateTime.MinValue;
            }  //Создали список всех дней
            //TODO: Записать в словарь часы посещения;
            foreach (var item in days)
            {
                if (item.WorkHours != 0)
                {
                    DaysVisit[item.NowDay] = item.WorkHours;
                }

            }
        }
        int GetMonth(List<string[]> list)
        {
            if (list.Count != 0)
            {
                if (list.Count > 1)
                {
                    string[] row = list[1];
                    string time = row[1];
                    double d = double.Parse(time);
                    DateTime date = DateTime.FromOADate(d);
                    int daysInMonth = DateTime.DaysInMonth(date.Year, date.Month);
                    return daysInMonth;
                }
                else
                {
                    string[] row = list[0];
                    string time = row[1];
                    double d = double.Parse(time);
                    DateTime date = DateTime.FromOADate(d);
                    int daysInMonth = DateTime.DaysInMonth(date.Year, date.Month);
                    return daysInMonth;
                }
            }
            return 0;
        }
    }
    public class Day
    {
        public int NowDay { get; set; }
        public int WorkHours { get; private set; }

        public Day(int day, DateTime timeIn, DateTime timeOut)
        {
            NowDay = day;
            DateTime none = DateTime.MinValue;

            if (timeIn != none && timeOut != none) // Если есть время прихода и ухода
            {
                if (timeIn < timeOut) //если пришел раньше, чем ушел ДНЕВНАЯ СМЕНА
                {
                    float hI = timeIn.Hour;
                    float mI = timeIn.Minute / 60.0f;
                    float tI = hI + mI;

                    float hO = timeOut.Hour;
                    float mO = timeOut.Minute / 60.0f;
                    float tO = hO + mO;

                    if (6.0f < tI & tI < 8.0f) 
                    {
                        tI = 7.5f;
                    }
                    if (15.5f < tO & tO < 16.5f)
                    {
                        tO = 16;
                    }
                    float Wt = tO - tI - 0.5f;
                    int Int = (int)Wt;
                    float razn = Wt - (float)Int;
                    int x = 0;
                    if (razn > 0.7f) x = 1;
                    WorkHours = (int)Wt + x;

                }
                else // НОЧНАЯ СМЕНА
                {
                    WorkHours = 1;
                }
                ///////////////////////////////////////////////////////////TODO: СДЕЛАТЬ ПРАВИЛЬНЫЙ РАСЧЕТ ВРЕМЕНИ////////////////////////////////////////////////
            }
            else if (timeIn != none) //если есть только дата прихода
            {
                if (timeIn.Hour > 16)
                    WorkHours = 1;
                else
                {

                    int h = 15 - timeIn.Hour;
                    WorkHours = h;
                }
            }
            else if (timeOut != none) //если есть только дата УХОДА
            {
                if (timeOut.Hour < 7)
                    WorkHours = 1;
                else
                {
                    int h = 8;

                    int H = timeOut.Hour - h;
                    WorkHours = H;
                }
            }
            else //если пусто
            {
                WorkHours = 0;
            }
        }
    }
}
