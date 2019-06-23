using Microsoft.Win32;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Windows;
//using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace Prog
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        List<Mark> marks;
        double n2, n3, n4, n5, count;
        public MainWindow()
        {
            n2 = 0;
            n3 = 0;
            n4 = 0;
            n5 = 0;
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                OpenFileDialog file = new OpenFileDialog();
                file.Filter = "JSON| *.json";
                if (file.ShowDialog() == true)
                {
                    string json = File.ReadAllText(file.FileName, Encoding.GetEncoding(1251));
                    json = json.Replace(",\"items\":\n", String.Empty);
                    json = json.Replace("}{", "},{");
                    json = json.Replace("]}", "]");
                    marks = JsonConvert.DeserializeObject<List<Mark>>(json);
                    Calculate();
                    ToExcel();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка: " + ex.Message, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        private void Calculate()
        {
            count = marks.Count;
            int mark;
            foreach (Mark m in marks)
            {
                if (int.TryParse(m.name, out mark))
                {
                    if (mark >= 90) n5++;
                    if (mark >= 75 && mark < 90) n4++;
                    if (mark <= 74) n3++;
                }
                else
                {
                    if (m.name == "зв.") count--;
                    if (m.name == "н.д.") n2++;
                }
            }
            n2 = Math.Round(n2 / count * 100, 2);
            n3 = Math.Round(n3 / count * 100, 2);
            n4 = Math.Round(n4 / count * 100, 2);
            n5 = Math.Round(n5 / count * 100, 2);
        }
        private void ToExcel()
        {
            SaveFileDialog saveFile = new SaveFileDialog();
            saveFile.Filter = "Excel| *.xls";
            if (saveFile.ShowDialog() == true)
            {
                try
                {
                    Excel.Application exApp = new Excel.Application();
                    Excel.Workbook exWorkbook;
                    Excel.Worksheet exWorksheet;
                    exWorkbook = exApp.Workbooks.Add(Type.Missing);
                    exWorksheet = (Excel.Worksheet)exWorkbook.Sheets[1];
                    exWorksheet.Name = "Результаты";
                    exWorksheet.Cells[1, 1] = "Двоек";
                    exWorksheet.Cells[1, 2] = n2;
                    exWorksheet.Cells[1, 3] = "%";
                    exWorksheet.Cells[2, 1] = "Троек";
                    exWorksheet.Cells[2, 2] = n3;
                    exWorksheet.Cells[2, 3] = "%";
                    exWorksheet.Cells[3, 1] = "Четверок";
                    exWorksheet.Cells[3, 2] = n4;
                    exWorksheet.Cells[3, 3] = "%";
                    exWorksheet.Cells[4, 1] = "Пятерок";
                    exWorksheet.Cells[4, 2] = n5;
                    exWorksheet.Cells[4, 3] = "%";
                    exWorksheet.UsedRange.Columns.AutoFit();
                    exWorkbook.SaveAs(saveFile.FileName);
                    exApp.Quit();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Ошибка: " + ex.Message, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }
    }
}
