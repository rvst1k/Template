using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using Microsoft.Win32;

namespace Template_4333
{
    /// <summary>
    /// Логика взаимодействия для _4333_Akhmetzyanov.xaml
    /// </summary>
    public partial class _4333_Akhmetzyanov : System.Windows.Window
    {
        public _4333_Akhmetzyanov()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog()
            {
                DefaultExt = "*.xls;*.xlsx",
                Filter = "файл Excel (Spisok.xlsx)|*.xlsx",
                Title = "Выберите файл базы данных"
            };
            if (!(ofd.ShowDialog() == true))
                return;
            string[,] list;

            Excel.Application ObjWorkExcel = new
            Excel.Application();
            Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(ofd.FileName);
            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1];
            var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
            int _columns = (int)lastCell.Column;
            int _rows = (int)lastCell.Row;
            list = new string[_rows, _columns];
            for (int j = 0; j < _columns; j++)
            {
                for (int i = 0; i < _rows; i++)
                {
                    list[i, j] = ObjWorkSheet.Cells[i + 1, j + 1].Text;
                }
            }
            ObjWorkBook.Close(false, Type.Missing, Type.Missing);
            ObjWorkExcel.Quit();
            GC.Collect();

            using (isrpo2Entities1 usersEntities = new isrpo2Entities1())
            {
                for (int i = 1; i < _rows; i++)
                {
                    usersEntities.isrpo.Add(new isrpo()
                    {
                        ID = list[i, 0],
                        НаименованиеУслуги = list[i, 1],
                        ВидУслуги = list[i, 2],
                        КодУслуги = list[i, 3],
                        Стоимость = list[i, 4],
                    });
                }
                usersEntities.SaveChanges();
            }
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            List<isrpo> AllService;
            using (isrpo2Entities1 UserEntities = new isrpo2Entities1())
            {
                AllService = UserEntities.isrpo.ToList();
            }
            var app = new Excel.Application();
            Excel.Workbook workbook = app.Workbooks.Add(Type.Missing);
            app.Visible = true;
            Excel.Worksheet worksheet1 = app.Worksheets.Add();
            worksheet1.Name = "Категория 1";
            Excel.Worksheet worksheet2 = app.Worksheets.Add();
            worksheet2.Name = "Категория 2";
            Excel.Worksheet worksheet3 = app.Worksheets.Add();
            worksheet3.Name = "Категория 3";
            worksheet1.Cells[1, 1] = "id";
            worksheet1.Cells[1, 2] = "Nazvanie Uslugi";
            worksheet1.Cells[1, 3] = "Vid Uslugi";
            worksheet1.Cells[1, 4] = "Stoimost";

            worksheet2.Cells[1, 1] = "id";
            worksheet2.Cells[1, 2] = "Nazvanie Uslugi";
            worksheet2.Cells[1, 3] = "Vid Uslugi";
            worksheet2.Cells[1, 4] = "Stoimost";

            worksheet3.Cells[1, 1] = "id";
            worksheet3.Cells[1, 2] = "Nazvanie Uslugi";
            worksheet3.Cells[1, 3] = "Vid Uslugi";
            worksheet3.Cells[1, 4] = "Stoimost";
            int rowindex1 = 2;
            int rowindex2 = 2;
            int rowindex3 = 2;

            foreach(var service in AllService)
            {
                if(Convert.ToDouble(service.Стоимость) < 350)
                {
                    worksheet1.Cells[rowindex1, 1] = service.ID;
                    worksheet1.Cells[rowindex1, 2] = service.НаименованиеУслуги;
                    worksheet1.Cells[rowindex1, 3] = service.ВидУслуги;
                    worksheet1.Cells[rowindex1, 4] = service.Стоимость;
                    rowindex1++;
                }
                else if (Convert.ToDouble(service.Стоимость) > 250 && Convert.ToInt32(service.Стоимость) < 800)
                {
                    worksheet2.Cells[rowindex2, 1] = service.ID;
                    worksheet2.Cells[rowindex2, 2] = service.НаименованиеУслуги;
                    worksheet2.Cells[rowindex2, 3] = service.ВидУслуги;
                    worksheet2.Cells[rowindex2, 4] = service.Стоимость;
                    rowindex2++;
                }
                else if (Convert.ToDouble(service.Стоимость) > 800)
                {
                    worksheet3.Cells[rowindex3, 1] = service.ID;
                    worksheet3.Cells[rowindex3, 2] = service.НаименованиеУслуги;
                    worksheet3.Cells[rowindex3, 3] = service.ВидУслуги;
                    worksheet3.Cells[rowindex3, 4] = service.Стоимость;
                    rowindex3++;
                }
                else
                {

                }

            }
        }
    }
}

