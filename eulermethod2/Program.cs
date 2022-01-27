using System;
using Excel = Microsoft.Office.Interop.Excel;// if you want to run it on your pc you need to enable this link
namespace Euler_method
{
    class Program
    {
        static void Main(string[] args)
        {
            

            //___________________________________________________________________________
            //Объявляем приложение
            Excel.Application ex = new Microsoft.Office.Interop.Excel.Application();
            //Отобразить Excel
            ex.Visible = true;
            //Количество листов в рабочей книге
            ex.SheetsInNewWorkbook = 2;
            //Добавить рабочую книгу
            Excel.Workbook workBook = ex.Workbooks.Add(Type.Missing);

            //Отключить отображение окон с сообщениями
            ex.DisplayAlerts = false;
            //Получаем первый лист документа (счет начинается с 1)
            Excel.Worksheet sheet = (Excel.Worksheet)ex.Worksheets.get_Item(1);
            //Название листа (вкладки снизу)
            //sheet.Name = "таблица1";
            //___________________________________________________________________________

            double x = 0;
            double y = 1;
            double x_1 = 0;
            double y_1 = 0;
            double delta_x=0;
            //начальное усл-е

           Console.WriteLine("Введите произвольный шаг от 0 до 1");
           delta_x = Convert.ToDouble(Console.ReadLine());// need to enter number from 0 to 1 through ,

            //function y'=x^2 - 2*y
            int iterations = Convert.ToInt32(Math.Floor(1 / delta_x));// nuber of steps
            Console.WriteLine("{0}- Число шагов, {1}-ваш произв. шаг", iterations, delta_x);// it outputs all number of steps (1/delta_x) and your step
            //first step
            x_1 = x + delta_x;
            x = x_1;
            y_1 = y + delta_x * (Math.Pow(x, 2) - 2 * y);// function y'=x^2-2y, modify here if you want
            y = y_1;
            Console.WriteLine("Шаг номер {0}\n x= {1}; y={2};", 1, x_1, y_1);// result after first step

            //___________________________________________________________________________
            sheet.Cells[1, 1] = String.Format("X");//printing in Excel-table labels and first step
            sheet.Cells[1, 2] = String.Format("Y");
            //__________________________________________________
            sheet.Cells[2, 1] = String.Format("{0}", x_1);
            sheet.Cells[2, 2] = String.Format("{0}", y_1);
            //___________________________________________________________________________
            for (int i = 1; i < iterations; i++) //another steps
            {
                
                    x = x + delta_x;
                    y = y + delta_x * (Math.Pow(x, 2) - 2 * y);// function y'=x^2-2y, modify here if you want
                    Console.WriteLine("Шаг номер {0}\n x= {1}; y={2};", i, x, y);
               
                //___________________________________________________________________________
                sheet.Cells[(i+2), 1] = String.Format("{0}", x);//printing in Excel-table file another steps
                sheet.Cells[(i+2), 2] = String.Format("{0}", y);
                //___________________________________________________________________________
            }
            //сохранение экселя saving Excel-table
            ex.Application.ActiveWorkbook.SaveAs("doc.xlsx", Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            //___________________________________________________________________________
           Console.ReadKey();
            //___________________________________________________________________________
        }
    }
}
