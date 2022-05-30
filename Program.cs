using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel  = Microsoft.Office.Interop.Excel;
/// <summary>
/// Решение уравнения Лапласа используя схему "крест"
/// </summary>
namespace Laplass
{
    class Program
    {
        static void Main(string[] args)
        {
            // Загрузить Excel, затем создать новую пустую рабочую книгу
         //   Excel.Application excelApp = new Excel.Application();
         //   // Сделать приложение Excel видимым
         //   excelApp.Visible = true;
         //   excelApp.Workbooks.Add();
         //   Excel._Worksheet workSheet = excelApp.ActiveSheet;
         //   string[] massivExcel = new string[] { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z" };
         //   int counter = 0;
            Console.WriteLine("Решение уравнения Лапласа: d^2*u/d^x*2 + d^2*u / d*y = 0, x = [0,1], y = [0,1]"); //вывод уравнения
            string stroka; //строка для хранения вводимых данных
            double x; //ввод x
            double y; //ввод y
            double stepByTime; //шаг по времени
            double stepBySpace; //шаг по пространству
            double Cu; //число Курента
            int n; //размер сетки
            Console.Write("Введите размер сетки: ");
            stroka = Console.ReadLine(); //считывание
            n = Convert.ToInt32(stroka); //конвертация из типа string в тип int
            double[,] massivResult = new double[n + 1, n + 1];
        Restart: //метка для рестарта если число x или y не подходят по условиям
            Console.Write("Введите x: ");
            stroka = Console.ReadLine();
            x = Convert.ToDouble(stroka);

            Console.Write("Введите y: ");
            stroka = Console.ReadLine();
            y = Convert.ToDouble(stroka);

            //условие проверки x и y
            if (x < 0 || x > 1 || y < 0 || y > 1)
            {
                Console.WriteLine("x и y должны быть больше 0 и меньше 1");
                goto Restart; //отправляем на повторный запрос если условия не подходят
            }
        NumberCurrent: //метка для повторного ввода если число курента > 0.5 или равно 0.5
            Console.Write("Введите шаг про времени: ");
            stroka = Console.ReadLine();
            stepByTime = Convert.ToDouble(stroka);

            Console.Write("Введите шаг по пространству: ");
            stroka = Console.ReadLine();
            stepBySpace = Convert.ToDouble(stroka);
            Cu = (stepByTime / (stepBySpace * stepBySpace));

            if (Cu == 0.5 || Cu > 0.5)
            {
                Console.WriteLine("Число вышло за пределы, измените шаг");
                goto NumberCurrent; //переход на Label если курент больше 0.5
            }
            //вывод числа курента и краевых условий
            Console.WriteLine("Число курента: {0}", Cu);
            Console.WriteLine("Краевые условия 1: u(0,y) = -7*y^2 - 5 * y + 3");
            Console.WriteLine("Краевые условия 2: u(1,y) = -7*y^2 - 21 * y + 13");
            Console.WriteLine("Краевые условия 3: u(x,0) = 6*x^2 + 4*x+3");
            Console.WriteLine("Краевые условия 4: u(x,1) = 6*x^2 - 12 * x - 9");

            //заполнение краевых условий
            for (int i = 0; i < n; i++)
            {
                massivResult[0, i] = -7 * Math.Pow(y, 2) - 5 * y + 3;
                massivResult[1, i] = -7 * Math.Pow(y, 2) - 21 * y + 13;
                massivResult[i, 0] = 6 * Math.Pow(x, 2) + 4 * x + 3;
                massivResult[i, 1] = 6 * Math.Pow(x, 2) - 12 * x - 9;
            }
            Console.WriteLine("Для решения используем схему крест");
            //расчёт схемой крест
            for (int j = 2; j < n; j++)
            {
                for (int i = 2; i < n; i++)
                {
                    massivResult[i, j] = 0.25 * (massivResult[i - 1, j] + massivResult[i + 1, j] + massivResult[i, j - 1] + massivResult[i, j + 1]) + Math.Pow(stepBySpace, 2);
                    int con = Convert.ToInt32(massivResult[i, j]);
                    massivResult[i, j] = con;
                }

            }
            Console.WriteLine();
            #region Вывод_Результата
            for (int j = n - 1; j > 0; j--)
            {
                Console.Write("|" + " ");

                for (int i = 0; i < n; i++)
                {

                    Console.Write(massivResult[i, j] + "\t" + "|");
                    //workSheet.Cells[i+1, j+1] = massivResult[i,j];
                }
                Console.WriteLine();
                Console.Write(new string('-', n * 10));
                Console.WriteLine();
            }
            Console.Write("|" + " ");

            for (int i = 0; i < n; i++)
            {
                Console.Write("{0}", massivResult[i, 0] + "\t" + "|");
             //   workSheet.Cells[i+1, 1] = massivResult[i, 0];
            }
            Console.WriteLine();
            Console.Write(new string('-', n * 10));


           
           // excelApp.DisplayAlerts = false;
           // workSheet.SaveAs(string.Format(@"C:\Users\User\source\repos\Laplass\1.xlsx"));
           //
           // excelApp.Quit();
            Console.ReadKey();
            #endregion Вывод_Результата
        }
    }
}
