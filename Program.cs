using System;
using Excel = Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.IO;

namespace Practica
{
    class Program
    {
        static void Main(string[] args)
        {
            //Количество пар кряхтения над этим: 9

            Console.Title = "Задание №36 «Статистический анализ» выполнено Васильевым Глебом";
            Excel.Application ObjWorkExcel = new Excel.Application(); //Эксель

            Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(Directory.GetCurrentDirectory() + "\\2019_copy.xlsx", 
            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing); // Открыть файл

            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[3]; // Третий лист

            short sum = 0;
            short kolychestvo = 0;
            double min = double.MaxValue;
            double max = double.MinValue;
            string date_max = "";
            string date_min = "";

            for (short i = 1; i < 744; i+=3) { // Перебираем все ячейки, находим их значения
                try
                {
                    sum += ObjWorkSheet.Cells[i, 3].Value;
                    kolychestvo++;
                } catch
                {
                    continue;
                }
            }
            Console.WriteLine($"Количество записей: {kolychestvo}");
            Console.WriteLine($"Сумма: {sum}");

            float m0 = (float) sum / kolychestvo;
            Console.WriteLine($"Среднее значение: {m0}");

            short disp_sum = 0;
            for (short i = 1; i<744; i+=3)
            {
                disp_sum += Math.Pow((ObjWorkSheet.Cells[i, 3].Value - m0), 2);
                if (ObjWorkSheet.Cells[i, 3].Value > max) { max = ObjWorkSheet.Cells[i, 3].Value; date_max = (string)ObjWorkSheet.Cells[i+1, 1].Value; }
                if (ObjWorkSheet.Cells[i, 3].Value < min) { min = ObjWorkSheet.Cells[i, 3].Value; date_min = (string) ObjWorkSheet.Cells[i+1, 1].Value; }
            }
            float d = disp_sum / kolychestvo - 1;
            Console.WriteLine($"Дисперсия: {d}");
            Console.WriteLine();
            Console.WriteLine($"Максимальное: {max}, Дата: {date_max} \nМинимальное: {min}, Дата: {date_min}");

            Console.WriteLine();
            Console.WriteLine("Средняя температура по дням: ");
            short time = 1;
            short time_temp = 0;
            short j = 1;
            var days = new List<double>();
            for (short i = 1; i<744; i+=3)
            {
                time_temp += ObjWorkSheet.Cells[i, 3].Value;
                if (time % 8 == 0 && time!=0)
                {
                    Console.WriteLine($"День {j}: { time_temp / 8.0}");
                    days.Add(time_temp/8.0);
                    time = 0;
                    time_temp = 0;
                    j++;
                }
                time++;
            }

            Console.WriteLine();
            Console.WriteLine("Введите имя файла отчета: ");
            string file_name = Console.ReadLine();

            Excel.Application ObjWorkExcel2 = new Excel.Application();
            ObjWorkExcel2.SheetsInNewWorkbook = 1;//создать 1 лист в этой книге
            Excel.Workbook ObjWorkBook2 = ObjWorkExcel2.Workbooks.Add(); // создал книгу
            Excel.Worksheet ObjWorkSheet2 = (Excel.Worksheet)ObjWorkBook2.Sheets[1]; //первый лист

            ObjWorkSheet2.Cells[1, 2].Value = "Количество записей:";
            ObjWorkSheet2.Cells[1, 3].Value = kolychestvo;

            ObjWorkSheet2.Cells[2, 2].Value = "Сумма:";
            ObjWorkSheet2.Cells[2, 3].Value = sum;

            ObjWorkSheet2.Cells[3, 2].Value = "Среднее значение:";
            ObjWorkSheet2.Cells[3, 3].Value = m0;

            ObjWorkSheet2.Cells[4, 2].Value = "Дисперсия:";
            ObjWorkSheet2.Cells[4, 3].Value = d;

            ObjWorkSheet2.Cells[5, 2].Value = "Максимальная температура:";
            ObjWorkSheet2.Cells[5, 3].Value = max;
            ObjWorkSheet2.Cells[5, 4].Value = "Дата: ";
            ObjWorkSheet2.Cells[5, 5].Value = date_max;

            ObjWorkSheet2.Cells[6, 2].Value = "Минимальная температура:";
            ObjWorkSheet2.Cells[6, 3].Value = min;
            ObjWorkSheet2.Cells[6, 4].Value = "Дата:";
            ObjWorkSheet2.Cells[6, 5].Value = date_min;
            ObjWorkBook2.SaveAs(Directory.GetCurrentDirectory() + $"\\report\\{file_name}.xlsx");//сохранить файл
            ObjWorkBook2.Close(true); ObjWorkExcel2.Quit();//обязательно закрыть и выйти
        }
    }
}
