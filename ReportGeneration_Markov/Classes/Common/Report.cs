using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using Microsoft.Win32;
using ReportGeneration_Markov.Pages;
using Excel = Microsoft.Office.Interop.Excel;

namespace ReportGeneration_Markov.Classes.Common
{
    public class Report
    {
        public static void Group(int IdGroup, Main Main)
        {
            SaveFileDialog SFD = new SaveFileDialog
            {
                InitialDirectory = @"C:\",
                Filter = "Excel (*.xlsx")|*.xlsx"
            };
            SFD.ShowDialog();
            if (SFD.FileName) != ""){
                GroupContext Group = Main.AllGroups.Find(x => x.Id == IdGroup);
                var ExcelApp = new Excel.Application();
                try
                {
                    ExcelApp.Visible = false;
                    Excel.Workbook Workbook = ExcelApp.Workbooks.Add(Type.Missing);
                    Excel.Worksheet Worksheet = Workbook.ActiveSheet;

                    (Worksheet.Cells[1, 1] as Excel.Range).Value = $"Отчёт о группе {Group.Name}";
                    // Объединяем ячейки A1 и E1
                    Worksheet.Range[Worksheet.Cells[1, 1], Worksheet.Cells[1, 5]].Merge();
                    // Создаём стили для ячейки A1
                    Styles(Worksheet.Cells[1, 1], 18);

                    // Обращаемся к ячейке A3 и указываем текст
                    (Worksheet.Cells[3, 1] as Excel.Range).Value = "Список группы";
                    // Объединяем ячейки A3 и E3
                    Worksheet.Range[Worksheet.Cells[3, 1], Worksheet.Cells[3, 5]].Merge();
                    // Создаём стили для ячейки A3
                    Styles(Worksheet.Cells[3, 1], 12, Excel.XlHAlign.xlHAlignLeft);

                    // Обращаемся к ячейке, указываем текст
                    (Worksheet.Cells[4, 1] as Excel.Range).Value = "ФИО";
                    // Создаём стили для ячейки
                    Styles(Worksheet.Cells[4, 1], 12, Excel.XlHAlign.xlHAlignCenter, true);
                    // Указываем ширину
                    (Worksheet.Cells[4, 1] as Excel.Range).ColumnWidth = 35.0;

                    // Обращаемся к ячейке, указываем текст
                    (Worksheet.Cells[4, 2] as Excel.Range).Value = "Кол-во не сданных практических";
                    // Создаём стили для ячейки
                    Styles(Worksheet.Cells[4, 2], 12, Excel.XlHAlign.xlHAlignCenter, true);

                    // Обращаемся к ячейке, указываем текст
                    (Worksheet.Cells[4, 3] as Excel.Range).Value = "Кол-во не сданных теоретических";
                    // Создаём стили для ячейки
                    Styles(Worksheet.Cells[4, 3], 12, Excel.XlHAlign.xlHAlignCenter, true);

                    // Обращаемся к ячейке, указываем текст
                    (Worksheet.Cells[4, 4] as Excel.Range).Value = "Отсутствовал на паре";
                    // Создаём стили для ячейки
                    Styles(Worksheet.Cells[4, 4], 12, Excel.XlHAlign.xlHAlignCenter, true);

                    // Обращаемся к ячейке, указываем текст
                    (Worksheet.Cells[4, 5] as Excel.Range).Value = "Опоздал";
                    // Создаём стили для ячейки
                    Styles(Worksheet.Cells[4, 5], 12, Excel.XlHAlign.xlHAlignCenter, true);

                    int Height = 5;
                    List<StudentContext> Students = Main.AllStudents.FindAll(x => x.IdGroup == IdGroup);
                    foreach (StudentContext Student in Students)
                    {
                        List<DisciplineContext> StudentsDisciplines = Main.AllDisciplines.FindAll(
                            x => x.IdGroup == Student.IdGroup);
                        int PracticeCount = 0;
                        int TheoryCount = 0;
                        int AbesnteeismCount = 0;
                        int LateCount = 0;

                        foreach (DisciplineContext StudentDiscipline in StudentDisciplines)
                        {
                            // Получаем работы студента
                            List<WorkContext> StudentWorks = Main.AllWorks.FindAll(x => x.IdDiscipline == StudentDiscipline.Id);

                            // Перебираем работы студента
                            foreach (WorkContext StudentWork in StudentWorks)
                            {
                                // Получаем оценку за работу
                                EvaluationContext Evaluation = Main.AllEvaluation.Find(x =>
                                    x.IdWork == StudentWork.Id &&
                                    x.IdStudent == student.Id);

                                // Если оценки нет, или она пустая, или равно 2
                                if ((Evaluation != null && (Evaluation.Value.Trim() == "" || Evaluation.Value.Trim() == "2"))
                                    || Evaluation == null)
                                {
                                    // Если практика
                                    if (StudentWork.IdType == 1)
                                        // Считаем не сданную работу
                                        PracticeCount++;
                                    // Если теория
                                    else if (StudentWork.IdType == 2)
                                        // Считаем не сданную работу
                                        TheoryCount++;
                                }

                                // Проверяем что оценка не отсутствует и стоит пропуск
                                if (Evaluation != null && Evaluation.Lateness.Trim() != "")
                                {
                                    // Если пропуск 90 минут
                                    if (Convert.ToInt32(Evaluation.Lateness) == 90)
                                        // Считаем как пропущенную пару
                                        AbsenteeismCount++;
                                    else
                                        // Считаем как опоздание
                                        LateCount++;
                                }
                            }
                        }
                        (Worksheet.Cells[Height, 1] as Excel.Range).Value = $"{Student.Lastname} {Student.Firstname}";
                        Styles(Worksheet.Cells[Height, 1]), 12, XlAlign.xlAlignLeft, true);
                        (Worksheet.Cells[Height, 2] as Excel.Range).Value = PracticeCount.ToString();
                        Styles(Worksheet.Cells[Height, 2]), 12, XlAlign.xlAlignCenter, true);
                        (Worksheet.Cells[Height, 3] as Excel.Range).Value = TheoryCount.ToString();
                        Styles(Worksheet.Cells[Height, 3]), 12, XlAlign.xlAlignLeft, true);
                        (Worksheet.Cells[Height, 4] as Excel.Range).Value = AbesnteeismCount.ToString();
                        Styles(Worksheet.Cells[Height, 4]), 12, XlAlign.xlAlignCenter, true);
                        (Worksheet.Cells[Height, 5] as Excel.Range).Value = LateCount.ToString();
                        Styles(Worksheet.Cells[Height, 5]), 12, XlAlign.xlAlignCenter, true);
                    }
                    Workbook.SaveAs2(SFD.FileName);
                    Workbook.Close();
                }
                catch (Exception exp) { };
                ExcelApp.Quit();
            }
        }
        public static void Styles(Excel.Range Cell,
        int FontSize,
        Excel.XlHAlign Position = Excel.XlHAlign.xlHAlignCenter,
        bool Border = false)
        {
            // Присваиваем шрифт
            Cell.Font.Name = "Bahnschrift Light Condensed";
            // Присваиваем размер
            Cell.Font.Size = FontSize;
            // Указываем горизонтальное центрирование
            Cell.HorizontalAlignment = Position;
            // Указываем вертикальное центрирование
            Cell.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;

            // Если границы
            if (Border)
            {
                // Получаем границу ячейки
                Excel.Borders border = Cell.Borders;
                // Задаём стиль линии
                border.LineStyle = Excel.XlLineStyle.xlContinuous;
                // Задаём ширину линии
                border.Weight = Excel.XlBorderWeight.xlThin;
            }

            // Включаем перенос текста
            Cell.WrapText = true;
        } 
    }
}

