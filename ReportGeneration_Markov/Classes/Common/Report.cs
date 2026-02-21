using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
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
                InitialDirectory = @"C:\Users\Student-A502.PERMAVIAT\Desktop\",
                Filter = "Excel (*.xlsx)|*.xlsx",
                FileName = "Отчет.xlsx"
            };

            if (SFD.ShowDialog() == true)
            {
                GroupContext Group = Main.AllGroups.Find(x => x.Id == IdGroup);
                var ExcelApp = new Excel.Application();
                try
                {
                    ExcelApp.Visible = false;
                    Excel.Workbook Workbook = ExcelApp.Workbooks.Add(Type.Missing);

                    // Создаем основной лист с общим отчетом
                    CreateMainSheet(Workbook, IdGroup, Main, Group);

                    // Для оценки "Отлично" - создаем отдельные листы для каждого студента
                    CreateStudentSheets(Workbook, IdGroup, Main);

                    Workbook.SaveAs2(SFD.FileName);
                    Workbook.Close();

                    MessageBox.Show("Отчет успешно создан!", "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                catch (Exception exp)
                {
                    MessageBox.Show($"Ошибка при создании отчета: {exp.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                }
                finally
                {
                    ExcelApp.Quit();
                }
            }
        }

        private static void CreateMainSheet(Excel.Workbook Workbook, int IdGroup, Main Main, GroupContext Group)
        {
            Excel.Worksheet Worksheet = Workbook.ActiveSheet;
            Worksheet.Name = "Общий отчет";

            (Worksheet.Cells[1, 1] as Excel.Range).Value = $"Отчёт о группе {Group.Name}";
            Worksheet.Range[Worksheet.Cells[1, 1], Worksheet.Cells[1, 7]].Merge();
            Styles(Worksheet.Cells[1, 1] as Excel.Range, 18);

            (Worksheet.Cells[3, 1] as Excel.Range).Value = "Список группы";
            Worksheet.Range[Worksheet.Cells[3, 1], Worksheet.Cells[3, 7]].Merge();
            Styles(Worksheet.Cells[3, 1] as Excel.Range, 12, Excel.XlHAlign.xlHAlignLeft);

            // Заголовки
            string[] headers = { "ФИО", "Не сдано практ.", "Не сдано теор.", "Пропуски (пар)", "Опоздания", "Успеваемость %", "Посещаемость %" };
            for (int i = 0; i < headers.Length; i++)
            {
                (Worksheet.Cells[4, i + 1] as Excel.Range).Value = headers[i];
                Styles(Worksheet.Cells[4, i + 1] as Excel.Range, 12, Excel.XlHAlign.xlHAlignCenter, true);
            }

            (Worksheet.Cells[4, 1] as Excel.Range).ColumnWidth = 35.0;

            int Height = 5;
            List<StudentContext> Students = Main.AllStudents.FindAll(x => x.IdGroup == IdGroup);

            // Словарь для хранения показателей студентов
            Dictionary<int, StudentStats> stats = new Dictionary<int, StudentStats>();

            foreach (StudentContext Student in Students)
            {
                var studentStats = CalculateStudentStats(Student, Main);
                stats[Student.Id] = studentStats;

                (Worksheet.Cells[Height, 1] as Excel.Range).Value = $"{Student.Lastname} {Student.Firstname}";
                Styles(Worksheet.Cells[Height, 1] as Excel.Range, 12, Excel.XlHAlign.xlHAlignLeft, true);

                (Worksheet.Cells[Height, 2] as Excel.Range).Value = studentStats.PracticeNotPassed.ToString();
                Styles(Worksheet.Cells[Height, 2] as Excel.Range, 12, Excel.XlHAlign.xlHAlignCenter, true);

                (Worksheet.Cells[Height, 3] as Excel.Range).Value = studentStats.TheoryNotPassed.ToString();
                Styles(Worksheet.Cells[Height, 3] as Excel.Range, 12, Excel.XlHAlign.xlHAlignCenter, true);

                (Worksheet.Cells[Height, 4] as Excel.Range).Value = studentStats.AbsenteeismCount.ToString();
                Styles(Worksheet.Cells[Height, 4] as Excel.Range, 12, Excel.XlHAlign.xlHAlignCenter, true);

                (Worksheet.Cells[Height, 5] as Excel.Range).Value = studentStats.LateCount.ToString();
                Styles(Worksheet.Cells[Height, 5] as Excel.Range, 12, Excel.XlHAlign.xlHAlignCenter, true);

                // Успеваемость (процент сданных работ)
                double successRate = studentStats.TotalWorks > 0
                    ? Math.Round((double)studentStats.PassedWorks / studentStats.TotalWorks * 100, 1)
                    : 0;
                (Worksheet.Cells[Height, 6] as Excel.Range).Value = successRate + "%";
                Styles(Worksheet.Cells[Height, 6] as Excel.Range, 12, Excel.XlHAlign.xlHAlignCenter, true);

                // Посещаемость (процент посещенных занятий)
                double attendanceRate = studentStats.TotalClasses > 0
                    ? Math.Round(100 - ((double)studentStats.TotalMissedMinutes / (studentStats.TotalClasses * 90) * 100), 1)
                    : 0;
                if (attendanceRate < 0) attendanceRate = 0;
                (Worksheet.Cells[Height, 7] as Excel.Range).Value = attendanceRate + "%";
                Styles(Worksheet.Cells[Height, 7] as Excel.Range, 12, Excel.XlHAlign.xlHAlignCenter, true);

                Height++;
            }

            Height += 2;

            // Находим самого успешного студента
            var bestStudent = FindBestStudent(Students, stats);
            if (bestStudent.Item1 != null)
            {
                (Worksheet.Cells[Height, 1] as Excel.Range).Value = "САМЫЙ УСПЕШНЫЙ СТУДЕНТ:";
                Styles(Worksheet.Cells[Height, 1] as Excel.Range, 14, Excel.XlHAlign.xlHAlignRight, true);
                Worksheet.Range[Worksheet.Cells[Height, 1], Worksheet.Cells[Height, 3]].Merge();

                (Worksheet.Cells[Height, 4] as Excel.Range).Value = $"{bestStudent.Item1.Lastname} {bestStudent.Item1.Firstname}";
                Styles(Worksheet.Cells[Height, 4] as Excel.Range, 14, Excel.XlHAlign.xlHAlignLeft, true, true);
                Worksheet.Range[Worksheet.Cells[Height, 4], Worksheet.Cells[Height, 7]].Merge();

                // Подсвечиваем строку лучшего студента зеленым
                Excel.Range bestRow = Worksheet.Rows[bestStudent.Item2 + 4] as Excel.Range;
                bestRow.Interior.Color = XlRgbColor.rgbLightGreen;
            }
        }

        private static void CreateStudentSheets(Excel.Workbook Workbook, int IdGroup, Main Main)
        {
            List<StudentContext> Students = Main.AllStudents.FindAll(x => x.IdGroup == IdGroup);

            foreach (StudentContext Student in Students)
            {
                Excel.Worksheet Worksheet = Workbook.Sheets.Add(After: Workbook.Sheets[Workbook.Sheets.Count]);
                Worksheet.Name = $"{Student.Lastname} {Student.Firstname}";

                int row = 1;

                // Заголовок
                (Worksheet.Cells[row, 1] as Excel.Range).Value = $"Детальный отчет по студенту: {Student.Lastname} {Student.Firstname}";
                Worksheet.Range[Worksheet.Cells[row, 1], Worksheet.Cells[row, 6]].Merge();
                Styles(Worksheet.Cells[row, 1] as Excel.Range, 16);
                row += 2;

                // Информация о студенте
                (Worksheet.Cells[row, 1] as Excel.Range).Value = "Группа:";
                Styles(Worksheet.Cells[row, 1] as Excel.Range, 12, Excel.XlHAlign.xlHAlignRight, true);

                string groupName = Main.AllGroups.Find(x => x.Id == Student.IdGroup)?.Name ?? "Неизвестно";
                (Worksheet.Cells[row, 2] as Excel.Range).Value = groupName;
                Styles(Worksheet.Cells[row, 2] as Excel.Range, 12, Excel.XlHAlign.xlHAlignLeft, true);

                (Worksheet.Cells[row, 4] as Excel.Range).Value = "Статус:";
                Styles(Worksheet.Cells[row, 4] as Excel.Range, 12, Excel.XlHAlign.xlHAlignRight, true);

                string status = Student.Expelled ? $"Отчислен ({Student.DateExpelled.ToShortDateString()})" : "Учится";
                (Worksheet.Cells[row, 5] as Excel.Range).Value = status;
                Styles(Worksheet.Cells[row, 5] as Excel.Range, 12, Excel.XlHAlign.xlHAlignLeft, true);
                row += 2;

                // Заголовки таблицы работ
                string[] workHeaders = { "Дисциплина", "Работа", "Тип", "Дата", "Оценка", "Опоздание (мин)" };
                for (int i = 0; i < workHeaders.Length; i++)
                {
                    (Worksheet.Cells[row, i + 1] as Excel.Range).Value = workHeaders[i];
                    Styles(Worksheet.Cells[row, i + 1] as Excel.Range, 12, Excel.XlHAlign.xlHAlignCenter, true);
                }
                row++;

                // Получаем все дисциплины студента
                List<DisciplineContext> StudentDisciplines = Main.AllDisciplines.FindAll(x => x.IdGroup == Student.IdGroup);

                int passedCount = 0;
                int notPassedCount = 0;
                int totalWorks = 0;

                foreach (DisciplineContext Discipline in StudentDisciplines)
                {
                    List<WorkContext> DisciplineWorks = Main.AllWorks.FindAll(x => x.IdDiscipline == Discipline.Id);

                    foreach (WorkContext Work in DisciplineWorks)
                    {
                        EvaluationContext Evaluation = Main.AllEvaluations.Find(x =>
                            x.IdWork == Work.Id && x.IdStudent == Student.Id);

                        string workType = GetWorkType(Work.IdType);
                        string evaluation = Evaluation?.Value ?? "";
                        string lateness = Evaluation?.Lateness ?? "";

                        // Определяем цвет строки в зависимости от сдачи
                        bool isPassed = Evaluation != null && evaluation.Trim() != "" && evaluation.Trim() != "2";
                        if (isPassed)
                            passedCount++;
                        else
                            notPassedCount++;
                        totalWorks++;

                        (Worksheet.Cells[row, 1] as Excel.Range).Value = Discipline.Name;
                        (Worksheet.Cells[row, 2] as Excel.Range).Value = Work.Name;
                        (Worksheet.Cells[row, 3] as Excel.Range).Value = workType;
                        (Worksheet.Cells[row, 4] as Excel.Range).Value = Work.Date.ToShortDateString();
                        (Worksheet.Cells[row, 5] as Excel.Range).Value = evaluation;
                        (Worksheet.Cells[row, 6] as Excel.Range).Value = lateness;

                        // Применяем стили ко всей строке
                        for (int i = 1; i <= 6; i++)
                        {
                            Styles(Worksheet.Cells[row, i] as Excel.Range, 11, Excel.XlHAlign.xlHAlignLeft, true);
                        }

                        // Подсвечиваем строки в зависимости от результата
                        Excel.Range workRow = Worksheet.Rows[row] as Excel.Range;
                        if (isPassed)
                            workRow.Interior.Color = XlRgbColor.rgbLightGreen;
                        else if (Evaluation != null && evaluation.Trim() == "2")
                            workRow.Interior.Color = XlRgbColor.rgbLightBlue;
                        else if (Evaluation == null || evaluation.Trim() == "")
                            workRow.Interior.Color = XlRgbColor.rgbLightYellow;

                        row++;
                    }
                }

                row += 2;

                // Итоговая статистика
                (Worksheet.Cells[row, 1] as Excel.Range).Value = "ИТОГОВАЯ СТАТИСТИКА:";
                Styles(Worksheet.Cells[row, 1] as Excel.Range, 14, Excel.XlHAlign.xlHAlignLeft, true);
                Worksheet.Range[Worksheet.Cells[row, 1], Worksheet.Cells[row, 3]].Merge();
                row++;

                (Worksheet.Cells[row, 1] as Excel.Range).Value = "Всего работ:";
                Styles(Worksheet.Cells[row, 1] as Excel.Range, 12, Excel.XlHAlign.xlHAlignRight, true);
                (Worksheet.Cells[row, 2] as Excel.Range).Value = totalWorks.ToString();
                Styles(Worksheet.Cells[row, 2] as Excel.Range, 12, Excel.XlHAlign.xlHAlignCenter, true);
                row++;

                (Worksheet.Cells[row, 1] as Excel.Range).Value = "Сдано работ:";
                Styles(Worksheet.Cells[row, 1] as Excel.Range, 12, Excel.XlHAlign.xlHAlignRight, true);
                (Worksheet.Cells[row, 2] as Excel.Range).Value = passedCount.ToString();
                Styles(Worksheet.Cells[row, 2] as Excel.Range, 12, Excel.XlHAlign.xlHAlignCenter, true);
                row++;

                (Worksheet.Cells[row, 1] as Excel.Range).Value = "Не сдано работ:";
                Styles(Worksheet.Cells[row, 1] as Excel.Range, 12, Excel.XlHAlign.xlHAlignRight, true);
                (Worksheet.Cells[row, 2] as Excel.Range).Value = notPassedCount.ToString();
                Styles(Worksheet.Cells[row, 2] as Excel.Range, 12, Excel.XlHAlign.xlHAlignCenter, true);

                // Настраиваем ширину колонок
                (Worksheet.Cells[1, 1] as Excel.Range).ColumnWidth = 20;
                (Worksheet.Cells[1, 2] as Excel.Range).ColumnWidth = 30;
                (Worksheet.Cells[1, 3] as Excel.Range).ColumnWidth = 15;
                (Worksheet.Cells[1, 4] as Excel.Range).ColumnWidth = 12;
                (Worksheet.Cells[1, 5] as Excel.Range).ColumnWidth = 10;
                (Worksheet.Cells[1, 6] as Excel.Range).ColumnWidth = 15;
            }
        }

        private class StudentStats
        {
            public int PracticeNotPassed { get; set; }
            public int TheoryNotPassed { get; set; }
            public int AbsenteeismCount { get; set; }
            public int LateCount { get; set; }
            public int PassedWorks { get; set; }
            public int TotalWorks { get; set; }
            public int TotalClasses { get; set; }
            public int TotalMissedMinutes { get; set; }
        }

        private static StudentStats CalculateStudentStats(StudentContext Student, Main Main)
        {
            var stats = new StudentStats();

            List<DisciplineContext> StudentsDisciplines = Main.AllDisciplines.FindAll(
                x => x.IdGroup == Student.IdGroup);

            foreach (DisciplineContext StudentDiscipline in StudentsDisciplines)
            {
                List<WorkContext> StudentWorks = Main.AllWorks.FindAll(x => x.IdDiscipline == StudentDiscipline.Id);

                foreach (WorkContext StudentWork in StudentWorks)
                {
                    EvaluationContext Evaluation = Main.AllEvaluations.Find(x =>
                        x.IdWork == StudentWork.Id &&
                        x.IdStudent == Student.Id);

                    // Учитываем только обязательные работы для статистики успеваемости
                    if (StudentWork.IdType == 1 || StudentWork.IdType == 2 || StudentWork.IdType == 3)
                    {
                        stats.TotalWorks++;

                        bool isPassed = Evaluation != null && Evaluation.Value.Trim() != "" && Evaluation.Value.Trim() != "2";
                        if (isPassed)
                        {
                            stats.PassedWorks++;
                        }
                        else
                        {
                            if (StudentWork.IdType == 1)
                                stats.PracticeNotPassed++;
                            else if (StudentWork.IdType == 2)
                                stats.TheoryNotPassed++;
                        }
                    }

                    // Учитываем посещаемость для всех занятий (кроме экзаменов)
                    if (StudentWork.IdType != 4 && StudentWork.IdType != 3)
                    {
                        stats.TotalClasses++;

                        if (Evaluation != null && Evaluation.Lateness.Trim() != "")
                        {
                            if (int.TryParse(Evaluation.Lateness, out int lateness))
                            {
                                stats.TotalMissedMinutes += lateness;

                                if (lateness == 90)
                                    stats.AbsenteeismCount++;
                                else
                                    stats.LateCount++;
                            }
                        }
                    }
                }
            }

            return stats;
        }

        private static Tuple<StudentContext, int> FindBestStudent(List<StudentContext> Students, Dictionary<int, StudentStats> stats)
        {
            if (Students.Count == 0) return new Tuple<StudentContext, int>(null, -1);

            double bestScore = -1;
            StudentContext bestStudent = null;
            int bestIndex = -1;

            for (int i = 0; i < Students.Count; i++)
            {
                var student = Students[i];
                var studentStats = stats[student.Id];

                if (student.Expelled) continue;

                double successRate = studentStats.TotalWorks > 0
                    ? (double)studentStats.PassedWorks / studentStats.TotalWorks * 60
                    : 0;

                double attendanceRate = studentStats.TotalClasses > 0
                    ? (1 - (double)studentStats.TotalMissedMinutes / (studentStats.TotalClasses * 90)) * 40
                    : 40;

                if (attendanceRate < 0) attendanceRate = 0;

                double totalScore = successRate + attendanceRate;

                if (totalScore > bestScore)
                {
                    bestScore = totalScore;
                    bestStudent = student;
                    bestIndex = i;
                }
            }

            return new Tuple<StudentContext, int>(bestStudent, bestIndex);
        }

        private static string GetWorkType(int typeId)
        {
            switch (typeId)
            {
                case 1: return "Практическая";
                case 2: return "Теоретическая";
                case 3: return "Экзамен";
                case 4: return "Лекция";
                default: return "Другое";
            }
        }

        public static void Styles(Excel.Range Cell,
        int FontSize,
        Excel.XlHAlign Position = Excel.XlHAlign.xlHAlignCenter,
        bool Border = false,
        bool Bold = false)
        {
            if (Cell == null) return;

            // Присваиваем шрифт
            Cell.Font.Name = "Arial";
            // Присваиваем размер
            Cell.Font.Size = FontSize;
            // Жирный шрифт
            Cell.Font.Bold = Bold;
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