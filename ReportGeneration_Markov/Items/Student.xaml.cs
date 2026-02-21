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
using System.Windows.Navigation;
using System.Windows.Shapes;
using ReportGeneration_Markov.Classes;
using ReportGeneration_Markov.Models;
using ReportGeneration_Markov.Pages;

namespace ReportGeneration_Markov.Items
{
    /// <summary>
    /// Логика взаимодействия для Student.xaml
    /// </summary>
    public partial class Student : UserControl
    {
        public Student()
        {
            InitializeComponent();
            TBFio.Text = $"({student.Lastname}) {student.Firstname}";

            // Активируем checkbox отчислен
            CBExpelled.IsChecked = student.Expelled;

            // Получаем дисциплины в которых участвует студент
            List<DisciplineContext> StudentDisciplines = Main.AllDisciplines.FindAll(
                x => x.IdGroup == student.IdGroup);

            // Создаём переменные отвечающие за расчёты
            // Обязательных работ
            int NecessarilyCount = 0;
            // Всего занятий
            int WorksCount = 0;
            // Выполненных работ
            int DoneCount = 0;
            // Пропущенных минут
            int MissedCount = 0;

            // Перебираем дисциплины
            foreach (DisciplineContext StudentDiscipline in StudentDisciplines)
            {
                // Получаем кол-во работ принадлежащих к группе студента
                // К обязательным работам относятся [теоретические тесты], [экзамены] и [практические работы]
                List<WorkContext> StudentWorks = Main.AllWorks.FindAll(x =>
                    (x.IdType == 1 || x.IdType == 2 || x.IdType == 3) &&
                    x.IdDiscipline == StudentDiscipline.Id);

                // Увеличиваем кол-во обязательных работ
                NecessarilyCount += StudentWorks.Count;

                // Перебираем обязательные работы
                foreach (WorkContext StudentWork in StudentWorks)
                {
                    EvaluationContext Evaluation = Main.AllEvaluation.Find(x =>
                        x.IdWork == StudentWork.Id &&
                        x.IdStudent == student.Id);

                    // Проверяем если есть оценка за занятие и она не пустая, и не стоит оценка 2
                    if (Evaluation != null && Evaluation.Value.Trim() != "" && Evaluation.Value.Trim() != "2")
                        // Значит работа сдана
                        DoneCount++;
                }
            }

            // Получаем все занятия, кроме экзамена и оценки за месяц
            List<WorkContext> StudentWorksAll = Main.AllWorks.FindAll(x =>
                x.IdType != 4 && x.IdType != 3);

            // Увеличиваем количество занятий
            WorksCount += StudentWorksAll.Count;

            // Перебираем занятия
            foreach (WorkContext StudentWork in StudentWorksAll)
            {
                // Получаем оценку к занятия с пропусками
                EvaluationContext Evaluation = Main.AllEvaluation.Find(x =>
                    x.IdWork == StudentWork.Id &&
                    x.IdStudent == student.Id);

                // Если оценка не пустая, и есть прогул
                if (Evaluation != null && Evaluation.Lateness.Trim() != "")
                    // Добавляем его в общее кол-во пропущенных минут
                    MissedCount += Convert.ToInt32(Evaluation.Lateness);
            }

            // Выводим в процесс бар по формуле 100/(кол-во занятий)*выполненные
            doneWorks.Value = (100f / (float)NecessarilyCount) * ((float)DoneCount);

            // Выводим в процесс бар по формуле 100/(кол-во занятий * 90 (пара))*пропущенное кол-во минут
            missedCount.Value = (100f / ((float)WorksCount * 90f)) * ((float)MissedCount);

            TBGroup.Text = Main.AllGroups.Find(x => x.Id == student.IdGroup).Name;
        }
    }
}
