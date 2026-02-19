using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MySql.Data.MySqlClient;
using ReportGeneration_Markov.Classes.Common;
using ReportGeneration_Markov.Models;

namespace ReportGeneration_Markov.Classes
{
    public class EvaluationContext : Evaluation
    {
        /// <summary> Конструктор для заполнения объекта
        /// </summary>
        public EvaluationContext(int Id, int IdWork, int IdStudent, string Value, string Lateness) :
            base(Id, IdWork, IdStudent, Value, Lateness)
        { }

        /// <summary> Получение оценок студентов
        /// </summary>
        public static List<EvaluationContext> AllEvaluations()
        {
            // Коллекция оценок
            List<EvaluationContext> allEvaluations = new List<EvaluationContext>();

            // Открываем соединение
            MySqlConnection connection = Connection.OpenConnection();

            // Выполняем запрос
            MySqlDataReader BDEvaluations = Connection.Query("SELECT * FROM `evaluation`;", connection);

            // Читаем данные из БД
            while (BDEvaluations.Read())
            {
                // Добавляем данные в коллекцию
                allEvaluations.Add(new EvaluationContext(
                    BDEvaluations.GetInt32(0),
                    BDEvaluations.GetInt32(1),
                    BDEvaluations.GetInt32(2),
                    BDEvaluations.GetString(3),
                    BDEvaluations.GetString(4)));
            }

            // Закрываем подключение
            Connection.CloseConnection(connection);

            // Возвращаем коллекцию оценок
            return allEvaluations;
        }
    }
}
