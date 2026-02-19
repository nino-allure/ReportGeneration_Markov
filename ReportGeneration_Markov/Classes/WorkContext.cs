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
    public class WorkContext : Work
    {
        /// <summary> Конструктор для заполнения объектов
        /// </summary>
        public WorkContext(int Id, int IdDiscipline, int IdType, DateTime Date, string Name, int Semester) :
            base(Id, IdDiscipline, IdType, Date, Name, Semester)
        { }

        /// <summary> Получение всех работ
        /// </summary>
        public static List<WorkContext> AllWorks()
        {
            List<WorkContext> allWorks = new List<WorkContext>();
            MySqlConnection connection = Connection.OpenConnection();
            MySqlDataReader BDworks = Connection.Query("SELECT * FROM `work` ORDER BY `Date`", connection);
            while (BDworks.Read())
            { 
                allWorks.Add(new WorkContext(
                    BDworks.GetInt32(0),
                    BDworks.GetInt32(1),
                    BDworks.GetInt32(2),
                    BDworks.GetDateTime(3),
                    BDworks.GetString(4),
                    BDworks.GetInt32(5)));
            }
            Connection.CloseConnection(connection);
            return allWorks;
        }
    }
}
