using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ReportGeneration_Markov.Classes.Common;
using ReportGeneration_Markov.Models;
using MySql.Data.MySqlClient;

namespace ReportGeneration_Markov.Classes
{
    public class GroupContext : Group
    {
        public GroupContext(int Id, string Name) : base(Id, Name) { }
        public static List<GroupContext> AllGroups()
        {
            List<GroupContext> allGroups = new List<GroupContext>();
            MySqlConnection connection = Connection.OpenConnection();
            MySqlDataReader BDGroups = Connection.Query("SELECT * FROM `group` ORDER BY `Name`", connection);
            while (BDGroups.Read())
            {
                allGroups.Add(new GroupContext(
                    BDGroups.GetInt32(0),
                    BDGroups.GetString(1)));
            }
            Connection.CloseConnection(connection);
            return allGroups;
        }
    }
}
