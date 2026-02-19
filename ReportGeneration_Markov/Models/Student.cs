using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportGeneration_Markov.Models
{
    public class Student
    {
        public int Id { get; set; }
        public string Firstname { get; set; }

        /// <summary> Фамилия
        public string Lastname { get; set; }

        /// <summary> Код группы
        public int IdGroup { get; set; }

        /// <summary> Отчислен
        public bool Expelled { get; set; }

        /// <summary> Дата отчисления
        public DateTime DateExpelled { get; set; }

        /// <summary> Конструктор для студента
        public Student(int Id, string Firstname, string Lastname, int IdGroup, bool Expelled, DateTime DateExpelled)
        {
            this.Id = Id;
            this.Firstname = Firstname;
            this.Lastname = Lastname;
            this.IdGroup = IdGroup;
            this.Expelled = Expelled;
            this.DateExpelled = DateExpelled;
        }
    }
}
