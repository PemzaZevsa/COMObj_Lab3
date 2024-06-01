using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DOTNET_Lab3
{
    public class Student : User
    {
        public List<uint> CoursesIds { get; set; }

        public Student()
        {
            CoursesIds = new List<uint>();
            UserType = 2;
        }
        public Student(string name, string surname)
        {
            Id = counter++;
            Name = name;
            Surname = surname;
            UserType = 2;
            CoursesIds = new List<uint>();
        }
        public Student(string name, string surname, string login, string password)
        {
            Id = counter++;
            Name = name;
            Surname = surname;
            Login = login;
            Password = password;
            UserType = 2;
            CoursesIds = new List<uint>();
        }
    }
}
