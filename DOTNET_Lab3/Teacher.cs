using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DOTNET_Lab3
{
    public class Teacher : User
    {
        public Teacher()
        {
            UserType = 1;
        }
        public Teacher(string name, string surname)
        {
            Id = counter++;
            Name = name;
            Surname = surname;
            UserType = 1;
        }
        public Teacher(string name, string surname, string login, string password)
        {
            Id = counter++;
            Name = name;
            Surname = surname;
            UserType = 1;
            Login = login;
            Password = password;
        }
    }
}
