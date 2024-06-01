using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DOTNET_Lab3
{
    public class Admin : User
    {
        public Admin()
        {
            UserType = 0;
        }
        public Admin(string name, string surname)
        {
            Id = counter++;
            Name = name;
            Surname = surname;
            UserType = 0;
        }
    }
}
