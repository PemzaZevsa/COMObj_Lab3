using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DOTNET_Lab3
{
    public abstract class User
    {
        private uint id;
        public uint Id { get; set; }
        private string login;
        public string Login
        {
            get => login;
            set
            {
                if (value is null)
                {
                    throw new ArgumentNullException(nameof(value));
                }
                if (value.Length < 4)
                {
                    throw new ArgumentException("Логін занадто короткий", nameof(value));
                }

                login = value;
            }
        }
        private string password;
        public string Password
        {
            get => password;
            set
            {
                if (value is null)
                {
                    throw new ArgumentNullException(nameof(value));
                }
                if (value.Length < 4)
                {
                    throw new ArgumentException("Пароль занадто короткий", nameof(value));
                }

                password = value;
            }
        }
        private string name;
        public string Name
        {
            get => name;
            set
            {
                if (value is null)
                {
                    throw new ArgumentNullException(nameof(value));
                }

                name = value;
            }
        }
        private string surname;
        public string Surname
        {
            get => surname;
            set
            {
                if (value is null)
                {
                    throw new ArgumentNullException(nameof(value));
                }

                surname = value;
            }
        }

        public static uint counter;
        public event Action<String> updatePassword;
        public short UserType { get; set; }

        public User()
        {
            Id = counter++;
        }
        public User(string name, string surname)
        {
            Id = counter++;
            Name = name;
            Surname = surname;
        }
        public User(uint id, string name, string surname, string login, string password)
        {
            Id = id;
            Name = name;
            Surname = surname;
            Login = login;
            Password = password;
        }

    }
}
