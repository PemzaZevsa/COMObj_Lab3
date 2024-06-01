using System.Diagnostics;

namespace DOTNET_Lab3
{
    internal class Program
    {
        static void Main(string[] args)
        {
            using (var x = new ExcelDocument())
            {
                Thread.Sleep(10000);
            }
        }
    }
}
