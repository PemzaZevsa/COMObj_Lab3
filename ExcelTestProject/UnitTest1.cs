using DOTNET_Lab3;
using System.Diagnostics;

namespace ExcelTestProject
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod("Start/stop excel com server")]
        public void Test01()
        {
            int c3, c2, c1;
            var d = new WordDocument();
            c1 = Process.GetProcessesByName("excel").Length;
            using (var x = new ExcelDocument())
            {
                c2 = Process.GetProcessesByName("excel").Length;
                Thread.Sleep(5000);
            }

            c3 = Process.GetProcessesByName("excel").Length;

            Assert.AreEqual(0, c1);
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        [TestMethod("Create new excel file")]
        public void Test02()
        {
            string fileName = "TestDoc";
            string fullName = $"{Directory.GetCurrentDirectory()}\\{fileName}.xlsx";

            bool flag = File.Exists(fullName);

            using (var x = new ExcelDocument())
            {
                x["A1"] = "Cell A1";
                x[2,2] = "Cell 2,2";
                x[3, 3] = "100.5";
                x[4, 4] = "100,5";
                x[5, 5] = "öóöóéó5";
                x.SaveAs(fullName);
            }

            Assert.IsTrue(File.Exists(fullName));
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        [TestMethod("Check content")]
        public void Test03()
        {
            string fileName = "TestDoc";
            string fullName = $"{Directory.GetCurrentDirectory()}\\{fileName}.xlsx";

            bool flag = File.Exists(fullName);

            using (var x = new ExcelDocument(fullName))
            {
                flag &= x[2, 2] == "Cell 2,2";
                flag &= x[3, 3] == "100.5";
                flag &= x[4, 4] == "100,5";
                flag &= x[5, 5] == "öóöóéó5";

                
               
            }

            Assert.IsTrue(flag); 
            GC.Collect();
            GC.WaitForPendingFinalizers();

        }
    }
}