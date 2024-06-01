using DOTNET_Lab3;
using System.Diagnostics;
using System.Runtime.Intrinsics.Arm;

namespace ExcelTestProject
{
    [TestClass]
    public class UnitTest1
    {
        static int c1 = 0, c3 = 0;
        [TestMethod("01 Start/stop excel com server")]
        public void Test01()
        {
            int  c2;
            var d = new WordDocument();
            c1 = Process.GetProcessesByName("excel").Length;
            using (var x = new ExcelDocument())
            {
                c2 = Process.GetProcessesByName("excel").Length;
                Thread.Sleep(1000);
            }

            c3 = Process.GetProcessesByName("excel").Length;

            Assert.AreEqual(0, c1);
        }

        [TestMethod("02 Create new excel file")]
        public void Test02()
        {
            string fileName = "TestDoc";
            string fullName = $"{Directory.GetCurrentDirectory()}\\{fileName}.xlsx";

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

        }

        [TestMethod("03 Check content")]
        public void Test03()
        {
            string fileName = "TestDoc";
            string fullName = $"{Directory.GetCurrentDirectory()}\\{fileName}.xlsx";

            bool flag = File.Exists(fullName);

            using (var x = new ExcelDocument(fullName))
            {
                flag &= x[2, 2] == "Cell 2,2";
                flag &= x[3, 3] == "100,5";
                flag &= x[4, 4] == "100,5";
                flag &= x[5, 5] == "öóöóéó5";

                
               
            }

            Assert.IsTrue(flag);
        }

        [TestMethod("04 Garbage collector")]
        public void Test04()
        {
            Thread.Sleep(4000);
             c3 = Process.GetProcessesByName("excel").Length;

            Assert.IsTrue(c1 == c3);
        }


        [TestMethod("11 Word")]
        public void Test11()
        {
            string fileName = "WordDoc";
            string fullName = $"{Directory.GetCurrentDirectory()}\\{fileName}.docx";

            using (var x = new WordDocument())
            {
                x[1] = "Text";
                x.SaveAs(fullName);
            }

            Assert.IsTrue(File.Exists(fullName));
        }
    }
}