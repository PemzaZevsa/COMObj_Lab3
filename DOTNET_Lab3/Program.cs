using System.Text.Json;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace DOTNET_Lab3
{
    internal class Program
    {
        static string coursesPath = "C:\\Users\\Pemza\\source\\repos\\CourseworkOOP\\OOP_Cursova\\CourseworkOOP\\CourseworkOOP\\bin\\Debug\\net8.0-windows\\Data\\Courses\\CoursesData.json";
        static string usersPath = "C:\\Users\\Pemza\\source\\repos\\CourseworkOOP\\OOP_Cursova\\CourseworkOOP\\CourseworkOOP\\bin\\Debug\\net8.0-windows\\Data\\Users\\UsersData.json";

        static void Main(string[] args)
        {
            List<Course> courses = new List<Course>();
            List<User> users = new List<User>();
            try
            {
                LoadCourses(courses);
                LoadUsers(users);
                WordReport(courses, users);
                ExcelReport(courses,users);
            }
            catch (Exception)
            {
            }
        }

        static public void WordReport(List<Course> courses, List<User> users)
        {
            string fileName = "WordReport2";
            string fullName = $"{Directory.GetCurrentDirectory()}\\{fileName}.docx";
            string str = "";

            foreach (var course in courses)
            {
                str += $"Назва: {course.Name}\nОпис: {course.Description}\nЦіна: {course.Cost}\tРейтинг {course.Rating}\n\n";
            }

            string str2 = "";

            foreach (var user in users)
            {
                str2 += $"Ім'я: {user.Name}\tПрізвище: {user.Surname}\nТип користувача: {user.UserType}\tЛогін {user.Login}\tПароль {user.Password}\n\n";
            }

            using (var x = new WordDocument())
            {
                x.AddParagraph("Інформація про платформу платних навчальних курсах", 20, Word.WdColor.wdColorBlack, Word.WdParagraphAlignment.wdAlignParagraphCenter);
                x.AddParagraph($"Звіт зроблено : {DateTime.Now}", 16, Word.WdColor.wdColorDarkRed, Word.WdParagraphAlignment.wdAlignParagraphCenter);
                x.AddParagraph("",  12, Word.WdColor.wdColorBlack, Word.WdParagraphAlignment.wdAlignParagraphJustify);
                x.AddParagraph($"Усього {courses.Count} курсів", 14, Word.WdColor.wdColorBlack, Word.WdParagraphAlignment.wdAlignParagraphJustify);
                x.AddParagraph(str,  9, Word.WdColor.wdColorGray875, Word.WdParagraphAlignment.wdAlignParagraphJustify);
                x.AddParagraph("", 12, Word.WdColor.wdColorBlack, Word.WdParagraphAlignment.wdAlignParagraphJustify);
                x.AddParagraph($"Усього {users.Count} користувачів", 14, Word.WdColor.wdColorBlack, Word.WdParagraphAlignment.wdAlignParagraphJustify);
                x.AddParagraph(str2, 9, Word.WdColor.wdColorGray875 , Word.WdParagraphAlignment.wdAlignParagraphJustify);
                x.SaveAs(fullName);
            }
        }
        static public void ExcelReport(List<Course> courses, List<User> users)
        {
            string fileName = "ExcelReport2";
            string fullName = $"{Directory.GetCurrentDirectory()}\\{fileName}.xlsx";
            string str = "";

            foreach (var course in courses)
            {
                str += $"Назва: {course.Name}\nОпис: {course.Description}\nЦіна: {course.Cost}\tРейтинг {course.Rating}\n\n";
            }

            string str2 = "";

            foreach (var user in users)
            {
                str2 += $"Ім'я: {user.Name}\tПрізвище: {user.Surname}\nТип користувача: {user.UserType}\tЛогін {user.Login}\tПароль {user.Password}\n\n";
            }

            using (var x = new ExcelDocument())
            {
                x.AddMergedCells(1, 1, 1, 10, "Інформація про платформу платних навчальних курсах", 20, true, Excel.XlRgbColor.rgbBlack, Excel.XlHAlign.xlHAlignCenter);
                x.AddMergedCells(2, 1, 2, 10, $"Звіт зроблено : {DateTime.Now}", 16, false, Excel.XlRgbColor.rgbRed, Excel.XlHAlign.xlHAlignCenter);

                int i = 3;

                i++;
                i++;

                x.AddMergedCells(i, 1, i, 5, "Назва", 12, true, Excel.XlRgbColor.rgbBlack, Excel.XlHAlign.xlHAlignCenter);
                x.AddMergedCells(i, 6, i,10, "Опис", 12, true, Excel.XlRgbColor.rgbBlack, Excel.XlHAlign.xlHAlignCenter);
                x.AddCell(i, 11, "Ціна", 12, true, Excel.XlRgbColor.rgbBlack, Excel.XlHAlign.xlHAlignCenter);
                x.AddCell(i, 12, "Рейтинг", 12, true, Excel.XlRgbColor.rgbBlack, Excel.XlHAlign.xlHAlignCenter);

                i++;

                foreach (var course in courses)
                {
                    x.AddMergedCells(i, 1,i,5, $"{course.Name}", 10, false, Excel.XlRgbColor.rgbBlack, Excel.XlHAlign.xlHAlignLeft);
                    x.AddMergedCells(i, 6,i,10, $"{course.Description}", 8, false, Excel.XlRgbColor.rgbBlack, Excel.XlHAlign.xlHAlignLeft);
                    x.AddCell(i, 11, $"{course.Cost}", 10, false, Excel.XlRgbColor.rgbBlack, Excel.XlHAlign.xlHAlignCenter);
                    x.AddCell(i, 12, $"{Math.Round(course.Rating,2)}", 10, false, Excel.XlRgbColor.rgbBlack, Excel.XlHAlign.xlHAlignCenter);
                    i++;
                }

                i++;
                i++;

                x.AddMergedCells(i, 1, i, 2, "Ім'я", 12, true, Excel.XlRgbColor.rgbBlack, Excel.XlHAlign.xlHAlignCenter);
                x.AddMergedCells(i, 3, i, 4, "Прізвище", 12, true, Excel.XlRgbColor.rgbBlack, Excel.XlHAlign.xlHAlignCenter);
                x.AddCell(i, 5, "Тип", 12, true, Excel.XlRgbColor.rgbBlack, Excel.XlHAlign.xlHAlignCenter);
                x.AddCell(i, 6, "Логін", 12, true, Excel.XlRgbColor.rgbBlack, Excel.XlHAlign.xlHAlignCenter);
                x.AddCell(i, 7, "Пароль", 12, true, Excel.XlRgbColor.rgbBlack, Excel.XlHAlign.xlHAlignCenter);

                i++;

                foreach (var user in users)
                {
                    x.AddMergedCells(i, 1, i, 2, $"{user.Name}", 10, false, Excel.XlRgbColor.rgbBlack, Excel.XlHAlign.xlHAlignLeft);
                    x.AddMergedCells(i, 3, i, 4, $"{user.Surname}", 10, false, Excel.XlRgbColor.rgbBlack, Excel.XlHAlign.xlHAlignLeft);
                    x.AddCell(i, 5, $"{(UserType)user.UserType}", 10, false, Excel.XlRgbColor.rgbBlue, Excel.XlHAlign.xlHAlignLeft);
                    x.AddCell(i, 6, $"{user.Login}", 10, false, Excel.XlRgbColor.rgbBlack, Excel.XlHAlign.xlHAlignLeft);
                    x.AddCell(i, 7, $"{user.Password}", 10, false, Excel.XlRgbColor.rgbBlack, Excel.XlHAlign.xlHAlignLeft);

                    i++;
                }

                x.SaveAs(fullName);
            }
        }

        static public void LoadCourses(List<Course> courses)
        {
            try
            {
                List<string> lines = File.ReadAllLines(coursesPath).ToList();

                foreach (var item in lines)
                {
                    Course? course = JsonSerializer.Deserialize<Course>(item);
                    if (course != null)
                    {
                        courses.Add(course);
                    }
                }
            }
            catch (IOException e)
            {
                Console.WriteLine(e.ToString());
            }
        }

        static public void LoadUsers(List<User> users)
        {
            try
            {
                List<string> lines = File.ReadAllLines(usersPath).ToList();

                foreach (var item in lines)
                {
                    User? user = null;
                    switch (item[^2])
                    {
                        case '0':
                            user = JsonSerializer.Deserialize<Admin>(item);
                            break;
                        case '1':
                            user = JsonSerializer.Deserialize<Teacher>(item);
                            break;
                        case '2':
                            user = JsonSerializer.Deserialize<Student>(item);
                            break;
                    }

                    if (user != null) users.Add(user);
                }
            }
            catch (IOException e)
            {
            }
        }
    }
}
