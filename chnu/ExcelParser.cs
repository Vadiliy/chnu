using chnu.Models;
using System.Collections.Generic;
using System.IO;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System.Text.RegularExpressions;

namespace chnu
{
    public class ExcelParser
    {
        public UniversityContext Context { get; set; }

        string fileName = "wwwroot/statement.xls";
        public void GetGroups()
        {
            using (FileStream fs = new FileStream(fileName, FileMode.Open, FileAccess.Read))
            {
                HSSFWorkbook book = new HSSFWorkbook(fs);
                int count = book.NumberOfSheets;
                Year year = new Year();
                Context.Years.Add(year);
                if(year.Groups == null)
                    year.Groups = new List<Models.Group>();

                for (int a = 0; a < count; a++)
                {
                    string nameOfGroup = book.GetSheetName(a);
                    Regex regex = new Regex("^\\d");
                    if (regex.IsMatch(nameOfGroup))
                    {
                        Models.Group group = new Models.Group()
                        {
                            NameGroup = nameOfGroup
                        };
                        year.Groups.Add(group);
                    }
                }

                GetStudents(year, book);

                ISheet sheet = book.GetSheet(year.Groups[0].NameGroup);
                string yearStr = sheet.GetRow(1).GetCell(1).StringCellValue;
                Regex regexYear = new Regex("\\d{4}\\/\\d{4}");
                year.Name = regexYear.Match(yearStr).Value;

                FilterSubjects();
                book.Close();
            }
            Context.SaveChanges();
        }

        void GetStudents(Year year, HSSFWorkbook book)
        {
            foreach (Models.Group group in year.Groups)
            {
                group.Year = year;
                if(group.Students == null)
                    group.Students = new List<Student>();

                ISheet sheet = book.GetSheet(group.NameGroup);
                int number = 0;
                while (true)
                {
                    string nameofStudent;
                    try
                    {
                        nameofStudent = sheet.GetRow(9 + number).GetCell(1).StringCellValue;
                    }
                    catch
                    {
                        number++;
                        continue;
                    }
                    if (string.IsNullOrEmpty(nameofStudent))
                    {
                        number++;
                        continue;
                    }
                    if (nameofStudent.Contains("Дата складання"))
                        break;

                    Student student = new Student()
                    {
                        Name = nameofStudent,
                        Group = group
                    };
                    group.Students.Add(student);
                    number++;
                }
            }
            GetSubjectInfo(year, book);
        }

        void GetSubjectInfo(Year year, HSSFWorkbook book)
        {
            foreach (Models.Group group in year.Groups)
            {
                ISheet sheet = book.GetSheet(group.NameGroup);
                foreach (Student student in group.Students)
                {
                    if(student.Subjects == null)
                        student.Subjects = new List<Subject>();

                    int number = 0;
                    while(true)
                    {
                        string nameOfSubject;
                        try
                        {
                            nameOfSubject = sheet.GetRow(7).GetCell(5 + number).StringCellValue;
                        }
                        catch
                        {
                            break;
                        }
                        if (nameOfSubject.Contains("Сер.зважений") || string.IsNullOrEmpty(nameOfSubject))
                            break;

                        Subject subject = new Subject
                        {
                            Name = nameOfSubject,
                            Lecture = sheet.GetRow(8).GetCell(5 + number).StringCellValue,
                            ExamType = sheet.GetRow(6).GetCell(5 + number).StringCellValue,
                            Hours = sheet.GetRow(5).GetCell(5 + number).NumericCellValue,
                            Credits = sheet.GetRow(4).GetCell(5 + number).NumericCellValue,
                            Student = student
                        };
                        student.Subjects.Add(subject);
                        number++;
                    }
                }
            }
            GetScores(year, book);
        }

        void GetScores(Year year, HSSFWorkbook book)
        {
            foreach (Models.Group group in year.Groups)
            {
                ISheet sheet = book.GetSheet(group.NameGroup);
                int studentPos = 0;
                foreach (Student student in group.Students)
                {
                    int subjectPos = 0;
                    foreach (Subject subject in student.Subjects)
                    {
                        int score = -1;
                        try
                        {
                            score = (int)sheet.GetRow(9 + studentPos).GetCell(5 + subjectPos).NumericCellValue;
                        }
                        catch
                        {
                            score = 0;
                        }
                        if (score < 60) student.Debt = true;
                        subject.Score = score;
                        subjectPos++;
                    }
                    studentPos++;
                }
            }
        }

        void FilterSubjects()
        {
            foreach(Year year in Context.Years.Local)
            {
                foreach(Models.Group gr in year.Groups)
                {
                    foreach(Student st in gr.Students)
                    {
                        st.Subjects.RemoveAll(x => x.Score >= 60);
                    }
                }
            }
        }
    }
}
