using chnu.Models;
using System.Collections.Generic;
using System.IO;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System.Text.RegularExpressions;
using NPOI.XWPF.UserModel;
using System;

namespace chnu
{
    public class ExcelParser
    {
        public UniversityContext Context { get; set; }

        #region Excel

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

        #endregion

        public void Word(Subject subj)
        {
            XWPFDocument doc = new XWPFDocument();

            #region ЧНУ імені Петра Могили
            XWPFParagraph para = doc.CreateParagraph();
            XWPFRun r0 = para.CreateRun();
            r0.SetText("ЧНУ імені Петра могили");
            para.Alignment = ParagraphAlignment.CENTER;
            r0.FontSize = 16;
            r0.IsBold = true;
            #endregion

            #region Направлення №1234
            para = doc.CreateParagraph();
            r0 = para.CreateRun();
            r0.SetText("Направлення № 123 на перескладання екзамену (заліку)");
            para.Alignment = ParagraphAlignment.CENTER;
            r0.FontSize = 16;
            r0.IsBold = true;
            #endregion

            #region Імя студента та назва предмету
            para = doc.CreateParagraph();
            para.Alignment = ParagraphAlignment.BOTH;
            r0 = para.CreateRun();
            XWPFRun r1 = para.CreateRun();
            XWPFRun r2 = para.CreateRun();
            XWPFRun r3 = para.CreateRun();
            r0.FontSize = 14;
            r1.FontSize = 14;
            r2.FontSize = 14;
            r3.FontSize = 14;

            r0.SetText("З дисципліни ");
            r1.SetText(subj.Name);
            r1.SetUnderline(UnderlinePatterns.Single);
            r2.SetText(" видане студенту ");
            r3.SetText(subj.Student.Name);
            r3.SetUnderline(UnderlinePatterns.Single);
            r3.AddCarriageReturn();
            #endregion

            #region Група
            para = doc.CreateParagraph();
            r0 = para.CreateRun();
            r1 = para.CreateRun();
            r2 = para.CreateRun();
            r3 = para.CreateRun();
            XWPFRun r4 = para.CreateRun();

            r0.FontSize = 14;
            r1.FontSize = 14;
            r2.FontSize = 14;
            r3.FontSize = 14;
            r4.FontSize = 14;

            r0.SetText("Курсу ");
            r1.SetText(subj.Student.Group.NameGroup[0].ToString());
            r1.SetUnderline(UnderlinePatterns.Single);
            r2.SetText(" групи ");
            r3.SetText(subj.Student.Group.NameGroup);
            r3.SetUnderline(UnderlinePatterns.Single);
            r4.SetText(" факультету комп'ютерних наук");
            #endregion

            #region Викладач
            XWPFTable table = doc.CreateTable(4, 5);
            table.SetBottomBorder(XWPFTable.XWPFBorderType.NONE, table.InsideHBorderSize, table.InsideHBorderSpace, "FFFFFF");
            table.SetLeftBorder(XWPFTable.XWPFBorderType.NONE, table.InsideHBorderSize, table.InsideHBorderSpace, "FFFFFF");
            table.SetRightBorder(XWPFTable.XWPFBorderType.NONE, table.InsideHBorderSize, table.InsideHBorderSpace, "FFFFFF");
            table.SetTopBorder(XWPFTable.XWPFBorderType.NONE, table.InsideHBorderSize, table.InsideHBorderSpace, "FFFFFF");
            table.SetInsideHBorder(XWPFTable.XWPFBorderType.NONE, table.InsideHBorderSize, table.InsideHBorderSpace, "FFFFFF");
            table.SetInsideVBorder(XWPFTable.XWPFBorderType.NONE, table.InsideHBorderSize, table.InsideHBorderSpace, "FFFFFF");
            table.Width = 5000;


            XWPFTableCell c1 = table.GetRow(0).GetCell(0);
            XWPFParagraph p1 = c1.AddParagraph();   //don't use doc.CreateParagraph
            r0 = p1.CreateRun();
            r1 = p1.CreateRun();
            r0.FontSize = 14;
            r1.FontSize = 14;
            r0.SetText("Викладач ");
            r1.SetText(subj.Lecture);
            r1.SetUnderline(UnderlinePatterns.Single);

            XWPFTableCell c2 = table.GetRow(0).GetCell(1);
            XWPFParagraph p2 = c2.AddParagraph();   //don't use doc.
            r0 = p2.CreateRun();
            r0.FontSize = 14;
            r0.SetText("Підпис ______________");
            p2.Alignment = ParagraphAlignment.RIGHT;
            #endregion

            #region Дата

            XWPFTableCell c3 = table.GetRow(1).GetCell(0);
            XWPFParagraph p3 = c3.AddParagraph();   //don't use doc.CreateParagraph
            r0 = p3.CreateRun();
            r1 = p3.CreateRun();
            r0.FontSize = 14;
            r1.FontSize = 14;
            r0.SetText("Дата видачі ");
            r1.SetText(DateTime.Now.ToShortDateString());
            r1.SetUnderline(UnderlinePatterns.Single);

            XWPFTableCell c4 = table.GetRow(1).GetCell(1);
            XWPFParagraph p4 = c4.AddParagraph();   //don't use doc.
            r0 = p4.CreateRun();
            r0.FontSize = 14;
            r0.SetText("Дата складання ______________");
            p4.Alignment = ParagraphAlignment.RIGHT;
            #endregion

            #region Оцінка

            XWPFTableCell c5 = table.GetRow(2).GetCell(0);
            XWPFParagraph p5 = c5.AddParagraph();   //don't use doc.CreateParagraph
            r0 = p5.CreateRun();
            r0.FontSize = 14;
            r0.SetText("Оцінка _______________________");

            XWPFTableCell c6 = table.GetRow(2).GetCell(1);
            XWPFParagraph p6 = c6.AddParagraph();   //don't use doc.
            r0 = p6.CreateRun();
            r0.FontSize = 14;
            r0.SetText("Кількість балів сем. ______+______=________");
            p6.Alignment = ParagraphAlignment.RIGHT;
            #endregion

            #region Декан

            XWPFTableCell c7 = table.GetRow(3).GetCell(0);
            XWPFParagraph p7 = c7.AddParagraph();   //don't use doc.CreateParagraph
            r0 = p7.CreateRun();
            r0.FontSize = 14;
            r0.SetText("Дійсне до _____________");
            p7.Alignment = ParagraphAlignment.CENTER;

            XWPFTableCell c8 = table.GetRow(3).GetCell(1);
            XWPFParagraph p8 = c8.AddParagraph();   //don't use doc.
            r0 = p8.CreateRun();
            r0.FontSize = 14;
            r0.SetText("Декан _____________");
            p8.Alignment = ParagraphAlignment.CENTER;
            #endregion

            FileStream out1 = new FileStream("E:\\simpleTable.docx", FileMode.Create);
            doc.Write(out1);
            out1.Close();
        }
    }
}
