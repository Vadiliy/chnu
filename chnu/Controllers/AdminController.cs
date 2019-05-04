using System.Linq;
using Microsoft.AspNetCore.Mvc;
using chnu.Models;
using System.Collections.Generic;
using Microsoft.AspNetCore.Identity;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Authorization;
using System.IO;
using NPOI.XWPF.UserModel;
using System;

namespace chnu.Controllers
{
    public class AdminController : Controller
    {
        UniversityContext context;
        private readonly SignInManager<IdentityUser> _signInManager;
        public AdminController(UniversityContext universityContext, SignInManager<IdentityUser> signInManager)
        {
            this.context = universityContext;
            _signInManager = signInManager;
        }

        public IActionResult Login()
        {
            if(User.IsInRole("admin"))
            {
                return View("~/Views/Admin/Admin.cshtml");
            }
            return View();
        }
        // Admin
        // 123Ert-123
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Login(string login, string pass)
        {
            var result = await _signInManager.PasswordSignInAsync(login, pass, true, false);
            if(result.Succeeded)
            {
                return View("~/Views/Admin/Admin.cshtml");
            }
            return View();
        }

        [HttpPost("{id}")]
        [Authorize(Roles = "admin")]
        public IActionResult Get(string discipline, string student, string group)
        {
            if (student != null)
            {
                List<Student> students = context.Students.Where(x => x.Name.Contains(student)).ToList();
                if (students == null || students.Count == 0)
                    return View("~/Views/Values/Error.cshtml", "Такого студента немає");

                if (students.Count == 1)
                {
                    Student stud = students.First(x => x.Name.Contains(student));
                    context.Entry(stud).Collection("Subjects").Load();
                    context.Entry(stud).Reference("Group").Load();
                    List<Subject> subjects = stud.Subjects;

                    subjects = FilterSubjects(subjects);

                    return View("~/Views/Admin/Admin.cshtml", subjects);
                }
                else
                {
                    if (group != null)
                    {
                        foreach (Student st in students)
                            context.Entry(st).Reference("Group").Load();

                        try
                        {
                            Student first = students.First(x => x.Group.NameGroup == group);
                            context.Entry(first).Collection("Subjects").Load();
                            context.Entry(first).Reference("Group").Load();
                            List<Subject> subjects = first.Subjects;

                            subjects = FilterSubjects(subjects);

                            return View("~/Views/Admin/Admin.cshtml", subjects);
                        }
                        catch
                        {

                        }
                    }
                    else
                    {
                        Student stud = context.Students.First(x => x.Name.Contains(student));
                        context.Entry(stud).Collection("Subjects").Load();
                        context.Entry(stud).Reference("Group").Load();
                        List<Subject> subjects = stud.Subjects;
                        subjects = FilterSubjects(subjects);
                        return View("~/Views/Admin/Admin.cshtml", subjects);
                    }
                }
            }

            if (group != null)
            {
                Group grou = null;
                try
                {
                    grou = context.Groups.First(x => x.NameGroup == group);
                }
                catch
                {
                    return View("~/Views/Values/Error.cshtml", "Такої групи немає");
                }

                List<Subject> subjects = new List<Subject>();
                context.Entry(grou).Collection("Students").Load();
                foreach (Student st in grou.Students)
                {
                    context.Entry(st).Collection("Subjects").Load();
                    context.Entry(st).Reference("Group").Load();
                    subjects.AddRange(st.Subjects);
                }
                subjects = FilterSubjects(subjects);
                return View("~/Views/Admin/Admin.cshtml", subjects);
            }

            if (discipline != null)
            {
                Subject subject = null;

                try
                {
                    subject = context.Subjects.First(x => x.Name.Contains(discipline));
                }
                catch
                {
                    return View("~/Views/Values/Error.cshtml", "Такого предмету немає");
                }

                List<Subject> subjects = context.Subjects.Where(s => s.Name == subject.Name).ToList();
                foreach (Subject sb in subjects)
                {
                    context.Entry(sb).Reference("Student").Load();
                    context.Entry(sb.Student).Reference("Group").Load();
                }
                subjects = FilterSubjects(subjects);
                return View("~/Views/Admin/Admin.cshtml", subjects);
            }

            return View("~/Views/Values/Error.cshtml", "Невідома помилка");
        }

        private List<Subject> FilterSubjects(List<Subject> subjects)
        {
            DateTime time = new DateTime();
            subjects = subjects.OrderBy(x => x.Student.Name).ToList();
            subjects.RemoveAll(x => x.IsOreded == false); // если бегунок не заказан
            subjects.RemoveAll(x => x.DebtIsClosed != time); // если долг не закрыт
            return subjects;
        }

        [HttpPost]
        [Authorize(Roles = "admin")]
        public IActionResult Print(int[] toPrint)
        {
            XWPFDocument doc;
            using (Stream fileStream = System.IO.File.OpenRead("wwwroot\\template.docx"))
            {
                doc = new XWPFDocument(fileStream);
                fileStream.Close();
            }
            ExcelParser excel = new ExcelParser();
            int score = 0;
            while(score < toPrint.Count())
            {
                for(int a = 0; a < 3; a++)
                {          
                    Subject subject = context.Subjects.First(x => x.Id == toPrint[score]);

                    //if (subject.IsPrinted)
                    //{
                    //    score++;
                    //    if (score == toPrint.Count())
                    //        break;
                    //    a--;
                    //    continue;
                    //}

                    context.Entry(subject).Reference("Student").Load();
                    Student student = subject.Student;
                    context.Entry(student).Reference("Group").Load();
                    excel.Create(doc, subject);
                    subject.IsPrinted = true;
                    XWPFParagraph para = doc.CreateParagraph();
                    score++;
                    if (score == toPrint.Count())
                        break;
                }
                if(score < toPrint.Count())
                {
                    XWPFParagraph para = doc.CreateParagraph();
                    XWPFRun run = para.CreateRun();
                    run.AddBreak();
                }
            }

            context.SaveChanges();
        
            using (FileStream fileStreamNew = System.IO.File.Create("wwwroot\\begunok.docx"))
            {
                doc.Write(fileStreamNew);
                fileStreamNew.Close();
            }

            // Тип файла - content-type
            string file_type = "application/docx";
            // Имя файла - необязательно
            string file_name = "begunok.docx";
            var filepath = Path.Combine("~/", "begunok.docx");

            return File(filepath, file_type,  file_name);
        }

        [HttpPost]
        [Authorize(Roles = "admin")]
        public string NewMark(int newMark, int disciplineId)
        {
            Subject subject = context.Subjects.First(x => x.Id == disciplineId);
            subject.Score = newMark;
            subject.DebtIsClosed = DateTime.Now;
            context.SaveChanges();
            return "ok";
        }
    }
}