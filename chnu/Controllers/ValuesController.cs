using System.Linq;
using Microsoft.AspNetCore.Mvc;
using chnu.Models;
using System.Collections.Generic;

namespace chnu.Controllers
{
    public class ValuesController : Controller
    {
        UniversityContext context;

        public ValuesController(UniversityContext context)
        {
            this.context = context;

            //ExcelParser excel = new ExcelParser();
            //excel.Context = context;
            //excel.GetGroups();
        }
        // GET api/valuesя
        [HttpGet]
        public IActionResult Index()
        {
            return View();
        }

        [HttpGet("{id}")]
        public IActionResult Get(string discipline, string student, string group, string time)
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
                    return View("~/Views/Values/byName.cshtml", stud);
                }
                else
                {
                    if(group != null)
                    {
                        foreach(Student st in students)                     
                            context.Entry(st).Reference("Group").Load();

                        try
                        {
                            Student first = students.First(x => x.Group.NameGroup == group);
                            context.Entry(first).Collection("Subjects").Load();
                            return View("~/Views/Values/byName.cshtml", first);
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
                        return View("~/Views/Values/byName.cshtml", stud);
                    }
                }
            }

            if(group != null)
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

                context.Entry(grou).Collection("Students").Load();
                foreach(Student st in grou.Students)               
                    context.Entry(st).Collection("Subjects").Load();


                grou.Students = grou.Students.OrderBy(x => x.Name).ToList();
                return View("~/Views/Values/byGroup.cshtml", grou);
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
                return View("~/Views/Values/byDiscipline.cshtml", subjects);

            }

            return null;
        }

        [HttpPost]
        public string[] Order(string date, int studentId, int disciplineId)
        {
            Subject subject = context.Subjects.First(x => x.Id == disciplineId);
            subject.IsOreded = true;
            subject.DateDebt = date;
            context.SaveChanges();

            return new string[] { date + "  " +  studentId + "  " + disciplineId};
        }
    }
}
