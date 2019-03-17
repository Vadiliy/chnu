using System.Linq;
using Microsoft.AspNetCore.Mvc;
using chnu.Models;
using System.Collections.Generic;
using Microsoft.EntityFrameworkCore;
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
        // GET api/values
        [HttpGet]
        public IActionResult Index()
        {
            return View();
        }

        [HttpGet("{id}")]
        // GET api/values/5
        public IActionResult Get(string discipline, string student, string group, string time)
        {
            if (student != null)
            {
                Student stud = context.Students.First(x => x.Name.Contains(student));
                context.Entry(stud).Collection("Subjects").Load();
                context.Entry(stud).Reference("Group").Load();
                return View("~/Views/Values/byName.cshtml", stud);
            }
            if(group != null)
            {
                Group grou = context.Years.SelectMany(g => g.Groups).First(x => x.NameGroup == group);
                context.Entry(grou).Collection("Students").Load();
                foreach(Student st in grou.Students)               
                    context.Entry(st).Collection("Subjects").Load();


                grou.Students = grou.Students.OrderBy(x => x.Name).ToList();
                return View("~/Views/Values/byGroup.cshtml", grou);
            }
            if(discipline != null)
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
    }
}
