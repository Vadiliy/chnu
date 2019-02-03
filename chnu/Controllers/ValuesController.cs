using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using chnu.Models;

namespace chnu.Controllers
{
    [Route("api/[controller]")]
    public class ValuesController : Controller
    {
        UniversityContext context;

        public ValuesController(UniversityContext context)
        {
            this.context = context;
        }
        // GET api/values
        [HttpGet]
        public IActionResult Index()
        {
            return View();
        }

        [HttpGet("{id}")]
        // GET api/values/5
        public string Get(string discipline, string student, string group, string time)
        {
            return "value";
        }
    }
}
