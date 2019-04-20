using System.Linq;
using Microsoft.AspNetCore.Mvc;
using chnu.Models;

namespace chnu.Controllers
{
    public class AdminController : Controller
    {
        UniversityContext context;

        public AdminController(UniversityContext universityContext)
        {
            this.context = universityContext;
        }

        [HttpPost]
        public IActionResult Auth(string login, string pass)
        {
            try
            {
                Models.User user = context.Admins.Where(x => x.Login == login && x.Pass == pass).First();
                return View();
            }
            catch
            {
                return View("~/Views/Values/Error.cshtml", "Невірний логін та пароль");
            }
        }
    }
}